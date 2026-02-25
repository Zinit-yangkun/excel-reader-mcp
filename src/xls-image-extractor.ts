import * as CFB from 'cfb';
import * as XLSX from 'xlsx';
import { existsSync, readFileSync } from 'fs';
import { inflateSync } from 'zlib';
import {
  ErrorCode,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import type { ExtractedImage, GetExcelImagesArgs, ImagePosition } from './types.js';

const MAX_IMAGES_SIZE = 10 * 1024 * 1024; // 10MB total base64 size limit

// BIFF record types
const BIFF_MSODRAWINGGROUP = 0x00eb;
const BIFF_MSODRAWING = 0x00ec;
const BIFF_CONTINUE = 0x003c;
const BIFF_BOF = 0x0809;
const BIFF_BOUNDSHEET = 0x0085;
const BIFF_EOF = 0x000a;
const BIFF_OBJ = 0x005d;

// Escher record types
const ESCHER_DGG_CONTAINER = 0xf000;
const ESCHER_BSTORE_CONTAINER = 0xf001;
const ESCHER_DG_CONTAINER = 0xf002;
const ESCHER_SPGR_CONTAINER = 0xf003;
const ESCHER_SP_CONTAINER = 0xf004;
const ESCHER_BSE = 0xf007;
const ESCHER_SP = 0xf00a;
const ESCHER_CLIENT_ANCHOR = 0xf010;
const ESCHER_CLIENT_DATA = 0xf011;

// BLIP types
const ESCHER_BLIP_EMF = 0xf01a;
const ESCHER_BLIP_WMF = 0xf01b;
const ESCHER_BLIP_PICT = 0xf01c;
const ESCHER_BLIP_JPEG = 0xf01d;
const ESCHER_BLIP_PNG = 0xf01e;
const ESCHER_BLIP_DIB = 0xf01f;
const ESCHER_BLIP_TIFF = 0xf029;
const ESCHER_BLIP_JPEG2 = 0xf02a;

interface BiffRecord {
  type: number;
  data: Buffer;
}

interface EscherRecord {
  version: number;
  instance: number;
  type: number;
  length: number;
  data: Buffer;
  offset: number; // offset within the container data
}

interface BlipInfo {
  index: number; // 1-based BSE index
  mimeType: string;
  extension: string;
  data: Buffer;
}

interface DrawingAnchor {
  sheetName: string;
  spId: number;
  blipIndex: number; // reference to BSE index
  fromCol: number;
  fromRow: number;
  toCol: number;
  toRow: number;
}

function getMimeAndExt(blipType: number): { mimeType: string; extension: string } {
  switch (blipType) {
    case ESCHER_BLIP_EMF: return { mimeType: 'image/x-emf', extension: '.emf' };
    case ESCHER_BLIP_WMF: return { mimeType: 'image/x-wmf', extension: '.wmf' };
    case ESCHER_BLIP_PICT: return { mimeType: 'image/pict', extension: '.pict' };
    case ESCHER_BLIP_JPEG:
    case ESCHER_BLIP_JPEG2: return { mimeType: 'image/jpeg', extension: '.jpg' };
    case ESCHER_BLIP_PNG: return { mimeType: 'image/png', extension: '.png' };
    case ESCHER_BLIP_DIB: return { mimeType: 'image/bmp', extension: '.bmp' };
    case ESCHER_BLIP_TIFF: return { mimeType: 'image/tiff', extension: '.tiff' };
    default: return { mimeType: 'application/octet-stream', extension: '.bin' };
  }
}

/**
 * Read BIFF records from a workbook stream buffer.
 * Handles CONTINUE records by merging them into the preceding record.
 */
function readBiffRecords(buf: Buffer): BiffRecord[] {
  const records: BiffRecord[] = [];
  let offset = 0;

  while (offset + 4 <= buf.length) {
    const type = buf.readUInt16LE(offset);
    const length = buf.readUInt16LE(offset + 2);
    offset += 4;

    if (offset + length > buf.length) break;

    const data = buf.subarray(offset, offset + length);
    offset += length;

    // Merge CONTINUE records into the previous record
    if (type === BIFF_CONTINUE && records.length > 0) {
      const prev = records[records.length - 1];
      prev.data = Buffer.concat([prev.data, data]);
    } else {
      records.push({ type, data: Buffer.from(data) });
    }
  }

  return records;
}

/**
 * Parse a single Escher record header from buffer at the given offset.
 */
function readEscherRecord(buf: Buffer, offset: number): EscherRecord | null {
  if (offset + 8 > buf.length) return null;

  const verAndInstance = buf.readUInt16LE(offset);
  const version = verAndInstance & 0x0f;
  const instance = (verAndInstance >> 4) & 0x0fff;
  const type = buf.readUInt16LE(offset + 2);
  const length = buf.readInt32LE(offset + 4);

  if (length < 0 || offset + 8 + length > buf.length) return null;

  const data = buf.subarray(offset + 8, offset + 8 + length);
  return { version, instance, type, length, data, offset };
}

/**
 * Iterate Escher records within a buffer (non-recursive, one level).
 */
function* iterEscherRecords(buf: Buffer): Generator<EscherRecord> {
  let offset = 0;
  while (offset < buf.length) {
    const rec = readEscherRecord(buf, offset);
    if (!rec) break;
    yield rec;
    offset += 8 + rec.length;
  }
}

/**
 * Check if an Escher record is a container (version == 0x0F).
 */
function isContainer(rec: EscherRecord): boolean {
  return rec.version === 0x0f;
}

/**
 * Extract BLIP images from MsoDrawingGroup data.
 * The structure is: DggContainer → BStoreContainer → BSE → BLIP
 */
function extractBlipsFromDrawingGroup(data: Buffer): BlipInfo[] {
  const blips: BlipInfo[] = [];

  for (const dggContainer of iterEscherRecords(data)) {
    if (dggContainer.type !== ESCHER_DGG_CONTAINER || !isContainer(dggContainer)) continue;

    for (const child of iterEscherRecords(dggContainer.data)) {
      if (child.type !== ESCHER_BSTORE_CONTAINER || !isContainer(child)) continue;

      let bseIndex = 0;
      for (const bse of iterEscherRecords(child.data)) {
        bseIndex++;
        if (bse.type !== ESCHER_BSE) continue;

        // BSE header: 1 byte btWin32 + 1 byte btMacOS + 16 bytes UID + 2 bytes tag + 4 bytes size
        //           + 4 bytes cRef + 4 bytes foDelay + 1 byte usage + 1 byte cbName + 1 byte unused + 1 byte unused
        // Total BSE header = 36 bytes, then optional name (cbName bytes), then BLIP data
        if (bse.data.length < 36) continue;

        const btWin32 = bse.data.readUInt8(0);
        const cbName = bse.data.readUInt8(33);
        const blipOffset = 36 + cbName;

        if (blipOffset >= bse.data.length) continue;

        // Parse the embedded BLIP record
        const blipRec = readEscherRecord(bse.data, blipOffset);
        if (!blipRec) continue;

        const blipType = blipRec.type;
        const { mimeType, extension } = getMimeAndExt(blipType);

        // Extract raw image data from BLIP
        // BLIP structure varies by type:
        // - EMF/WMF/PICT: 16 bytes UID + 16 bytes UID2(if instance==0x6E1/0x2160) + 34 bytes metafile header
        // - JPEG/PNG/DIB: 16 bytes UID + (optional 16 bytes UID2) + 1 byte tag
        let imageDataOffset = 0;
        const blipData = blipRec.data;

        if (blipType === ESCHER_BLIP_EMF || blipType === ESCHER_BLIP_WMF || blipType === ESCHER_BLIP_PICT) {
          // Metafile BLIP: UID(16) + possibly UID2(16) + cb(4) + rcBounds(16) + ptSize(8) + cbSave(4) + compression(1) + filter(1)
          // instance value tells us if there's a second UID
          const hasUID2 = (blipRec.instance === 0x3d5 || blipRec.instance === 0x217 || blipRec.instance === 0x543);
          imageDataOffset = 16 + (hasUID2 ? 16 : 0) + 34;
        } else if (blipType === ESCHER_BLIP_JPEG || blipType === ESCHER_BLIP_JPEG2) {
          const hasUID2 = (blipRec.instance === 0x46b || blipRec.instance === 0x6e3);
          imageDataOffset = 16 + (hasUID2 ? 16 : 0) + 1;
        } else if (blipType === ESCHER_BLIP_PNG) {
          const hasUID2 = (blipRec.instance === 0x6e1);
          imageDataOffset = 16 + (hasUID2 ? 16 : 0) + 1;
        } else if (blipType === ESCHER_BLIP_DIB) {
          const hasUID2 = (blipRec.instance === 0x7a9);
          imageDataOffset = 16 + (hasUID2 ? 16 : 0) + 1;
        } else if (blipType === ESCHER_BLIP_TIFF) {
          const hasUID2 = (blipRec.instance === 0x6e5);
          imageDataOffset = 16 + (hasUID2 ? 16 : 0) + 1;
        } else {
          // Unknown BLIP type, try to extract with minimal header skip
          imageDataOffset = 17;
        }

        if (imageDataOffset >= blipData.length) continue;

        let imageData = blipData.subarray(imageDataOffset);

        // For metafile types (EMF/WMF/PICT), data may be zlib-compressed
        if (blipType === ESCHER_BLIP_EMF || blipType === ESCHER_BLIP_WMF || blipType === ESCHER_BLIP_PICT) {
          try {
            imageData = inflateSync(imageData);
          } catch {
            // Not compressed or decompression failed, use raw data
          }
        }

        blips.push({
          index: bseIndex,
          mimeType,
          extension,
          data: imageData,
        });
      }
    }
  }

  return blips;
}

/**
 * Parse drawing records per sheet to extract anchor positions and BLIP references.
 */
function parseSheetDrawings(drawingData: Buffer, sheetName: string): DrawingAnchor[] {
  const anchors: DrawingAnchor[] = [];

  // Parse the Escher records in the MsoDrawing data
  // Structure: DgContainer → SpgrContainer → SpContainer(s)
  // Each SpContainer can have: SP + ClientAnchor + ClientData

  function processContainer(buf: Buffer): void {
    for (const rec of iterEscherRecords(buf)) {
      if (isContainer(rec)) {
        processContainer(rec.data);
        continue;
      }

      if (rec.type === ESCHER_SP_CONTAINER) {
        // This shouldn't happen since SP_CONTAINER is a container, but handle defensively
        processContainer(rec.data);
        continue;
      }
    }

    // Look for SP_CONTAINERs that have both SP and ClientAnchor
    // We need to process SpContainers as groups
    const records: EscherRecord[] = [];
    for (const rec of iterEscherRecords(buf)) {
      records.push(rec);
    }

    // Process SpContainers
    for (const rec of records) {
      if (rec.type === ESCHER_SP_CONTAINER && isContainer(rec)) {
        const spChildren: EscherRecord[] = [];
        for (const child of iterEscherRecords(rec.data)) {
          spChildren.push(child);
        }

        let spId = 0;
        let blipIndex = 0;
        let anchor: { fromCol: number; fromRow: number; toCol: number; toRow: number } | null = null;

        for (const child of spChildren) {
          if (child.type === ESCHER_SP && child.data.length >= 4) {
            spId = child.data.readInt32LE(0);
          }

          if (child.type === ESCHER_CLIENT_ANCHOR && child.data.length >= 18) {
            // ClientAnchor structure:
            // 2 bytes flags
            // 2 bytes col1, 2 bytes dx1, 2 bytes row1, 2 bytes dy1
            // 2 bytes col2, 2 bytes dx2, 2 bytes row2, 2 bytes dy2
            const col1 = child.data.readUInt16LE(2);
            const row1 = child.data.readUInt16LE(6);
            const col2 = child.data.readUInt16LE(10);
            const row2 = child.data.readUInt16LE(14);
            anchor = { fromCol: col1, fromRow: row1, toCol: col2, toRow: row2 };
          }
        }

        // Try to find the OPT record to get the BLIP reference
        for (const child of spChildren) {
          // OPT/FOPT records have type 0xF00B or 0xF122
          if ((child.type === 0xf00b || child.type === 0xf122) && child.data.length >= 6) {
            const numProps = child.instance;
            let propOffset = 0;
            for (let i = 0; i < numProps && propOffset + 6 <= child.data.length; i++) {
              const propId = child.data.readUInt16LE(propOffset);
              const propValue = child.data.readInt32LE(propOffset + 2);
              propOffset += 6;

              // Property 0x0104 = blipId (pib - picture index in BSE)
              if ((propId & 0x3fff) === 0x0104) {
                blipIndex = propValue;
              }
            }
          }
        }

        if (blipIndex > 0 && anchor) {
          anchors.push({
            sheetName,
            spId,
            blipIndex,
            ...anchor,
          });
        }
      }
    }
  }

  processContainer(drawingData);
  return anchors;
}

export async function extractXlsImages(args: GetExcelImagesArgs): Promise<{
  images: ExtractedImage[];
  truncated: boolean;
}> {
  const { filePath, sheetName } = args;

  if (!existsSync(filePath)) {
    throw new McpError(ErrorCode.InvalidRequest, `File not found: ${filePath}`);
  }

  const buffer = readFileSync(filePath);

  // Parse with SheetJS to get sheet names
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetNames = workbook.SheetNames;

  if (sheetName && !sheetNames.includes(sheetName)) {
    throw new McpError(ErrorCode.InvalidRequest, `Sheet not found: ${sheetName}`);
  }

  // Parse CFB to get the Workbook stream
  const cfb = CFB.read(buffer, { type: 'buffer' });
  const workbookEntry = CFB.find(cfb, '/Workbook') || CFB.find(cfb, '/Book');
  if (!workbookEntry) {
    throw new McpError(ErrorCode.InvalidRequest, 'Cannot find Workbook stream in .xls file');
  }

  const wbBuf = Buffer.from(workbookEntry.content);
  const records = readBiffRecords(wbBuf);

  // 1. Collect BOUNDSHEET records (in order) to map sheet index → name
  // 2. Collect MsoDrawingGroup data (global images)
  // 3. Collect per-sheet MsoDrawing data (positions)

  const drawingGroupBuffers: Buffer[] = [];
  const sheetDrawings: Map<number, Buffer[]> = new Map(); // sheetIndex → MsoDrawing buffers

  // BIFF8 structure: BOF(globals)...EOF, BOF(sheet0)...EOF, BOF(sheet1)...EOF, ...
  // All substreams are sequential at the same level, NOT nested.
  let substreamIndex = -1; // -1 = not yet started, 0 = globals, 1+ = sheets
  let inSubstream = false;

  for (const rec of records) {
    if (rec.type === BIFF_BOF) {
      substreamIndex++;
      inSubstream = true;
    }

    if (rec.type === BIFF_EOF) {
      inSubstream = false;
    }

    // Collect all MsoDrawingGroup records (global scope, substreamIndex === 0)
    if (rec.type === BIFF_MSODRAWINGGROUP) {
      drawingGroupBuffers.push(rec.data);
    }

    // Collect MsoDrawing records per sheet (substreamIndex >= 1 → sheet index = substreamIndex - 1)
    if (rec.type === BIFF_MSODRAWING && inSubstream && substreamIndex >= 1) {
      const sheetIdx = substreamIndex - 1;
      const existing = sheetDrawings.get(sheetIdx) || [];
      existing.push(rec.data);
      sheetDrawings.set(sheetIdx, existing);
    }
  }

  // Concatenate all MsoDrawingGroup data
  const drawingGroupData = drawingGroupBuffers.length > 0
    ? Buffer.concat(drawingGroupBuffers)
    : null;

  if (!drawingGroupData || drawingGroupData.length === 0) {
    return { images: [], truncated: false };
  }

  // Extract BLIP data from the drawing group
  const blips = extractBlipsFromDrawingGroup(drawingGroupData);

  if (blips.length === 0) {
    return { images: [], truncated: false };
  }

  // Extract position anchors from per-sheet drawings
  const allAnchors: DrawingAnchor[] = [];

  for (const [sheetIdx, drawingBuffers] of sheetDrawings.entries()) {
    if (sheetIdx >= sheetNames.length) continue;
    const sName = sheetNames[sheetIdx];
    if (sheetName && sName !== sheetName) continue;

    // Concatenate all MsoDrawing buffers for this sheet
    const combinedDrawing = Buffer.concat(drawingBuffers);
    const anchors = parseSheetDrawings(combinedDrawing, sName);
    allAnchors.push(...anchors);
  }

  // Map BLIP index → positions
  const blipPositions: Map<number, ImagePosition[]> = new Map();
  for (const anchor of allAnchors) {
    const positions = blipPositions.get(anchor.blipIndex) || [];
    positions.push({
      sheet: anchor.sheetName,
      fromRow: anchor.fromRow,
      fromCol: anchor.fromCol,
      toRow: anchor.toRow,
      toCol: anchor.toCol,
    });
    blipPositions.set(anchor.blipIndex, positions);
  }

  // Build result
  const images: ExtractedImage[] = [];
  let totalSize = 0;
  let truncated = false;

  for (const blip of blips) {
    const base64 = blip.data.toString('base64');
    if (totalSize + base64.length > MAX_IMAGES_SIZE) {
      truncated = true;
      break;
    }
    totalSize += base64.length;

    const positions = blipPositions.get(blip.index) || [];

    // If filtering by sheet and this image has no positions in the target sheet, skip
    if (sheetName && positions.length === 0) continue;

    images.push({
      name: `image${blip.index}${blip.extension}`,
      mimeType: blip.mimeType,
      data: base64,
      positions,
    });
  }

  return { images, truncated };
}
