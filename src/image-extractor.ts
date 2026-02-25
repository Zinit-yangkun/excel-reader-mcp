import { existsSync, readFileSync } from "node:fs";
import { ErrorCode, McpError } from "@modelcontextprotocol/sdk/types.js";
import JSZip from "jszip";
import * as XLSX from "xlsx";
import type { ExtractedImage, GetExcelImagesArgs, ImagePosition } from "./types.js";
import { extractXlsImages } from "./xls-image-extractor.js";

const MIME_TYPES: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".bmp": "image/bmp",
  ".tiff": "image/tiff",
  ".tif": "image/tiff",
  ".emf": "image/x-emf",
  ".wmf": "image/x-wmf",
  ".svg": "image/svg+xml",
};

const MAX_IMAGES_SIZE = 10 * 1024 * 1024; // 10MB total base64 size limit

function resolveRelativePath(baseDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = baseDir.split("/");
  const targetParts = target.split("/");
  for (const part of targetParts) {
    if (part === "..") {
      parts.pop();
    } else if (part !== ".") {
      parts.push(part);
    }
  }
  return parts.join("/");
}

export async function extractImages(args: GetExcelImagesArgs): Promise<{
  images: ExtractedImage[];
  truncated: boolean;
}> {
  const { filePath, sheetName } = args;
  if (!existsSync(filePath)) {
    throw new McpError(ErrorCode.InvalidRequest, `File not found: ${filePath}`);
  }

  const buffer = readFileSync(filePath);

  // Detect file format by magic bytes
  if (buffer.length < 4) {
    throw new McpError(ErrorCode.InvalidRequest, "File is too small to be a valid Excel file.");
  }

  const isZip = buffer[0] === 0x50 && buffer[1] === 0x4b; // PK = ZIP (.xlsx)
  const isOle2 = buffer[0] === 0xd0 && buffer[1] === 0xcf && buffer[2] === 0x11 && buffer[3] === 0xe0; // OLE2 (.xls)

  if (!isZip && !isOle2) {
    throw new McpError(
      ErrorCode.InvalidRequest,
      "File is not a valid Excel file. Expected .xlsx (ZIP) or .xls (OLE2) format.",
    );
  }

  // Dispatch to xls extractor for OLE2 format
  if (isOle2) {
    return extractXlsImages(args);
  }

  const zip = await JSZip.loadAsync(buffer);

  // Get sheet names from the workbook via SheetJS
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheetNames = workbook.SheetNames;

  if (sheetName && !sheetNames.includes(sheetName)) {
    throw new McpError(ErrorCode.InvalidRequest, `Sheet not found: ${sheetName}`);
  }

  // 1. Collect all images from xl/media/
  const mediaFiles: Map<string, { data: string; mimeType: string }> = new Map();
  for (const [path, file] of Object.entries(zip.files)) {
    if (path.startsWith("xl/media/") && !file.dir) {
      const ext = `.${path.split(".").pop()?.toLowerCase()}`;
      const mimeType = MIME_TYPES[ext] || "application/octet-stream";
      const base64 = await file.async("base64");
      const fileName = path.split("/").pop()!;
      mediaFiles.set(path, { data: base64, mimeType });
      // Also index by relative path from xl/ for rels matching
      mediaFiles.set(`media/${fileName}`, { data: base64, mimeType });
    }
  }

  if (mediaFiles.size === 0) {
    return { images: [], truncated: false };
  }

  // 2. Parse workbook.xml.rels to find sheet→file mapping
  const workbookRelsFile = zip.files["xl/_rels/workbook.xml.rels"];
  const sheetFileMap: Map<string, number> = new Map(); // e.g. "xl/worksheets/sheet1.xml" → 0 (index in sheetNames)

  if (workbookRelsFile) {
    const workbookRelsXml = await workbookRelsFile.async("text");
    // Match sheet relationships
    const sheetRels = [...workbookRelsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)];

    // Parse workbook.xml to get sheet name → rId mapping
    const workbookFile = zip.files["xl/workbook.xml"];
    if (workbookFile) {
      const workbookXml = await workbookFile.async("text");
      const sheetEntries = [...workbookXml.matchAll(/<sheet[^>]+name="([^"]+)"[^>]+r:id="(rId\d+)"/g)];

      for (const sheetEntry of sheetEntries) {
        const name = sheetEntry[1];
        const rId = sheetEntry[2];
        const rel = sheetRels.find((r) => r[1] === rId);
        if (rel) {
          const target = rel[2].startsWith("/") ? rel[2].slice(1) : `xl/${rel[2]}`;
          const sheetIdx = sheetNames.indexOf(name);
          if (sheetIdx !== -1) {
            sheetFileMap.set(target, sheetIdx);
          }
        }
      }
    }
  }

  // 3. For each sheet, parse drawing relationships to map images to positions
  type ImageMapping = {
    imagePath: string;
    position: ImagePosition;
  };
  const imageMappings: ImageMapping[] = [];

  for (const [sheetPath, sheetIdx] of sheetFileMap.entries()) {
    const currentSheetName = sheetNames[sheetIdx];
    if (sheetName && currentSheetName !== sheetName) continue;

    // Find the sheet's rels file
    const sheetFileName = sheetPath.split("/").pop()!;
    const sheetDir = sheetPath.substring(0, sheetPath.lastIndexOf("/"));
    const sheetRelsPath = `${sheetDir}/_rels/${sheetFileName}.rels`;
    const sheetRelsFile = zip.files[sheetRelsPath];
    if (!sheetRelsFile) continue;

    const sheetRelsXml = await sheetRelsFile.async("text");
    // Find drawing relationships
    const drawingRels = [...sheetRelsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]*drawing[^"]*)"/gi)];

    for (const drawingRel of drawingRels) {
      const drawingTarget = drawingRel[2];
      // Resolve "../drawings/drawing1.xml" relative to sheet dir
      const normalizedDrawingPath = resolveRelativePath(sheetDir, drawingTarget);

      const drawingFile = zip.files[normalizedDrawingPath];
      if (!drawingFile) continue;

      const drawingXml = await drawingFile.async("text");

      // Parse the drawing's own rels to map rId → image file
      const drawingFileName = normalizedDrawingPath.split("/").pop()!;
      const drawingDir = normalizedDrawingPath.substring(0, normalizedDrawingPath.lastIndexOf("/"));
      const drawingRelsPath = `${drawingDir}/_rels/${drawingFileName}.rels`;
      const drawingRelsFile = zip.files[drawingRelsPath];
      const imageRIdMap: Map<string, string> = new Map(); // rId → full image path

      if (drawingRelsFile) {
        const drawingRelsXml = await drawingRelsFile.async("text");
        const imageRels = [...drawingRelsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)];
        for (const imageRel of imageRels) {
          const fullPath = resolveRelativePath(drawingDir, imageRel[2]);
          imageRIdMap.set(imageRel[1], fullPath);
        }
      }

      // Parse twoCellAnchor and oneCellAnchor elements
      const anchorPatterns = [
        // twoCellAnchor: has from and to
        /<xdr:twoCellAnchor[^>]*>([\s\S]*?)<\/xdr:twoCellAnchor>/g,
        // oneCellAnchor: has from only
        /<xdr:oneCellAnchor[^>]*>([\s\S]*?)<\/xdr:oneCellAnchor>/g,
      ];

      for (const pattern of anchorPatterns) {
        const isTwoCell = pattern.source.includes("twoCellAnchor");
        let match: RegExpExecArray | null;
        while ((match = pattern.exec(drawingXml)) !== null) {
          const anchorContent = match[1];

          // Extract ALL image rIds from blipFill (handles grouped images)
          const allBlipMatches = [...anchorContent.matchAll(/r:embed="(rId\d+)"/g)];
          if (allBlipMatches.length === 0) continue;

          // Extract from position
          const fromMatch = anchorContent.match(
            /<xdr:from>\s*<xdr:col>(\d+)<\/xdr:col>\s*<xdr:colOff>\d+<\/xdr:colOff>\s*<xdr:row>(\d+)<\/xdr:row>/,
          );
          const fromRow = fromMatch ? parseInt(fromMatch[2], 10) : 0;
          const fromCol = fromMatch ? parseInt(fromMatch[1], 10) : 0;

          let toRow = fromRow;
          let toCol = fromCol;
          if (isTwoCell) {
            const toMatch = anchorContent.match(
              /<xdr:to>\s*<xdr:col>(\d+)<\/xdr:col>\s*<xdr:colOff>\d+<\/xdr:colOff>\s*<xdr:row>(\d+)<\/xdr:row>/,
            );
            toRow = toMatch ? parseInt(toMatch[2], 10) : fromRow;
            toCol = toMatch ? parseInt(toMatch[1], 10) : fromCol;
          }

          for (const blipMatch of allBlipMatches) {
            const imageFullPath = imageRIdMap.get(blipMatch[1]);
            if (!imageFullPath) continue;

            imageMappings.push({
              imagePath: imageFullPath,
              position: {
                sheet: currentSheetName,
                fromRow,
                fromCol,
                toRow,
                toCol,
              },
            });
          }
        }
      }
    }
  }

  // 4. Build result: combine media files with position mappings
  // Deduplicate: group positions by image path
  const imageMap: Map<string, ImagePosition[]> = new Map();
  for (const mapping of imageMappings) {
    const existing = imageMap.get(mapping.imagePath);
    if (existing) {
      existing.push(mapping.position);
    } else {
      imageMap.set(mapping.imagePath, [mapping.position]);
    }
  }

  const images: ExtractedImage[] = [];
  let totalSize = 0;
  let truncated = false;

  // First add images with position info (deduplicated)
  for (const [imagePath, positions] of imageMap.entries()) {
    const media = mediaFiles.get(imagePath);
    if (!media) continue;

    if (totalSize + media.data.length > MAX_IMAGES_SIZE) {
      truncated = true;
      break;
    }
    totalSize += media.data.length;

    const fileName = imagePath.split("/").pop()!;
    images.push({
      name: fileName,
      mimeType: media.mimeType,
      data: media.data,
      positions,
    });
  }

  // Then add unmapped images (if not filtering by sheet)
  if (!truncated && !sheetName) {
    for (const [path, media] of mediaFiles.entries()) {
      if (!path.startsWith("xl/media/")) continue;
      if (imageMap.has(path)) continue;

      if (totalSize + media.data.length > MAX_IMAGES_SIZE) {
        truncated = true;
        break;
      }
      totalSize += media.data.length;

      const fileName = path.split("/").pop()!;
      images.push({
        name: fileName,
        mimeType: media.mimeType,
        data: media.data,
        positions: [],
      });
    }
  }

  return { images, truncated };
}
