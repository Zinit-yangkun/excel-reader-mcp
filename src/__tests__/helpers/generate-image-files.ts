import { writeFileSync } from "node:fs";
import { join } from "node:path";
import * as CFB from "cfb";
import JSZip from "jszip";

// A minimal 1x1 red PNG (68 bytes)
const TINY_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==",
  "base64",
);

// A minimal 1x1 blue JPEG
const TINY_JPEG = Buffer.from(
  "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/xAAUAQEAAAAAAAAAAAAAAAAAAAAA/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8AKwA=",
  "base64",
);

export async function createXlsxWithImages(testDir: string): Promise<string> {
  const zip = new JSZip();

  // [Content_Types].xml
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`,
  );

  // _rels/.rels
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  );

  // xl/workbook.xml
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
  );

  // xl/_rels/workbook.xml.rels
  zip.file(
    "xl/_rels/workbook.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
  );

  // xl/worksheets/sheet1.xml - minimal sheet with data
  zip.file(
    "xl/worksheets/sheet1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`,
  );

  // xl/worksheets/sheet2.xml
  zip.file(
    "xl/worksheets/sheet2.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>`,
  );

  // xl/worksheets/_rels/sheet1.xml.rels
  zip.file(
    "xl/worksheets/_rels/sheet1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
  );

  // xl/worksheets/_rels/sheet2.xml.rels
  zip.file(
    "xl/worksheets/_rels/sheet2.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>
</Relationships>`,
  );

  // xl/drawings/drawing1.xml - two images on Sheet1
  zip.file(
    "xl/drawings/drawing1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="image1.png"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="image2.jpeg"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId2"/></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
  );

  // xl/drawings/drawing2.xml - one image on Sheet2
  zip.file(
    "xl/drawings/drawing2.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="image1.png"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`,
  );

  // xl/drawings/_rels/drawing1.xml.rels
  zip.file(
    "xl/drawings/_rels/drawing1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.jpeg"/>
</Relationships>`,
  );

  // xl/drawings/_rels/drawing2.xml.rels
  zip.file(
    "xl/drawings/_rels/drawing2.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`,
  );

  // xl/sharedStrings.xml
  zip.file(
    "xl/sharedStrings.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Data</t></si>
</sst>`,
  );

  // Media files
  zip.file("xl/media/image1.png", TINY_PNG);
  zip.file("xl/media/image2.jpeg", TINY_JPEG);

  const buf = await zip.generateAsync({ type: "nodebuffer" });
  const filePath = join(testDir, "with-images.xlsx");
  writeFileSync(filePath, buf);
  return filePath;
}

export function createXlsWithImages(testDir: string): string {
  // Build a minimal BIFF8 .xls file with embedded PNG image using Escher records.
  //
  // Structure:
  //   Workbook globals: BOF, BOUNDSHEET(x1), MsoDrawingGroup, EOF
  //   Sheet substream:  BOF, MsoDrawing, OBJ, EOF

  const parts: Buffer[] = [];

  function writeBiffRecord(type: number, data: Buffer): Buffer {
    // BIFF records can be at most 8224 bytes of data.
    // If data is larger, we need CONTINUE records.
    const MAX_REC_DATA = 8224;
    const chunks: Buffer[] = [];

    const firstChunkLen = Math.min(data.length, MAX_REC_DATA);
    const header = Buffer.alloc(4);
    header.writeUInt16LE(type, 0);
    header.writeUInt16LE(firstChunkLen, 2);
    chunks.push(header, data.subarray(0, firstChunkLen));

    let offset = firstChunkLen;
    while (offset < data.length) {
      const chunkLen = Math.min(data.length - offset, MAX_REC_DATA);
      const contHeader = Buffer.alloc(4);
      contHeader.writeUInt16LE(BIFF_CONTINUE, 0);
      contHeader.writeUInt16LE(chunkLen, 2);
      chunks.push(contHeader, data.subarray(offset, offset + chunkLen));
      offset += chunkLen;
    }

    return Buffer.concat(chunks);
  }

  function writeEscherRecord(type: number, version: number, instance: number, data: Buffer): Buffer {
    const header = Buffer.alloc(8);
    header.writeUInt16LE((instance << 4) | (version & 0x0f), 0);
    header.writeUInt16LE(type, 2);
    header.writeInt32LE(data.length, 4);
    return Buffer.concat([header, data]);
  }

  // ---- Build the Escher DggContainer with one PNG BLIP ----

  // PNG BLIP: UID(16) + tag(1) + imageData
  const uid = Buffer.alloc(16, 0x42); // dummy UID
  const blipTag = Buffer.alloc(1, 0xff);
  const blipData = Buffer.concat([uid, blipTag, TINY_PNG]);
  // BLIP_PNG record: instance=0x6E0 (no UID2), version=0
  const blipRecord = writeEscherRecord(0xf01e, 0, 0x6e0, blipData);

  // BSE record wrapping the BLIP
  // BSE header: btWin32(1) + btMacOS(1) + UID(16) + tag(2) + size(4) + cRef(4) + foDelay(4) + usage(1) + cbName(1) + unused(2)
  const bseHeader = Buffer.alloc(36);
  bseHeader.writeUInt8(6, 0); // btWin32 = PNG
  bseHeader.writeUInt8(6, 1); // btMacOS = PNG
  uid.copy(bseHeader, 2); // UID
  bseHeader.writeUInt16LE(0xff, 18); // tag
  bseHeader.writeInt32LE(blipRecord.length, 20); // size of BLIP
  bseHeader.writeInt32LE(1, 24); // cRef = 1
  bseHeader.writeInt32LE(0, 28); // foDelay
  bseHeader.writeUInt8(0, 32); // usage
  bseHeader.writeUInt8(0, 33); // cbName
  const bseData = Buffer.concat([bseHeader, blipRecord]);
  const bseRecord = writeEscherRecord(0xf007, 2, 6, bseData); // BSE, instance=btWin32=6

  // BStoreContainer
  const bstoreRecord = writeEscherRecord(0xf001, 0x0f, 1, bseRecord); // instance=count=1

  // DGG atom: spidMax(4) + cidcl(4) + cspSaved(4) + cdgSaved(4) + FIDCL[](8 each)
  const dggAtomData = Buffer.alloc(24);
  dggAtomData.writeInt32LE(1026, 0); // spidMax
  dggAtomData.writeInt32LE(2, 4); // cidcl (number of ID clusters + 1)
  dggAtomData.writeInt32LE(1, 8); // cspSaved
  dggAtomData.writeInt32LE(1, 12); // cdgSaved
  // FIDCL entry: dgid(4) + cspidCur(4)
  dggAtomData.writeInt32LE(1, 16); // dgid
  dggAtomData.writeInt32LE(2, 20); // cspidCur
  const dggAtom = writeEscherRecord(0xf006, 0, 0, dggAtomData);

  // DggContainer = dggAtom + bstoreContainer
  const dggContainerData = Buffer.concat([dggAtom, bstoreRecord]);
  const dggContainer = writeEscherRecord(0xf000, 0x0f, 0, dggContainerData);

  // ---- Build per-sheet Escher DgContainer with SpContainer ----

  // DG atom
  const dgAtomData = Buffer.alloc(8);
  dgAtomData.writeInt32LE(1, 0); // csp (number of shapes including group shape)
  dgAtomData.writeInt32LE(1025, 4); // spidCur (last SPID)
  const dgAtom = writeEscherRecord(0xf008, 0, 1, dgAtomData); // instance=1 (dgId)

  // Group shape SpContainer (required as first shape)
  const spGroupData = Buffer.alloc(8);
  spGroupData.writeInt32LE(1024, 0); // spid
  spGroupData.writeInt32LE(0x05, 4); // flags: group + patriarch
  const spGroupRecord = writeEscherRecord(0xf00a, 2, 0, spGroupData);

  // SpGR record (group coordinate system) - 16 bytes of zeros
  const spgrData = Buffer.alloc(16);
  const spgrRecord = writeEscherRecord(0xf009, 1, 0, spgrData);

  const groupSpContainer = writeEscherRecord(0xf004, 0x0f, 0, Buffer.concat([spgrRecord, spGroupRecord]));

  // Image SpContainer
  // SP record
  const spData = Buffer.alloc(8);
  spData.writeInt32LE(1025, 0); // spid
  spData.writeInt32LE(0x0a00, 4); // flags: hasAnchor + hasShapeType
  const spRecord = writeEscherRecord(0xf00a, 2, 75, spData); // instance=75 (msosptPictureFrame)

  // OPT record with blipId property (0x0104)
  // Property: id(2) + value(4) = 6 bytes per property
  const optPropData = Buffer.alloc(6);
  optPropData.writeUInt16LE(0x4104, 0); // propId 0x0104 with fComplex=0, fBid=1
  optPropData.writeInt32LE(1, 2); // blipIndex = 1 (1-based BSE index)
  const optRecord = writeEscherRecord(0xf00b, 3, 1, optPropData); // version=3, instance=numProps=1

  // ClientAnchor: flags(2) + col1(2) + dx1(2) + row1(2) + dy1(2) + col2(2) + dx2(2) + row2(2) + dy2(2) = 18 bytes
  const anchorData = Buffer.alloc(18);
  anchorData.writeUInt16LE(0, 0); // flags
  anchorData.writeUInt16LE(0, 2); // col1 (fromCol)
  anchorData.writeUInt16LE(0, 4); // dx1
  anchorData.writeUInt16LE(0, 6); // row1 (fromRow)
  anchorData.writeUInt16LE(0, 8); // dy1
  anchorData.writeUInt16LE(3, 10); // col2 (toCol)
  anchorData.writeUInt16LE(0, 12); // dx2
  anchorData.writeUInt16LE(4, 14); // row2 (toRow)
  anchorData.writeUInt16LE(0, 16); // dy2
  const clientAnchor = writeEscherRecord(0xf010, 0, 0, anchorData);

  // ClientData
  const clientDataRecord = writeEscherRecord(0xf011, 0, 0, Buffer.alloc(0));

  const imageSpContainer = writeEscherRecord(
    0xf004,
    0x0f,
    0,
    Buffer.concat([spRecord, optRecord, clientAnchor, clientDataRecord]),
  );

  // SpGR container wrapping both shape containers
  const spgrContainer = writeEscherRecord(0xf003, 0x0f, 0, Buffer.concat([groupSpContainer, imageSpContainer]));

  // DgContainer
  const dgContainer = writeEscherRecord(0xf002, 0x0f, 0, Buffer.concat([dgAtom, spgrContainer]));

  // ---- Build the BIFF8 workbook stream ----

  // BOF record for globals (BIFF8 workbook globals)
  const bofData = Buffer.alloc(16);
  bofData.writeUInt16LE(0x0600, 0); // version: BIFF8
  bofData.writeUInt16LE(0x0005, 2); // type: workbook globals
  bofData.writeUInt16LE(0x0dbb, 4); // build identifier
  bofData.writeUInt16LE(0x07cc, 6); // build year
  parts.push(writeBiffRecord(BIFF_BOF, bofData));

  // BOUNDSHEET record for "ImageSheet"
  const sheetNameStr = "ImageSheet";
  const boundSheetData = Buffer.alloc(8 + sheetNameStr.length);
  boundSheetData.writeInt32LE(0, 0); // BOF position (will be patched)
  boundSheetData.writeUInt8(0x00, 4); // visibility: visible
  boundSheetData.writeUInt8(0x00, 5); // sheet type: worksheet
  boundSheetData.writeUInt8(sheetNameStr.length, 6); // name length
  boundSheetData.writeUInt8(0x00, 7); // flag: 0 = compressed (Latin1)
  Buffer.from(sheetNameStr, "latin1").copy(boundSheetData, 8);
  parts.push(writeBiffRecord(BIFF_BOUNDSHEET, boundSheetData));

  // MsoDrawingGroup
  parts.push(writeBiffRecord(BIFF_MSODRAWINGGROUP, dggContainer));

  // EOF
  parts.push(writeBiffRecord(BIFF_EOF, Buffer.alloc(0)));

  // Record the offset where the sheet substream starts (for BOUNDSHEET patching)
  const globalsLen = parts.reduce((sum, b) => sum + b.length, 0);

  // BOF for sheet substream
  const sheetBofData = Buffer.alloc(16);
  sheetBofData.writeUInt16LE(0x0600, 0); // version: BIFF8
  sheetBofData.writeUInt16LE(0x0010, 2); // type: worksheet
  sheetBofData.writeUInt16LE(0x0dbb, 4);
  sheetBofData.writeUInt16LE(0x07cc, 6);
  parts.push(writeBiffRecord(BIFF_BOF, sheetBofData));

  // MsoDrawing
  parts.push(writeBiffRecord(BIFF_MSODRAWING, dgContainer));

  // OBJ record (required after MsoDrawing, minimal)
  const objData = Buffer.alloc(26);
  // ftCmo sub-record: ft(2)=0x15 + cb(2)=0x12 + ot(2)=0x08(Picture) + id(2) + flags(2) + reserved(12)
  objData.writeUInt16LE(0x15, 0); // ft = ftCmo
  objData.writeUInt16LE(0x12, 2); // cb = 18
  objData.writeUInt16LE(0x08, 4); // ot = Picture
  objData.writeUInt16LE(1, 6); // id = 1
  objData.writeUInt16LE(0x6011, 8); // flags
  // ftEnd sub-record
  objData.writeUInt16LE(0x00, 22); // ft = ftEnd
  objData.writeUInt16LE(0x00, 24); // cb = 0
  parts.push(writeBiffRecord(BIFF_OBJ, objData));

  // EOF
  parts.push(writeBiffRecord(BIFF_EOF, Buffer.alloc(0)));

  // Concatenate all parts
  const workbookStream = Buffer.concat(parts);

  // Patch the BOUNDSHEET record's BOF offset
  // BOUNDSHEET is the second record. Find its data offset.
  // First record (BOF): 4 (header) + 16 (data) = 20
  // BOUNDSHEET header starts at offset 20, data at 24
  workbookStream.writeInt32LE(globalsLen, 24);

  // Build CFB container
  const cfbFile = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfbFile, "/Workbook", workbookStream);
  const cfbBuf = CFB.write(cfbFile, { type: "buffer" }) as Buffer;

  const filePath = join(testDir, "with-images.xls");
  writeFileSync(filePath, cfbBuf);
  return filePath;
}

const BIFF_CONTINUE = 0x003c;
const BIFF_BOF = 0x0809;
const BIFF_BOUNDSHEET = 0x0085;
const BIFF_MSODRAWINGGROUP = 0x00eb;
const BIFF_MSODRAWING = 0x00ec;
const BIFF_EOF = 0x000a;
const BIFF_OBJ = 0x005d;
