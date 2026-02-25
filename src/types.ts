export interface ExcelChunk {
  rowStart: number;
  rowEnd: number;
  columns: string[];
  data: Record<string, unknown>[];
}

export interface ExcelSheetData {
  name: string;
  totalRows: number;
  totalColumns: number;
  chunk: ExcelChunk;
  hasMore: boolean;
  nextChunk?: {
    rowStart: number;
    columns: string[];
  };
}

export interface ExcelData {
  fileName: string;
  totalSheets: number;
  currentSheet: ExcelSheetData;
}

export interface ReadExcelArgs {
  filePath: string;
  sheetName?: string;
  startRow?: number;
  maxRows?: number;
}

export interface ListSheetsArgs {
  filePath: string;
}

export interface ImagePosition {
  sheet: string;
  fromRow: number;
  fromCol: number;
  toRow: number;
  toCol: number;
}

export interface ExtractedImage {
  name: string;
  mimeType: string;
  data: string;
  positions: ImagePosition[];
}

export interface GetExcelImagesArgs {
  filePath: string;
  sheetName?: string;
}
