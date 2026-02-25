import { existsSync, readFileSync } from "node:fs";
import * as XLSX from "xlsx";
import type { ExcelData, ListSheetsArgs, ReadExcelArgs } from "./types.js";

const MAX_RESPONSE_SIZE = 100 * 1024; // 100KB default max response size

export const estimateJsonSize = (obj: any): number => {
  const str = JSON.stringify(obj);
  return str.length * 2; // Rough estimate, multiply by 2 for unicode
};

export const calculateChunkSize = (data: any[], maxSize: number): number => {
  const singleRowSize = estimateJsonSize(data[0]);
  return Math.max(1, Math.floor(maxSize / singleRowSize));
};

export function readExcelFile(args: ReadExcelArgs): ExcelData {
  const { filePath, sheetName, startRow = 0, maxRows } = args;
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const data = readFileSync(filePath);
  const workbook = XLSX.read(data, {
    type: "buffer",
    cellDates: true,
    cellNF: false,
    cellText: false,
    dateNF: "yyyy-mm-dd",
  });
  const fileName = filePath.split(/[\\/]/).pop() || "";
  const selectedSheetName = sheetName || workbook.SheetNames[0];
  const worksheet = workbook.Sheets[selectedSheetName];
  const allData = XLSX.utils.sheet_to_json(worksheet, {
    raw: true,
    dateNF: "yyyy-mm-dd",
  }) as Record<string, any>[];

  const totalRows = allData.length;
  const columns = totalRows > 0 ? Object.keys(allData[0] as object) : [];
  const totalColumns = columns.length;

  let effectiveMaxRows = maxRows;
  if (!effectiveMaxRows) {
    const initialChunk = allData.slice(0, 100);
    if (initialChunk.length > 0) {
      effectiveMaxRows = calculateChunkSize(initialChunk, MAX_RESPONSE_SIZE);
    } else {
      effectiveMaxRows = 100;
    }
  }

  const endRow = Math.min(startRow + effectiveMaxRows, totalRows);
  const chunkData = allData.slice(startRow, endRow);

  const hasMore = endRow < totalRows;
  const nextChunk = hasMore
    ? {
        rowStart: endRow,
        columns,
      }
    : undefined;

  return {
    fileName,
    totalSheets: workbook.SheetNames.length,
    currentSheet: {
      name: selectedSheetName,
      totalRows,
      totalColumns,
      chunk: {
        rowStart: startRow,
        rowEnd: endRow,
        columns,
        data: chunkData,
      },
      hasMore,
      nextChunk,
    },
  };
}

export interface ListSheetsResult {
  fileName: string;
  sheets: string[];
}

export function listSheets(args: ListSheetsArgs): ListSheetsResult {
  const { filePath } = args;
  if (!existsSync(filePath)) {
    throw new Error(`File not found: ${filePath}`);
  }

  const data = readFileSync(filePath);
  const workbook = XLSX.read(data, { type: "buffer" });
  const fileName = filePath.split(/[\\/]/).pop() || "";
  return {
    fileName,
    sheets: workbook.SheetNames,
  };
}
