import * as XLSX from 'xlsx';
import { writeFileSync, mkdirSync, rmSync, existsSync } from 'fs';
import { join } from 'path';
import { tmpdir } from 'os';

export function setupTestFiles(): string {
  const testDir = join(tmpdir(), `excel-reader-mcp-test-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`);
  mkdirSync(testDir, { recursive: true });

  createBasicXlsx(testDir);
  createMultiSheetXlsx(testDir);
  createLargeXlsx(testDir);
  createEmptyXlsx(testDir);

  return testDir;
}

export function cleanupTestFiles(testDir: string): void {
  if (testDir && existsSync(testDir)) {
    rmSync(testDir, { recursive: true, force: true });
  }
}

function writeWorkbook(wb: XLSX.WorkBook, testDir: string, filename: string): string {
  const filePath = join(testDir, filename);
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  writeFileSync(filePath, buf);
  return filePath;
}

function createBasicXlsx(testDir: string): void {
  const wb = XLSX.utils.book_new();
  const data = [
    ['Name', 'Age', 'Date', 'Score'],
    ['Alice', 30, new Date('2024-01-15'), 95.5],
    ['Bob', 25, new Date('2024-02-20'), 88.0],
    ['Charlie', 35, new Date('2024-03-10'), 72.3],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data, { cellDates: true });
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  writeWorkbook(wb, testDir, 'basic.xlsx');
}

function createMultiSheetXlsx(testDir: string): void {
  const wb = XLSX.utils.book_new();

  const data1 = [
    ['Product', 'Price'],
    ['Apple', 1.5],
    ['Banana', 0.75],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(data1);
  XLSX.utils.book_append_sheet(wb, ws1, 'Products');

  const data2 = [
    ['City', 'Population'],
    ['Tokyo', 13960000],
    ['London', 8982000],
    ['Paris', 2161000],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(data2);
  XLSX.utils.book_append_sheet(wb, ws2, 'Cities');

  const data3 = [
    ['Color', 'Hex'],
    ['Red', '#FF0000'],
    ['Green', '#00FF00'],
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(data3);
  XLSX.utils.book_append_sheet(wb, ws3, 'Colors');

  writeWorkbook(wb, testDir, 'multi-sheet.xlsx');
}

function createLargeXlsx(testDir: string): void {
  const wb = XLSX.utils.book_new();
  // Create wide rows (20 columns) to ensure total data exceeds 100KB chunk limit
  const headers = Array.from({ length: 20 }, (_, i) => `Column_${i}`);
  const data: any[][] = [headers];
  for (let i = 1; i <= 600; i++) {
    const row = Array.from({ length: 20 }, (_, j) => `row_${i}_col_${j}_value_data`);
    data.push(row);
  }
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  writeWorkbook(wb, testDir, 'large.xlsx');
}

function createEmptyXlsx(testDir: string): void {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([]);
  XLSX.utils.book_append_sheet(wb, ws, 'Empty');
  writeWorkbook(wb, testDir, 'empty.xlsx');
}
