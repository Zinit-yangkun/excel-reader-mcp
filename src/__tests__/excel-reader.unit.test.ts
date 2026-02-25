import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { join } from 'path';
import { readExcelFile, listSheets } from '../excel-reader.js';
import { setupTestFiles, cleanupTestFiles } from './helpers/generate-test-files.js';

let testDir: string;

beforeAll(() => {
  testDir = setupTestFiles();
});

afterAll(() => {
  cleanupTestFiles(testDir);
});

describe('readExcelFile', () => {
  it('should read a basic file and return correct structure', () => {
    const data = readExcelFile({ filePath: join(testDir, 'basic.xlsx') });

    expect(data.fileName).toBe('basic.xlsx');
    expect(data.totalSheets).toBe(1);
    expect(data.currentSheet.name).toBe('Sheet1');
    expect(data.currentSheet.totalRows).toBe(3);
    expect(data.currentSheet.totalColumns).toBe(4);
    expect(data.currentSheet.chunk.columns).toEqual(['Name', 'Age', 'Date', 'Score']);
  });

  it('should have correct data content', () => {
    const data = readExcelFile({ filePath: join(testDir, 'basic.xlsx') });
    const rows = data.currentSheet.chunk.data;

    expect(rows[0].Name).toBe('Alice');
    expect(rows[0].Age).toBe(30);
    expect(rows[0].Score).toBe(95.5);
    expect(rows[1].Name).toBe('Bob');
    expect(rows[1].Age).toBe(25);
  });

  it('should read a specific sheet by name', () => {
    const data = readExcelFile({
      filePath: join(testDir, 'multi-sheet.xlsx'),
      sheetName: 'Cities',
    });

    expect(data.currentSheet.name).toBe('Cities');
    expect(data.currentSheet.chunk.data[0].City).toBe('Tokyo');
    expect(data.currentSheet.chunk.data[0].Population).toBe(13960000);
  });

  it('should support startRow and maxRows pagination', () => {
    const data = readExcelFile({
      filePath: join(testDir, 'basic.xlsx'),
      startRow: 1,
      maxRows: 1,
    });

    expect(data.currentSheet.chunk.rowStart).toBe(1);
    expect(data.currentSheet.chunk.rowEnd).toBe(2);
    expect(data.currentSheet.chunk.data).toHaveLength(1);
    expect(data.currentSheet.chunk.data[0].Name).toBe('Bob');
    expect(data.currentSheet.hasMore).toBe(true);
  });

  it('should auto-chunk a large file with hasMore=true', () => {
    const data = readExcelFile({ filePath: join(testDir, 'large.xlsx') });

    expect(data.currentSheet.totalRows).toBe(600);
    expect(data.currentSheet.chunk.data.length).toBeLessThan(600);
    expect(data.currentSheet.hasMore).toBe(true);
    expect(data.currentSheet.nextChunk).toBeDefined();
    expect(data.currentSheet.nextChunk!.rowStart).toBe(data.currentSheet.chunk.rowEnd);
  });

  it('should handle an empty file', () => {
    const data = readExcelFile({ filePath: join(testDir, 'empty.xlsx') });

    expect(data.currentSheet.totalRows).toBe(0);
    expect(data.currentSheet.chunk.data).toHaveLength(0);
    expect(data.currentSheet.hasMore).toBe(false);
  });

  it('should throw for a non-existent file', () => {
    expect(() =>
      readExcelFile({ filePath: '/non/existent/file.xlsx' })
    ).toThrow(/File not found/);
  });
});

describe('listSheets', () => {
  it('should list sheets for a single-sheet file', () => {
    const result = listSheets({ filePath: join(testDir, 'basic.xlsx') });

    expect(result.fileName).toBe('basic.xlsx');
    expect(result.sheets).toEqual(['Sheet1']);
  });

  it('should list all sheets for a multi-sheet file', () => {
    const result = listSheets({ filePath: join(testDir, 'multi-sheet.xlsx') });

    expect(result.sheets).toEqual(['Products', 'Cities', 'Colors']);
  });

  it('should throw for a non-existent file', () => {
    expect(() =>
      listSheets({ filePath: '/non/existent/file.xlsx' })
    ).toThrow(/File not found/);
  });
});
