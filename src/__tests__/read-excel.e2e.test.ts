import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { join } from 'path';
import { setupTestFiles, cleanupTestFiles } from './helpers/generate-test-files.js';

let client: Client;
let transport: StdioClientTransport;
let testDir: string;

beforeAll(async () => {
  testDir = setupTestFiles();

  transport = new StdioClientTransport({
    command: 'node',
    args: [join(import.meta.dirname, '../../build/index.js')],
  });
  client = new Client({ name: 'test-client', version: '1.0.0' });
  await client.connect(transport);
});

afterAll(async () => {
  await client.close();
  cleanupTestFiles(testDir);
});

describe('read_excel', () => {
  it('should read a basic file and return correct structure', async () => {
    const result = await client.callTool({
      name: 'read_excel',
      arguments: { filePath: join(testDir, 'basic.xlsx') },
    });

    const content = result.content as { type: string; text: string }[];
    expect(content).toHaveLength(1);
    expect(content[0].type).toBe('text');

    const data = JSON.parse(content[0].text);
    expect(data.fileName).toBe('basic.xlsx');
    expect(data.totalSheets).toBe(1);
    expect(data.currentSheet.name).toBe('Sheet1');
    expect(data.currentSheet.totalRows).toBe(3); // 3 data rows (header becomes keys)
    expect(data.currentSheet.totalColumns).toBe(4);
    expect(data.currentSheet.chunk.columns).toEqual(['Name', 'Age', 'Date', 'Score']);
  });

  it('should have correct data content (strings, numbers)', async () => {
    const result = await client.callTool({
      name: 'read_excel',
      arguments: { filePath: join(testDir, 'basic.xlsx') },
    });

    const data = JSON.parse((result.content as any)[0].text);
    const rows = data.currentSheet.chunk.data;

    expect(rows[0].Name).toBe('Alice');
    expect(rows[0].Age).toBe(30);
    expect(rows[0].Score).toBe(95.5);

    expect(rows[1].Name).toBe('Bob');
    expect(rows[1].Age).toBe(25);
  });

  it('should read a specific sheet by name', async () => {
    const result = await client.callTool({
      name: 'read_excel',
      arguments: {
        filePath: join(testDir, 'multi-sheet.xlsx'),
        sheetName: 'Cities',
      },
    });

    const data = JSON.parse((result.content as any)[0].text);
    expect(data.currentSheet.name).toBe('Cities');
    expect(data.currentSheet.chunk.data[0].City).toBe('Tokyo');
    expect(data.currentSheet.chunk.data[0].Population).toBe(13960000);
  });

  it('should support startRow and maxRows pagination', async () => {
    const result = await client.callTool({
      name: 'read_excel',
      arguments: {
        filePath: join(testDir, 'basic.xlsx'),
        startRow: 1,
        maxRows: 1,
      },
    });

    const data = JSON.parse((result.content as any)[0].text);
    expect(data.currentSheet.chunk.rowStart).toBe(1);
    expect(data.currentSheet.chunk.rowEnd).toBe(2);
    expect(data.currentSheet.chunk.data).toHaveLength(1);
    expect(data.currentSheet.chunk.data[0].Name).toBe('Bob');
    expect(data.currentSheet.hasMore).toBe(true);
  });

  it('should auto-chunk a large file with hasMore=true', async () => {
    const result = await client.callTool({
      name: 'read_excel',
      arguments: { filePath: join(testDir, 'large.xlsx') },
    });

    const data = JSON.parse((result.content as any)[0].text);
    expect(data.currentSheet.totalRows).toBe(600);
    // auto-chunking should limit the rows returned
    expect(data.currentSheet.chunk.data.length).toBeLessThan(600);
    expect(data.currentSheet.hasMore).toBe(true);
    expect(data.currentSheet.nextChunk).toBeDefined();
    expect(data.currentSheet.nextChunk.rowStart).toBe(data.currentSheet.chunk.rowEnd);
  });

  it('should handle an empty file', async () => {
    const result = await client.callTool({
      name: 'read_excel',
      arguments: { filePath: join(testDir, 'empty.xlsx') },
    });

    const data = JSON.parse((result.content as any)[0].text);
    expect(data.currentSheet.totalRows).toBe(0);
    expect(data.currentSheet.chunk.data).toHaveLength(0);
    expect(data.currentSheet.hasMore).toBe(false);
  });

  it('should throw an error for a non-existent file', async () => {
    await expect(
      client.callTool({
        name: 'read_excel',
        arguments: { filePath: '/non/existent/file.xlsx' },
      })
    ).rejects.toThrow(/File not found/);
  });
});
