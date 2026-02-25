import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { createXlsWithImages, createXlsxWithImages } from "./helpers/generate-image-files.js";
import { cleanupTestFiles, setupTestFiles } from "./helpers/generate-test-files.js";

let client: Client;
let transport: StdioClientTransport;
let testDir: string;
let xlsxImagePath: string;
let xlsImagePath: string;

beforeAll(async () => {
  testDir = setupTestFiles();
  xlsxImagePath = await createXlsxWithImages(testDir);
  xlsImagePath = createXlsWithImages(testDir);

  transport = new StdioClientTransport({
    command: "node",
    args: [join(import.meta.dirname, "../../build/index.js")],
  });
  client = new Client({ name: "test-client", version: "1.0.0" });
  await client.connect(transport);
});

afterAll(async () => {
  await client.close();
  cleanupTestFiles(testDir);
});

describe("get_excel_images - xlsx", () => {
  it("should extract images from xlsx file", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: xlsxImagePath },
    });

    const content = result.content as any[];
    // First content block is metadata JSON
    const metadata = JSON.parse(content[0].text);
    expect(metadata.imageCount).toBe(2);
    expect(metadata.images).toHaveLength(2);

    // Check image names and mime types
    const imageNames = metadata.images.map((img: any) => img.name).sort();
    expect(imageNames).toContain("image1.png");
    expect(imageNames).toContain("image2.jpeg");

    const pngImage = metadata.images.find((img: any) => img.name === "image1.png");
    expect(pngImage.mimeType).toBe("image/png");

    const jpegImage = metadata.images.find((img: any) => img.name === "image2.jpeg");
    expect(jpegImage.mimeType).toBe("image/jpeg");
  });

  it("should include position information for xlsx images", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: xlsxImagePath },
    });

    const metadata = JSON.parse((result.content as any)[0].text);

    // image1.png appears on both sheets
    const pngImage = metadata.images.find((img: any) => img.name === "image1.png");
    expect(pngImage.positions.length).toBeGreaterThanOrEqual(1);
    // At least one position should be on Sheet1
    const sheet1Pos = pngImage.positions.find((p: any) => p.sheet === "Sheet1");
    expect(sheet1Pos).toBeDefined();
    expect(sheet1Pos.fromRow).toBe(0);
    expect(sheet1Pos.fromCol).toBe(0);
    expect(sheet1Pos.toRow).toBe(3);
    expect(sheet1Pos.toCol).toBe(2);
  });

  it("should return image data as base64 content blocks", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: xlsxImagePath },
    });

    const content = result.content as any[];
    // Should have metadata + image content blocks
    const imageBlocks = content.filter((c: any) => c.type === "image");
    expect(imageBlocks.length).toBeGreaterThanOrEqual(2);
    for (const block of imageBlocks) {
      expect(block.data).toBeTruthy();
      expect(block.mimeType).toMatch(/^image\//);
    }
  });

  it("should filter images by sheetName", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: xlsxImagePath, sheetName: "Sheet2" },
    });

    const metadata = JSON.parse((result.content as any)[0].text);
    // Only images referenced on Sheet2 should be returned
    for (const img of metadata.images) {
      if (img.positions.length > 0) {
        expect(img.positions.every((p: any) => p.sheet === "Sheet2")).toBe(true);
      }
    }
  });
});

describe("get_excel_images - xls", () => {
  it("should extract images from xls file", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: xlsImagePath },
    });

    const content = result.content as any[];
    const metadata = JSON.parse(content[0].text);
    expect(metadata.imageCount).toBeGreaterThanOrEqual(1);

    const pngImage = metadata.images.find((img: any) => img.mimeType === "image/png");
    expect(pngImage).toBeDefined();
  });

  it("should include position information for xls images", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: xlsImagePath },
    });

    const metadata = JSON.parse((result.content as any)[0].text);
    const image = metadata.images[0];
    expect(image.positions).toBeDefined();
    expect(image.positions.length).toBeGreaterThanOrEqual(1);

    const pos = image.positions[0];
    expect(pos.sheet).toBe("ImageSheet");
    expect(typeof pos.fromRow).toBe("number");
    expect(typeof pos.fromCol).toBe("number");
    expect(typeof pos.toRow).toBe("number");
    expect(typeof pos.toCol).toBe("number");
  });
});

describe("get_excel_images - edge cases", () => {
  it("should return empty images for a file without images", async () => {
    const result = await client.callTool({
      name: "get_excel_images",
      arguments: { filePath: join(testDir, "basic.xlsx") },
    });

    const metadata = JSON.parse((result.content as any)[0].text);
    expect(metadata.imageCount).toBe(0);
    expect(metadata.images).toHaveLength(0);
  });

  it("should throw an error for a non-existent file", async () => {
    await expect(
      client.callTool({
        name: "get_excel_images",
        arguments: { filePath: "/non/existent/file.xlsx" },
      }),
    ).rejects.toThrow(/File not found/);
  });
});
