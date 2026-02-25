import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { extractImages } from "../image-extractor.js";
import { createXlsWithImages, createXlsxWithImages } from "./helpers/generate-image-files.js";
import { cleanupTestFiles, setupTestFiles } from "./helpers/generate-test-files.js";

let testDir: string;
let xlsxImagePath: string;
let xlsImagePath: string;

beforeAll(async () => {
  testDir = setupTestFiles();
  xlsxImagePath = await createXlsxWithImages(testDir);
  xlsImagePath = createXlsWithImages(testDir);
});

afterAll(() => {
  cleanupTestFiles(testDir);
});

describe("extractImages - xlsx", () => {
  it("should extract images with correct count and metadata", async () => {
    const { images, truncated } = await extractImages({ filePath: xlsxImagePath });

    expect(truncated).toBe(false);
    expect(images).toHaveLength(2);

    const names = images.map((img) => img.name).sort();
    expect(names).toContain("image1.png");
    expect(names).toContain("image2.jpeg");

    const pngImage = images.find((img) => img.name === "image1.png");
    expect(pngImage?.mimeType).toBe("image/png");

    const jpegImage = images.find((img) => img.name === "image2.jpeg");
    expect(jpegImage?.mimeType).toBe("image/jpeg");
  });

  it("should include position information", async () => {
    const { images } = await extractImages({ filePath: xlsxImagePath });

    const pngImage = images.find((img) => img.name === "image1.png")!;
    expect(pngImage.positions.length).toBeGreaterThanOrEqual(1);

    const sheet1Pos = pngImage.positions.find((p) => p.sheet === "Sheet1");
    expect(sheet1Pos).toBeDefined();
    expect(sheet1Pos?.fromRow).toBe(0);
    expect(sheet1Pos?.fromCol).toBe(0);
    expect(sheet1Pos?.toRow).toBe(3);
    expect(sheet1Pos?.toCol).toBe(2);
  });

  it("should return base64 image data", async () => {
    const { images } = await extractImages({ filePath: xlsxImagePath });

    for (const img of images) {
      expect(img.data).toBeTruthy();
      expect(typeof img.data).toBe("string");
    }
  });

  it("should filter images by sheetName", async () => {
    const { images } = await extractImages({ filePath: xlsxImagePath, sheetName: "Sheet2" });

    for (const img of images) {
      if (img.positions.length > 0) {
        expect(img.positions.every((p) => p.sheet === "Sheet2")).toBe(true);
      }
    }
  });
});

describe("extractImages - xls", () => {
  it("should extract images from xls file", async () => {
    const { images } = await extractImages({ filePath: xlsImagePath });

    expect(images.length).toBeGreaterThanOrEqual(1);
    const pngImage = images.find((img) => img.mimeType === "image/png");
    expect(pngImage).toBeDefined();
  });

  it("should include position information for xls images", async () => {
    const { images } = await extractImages({ filePath: xlsImagePath });

    const image = images[0];
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

describe("extractImages - edge cases", () => {
  it("should return empty images for a file without images", async () => {
    const { images } = await extractImages({ filePath: join(testDir, "basic.xlsx") });

    expect(images).toHaveLength(0);
  });

  it("should throw for a non-existent file", async () => {
    await expect(extractImages({ filePath: "/non/existent/file.xlsx" })).rejects.toThrow(/File not found/);
  });
});
