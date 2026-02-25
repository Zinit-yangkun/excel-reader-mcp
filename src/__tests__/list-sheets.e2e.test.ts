import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { cleanupTestFiles, setupTestFiles } from "./helpers/generate-test-files.js";

let client: Client;
let transport: StdioClientTransport;
let testDir: string;

beforeAll(async () => {
  testDir = setupTestFiles();

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

describe("list_sheets", () => {
  it("should list sheets for a single-sheet file", async () => {
    const result = await client.callTool({
      name: "list_sheets",
      arguments: { filePath: join(testDir, "basic.xlsx") },
    });

    const content = result.content as { type: string; text: string }[];
    const data = JSON.parse(content[0].text);
    expect(data.fileName).toBe("basic.xlsx");
    expect(data.sheets).toEqual(["Sheet1"]);
  });

  it("should list all sheets for a multi-sheet file", async () => {
    const result = await client.callTool({
      name: "list_sheets",
      arguments: { filePath: join(testDir, "multi-sheet.xlsx") },
    });

    const data = JSON.parse((result.content as { type: string; text: string }[])[0].text);
    expect(data.sheets).toEqual(["Products", "Cities", "Colors"]);
  });

  it("should throw an error for a non-existent file", async () => {
    await expect(
      client.callTool({
        name: "list_sheets",
        arguments: { filePath: "/non/existent/file.xlsx" },
      }),
    ).rejects.toThrow(/File not found/);
  });
});
