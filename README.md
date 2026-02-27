# Excel Reader MCP

A [Model Context Protocol](https://modelcontextprotocol.io/) server for reading Excel files. Supports `.xlsx` and `.xls` formats with automatic chunking, pagination, and image extraction.

## Features

- Read Excel files with automatic chunking for large datasets
- List all sheets in a workbook
- Extract embedded images with position information
- Proper date handling
- Row pagination support

## Usage

Run directly via npx:

```bash
npx excel-reader-mcp
```

## MCP Configuration

Add to your MCP settings:

```json
{
  "mcpServers": {
    "excel-reader-mcp": {
      "command": "npx",
      "args": [
        "-y",
        "excel-reader-mcp"
      ],
      "env": {}
    }
  }
}
```

## Tools

### `read_excel`

Read an Excel file and return its contents as structured data.

| Parameter   | Required | Description                          |
| ----------- | -------- | ------------------------------------ |
| `filePath`  | Yes      | Path to the Excel file               |
| `sheetName` | No       | Sheet name (defaults to first sheet) |
| `startRow`  | No       | Starting row index for pagination    |
| `maxRows`   | No       | Maximum number of rows to read       |

Large files are automatically split into chunks (~100KB). The response includes `hasMore` and `nextChunk` fields for pagination.

### `list_sheets`

List all sheet names in an Excel file.

| Parameter  | Required | Description            |
| ---------- | -------- | ---------------------- |
| `filePath` | Yes      | Path to the Excel file |

### `get_excel_images`

Extract embedded images from an Excel file, including position information (sheet, row, column). Returns base64-encoded image data.

| Parameter   | Required | Description                                 |
| ----------- | -------- | ------------------------------------------- |
| `filePath`  | Yes      | Path to the Excel file (.xlsx or .xls)      |
| `sheetName` | No       | Only return images from the specified sheet |

## Installation

```bash
npm install -g excel-reader-mcp
```

## Development

```bash
npm install
npm run build
npm test
```

## License

MIT

## Acknowledgments

- Originally forked from [ArchimedesCrypto/excel-reader-mcp](https://github.com/ArchimdesCrypto/excel-reader-mcp) — thanks for the foundation!
- Built with [SheetJS](https://sheetjs.com/)
