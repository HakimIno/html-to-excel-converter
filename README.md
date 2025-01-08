# HTML to Excel Converter

A Node.js library for converting HTML tables to Excel files with support for multiple documents.

## Prerequisites

- Node.js >= 12.0.0
- Python >= 3.7.0

## Installation

```bash
npm install html-to-excel-converter
```

Python dependencies will be installed automatically during npm installation.

## Usage

```javascript
const HTMLToExcelConverter = require('html-to-excel-converter');

async function convertHtmlToExcel() {
    try {
        const converter = new HTMLToExcelConverter();
        
        // Your HTML content
        const htmlContent = `
            <!DOCTYPE html>
            <html>
                <!-- Your HTML content here -->
            </html>
        `;
        
        // Convert HTML to Excel
        const excelBuffer = await converter.convertHtmlToExcel(htmlContent);
        
        // Save to file
        require('fs').writeFileSync('output.xlsx', excelBuffer);
        
        console.log('Conversion completed successfully!');
    } catch (error) {
        console.error('Conversion failed:', error);
    }
}

convertHtmlToExcel();
```

## Features

- Converts HTML tables to Excel format
- Supports multiple HTML documents in sequence
- Preserves table formatting and styles
- Handles headers, footers, and customer information
- Maintains cell alignment and borders

## API

### `HTMLToExcelConverter`

#### `constructor()`

Creates a new instance of the converter.

#### `convertHtmlToExcel(htmlContent, options = {})`

Converts HTML content to Excel format.

- `htmlContent` (string): The HTML content to convert
- `options` (object):
  - `timeout` (number): Timeout in milliseconds (default: 30000)

Returns: Promise<Buffer> - The Excel file as a buffer

## License

MIT
