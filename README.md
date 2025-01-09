# HTML to Excel Converter

Convert HTML tables to Excel files with support for styling and multiple documents.

## Prerequisites

- Node.js >= 14.0.0
- Python 3.x

## Installation

```bash
npm install html-to-excel-converter
```

The package will automatically set up a Python virtual environment and install required dependencies during the installation process.

## Usage

```javascript
const HTMLToExcelConverter = require('html-to-excel-converter');

// Create a new instance
const converter = new HTMLToExcelConverter();

// HTML content with tables
const html = `
<table>
    <tr>
        <th>Name</th>
        <th>Age</th>
    </tr>
    <tr>
        <td>John</td>
        <td>30</td>
    </tr>
</table>
`;

// Convert HTML to Excel
try {
    const excelBuffer = await converter.convert(html);
    // excelBuffer is a Buffer containing the Excel file
    // You can write it to a file or send it as a response
    fs.writeFileSync('output.xlsx', excelBuffer);
} catch (error) {
    console.error('Conversion failed:', error);
}
```

## Features

- Converts HTML tables to Excel (.xlsx) format
- Preserves table styling (colors, borders, etc.)
- Supports multiple tables in a single document
- Handles complex HTML structures
- Memory efficient for large tables

## Options

You can customize the converter behavior by passing options:

```javascript
const converter = new HTMLToExcelConverter({
    maxChunkSize: 5 * 1024 * 1024, // 5MB
    timeout: 10 * 60 * 1000,       // 10 minutes
    maxBuffer: 50 * 1024 * 1024    // 50MB
});
```

## API

### `constructor(options)`

Creates a new converter instance.

Options:
- `maxChunkSize`: Maximum size of data chunks (default: 5MB)
- `timeout`: Timeout for conversion process (default: 10 minutes)
- `maxBuffer`: Maximum buffer size for Python process (default: 50MB)

### `convert(html)`

Converts HTML content to Excel format.

Parameters:
- `html`: String containing HTML content with tables

Returns:
- Promise that resolves with a Buffer containing the Excel file

## License

MIT License - see LICENSE file for details
