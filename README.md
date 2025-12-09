# Authorization Service

[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Simple authorization service with XLSX hyperlink conversion functionality.

## Features

- User authorization and authentication
- XLSX hyperlink converter for MS Office compatibility
- Google Drive API integration
- File processing and validation

## XLSX Hyperlink Converter

Converts Excel files with HYPERLINK formulas into files with clickable hyperlinks, compatible with both MS Office and LibreOffice.

### Usage

```javascript
const xlsxService = require('./src/server/services/xlsxService.js');

// Main method (automatically selects best conversion method)
const result = await xlsxService.doConvertXlsx('path/to/input.xlsx');

// Direct methods
const result1 = xlsxService.convertWithSheetJS('input.xlsx', 'output.xlsx'); // Recommended
const result2 = xlsxService.convertToHyperlinks('input.xlsx', 'output.xlsx'); // Fallback
```

### Dependencies

- `xlsx` (SheetJS) - main library for maximum MS Office compatibility
- `adm-zip` - for working with XLSX archives

### XLSX Converter Features

- Dynamic hyperlink extraction from any files
- MS Office and LibreOffice support  
- Preserves all original file data
- Automatic selection of most compatible conversion method

## Requirements

- Node.js >= 18.x
- npm or yarn package manager

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/SvetlanaKrugovykh/AuthorizationService.git
   cd AuthorizationService
   ```

2. Install dependencies:

   ```bash
   npm install
   ```
