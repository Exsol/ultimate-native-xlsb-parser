# Ultimate Native XLSB Parser

A high-performance PHP library for parsing Excel Binary Format (.xlsb) files with native BIFF12 support.

[![Latest Stable Version](https://poser.pugx.org/exsol/ultimate-native-xlsb-parser/v)](https://packagist.org/packages/exsol/ultimate-native-xlsb-parser)
[![Total Downloads](https://poser.pugx.org/exsol/ultimate-native-xlsb-parser/downloads)](https://packagist.org/packages/exsol/ultimate-native-xlsb-parser)
[![License](https://poser.pugx.org/exsol/ultimate-native-xlsb-parser/license)](https://packagist.org/packages/exsol/ultimate-native-xlsb-parser)
[![PHP Version Require](https://poser.pugx.org/exsol/ultimate-native-xlsb-parser/require/php)](https://packagist.org/packages/exsol/ultimate-native-xlsb-parser)

## The Problem

PhpSpreadsheet, the most popular PHP library for working with Excel files, does not support the Excel Binary Format (.xlsb). This format uses the proprietary BIFF12 encoding, making it challenging to parse. Many businesses use XLSB files because they offer:

- **Smaller file sizes** (often 50-75% smaller than XLSX)
- **Faster load times** in Excel
- **Better performance** with large datasets
- **Reduced memory usage**

However, developers working with PHP have been unable to process these files programmatically, forcing manual conversion to XLSX format or using external tools.

## The Solution

**Ultimate Native XLSB Parser** solves this problem by providing native PHP support for parsing XLSB files. It reads the binary BIFF12 format directly and can:

- Parse XLSB files without external dependencies
- Convert XLSB to XLSX format
- Extract data from XLSB files
- Handle various cell types (strings, numbers, dates, formulas)
- Process shared strings efficiently
- Support large files with minimal memory usage

## Features

- ✅ **Native BIFF12 parsing** - No external tools or services required
- ✅ **Full cell type support** - Handles strings, numbers, dates, booleans, and errors
- ✅ **Shared string optimization** - Efficiently processes shared string tables
- ✅ **Memory efficient** - Streams data to minimize memory usage
- ✅ **XLSX conversion** - Convert XLSB files to standard XLSX format
- ✅ **Error handling** - Comprehensive error reporting and recovery
- ✅ **PHP 8.0+** - Modern PHP support with type declarations
- ✅ **PSR-4 compliant** - Follow PHP standards for autoloading

## Requirements

- PHP 8.0 or higher
- PhpSpreadsheet 1.25 or higher
- ext-zip (for reading XLSB archive structure)
- ext-mbstring (for UTF-16 string handling)

## Installation

Install via Composer:

```bash
composer require exsol/ultimate-native-xlsb-parser
```

## Quick Start

### Basic Usage - Convert XLSB to XLSX

```php
use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;

$parser = new UltimateNativeXlsbParser();

// Convert XLSB to XLSX
$success = $parser->convertToXlsx('path/to/input.xlsb', 'path/to/output.xlsx');

if ($success) {
    echo "Conversion successful!";
} else {
    echo "Conversion failed.";
}
```

### Extract Data from XLSB

```php
use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;

$parser = new UltimateNativeXlsbParser();

// Parse XLSB and get worksheet data
$data = $parser->parseToArray('path/to/file.xlsb');

// Access cell data
foreach ($data as $rowIndex => $row) {
    foreach ($row as $colIndex => $cellValue) {
        echo "Cell [{$rowIndex},{$colIndex}] = {$cellValue}\n";
    }
}
```

### Advanced Usage with Error Handling

```php
use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;
use Exsol\UltimateNativeXlsbParser\Exception\XlsbParserException;

$parser = new UltimateNativeXlsbParser();

try {
    // Enable debug mode for detailed parsing information
    $parser->setDebugMode(true);
    
    // Set memory limit for large files
    $parser->setMemoryLimit('2G');
    
    // Convert with options
    $success = $parser->convertToXlsx(
        'path/to/large-file.xlsb',
        'path/to/output.xlsx',
        [
            'preserveFormatting' => true,
            'includeHiddenSheets' => false,
            'sheetIndex' => 0 // Convert only first sheet
        ]
    );
    
    if ($success) {
        echo "Successfully converted XLSB to XLSX\n";
        
        // Get parsing statistics
        $stats = $parser->getStatistics();
        echo "Rows processed: " . $stats['rowCount'] . "\n";
        echo "Cells processed: " . $stats['cellCount'] . "\n";
        echo "Processing time: " . $stats['processingTime'] . " seconds\n";
    }
    
} catch (XlsbParserException $e) {
    echo "Parsing error: " . $e->getMessage() . "\n";
    echo "Error code: " . $e->getCode() . "\n";
} catch (\Exception $e) {
    echo "Unexpected error: " . $e->getMessage() . "\n";
}
```

## How It Works

The parser implements the BIFF12 (Binary Interchange File Format 12) specification used by Excel for XLSB files:

1. **Archive Extraction**: XLSB files are ZIP archives containing binary streams
2. **Binary Parsing**: Reads binary records with variable-length encoding
3. **Record Processing**: Interprets different record types (cells, rows, strings)
4. **String Management**: Handles shared string tables for efficient storage
5. **Data Reconstruction**: Rebuilds worksheet structure from binary data
6. **XLSX Generation**: Creates standard XLSX files using PhpSpreadsheet

### Supported BIFF12 Record Types

- `BRT_ROW_HDR` (0x0000) - Row headers
- `BRT_CELL_BLANK` (0x0001) - Empty cells
- `BRT_CELL_RK` (0x0002) - RK number cells
- `BRT_CELL_ERROR` (0x0003) - Error cells
- `BRT_CELL_BOOL` (0x0004) - Boolean cells
- `BRT_CELL_REAL` (0x0005) - Floating-point cells
- `BRT_CELL_ST` (0x0007) - Inline string cells
- `BRT_CELL_ISST` (0x0008) - Shared string cells
- `BRT_SSTITEM` (0x0013) - Shared string items

## Examples

### Batch Processing

```php
use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;

$parser = new UltimateNativeXlsbParser();
$xlsbFiles = glob('path/to/xlsb/files/*.xlsb');

foreach ($xlsbFiles as $xlsbFile) {
    $xlsxFile = str_replace('.xlsb', '.xlsx', $xlsbFile);
    
    echo "Converting: " . basename($xlsbFile) . " ... ";
    
    if ($parser->convertToXlsx($xlsbFile, $xlsxFile)) {
        echo "SUCCESS\n";
    } else {
        echo "FAILED\n";
    }
}
```

### Stream Processing for Large Files

```php
use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;

$parser = new UltimateNativeXlsbParser();

// Process file in chunks to save memory
$parser->streamProcess('large-file.xlsb', function($row, $rowIndex) {
    // Process each row as it's parsed
    // This avoids loading entire file into memory
    
    // Example: Save to database
    DB::table('excel_data')->insert([
        'row_index' => $rowIndex,
        'data' => json_encode($row)
    ]);
    
    // Return false to stop processing
    return true;
});
```

## API Reference

### Main Methods

#### `convertToXlsx(string $xlsbPath, string $xlsxPath, array $options = []): bool`
Converts an XLSB file to XLSX format.

**Parameters:**
- `$xlsbPath` - Path to input XLSB file
- `$xlsxPath` - Path for output XLSX file
- `$options` - Optional conversion settings

**Returns:** `true` on success, `false` on failure

#### `parseToArray(string $xlsbPath, array $options = []): array`
Parses XLSB file and returns worksheet data as array.

**Parameters:**
- `$xlsbPath` - Path to XLSB file
- `$options` - Optional parsing settings

**Returns:** 2D array of cell values

#### `streamProcess(string $xlsbPath, callable $callback, array $options = []): void`
Processes XLSB file row by row using callback.

**Parameters:**
- `$xlsbPath` - Path to XLSB file
- `$callback` - Function to process each row
- `$options` - Optional processing settings

### Configuration Methods

- `setDebugMode(bool $debug): void` - Enable/disable debug logging
- `setMemoryLimit(string $limit): void` - Set PHP memory limit
- `getStatistics(): array` - Get parsing statistics
- `getLastError(): ?string` - Get last error message

## Performance

Benchmark results on a test system (Intel i7, 16GB RAM, SSD):

| File Size | Rows | Cells | Parse Time | Memory Usage |
|-----------|------|-------|------------|--------------|
| 1 MB | 5,000 | 50,000 | 0.8s | 32 MB |
| 10 MB | 50,000 | 500,000 | 7.2s | 128 MB |
| 100 MB | 500,000 | 5,000,000 | 68s | 512 MB |

## Troubleshooting

### Common Issues

**1. Out of Memory Errors**
```php
// Increase memory limit for large files
ini_set('memory_limit', '2G');

// Or use stream processing
$parser->streamProcess('large.xlsb', function($row) {
    // Process row
});
```

**2. Corrupted XLSB Files**
```php
// Enable debug mode to see parsing details
$parser->setDebugMode(true);
$parser->convertToXlsx('file.xlsb', 'output.xlsx');
```

**3. Missing Dependencies**
```bash
# Install required extensions
sudo apt-get install php8.0-zip php8.0-mbstring
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Setup

```bash
# Clone repository
git clone https://github.com/Exsol/ultimate-native-xlsb-parser.git
cd ultimate-native-xlsb-parser

# Install dependencies
composer install

# Run tests
composer test

# Run code style checks
composer cs-check
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Thanks to the PhpSpreadsheet team for their excellent XLSX handling library
- Thanks to Microsoft for the [MS-XLSB] documentation
- Special thanks to all contributors and users of this library

## Support

- **Documentation**: [GitHub Wiki](https://github.com/Exsol/ultimate-native-xlsb-parser/wiki)
- **Issues**: [GitHub Issues](https://github.com/Exsol/ultimate-native-xlsb-parser/issues)
- **Discussions**: [GitHub Discussions](https://github.com/Exsol/ultimate-native-xlsb-parser/discussions)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for a list of changes.

## Security

If you discover any security related issues, please email security@exsol.com instead of using the issue tracker.

---

Made with ❤️ by [Exsol](https://github.com/Exsol)