# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-01-25

### Added
- Initial release of Ultimate Native XLSB Parser
- Native BIFF12 format parsing without external dependencies
- Support for all major cell types (strings, numbers, dates, booleans, errors)
- Shared string table optimization for efficient memory usage
- XLSB to XLSX conversion functionality
- Data extraction to PHP arrays
- Stream processing for large files
- Comprehensive error handling with custom exceptions
- Debug mode for troubleshooting
- Memory limit configuration
- Processing statistics
- Full PHP 8.0+ support with type declarations
- PSR-4 autoloading compliance
- Extensive documentation and examples

### Technical Details
- Implements BIFF12 record types: BRT_ROW_HDR, BRT_CELL_BLANK, BRT_CELL_RK, BRT_CELL_ERROR, BRT_CELL_BOOL, BRT_CELL_REAL, BRT_CELL_ST, BRT_CELL_ISST, BRT_SSTITEM
- Variable-length encoding support for BIFF12 records
- UTF-16LE to UTF-8 string conversion
- RK value decompression for numeric data
- Automatic worksheet detection (sheet1.bin/Sheet1.bin)

[1.0.0]: https://github.com/Exsol/ultimate-native-xlsb-parser/releases/tag/v1.0.0