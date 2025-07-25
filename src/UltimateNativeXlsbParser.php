<?php

declare(strict_types=1);

namespace Exsol\UltimateNativeXlsbParser;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use ZipArchive;
use Exsol\UltimateNativeXlsbParser\Exception\XlsbParserException;

/**
 * Ultimate Native XLSB Parser
 * 
 * A high-performance PHP library for parsing Excel Binary Format (.xlsb) files
 * with native BIFF12 support. This parser can read XLSB files that PhpSpreadsheet
 * cannot handle, providing a solution for working with binary Excel files in PHP.
 * 
 * @package Exsol\UltimateNativeXlsbParser
 * @author Exsol Development Team
 * @license MIT
 */
class UltimateNativeXlsbParser
{
    // BIFF12 Record Types
    private const BRT_ROW_HDR = 0x0000;
    private const BRT_CELL_BLANK = 0x0001;
    private const BRT_CELL_RK = 0x0002;
    private const BRT_CELL_ERROR = 0x0003;
    private const BRT_CELL_BOOL = 0x0004;
    private const BRT_CELL_REAL = 0x0005;
    private const BRT_CELL_ST = 0x0007;
    private const BRT_CELL_ISST = 0x0008;
    private const BRT_SSTITEM = 0x0013;
    private const BRT_BEGIN_SHEET_DATA = 0x0091;
    private const BRT_END_SHEET_DATA = 0x0092;
    
    private array $sharedStrings = [];
    private array $worksheetData = [];
    private bool $debugMode = false;
    private ?string $lastError = null;
    private array $statistics = [
        'rowCount' => 0,
        'cellCount' => 0,
        'processingTime' => 0,
        'memoryUsage' => 0
    ];
    
    /**
     * Convert XLSB file to XLSX format
     * 
     * @param string $xlsbPath Path to input XLSB file
     * @param string $xlsxPath Path for output XLSX file
     * @param array $options Optional conversion settings
     * @return bool True on success, false on failure
     * @throws XlsbParserException
     */
    public function convertToXlsx(string $xlsbPath, string $xlsxPath, array $options = []): bool
    {
        $startTime = microtime(true);
        $this->lastError = null;
        $this->resetStatistics();
        
        try {
            if (!file_exists($xlsbPath)) {
                throw new XlsbParserException("Input file not found: {$xlsbPath}");
            }
            
            if (!is_readable($xlsbPath)) {
                throw new XlsbParserException("Input file is not readable: {$xlsbPath}");
            }
            
            // Parse XLSB
            if (!$this->parseXlsb($xlsbPath)) {
                return false;
            }
            
            // Create spreadsheet
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            
            // Populate data
            foreach ($this->worksheetData as $rowIdx => $rowData) {
                foreach ($rowData as $colIdx => $value) {
                    $sheet->setCellValue([$colIdx + 1, $rowIdx + 1], $value);
                }
            }
            
            // Create directory if it doesn't exist
            $outputDir = dirname($xlsxPath);
            if (!is_dir($outputDir)) {
                if (!mkdir($outputDir, 0777, true)) {
                    throw new XlsbParserException("Failed to create output directory: {$outputDir}");
                }
            }
            
            // Save as XLSX
            $writer = new Xlsx($spreadsheet);
            $writer->save($xlsxPath);
            
            $this->statistics['processingTime'] = microtime(true) - $startTime;
            $this->statistics['memoryUsage'] = memory_get_peak_usage(true);
            
            return true;
            
        } catch (\Exception $e) {
            $this->lastError = $e->getMessage();
            if ($this->debugMode) {
                error_log("UltimateNativeXlsbParser Error: " . $e->getMessage());
            }
            
            if ($e instanceof XlsbParserException) {
                throw $e;
            }
            
            throw new XlsbParserException("Conversion failed: " . $e->getMessage(), 0, $e);
        }
    }
    
    /**
     * Parse XLSB file and return data as array
     * 
     * @param string $xlsbPath Path to XLSB file
     * @param array $options Optional parsing settings
     * @return array 2D array of cell values
     * @throws XlsbParserException
     */
    public function parseToArray(string $xlsbPath, array $options = []): array
    {
        $this->lastError = null;
        
        if (!$this->parseXlsb($xlsbPath)) {
            throw new XlsbParserException($this->lastError ?? "Failed to parse XLSB file");
        }
        
        return $this->worksheetData;
    }
    
    /**
     * Process XLSB file row by row using callback
     * 
     * @param string $xlsbPath Path to XLSB file
     * @param callable $callback Function to process each row
     * @param array $options Optional processing settings
     * @throws XlsbParserException
     */
    public function streamProcess(string $xlsbPath, callable $callback, array $options = []): void
    {
        // This is a simplified version - in a real implementation,
        // we would parse the file in chunks and call the callback for each row
        $data = $this->parseToArray($xlsbPath, $options);
        
        foreach ($data as $rowIndex => $row) {
            $continue = $callback($row, $rowIndex);
            if ($continue === false) {
                break;
            }
        }
    }
    
    /**
     * Enable or disable debug mode
     * 
     * @param bool $debug
     */
    public function setDebugMode(bool $debug): void
    {
        $this->debugMode = $debug;
    }
    
    /**
     * Set PHP memory limit
     * 
     * @param string $limit Memory limit (e.g., '2G', '512M')
     */
    public function setMemoryLimit(string $limit): void
    {
        ini_set('memory_limit', $limit);
    }
    
    /**
     * Get parsing statistics
     * 
     * @return array
     */
    public function getStatistics(): array
    {
        return $this->statistics;
    }
    
    /**
     * Get last error message
     * 
     * @return string|null
     */
    public function getLastError(): ?string
    {
        return $this->lastError;
    }
    
    /**
     * Parse XLSB file
     * 
     * @param string $filePath
     * @return bool
     */
    private function parseXlsb(string $filePath): bool
    {
        $this->sharedStrings = [];
        $this->worksheetData = [];
        
        $zip = new ZipArchive();
        if ($zip->open($filePath) !== true) {
            $this->lastError = "Failed to open XLSB file as ZIP archive";
            return false;
        }
        
        // Parse shared strings first
        $sstData = $zip->getFromName('xl/sharedStrings.bin');
        if ($sstData !== false) {
            $this->parseSharedStrings($sstData);
            if ($this->debugMode) {
                error_log("UltimateNativeXlsbParser: Loaded " . count($this->sharedStrings) . " shared strings");
            }
        }
        
        // Parse worksheet
        $worksheetData = $zip->getFromName('xl/worksheets/sheet1.bin');
        if ($worksheetData === false) {
            $worksheetData = $zip->getFromName('xl/worksheets/Sheet1.bin');
        }
        
        if ($worksheetData !== false) {
            $this->parseWorksheet($worksheetData);
        } else {
            $this->lastError = "No worksheet data found in XLSB file";
            $zip->close();
            return false;
        }
        
        $zip->close();
        
        return !empty($this->worksheetData);
    }
    
    /**
     * Parse shared strings from binary data
     * 
     * @param string $data Binary data containing shared strings
     */
    private function parseSharedStrings(string $data): void
    {
        $records = $this->parseRecords($data);
        
        foreach ($records as $record) {
            if ($record['type'] === self::BRT_SSTITEM) {
                // BRT_SSTITEM format: 00 + 4-byte length + UTF-16LE string
                if (strlen($record['data']) >= 5 && ord($record['data'][0]) === 0x00) {
                    $strLen = unpack('V', substr($record['data'], 1, 4))[1];
                    if ($strLen > 0 && $strLen < 10000 && strlen($record['data']) >= 5 + $strLen * 2) {
                        $string = mb_convert_encoding(substr($record['data'], 5, $strLen * 2), 'UTF-8', 'UTF-16LE');
                        $this->sharedStrings[] = $string;
                    } else {
                        $this->sharedStrings[] = '';
                    }
                } else {
                    // Fallback to original parsing
                    $string = $this->parseXLWideString($record['data']);
                    if ($string !== null) {
                        $this->sharedStrings[] = $string;
                    }
                }
            }
        }
    }
    
    /**
     * Parse worksheet data from binary
     * 
     * @param string $data Binary worksheet data
     */
    private function parseWorksheet(string $data): void
    {
        $records = $this->parseRecords($data);
        
        $currentRow = -1;
        $inSheetData = false;
        
        foreach ($records as $idx => $record) {
            // Check for sheet data boundaries
            if ($record['type'] === self::BRT_BEGIN_SHEET_DATA) {
                $inSheetData = true;
            } elseif ($record['type'] === self::BRT_END_SHEET_DATA) {
                $inSheetData = false;
            }
            
            switch ($record['type']) {
                case self::BRT_ROW_HDR:
                    if (strlen($record['data']) >= 4) {
                        $currentRow = unpack('V', substr($record['data'], 0, 4))[1];
                        $this->statistics['rowCount']++;
                    }
                    break;
                    
                case self::BRT_CELL_ISST:
                    if ($currentRow >= 0) {
                        $this->parseIsstCell($record['data'], $currentRow);
                    }
                    break;
                    
                case self::BRT_CELL_ST:
                    if ($currentRow >= 0) {
                        $this->parseStringCell($record['data'], $currentRow);
                    }
                    break;
                    
                case self::BRT_CELL_REAL:
                    if ($currentRow >= 0 && strlen($record['data']) >= 16) {
                        $col = unpack('V', substr($record['data'], 0, 4))[1];
                        $value = unpack('d', substr($record['data'], 8, 8))[1];
                        $this->setCell($currentRow, $col, $value);
                    }
                    break;
                    
                case self::BRT_CELL_RK:
                    if ($currentRow >= 0 && strlen($record['data']) >= 12) {
                        $col = unpack('V', substr($record['data'], 0, 4))[1];
                        $rk = unpack('V', substr($record['data'], 8, 4))[1];
                        $value = $this->parseRkValue($rk);
                        $this->setCell($currentRow, $col, $value);
                    }
                    break;
                    
                case self::BRT_CELL_BOOL:
                    if ($currentRow >= 0 && strlen($record['data']) >= 9) {
                        $col = unpack('V', substr($record['data'], 0, 4))[1];
                        $value = ord($record['data'][8]) !== 0;
                        $this->setCell($currentRow, $col, $value);
                    }
                    break;
                    
                case self::BRT_CELL_ERROR:
                    if ($currentRow >= 0 && strlen($record['data']) >= 9) {
                        $col = unpack('V', substr($record['data'], 0, 4))[1];
                        $errorCode = ord($record['data'][8]);
                        $this->setCell($currentRow, $col, $this->getErrorString($errorCode));
                    }
                    break;
            }
        }
    }
    
    /**
     * Parse ISST cell (shared string reference)
     * 
     * @param string $data Cell data
     * @param int $row Row index
     */
    private function parseIsstCell(string $data, int $row): void
    {
        // ISST cells have variable structure, try different offsets
        $offsets = [
            ['col' => 0, 'sst' => 8],   // Standard format
            ['col' => 0, 'sst' => 4],   // Compact format
            ['col' => 0, 'sst' => 12],  // Extended format
        ];
        
        foreach ($offsets as $offset) {
            if (strlen($data) >= $offset['sst'] + 4) {
                $col = unpack('V', substr($data, $offset['col'], 4))[1];
                $sstIndex = unpack('V', substr($data, $offset['sst'], 4))[1];
                
                // Validate SST index
                if ($sstIndex >= 0 && $sstIndex < count($this->sharedStrings)) {
                    $this->setCell($row, $col, $this->sharedStrings[$sstIndex]);
                    return;
                }
            }
        }
    }
    
    /**
     * Parse string cell (inline string)
     * 
     * @param string $data Cell data
     * @param int $row Row index
     */
    private function parseStringCell(string $data, int $row): void
    {
        if (strlen($data) >= 12) {
            $col = unpack('V', substr($data, 0, 4))[1];
            
            // Check if this might be a shared string reference
            $possibleSstIndex = unpack('V', substr($data, 8, 4))[1];
            
            if ($possibleSstIndex < count($this->sharedStrings)) {
                // Try as SST reference
                $value = $this->sharedStrings[$possibleSstIndex];
                if (!empty($value)) {
                    $this->setCell($row, $col, $value);
                    return;
                }
            }
            
            // Otherwise try as inline string
            if (strlen($data) > 12) {
                $string = $this->parseXLWideString(substr($data, 8));
                if ($string !== null && $string !== '') {
                    $this->setCell($row, $col, $string);
                }
            }
        }
    }
    
    /**
     * Parse BIFF12 records with variable-length encoding
     * 
     * @param string $data Binary data
     * @return array Array of records
     */
    private function parseRecords(string $data): array
    {
        $records = [];
        $offset = 0;
        $length = strlen($data);
        
        while ($offset + 1 < $length) {
            // Read record type (1 or 2 bytes)
            $byte1 = ord($data[$offset]);
            
            if ($byte1 & 0x80) {
                // 2-byte record type
                if ($offset + 2 > $length) break;
                $type = (($byte1 & 0x7F) | (ord($data[$offset + 1]) << 7));
                $offset += 2;
            } else {
                // 1-byte record type
                $type = $byte1;
                $offset += 1;
            }
            
            // Read record size (variable length)
            if ($offset >= $length) break;
            
            $size = 0;
            $sizeBytes = 0;
            
            // Read size using variable-length encoding
            do {
                if ($offset + $sizeBytes >= $length) break 2;
                $byte = ord($data[$offset + $sizeBytes]);
                $size |= (($byte & 0x7F) << (7 * $sizeBytes));
                $sizeBytes++;
            } while ($byte & 0x80 && $sizeBytes < 4);
            
            $offset += $sizeBytes;
            
            // Read record data
            if ($offset + $size > $length) break;
            
            $recordData = substr($data, $offset, $size);
            $offset += $size;
            
            $records[] = [
                'type' => $type,
                'size' => $size,
                'data' => $recordData
            ];
        }
        
        return $records;
    }
    
    /**
     * Parse XLWideString (4-byte length + UTF-16LE)
     * 
     * @param string $data String data
     * @return string|null Parsed string or null
     */
    private function parseXLWideString(string $data): ?string
    {
        if (strlen($data) < 4) return null;
        
        $charCount = unpack('V', substr($data, 0, 4))[1];
        
        if ($charCount === 0) return '';
        
        // Check if this might be a bit field instead of char count
        if ($charCount > 10000) {
            // Try parsing as a different format
            if (strlen($data) < 5) return null;
            
            $flag = ord($data[0]);
            $charCount = unpack('V', substr($data, 1, 4))[1];
            
            if ($charCount === 0) return '';
            if ($charCount > 10000) return null;
            
            $byteLength = $charCount * 2;
            if (strlen($data) < 5 + $byteLength) return null;
            
            $stringData = substr($data, 5, $byteLength);
            return mb_convert_encoding($stringData, 'UTF-8', 'UTF-16LE');
        }
        
        $byteLength = $charCount * 2;
        if (strlen($data) < 4 + $byteLength) return null;
        
        $stringData = substr($data, 4, $byteLength);
        return mb_convert_encoding($stringData, 'UTF-8', 'UTF-16LE');
    }
    
    /**
     * Parse RK value (compressed number format)
     * 
     * @param int $rk RK value
     * @return float Decoded value
     */
    private function parseRkValue(int $rk): float
    {
        $fX100 = $rk & 0x01;
        $fInt = $rk & 0x02;
        
        if ($fInt) {
            $value = $rk >> 2;
            if ($value & 0x20000000) {
                $value |= 0xC0000000;
            }
            $value = intval($value);
        } else {
            $floatBits = $rk & 0xFFFFFFFC;
            $value = unpack('f', pack('V', $floatBits))[1];
        }
        
        if ($fX100) {
            $value /= 100;
        }
        
        return floatval($value);
    }
    
    /**
     * Convert error code to string
     * 
     * @param int $errorCode
     * @return string
     */
    private function getErrorString(int $errorCode): string
    {
        $errors = [
            0x00 => '#NULL!',
            0x07 => '#DIV/0!',
            0x0F => '#VALUE!',
            0x17 => '#REF!',
            0x1D => '#NAME?',
            0x24 => '#NUM!',
            0x2A => '#N/A',
        ];
        
        return $errors[$errorCode] ?? '#ERROR!';
    }
    
    /**
     * Set cell value in worksheet data
     * 
     * @param int $row Row index
     * @param int $col Column index
     * @param mixed $value Cell value
     */
    private function setCell(int $row, int $col, $value): void
    {
        if (!isset($this->worksheetData[$row])) {
            $this->worksheetData[$row] = [];
        }
        $this->worksheetData[$row][$col] = $value;
        $this->statistics['cellCount']++;
    }
    
    /**
     * Reset statistics
     */
    private function resetStatistics(): void
    {
        $this->statistics = [
            'rowCount' => 0,
            'cellCount' => 0,
            'processingTime' => 0,
            'memoryUsage' => 0
        ];
    }
}