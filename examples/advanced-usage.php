<?php

/**
 * Advanced XLSB Parser Usage Example
 * 
 * This example demonstrates advanced features including:
 * - Debug mode for troubleshooting
 * - Memory limit configuration
 * - Data extraction without conversion
 * - Stream processing for large files
 * - Error handling
 */

require_once __DIR__ . '/../vendor/autoload.php';

use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;
use Exsol\UltimateNativeXlsbParser\Exception\XlsbParserException;

// Example 1: Debug Mode and Memory Configuration
echo "=== Example 1: Debug Mode and Memory Configuration ===\n";

$parser = new UltimateNativeXlsbParser();

// Enable debug mode to see detailed parsing information
$parser->setDebugMode(true);

// Set memory limit for large files
$parser->setMemoryLimit('2G');

$xlsbFile = 'path/to/large-file.xlsb';
$xlsxFile = 'path/to/output.xlsx';

try {
    echo "Converting large XLSB file with debug mode enabled...\n";
    
    $success = $parser->convertToXlsx($xlsbFile, $xlsxFile);
    
    if ($success) {
        echo "✓ Conversion successful!\n";
        
        $stats = $parser->getStatistics();
        echo "\nDetailed Statistics:\n";
        echo "- Total rows: " . number_format($stats['rowCount']) . "\n";
        echo "- Total cells: " . number_format($stats['cellCount']) . "\n";
        echo "- Cells per row (avg): " . round($stats['cellCount'] / $stats['rowCount'], 2) . "\n";
        echo "- Processing time: " . round($stats['processingTime'], 2) . " seconds\n";
        echo "- Processing speed: " . number_format($stats['cellCount'] / $stats['processingTime']) . " cells/second\n";
        echo "- Peak memory: " . round($stats['memoryUsage'] / 1024 / 1024, 2) . " MB\n";
    }
    
} catch (XlsbParserException $e) {
    echo "Parser error: " . $e->getMessage() . "\n";
} catch (Exception $e) {
    echo "Unexpected error: " . $e->getMessage() . "\n";
}

// Example 2: Extract Data Without Creating XLSX
echo "\n\n=== Example 2: Extract Data to Array ===\n";

$parser = new UltimateNativeXlsbParser();
$xlsbFile = 'path/to/data.xlsb';

try {
    echo "Parsing XLSB file to array...\n";
    
    $data = $parser->parseToArray($xlsbFile);
    
    echo "✓ Successfully parsed " . count($data) . " rows\n";
    
    // Display first 5 rows
    echo "\nFirst 5 rows of data:\n";
    $rowCount = 0;
    foreach ($data as $rowIndex => $row) {
        if ($rowCount >= 5) break;
        
        echo "Row {$rowIndex}: ";
        $cellValues = [];
        foreach ($row as $colIndex => $value) {
            $cellValues[] = "[{$colIndex}]={$value}";
        }
        echo implode(', ', $cellValues) . "\n";
        $rowCount++;
    }
    
    // Example: Find specific data
    echo "\nSearching for specific values...\n";
    $searchTerm = 'example';
    $found = 0;
    
    foreach ($data as $rowIndex => $row) {
        foreach ($row as $colIndex => $value) {
            if (stripos((string)$value, $searchTerm) !== false) {
                echo "Found '{$searchTerm}' at cell [{$rowIndex},{$colIndex}]: {$value}\n";
                $found++;
                if ($found >= 3) break 2; // Show only first 3 matches
            }
        }
    }
    
    if ($found === 0) {
        echo "No matches found for '{$searchTerm}'\n";
    }
    
} catch (XlsbParserException $e) {
    echo "Failed to parse XLSB: " . $e->getMessage() . "\n";
}

// Example 3: Stream Processing for Large Files
echo "\n\n=== Example 3: Stream Processing ===\n";

$parser = new UltimateNativeXlsbParser();
$xlsbFile = 'path/to/very-large-file.xlsb';

try {
    echo "Processing XLSB file in streaming mode...\n";
    
    $rowCount = 0;
    $totalValue = 0;
    $columnSums = [];
    
    // Process each row without loading entire file into memory
    $parser->streamProcess($xlsbFile, function($row, $rowIndex) use (&$rowCount, &$totalValue, &$columnSums) {
        $rowCount++;
        
        // Example: Calculate sum of numeric values
        foreach ($row as $colIndex => $value) {
            if (is_numeric($value)) {
                $totalValue += $value;
                
                if (!isset($columnSums[$colIndex])) {
                    $columnSums[$colIndex] = 0;
                }
                $columnSums[$colIndex] += $value;
            }
        }
        
        // Show progress every 1000 rows
        if ($rowCount % 1000 === 0) {
            echo "Processed {$rowCount} rows...\n";
        }
        
        // Example: Stop after 10000 rows
        if ($rowCount >= 10000) {
            echo "Stopping at 10000 rows (demo limit)\n";
            return false; // Return false to stop processing
        }
        
        return true; // Continue processing
    });
    
    echo "\n✓ Stream processing complete!\n";
    echo "- Total rows processed: " . number_format($rowCount) . "\n";
    echo "- Sum of all numeric values: " . number_format($totalValue, 2) . "\n";
    echo "\nColumn sums:\n";
    foreach ($columnSums as $col => $sum) {
        echo "  Column {$col}: " . number_format($sum, 2) . "\n";
    }
    
} catch (XlsbParserException $e) {
    echo "Stream processing failed: " . $e->getMessage() . "\n";
}

// Example 4: Batch Processing Multiple Files
echo "\n\n=== Example 4: Batch Processing ===\n";

$parser = new UltimateNativeXlsbParser();
$xlsbFiles = [
    'file1.xlsb' => 'output1.xlsx',
    'file2.xlsb' => 'output2.xlsx',
    'file3.xlsb' => 'output3.xlsx',
];

$successful = 0;
$failed = 0;

foreach ($xlsbFiles as $xlsbFile => $xlsxFile) {
    echo "\nProcessing: {$xlsbFile}\n";
    
    try {
        if ($parser->convertToXlsx($xlsbFile, $xlsxFile)) {
            echo "  ✓ Success -> {$xlsxFile}\n";
            $successful++;
        } else {
            echo "  ✗ Failed: " . ($parser->getLastError() ?? 'Unknown error') . "\n";
            $failed++;
        }
    } catch (Exception $e) {
        echo "  ✗ Exception: " . $e->getMessage() . "\n";
        $failed++;
    }
}

echo "\n\nBatch processing summary:\n";
echo "- Successful: {$successful}\n";
echo "- Failed: {$failed}\n";
echo "- Total: " . ($successful + $failed) . "\n";

// Example 5: Custom Error Handling
echo "\n\n=== Example 5: Custom Error Handling ===\n";

class CustomXlsbProcessor {
    private $parser;
    private $logFile;
    
    public function __construct(string $logFile = 'xlsb-errors.log') {
        $this->parser = new UltimateNativeXlsbParser();
        $this->logFile = $logFile;
    }
    
    public function processWithLogging(string $xlsbFile, string $xlsxFile): bool {
        try {
            echo "Processing with custom error handling: {$xlsbFile}\n";
            
            // Enable debug mode
            $this->parser->setDebugMode(true);
            
            // Attempt conversion
            $result = $this->parser->convertToXlsx($xlsbFile, $xlsxFile);
            
            if ($result) {
                $this->log("SUCCESS", $xlsbFile, "Converted successfully");
                return true;
            } else {
                $error = $this->parser->getLastError() ?? 'Unknown error';
                $this->log("FAILURE", $xlsbFile, $error);
                return false;
            }
            
        } catch (XlsbParserException $e) {
            $this->log("PARSER_ERROR", $xlsbFile, $e->getMessage());
            throw $e;
        } catch (Exception $e) {
            $this->log("SYSTEM_ERROR", $xlsbFile, $e->getMessage());
            throw $e;
        }
    }
    
    private function log(string $level, string $file, string $message): void {
        $timestamp = date('Y-m-d H:i:s');
        $logEntry = "[{$timestamp}] [{$level}] {$file}: {$message}\n";
        
        file_put_contents($this->logFile, $logEntry, FILE_APPEND);
        echo "  Logged: {$level} - {$message}\n";
    }
}

// Use custom processor
$processor = new CustomXlsbProcessor();

try {
    $processor->processWithLogging('test.xlsb', 'test.xlsx');
} catch (Exception $e) {
    echo "Processing failed with exception: " . $e->getMessage() . "\n";
}

echo "\n✓ All examples completed!\n";