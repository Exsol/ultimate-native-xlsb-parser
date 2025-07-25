<?php

/**
 * Basic XLSB to XLSX Conversion Example
 * 
 * This example demonstrates the simplest way to convert an XLSB file to XLSX format
 */

require_once __DIR__ . '/../vendor/autoload.php';

use Exsol\UltimateNativeXlsbParser\UltimateNativeXlsbParser;

// Create parser instance
$parser = new UltimateNativeXlsbParser();

// Input and output file paths
$xlsbFile = 'path/to/your/file.xlsb';
$xlsxFile = 'path/to/output/file.xlsx';

try {
    // Convert XLSB to XLSX
    echo "Converting {$xlsbFile} to XLSX format...\n";
    
    $success = $parser->convertToXlsx($xlsbFile, $xlsxFile);
    
    if ($success) {
        echo "✓ Conversion successful!\n";
        echo "Output file: {$xlsxFile}\n";
        
        // Get statistics
        $stats = $parser->getStatistics();
        echo "\nConversion Statistics:\n";
        echo "- Rows processed: " . number_format($stats['rowCount']) . "\n";
        echo "- Cells processed: " . number_format($stats['cellCount']) . "\n";
        echo "- Processing time: " . round($stats['processingTime'], 2) . " seconds\n";
        echo "- Memory used: " . round($stats['memoryUsage'] / 1024 / 1024, 2) . " MB\n";
    } else {
        echo "✗ Conversion failed!\n";
        $error = $parser->getLastError();
        if ($error) {
            echo "Error: {$error}\n";
        }
    }
    
} catch (Exception $e) {
    echo "Error: " . $e->getMessage() . "\n";
    exit(1);
}