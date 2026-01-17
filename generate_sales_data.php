<?php
/**
 * Sales Data CSV Generator
 * 
 * This script generates a CSV file containing summarized sales data
 * for operators and their product groups.
 */

// Define output file path
$outputFile = __DIR__ . '/exports/sales_data.csv';

// Define operators
$operators = ['Operator_A', 'Operator_B', 'Operator_C'];

// Define periods
$periods = ['zeszly_rok', 'ostatnie_4_miesiace'];

// Define product groups
$productGroups = ['TOMC', 'TOMK', 'TOMM', 'TOMS', 'TOMT'];

// Define CSV headers
$headers = [
    'operator',
    'period',
    'product_group',
    'net_value',
    'number_of_positions',
    'returns_net_value',
    'margin',
    'margin_percentage',
    'quantity'
];

// Open file for writing
$file = fopen($outputFile, 'w');
if ($file === false) {
    die("Error: Cannot create file $outputFile\n");
}

// Write headers
fputcsv($file, $headers);

// Generate sample sales data
foreach ($operators as $operator) {
    foreach ($periods as $period) {
        foreach ($productGroups as $productGroup) {
            // Generate sample data for each combination
            $netValue = round(rand(10000, 100000) / 100, 2);
            $numberOfPositions = rand(10, 100);
            $returnsNetValue = round(rand(100, 5000) / 100, 2);
            $margin = round(rand(1000, 20000) / 100, 2);
            $marginPercentage = round(($margin / $netValue) * 100, 2);
            $quantity = rand(50, 500);
            
            // Write row to CSV
            $row = [
                $operator,
                $period,
                $productGroup,
                $netValue,
                $numberOfPositions,
                $returnsNetValue,
                $margin,
                $marginPercentage,
                $quantity
            ];
            
            fputcsv($file, $row);
        }
    }
}

// Close file
fclose($file);

echo "CSV file generated successfully: $outputFile\n";
echo "Total rows: " . (count($operators) * count($periods) * count($productGroups)) . " (plus header)\n";
