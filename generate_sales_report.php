<?php
/**
 * Excel Sales Report Generator
 * Generates an Excel file with sales data for multiple operators
 */

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// Sample sales data structure
// In production, this would come from a database or API
$salesData = [
    'AR' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 125000, 'wartosc_netto_zwroty' => 5000, 'marza' => 25000, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 98000, 'wartosc_netto_zwroty' => 3500, 'marza' => 18000, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 156000, 'wartosc_netto_zwroty' => 6200, 'marza' => 31200, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 87000, 'wartosc_netto_zwroty' => 2900, 'marza' => 15000, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 203000, 'wartosc_netto_zwroty' => 8100, 'marza' => 40600, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 45000, 'wartosc_netto_zwroty' => 1800, 'marza' => 9000, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 35000, 'wartosc_netto_zwroty' => 1200, 'marza' => 6500, 'marza_procent' => 18.6],
            'TOMM' => ['wartosc_netto' => 56000, 'wartosc_netto_zwroty' => 2200, 'marza' => 11200, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 31000, 'wartosc_netto_zwroty' => 1000, 'marza' => 5400, 'marza_procent' => 17.4],
            'TOMT' => ['wartosc_netto' => 73000, 'wartosc_netto_zwroty' => 2900, 'marza' => 14600, 'marza_procent' => 20],
        ],
    ],
    'TR' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 145000, 'wartosc_netto_zwroty' => 6000, 'marza' => 29000, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 112000, 'wartosc_netto_zwroty' => 4200, 'marza' => 20500, 'marza_procent' => 18.3],
            'TOMM' => ['wartosc_netto' => 178000, 'wartosc_netto_zwroty' => 7100, 'marza' => 35600, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 99000, 'wartosc_netto_zwroty' => 3300, 'marza' => 17000, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 231000, 'wartosc_netto_zwroty' => 9200, 'marza' => 46200, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 52000, 'wartosc_netto_zwroty' => 2100, 'marza' => 10400, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 40000, 'wartosc_netto_zwroty' => 1400, 'marza' => 7400, 'marza_procent' => 18.5],
            'TOMM' => ['wartosc_netto' => 64000, 'wartosc_netto_zwroty' => 2500, 'marza' => 12800, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 35000, 'wartosc_netto_zwroty' => 1100, 'marza' => 6100, 'marza_procent' => 17.4],
            'TOMT' => ['wartosc_netto' => 83000, 'wartosc_netto_zwroty' => 3300, 'marza' => 16600, 'marza_procent' => 20],
        ],
    ],
    'MŁ' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 98000, 'wartosc_netto_zwroty' => 4000, 'marza' => 19600, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 76000, 'wartosc_netto_zwroty' => 2800, 'marza' => 14000, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 120000, 'wartosc_netto_zwroty' => 4800, 'marza' => 24000, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 67000, 'wartosc_netto_zwroty' => 2200, 'marza' => 11500, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 156000, 'wartosc_netto_zwroty' => 6200, 'marza' => 31200, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 35000, 'wartosc_netto_zwroty' => 1400, 'marza' => 7000, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 27000, 'wartosc_netto_zwroty' => 950, 'marza' => 5000, 'marza_procent' => 18.5],
            'TOMM' => ['wartosc_netto' => 43000, 'wartosc_netto_zwroty' => 1700, 'marza' => 8600, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 24000, 'wartosc_netto_zwroty' => 800, 'marza' => 4200, 'marza_procent' => 17.5],
            'TOMT' => ['wartosc_netto' => 56000, 'wartosc_netto_zwroty' => 2200, 'marza' => 11200, 'marza_procent' => 20],
        ],
    ],
    'GJ' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 112000, 'wartosc_netto_zwroty' => 4500, 'marza' => 22400, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 87000, 'wartosc_netto_zwroty' => 3200, 'marza' => 16000, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 137000, 'wartosc_netto_zwroty' => 5500, 'marza' => 27400, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 76000, 'wartosc_netto_zwroty' => 2500, 'marza' => 13100, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 178000, 'wartosc_netto_zwroty' => 7100, 'marza' => 35600, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 40000, 'wartosc_netto_zwroty' => 1600, 'marza' => 8000, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 31000, 'wartosc_netto_zwroty' => 1100, 'marza' => 5700, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 49000, 'wartosc_netto_zwroty' => 1900, 'marza' => 9800, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 27000, 'wartosc_netto_zwroty' => 900, 'marza' => 4700, 'marza_procent' => 17.4],
            'TOMT' => ['wartosc_netto' => 64000, 'wartosc_netto_zwroty' => 2500, 'marza' => 12800, 'marza_procent' => 20],
        ],
    ],
    'MM' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 134000, 'wartosc_netto_zwroty' => 5400, 'marza' => 26800, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 104000, 'wartosc_netto_zwroty' => 3800, 'marza' => 19100, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 164000, 'wartosc_netto_zwroty' => 6600, 'marza' => 32800, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 91000, 'wartosc_netto_zwroty' => 3000, 'marza' => 15700, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 213000, 'wartosc_netto_zwroty' => 8500, 'marza' => 42600, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 48000, 'wartosc_netto_zwroty' => 1900, 'marza' => 9600, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 37000, 'wartosc_netto_zwroty' => 1300, 'marza' => 6800, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 59000, 'wartosc_netto_zwroty' => 2300, 'marza' => 11800, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 33000, 'wartosc_netto_zwroty' => 1000, 'marza' => 5700, 'marza_procent' => 17.3],
            'TOMT' => ['wartosc_netto' => 76000, 'wartosc_netto_zwroty' => 3000, 'marza' => 15200, 'marza_procent' => 20],
        ],
    ],
    'GS' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 89000, 'wartosc_netto_zwroty' => 3600, 'marza' => 17800, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 69000, 'wartosc_netto_zwroty' => 2500, 'marza' => 12700, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 109000, 'wartosc_netto_zwroty' => 4400, 'marza' => 21800, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 61000, 'wartosc_netto_zwroty' => 2000, 'marza' => 10500, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 142000, 'wartosc_netto_zwroty' => 5700, 'marza' => 28400, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 32000, 'wartosc_netto_zwroty' => 1300, 'marza' => 6400, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 25000, 'wartosc_netto_zwroty' => 900, 'marza' => 4600, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 39000, 'wartosc_netto_zwroty' => 1500, 'marza' => 7800, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 22000, 'wartosc_netto_zwroty' => 700, 'marza' => 3800, 'marza_procent' => 17.3],
            'TOMT' => ['wartosc_netto' => 51000, 'wartosc_netto_zwroty' => 2000, 'marza' => 10200, 'marza_procent' => 20],
        ],
    ],
    'WEB' => [
        'zeszly_rok' => [
            'TOMC' => ['wartosc_netto' => 178000, 'wartosc_netto_zwroty' => 7100, 'marza' => 35600, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 138000, 'wartosc_netto_zwroty' => 5000, 'marza' => 25400, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 217000, 'wartosc_netto_zwroty' => 8700, 'marza' => 43400, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 121000, 'wartosc_netto_zwroty' => 4000, 'marza' => 20800, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 282000, 'wartosc_netto_zwroty' => 11300, 'marza' => 56400, 'marza_procent' => 20],
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => ['wartosc_netto' => 64000, 'wartosc_netto_zwroty' => 2500, 'marza' => 12800, 'marza_procent' => 20],
            'TOMK' => ['wartosc_netto' => 49000, 'wartosc_netto_zwroty' => 1800, 'marza' => 9000, 'marza_procent' => 18.4],
            'TOMM' => ['wartosc_netto' => 78000, 'wartosc_netto_zwroty' => 3100, 'marza' => 15600, 'marza_procent' => 20],
            'TOMS' => ['wartosc_netto' => 43000, 'wartosc_netto_zwroty' => 1400, 'marza' => 7400, 'marza_procent' => 17.2],
            'TOMT' => ['wartosc_netto' => 101000, 'wartosc_netto_zwroty' => 4000, 'marza' => 20200, 'marza_procent' => 20],
        ],
    ],
];

$operators = ['AR', 'TR', 'MŁ', 'GJ', 'MM', 'GS', 'WEB'];
$productGroups = ['TOMC', 'TOMK', 'TOMM', 'TOMS', 'TOMT'];
$periods = ['zeszly_rok', 'ostatnie_4_miesiace'];

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();

// Remove default sheet
$spreadsheet->removeSheetByIndex(0);

/**
 * Apply header styling to a range
 */
function styleHeader($sheet, $range) {
    $sheet->getStyle($range)->applyFromArray([
        'font' => [
            'bold' => true,
            'color' => ['rgb' => 'FFFFFF'],
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['rgb' => '4472C4'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
            ],
        ],
    ]);
}

/**
 * Apply cell borders
 */
function styleBorders($sheet, $range) {
    $sheet->getStyle($range)->applyFromArray([
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => '000000'],
            ],
        ],
    ]);
}

/**
 * Create operator sheet with detailed data
 */
function createOperatorSheet($spreadsheet, $operator, $data, $productGroups, $periods) {
    $sheet = $spreadsheet->createSheet();
    $sheet->setTitle($operator);
    
    // Set column widths
    $sheet->getColumnDimension('A')->setWidth(25);
    $sheet->getColumnDimension('B')->setWidth(18);
    $sheet->getColumnDimension('C')->setWidth(18);
    $sheet->getColumnDimension('D')->setWidth(15);
    $sheet->getColumnDimension('E')->setWidth(18);
    
    // Title
    $sheet->setCellValue('A1', "Raport sprzedaży - Operator: $operator");
    $sheet->mergeCells('A1:E1');
    $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
    $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    
    $row = 3;
    
    foreach ($periods as $period) {
        $periodName = $period === 'zeszly_rok' ? 'Zeszły rok' : 'Ostatnie 4 miesiące';
        
        // Period header
        $sheet->setCellValue("A$row", $periodName);
        $sheet->mergeCells("A$row:E$row");
        $sheet->getStyle("A$row")->getFont()->setBold(true)->setSize(12);
        $sheet->getStyle("A$row")->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('D9E1F2');
        $row++;
        
        // Column headers
        $sheet->setCellValue("A$row", 'Grupa produktowa');
        $sheet->setCellValue("B$row", 'Wartość netto');
        $sheet->setCellValue("C$row", 'Zwroty netto');
        $sheet->setCellValue("D$row", 'Marża');
        $sheet->setCellValue("E$row", 'Marża (%)');
        styleHeader($sheet, "A$row:E$row");
        $row++;
        
        // Data rows
        $periodTotal = [
            'wartosc_netto' => 0,
            'wartosc_netto_zwroty' => 0,
            'marza' => 0,
        ];
        
        foreach ($productGroups as $group) {
            if (isset($data[$period][$group])) {
                $groupData = $data[$period][$group];
                
                $sheet->setCellValue("A$row", $group);
                $sheet->setCellValue("B$row", $groupData['wartosc_netto']);
                $sheet->setCellValue("C$row", $groupData['wartosc_netto_zwroty']);
                $sheet->setCellValue("D$row", $groupData['marza']);
                $sheet->setCellValue("E$row", $groupData['marza_procent']);
                
                // Number formatting
                $sheet->getStyle("B$row:D$row")->getNumberFormat()->setFormatCode('#,##0.00');
                $sheet->getStyle("E$row")->getNumberFormat()->setFormatCode('0.00"%"');
                
                styleBorders($sheet, "A$row:E$row");
                
                $periodTotal['wartosc_netto'] += $groupData['wartosc_netto'];
                $periodTotal['wartosc_netto_zwroty'] += $groupData['wartosc_netto_zwroty'];
                $periodTotal['marza'] += $groupData['marza'];
                
                $row++;
            }
        }
        
        // Period total
        $sheet->setCellValue("A$row", 'RAZEM');
        $sheet->setCellValue("B$row", $periodTotal['wartosc_netto']);
        $sheet->setCellValue("C$row", $periodTotal['wartosc_netto_zwroty']);
        $sheet->setCellValue("D$row", $periodTotal['marza']);
        if ($periodTotal['wartosc_netto'] > 0) {
            $avgMarginPercent = ($periodTotal['marza'] / $periodTotal['wartosc_netto']) * 100;
            $sheet->setCellValue("E$row", $avgMarginPercent);
        }
        
        $sheet->getStyle("A$row:E$row")->getFont()->setBold(true);
        $sheet->getStyle("B$row:D$row")->getNumberFormat()->setFormatCode('#,##0.00');
        $sheet->getStyle("E$row")->getNumberFormat()->setFormatCode('0.00"%"');
        $sheet->getStyle("A$row:E$row")->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB('E2EFDA');
        styleBorders($sheet, "A$row:E$row");
        
        $row += 2;
    }
}

// Create individual sheets for each operator
foreach ($operators as $operator) {
    if (isset($salesData[$operator])) {
        createOperatorSheet($spreadsheet, $operator, $salesData[$operator], $productGroups, $periods);
    }
}

// Create consolidated summary sheet
$summarySheet = $spreadsheet->createSheet();
$summarySheet->setTitle('Podsumowanie TOMEX');
$spreadsheet->setActiveSheetIndexByName('Podsumowanie TOMEX');

// Set column widths for summary
$summarySheet->getColumnDimension('A')->setWidth(25);
$summarySheet->getColumnDimension('B')->setWidth(18);
$summarySheet->getColumnDimension('C')->setWidth(18);
$summarySheet->getColumnDimension('D')->setWidth(15);
$summarySheet->getColumnDimension('E')->setWidth(18);

// Title
$summarySheet->setCellValue('A1', 'Podsumowanie TOMEX - Wszyscy operatorzy');
$summarySheet->mergeCells('A1:E1');
$summarySheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
$summarySheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

$row = 3;

foreach ($periods as $period) {
    $periodName = $period === 'zeszly_rok' ? 'Zeszły rok' : 'Ostatnie 4 miesiące';
    
    // Period header
    $summarySheet->setCellValue("A$row", $periodName);
    $summarySheet->mergeCells("A$row:E$row");
    $summarySheet->getStyle("A$row")->getFont()->setBold(true)->setSize(12);
    $summarySheet->getStyle("A$row")->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setRGB('D9E1F2');
    $row++;
    
    // Column headers
    $summarySheet->setCellValue("A$row", 'Grupa produktowa');
    $summarySheet->setCellValue("B$row", 'Wartość netto');
    $summarySheet->setCellValue("C$row", 'Zwroty netto');
    $summarySheet->setCellValue("D$row", 'Marża');
    $summarySheet->setCellValue("E$row", 'Marża (%)');
    styleHeader($summarySheet, "A$row:E$row");
    $row++;
    
    // Aggregate data across all operators
    $consolidatedData = [];
    foreach ($productGroups as $group) {
        $consolidatedData[$group] = [
            'wartosc_netto' => 0,
            'wartosc_netto_zwroty' => 0,
            'marza' => 0,
        ];
        
        foreach ($operators as $operator) {
            if (isset($salesData[$operator][$period][$group])) {
                $groupData = $salesData[$operator][$period][$group];
                $consolidatedData[$group]['wartosc_netto'] += $groupData['wartosc_netto'];
                $consolidatedData[$group]['wartosc_netto_zwroty'] += $groupData['wartosc_netto_zwroty'];
                $consolidatedData[$group]['marza'] += $groupData['marza'];
            }
        }
    }
    
    // Display consolidated data
    $periodTotal = [
        'wartosc_netto' => 0,
        'wartosc_netto_zwroty' => 0,
        'marza' => 0,
    ];
    
    foreach ($productGroups as $group) {
        $groupData = $consolidatedData[$group];
        
        $summarySheet->setCellValue("A$row", $group);
        $summarySheet->setCellValue("B$row", $groupData['wartosc_netto']);
        $summarySheet->setCellValue("C$row", $groupData['wartosc_netto_zwroty']);
        $summarySheet->setCellValue("D$row", $groupData['marza']);
        
        if ($groupData['wartosc_netto'] > 0) {
            $marzaPercent = ($groupData['marza'] / $groupData['wartosc_netto']) * 100;
            $summarySheet->setCellValue("E$row", $marzaPercent);
        }
        
        // Number formatting
        $summarySheet->getStyle("B$row:D$row")->getNumberFormat()->setFormatCode('#,##0.00');
        $summarySheet->getStyle("E$row")->getNumberFormat()->setFormatCode('0.00"%"');
        
        styleBorders($summarySheet, "A$row:E$row");
        
        $periodTotal['wartosc_netto'] += $groupData['wartosc_netto'];
        $periodTotal['wartosc_netto_zwroty'] += $groupData['wartosc_netto_zwroty'];
        $periodTotal['marza'] += $groupData['marza'];
        
        $row++;
    }
    
    // Period total
    $summarySheet->setCellValue("A$row", 'RAZEM');
    $summarySheet->setCellValue("B$row", $periodTotal['wartosc_netto']);
    $summarySheet->setCellValue("C$row", $periodTotal['wartosc_netto_zwroty']);
    $summarySheet->setCellValue("D$row", $periodTotal['marza']);
    if ($periodTotal['wartosc_netto'] > 0) {
        $avgMarginPercent = ($periodTotal['marza'] / $periodTotal['wartosc_netto']) * 100;
        $summarySheet->setCellValue("E$row", $avgMarginPercent);
    }
    
    $summarySheet->getStyle("A$row:E$row")->getFont()->setBold(true);
    $summarySheet->getStyle("B$row:D$row")->getNumberFormat()->setFormatCode('#,##0.00');
    $summarySheet->getStyle("E$row")->getNumberFormat()->setFormatCode('0.00"%"');
    $summarySheet->getStyle("A$row:E$row")->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setRGB('E2EFDA');
    styleBorders($summarySheet, "A$row:E$row");
    
    $row += 2;
}

// Add operator breakdown section
$row++;
$summarySheet->setCellValue("A$row", 'Zestawienie operatorów');
$summarySheet->mergeCells("A$row:E$row");
$summarySheet->getStyle("A$row")->getFont()->setBold(true)->setSize(12);
$summarySheet->getStyle("A$row")->getFill()
    ->setFillType(Fill::FILL_SOLID)
    ->getStartColor()->setRGB('FFE699');
$row++;

foreach ($periods as $period) {
    $periodName = $period === 'zeszly_rok' ? 'Zeszły rok' : 'Ostatnie 4 miesiące';
    
    $summarySheet->setCellValue("A$row", $periodName);
    $summarySheet->mergeCells("A$row:E$row");
    $summarySheet->getStyle("A$row")->getFont()->setBold(true);
    $row++;
    
    // Column headers
    $summarySheet->setCellValue("A$row", 'Operator');
    $summarySheet->setCellValue("B$row", 'Wartość netto');
    $summarySheet->setCellValue("C$row", 'Zwroty netto');
    $summarySheet->setCellValue("D$row", 'Marża');
    $summarySheet->setCellValue("E$row", 'Marża (%)');
    styleHeader($summarySheet, "A$row:E$row");
    $row++;
    
    foreach ($operators as $operator) {
        if (isset($salesData[$operator][$period])) {
            $operatorTotal = [
                'wartosc_netto' => 0,
                'wartosc_netto_zwroty' => 0,
                'marza' => 0,
            ];
            
            foreach ($productGroups as $group) {
                if (isset($salesData[$operator][$period][$group])) {
                    $groupData = $salesData[$operator][$period][$group];
                    $operatorTotal['wartosc_netto'] += $groupData['wartosc_netto'];
                    $operatorTotal['wartosc_netto_zwroty'] += $groupData['wartosc_netto_zwroty'];
                    $operatorTotal['marza'] += $groupData['marza'];
                }
            }
            
            $summarySheet->setCellValue("A$row", $operator);
            $summarySheet->setCellValue("B$row", $operatorTotal['wartosc_netto']);
            $summarySheet->setCellValue("C$row", $operatorTotal['wartosc_netto_zwroty']);
            $summarySheet->setCellValue("D$row", $operatorTotal['marza']);
            
            if ($operatorTotal['wartosc_netto'] > 0) {
                $marzaPercent = ($operatorTotal['marza'] / $operatorTotal['wartosc_netto']) * 100;
                $summarySheet->setCellValue("E$row", $marzaPercent);
            }
            
            $summarySheet->getStyle("B$row:D$row")->getNumberFormat()->setFormatCode('#,##0.00');
            $summarySheet->getStyle("E$row")->getNumberFormat()->setFormatCode('0.00"%"');
            styleBorders($summarySheet, "A$row:E$row");
            
            $row++;
        }
    }
    
    $row++;
}

// Save the file
$outputFilename = 'raport_sprzedazy_TOMEX_' . date('Y-m-d_H-i-s') . '.xlsx';
$writer = new Xlsx($spreadsheet);
$writer->save($outputFilename);

echo "Raport został wygenerowany: $outputFilename\n";
echo "Plik zawiera:\n";
echo "- Arkusz podsumowania TOMEX (wszystkie operatory)\n";
foreach ($operators as $operator) {
    echo "- Arkusz szczegółowy dla operatora: $operator\n";
}
