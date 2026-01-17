<?php

require_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

// Sample sales data structure
// NOTE: This is placeholder data for demonstration purposes.
// In production, replace this with actual data from your database, API, or other data source.
$salesData = [
    'AR' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 125000.50,
                'brutto' => 153750.62,
                'marza' => 31250.12,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 150,
                'ilosc' => 1200
            ],
            'TOMK' => [
                'netto' => 89000.00,
                'brutto' => 109470.00,
                'marza' => 22250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 120,
                'ilosc' => 950
            ],
            'TOMM' => [
                'netto' => 67500.00,
                'brutto' => 83025.00,
                'marza' => 16875.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 90,
                'ilosc' => 720
            ],
            'TOMS' => [
                'netto' => 45000.00,
                'brutto' => 55350.00,
                'marza' => 11250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 60,
                'ilosc' => 480
            ],
            'TOMT' => [
                'netto' => 34500.00,
                'brutto' => 42435.00,
                'marza' => 8625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 45,
                'ilosc' => 360
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 45000.00,
                'brutto' => 55350.00,
                'marza' => 11250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 55,
                'ilosc' => 440
            ],
            'TOMK' => [
                'netto' => 32000.00,
                'brutto' => 39360.00,
                'marza' => 8000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 45,
                'ilosc' => 350
            ],
            'TOMM' => [
                'netto' => 24000.00,
                'brutto' => 29520.00,
                'marza' => 6000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 35,
                'ilosc' => 280
            ],
            'TOMS' => [
                'netto' => 16000.00,
                'brutto' => 19680.00,
                'marza' => 4000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 25,
                'ilosc' => 200
            ],
            'TOMT' => [
                'netto' => 12000.00,
                'brutto' => 14760.00,
                'marza' => 3000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 18,
                'ilosc' => 150
            ]
        ]
    ],
    'TR' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 110000.00,
                'brutto' => 135300.00,
                'marza' => 27500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 135,
                'ilosc' => 1080
            ],
            'TOMK' => [
                'netto' => 78000.00,
                'brutto' => 95940.00,
                'marza' => 19500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 105,
                'ilosc' => 840
            ],
            'TOMM' => [
                'netto' => 59000.00,
                'brutto' => 72570.00,
                'marza' => 14750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 78,
                'ilosc' => 625
            ],
            'TOMS' => [
                'netto' => 39500.00,
                'brutto' => 48585.00,
                'marza' => 9875.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 52,
                'ilosc' => 420
            ],
            'TOMT' => [
                'netto' => 30000.00,
                'brutto' => 36900.00,
                'marza' => 7500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 40,
                'ilosc' => 320
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 40000.00,
                'brutto' => 49200.00,
                'marza' => 10000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 50,
                'ilosc' => 400
            ],
            'TOMK' => [
                'netto' => 28000.00,
                'brutto' => 34440.00,
                'marza' => 7000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 40,
                'ilosc' => 320
            ],
            'TOMM' => [
                'netto' => 21000.00,
                'brutto' => 25830.00,
                'marza' => 5250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 30,
                'ilosc' => 240
            ],
            'TOMS' => [
                'netto' => 14000.00,
                'brutto' => 17220.00,
                'marza' => 3500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 20,
                'ilosc' => 160
            ],
            'TOMT' => [
                'netto' => 10500.00,
                'brutto' => 12915.00,
                'marza' => 2625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 15,
                'ilosc' => 120
            ]
        ]
    ],
    'MŁ' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 95000.00,
                'brutto' => 116850.00,
                'marza' => 23750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 115,
                'ilosc' => 920
            ],
            'TOMK' => [
                'netto' => 67000.00,
                'brutto' => 82410.00,
                'marza' => 16750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 90,
                'ilosc' => 720
            ],
            'TOMM' => [
                'netto' => 50500.00,
                'brutto' => 62115.00,
                'marza' => 12625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 67,
                'ilosc' => 540
            ],
            'TOMS' => [
                'netto' => 33500.00,
                'brutto' => 41205.00,
                'marza' => 8375.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 45,
                'ilosc' => 360
            ],
            'TOMT' => [
                'netto' => 25500.00,
                'brutto' => 31365.00,
                'marza' => 6375.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 34,
                'ilosc' => 270
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 35000.00,
                'brutto' => 43050.00,
                'marza' => 8750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 42,
                'ilosc' => 340
            ],
            'TOMK' => [
                'netto' => 24000.00,
                'brutto' => 29520.00,
                'marza' => 6000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 33,
                'ilosc' => 260
            ],
            'TOMM' => [
                'netto' => 18000.00,
                'brutto' => 22140.00,
                'marza' => 4500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 25,
                'ilosc' => 200
            ],
            'TOMS' => [
                'netto' => 12000.00,
                'brutto' => 14760.00,
                'marza' => 3000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 17,
                'ilosc' => 135
            ],
            'TOMT' => [
                'netto' => 9000.00,
                'brutto' => 11070.00,
                'marza' => 2250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 12,
                'ilosc' => 100
            ]
        ]
    ],
    'GJ' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 85000.00,
                'brutto' => 104550.00,
                'marza' => 21250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 100,
                'ilosc' => 800
            ],
            'TOMK' => [
                'netto' => 60000.00,
                'brutto' => 73800.00,
                'marza' => 15000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 80,
                'ilosc' => 640
            ],
            'TOMM' => [
                'netto' => 45000.00,
                'brutto' => 55350.00,
                'marza' => 11250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 60,
                'ilosc' => 480
            ],
            'TOMS' => [
                'netto' => 30000.00,
                'brutto' => 36900.00,
                'marza' => 7500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 40,
                'ilosc' => 320
            ],
            'TOMT' => [
                'netto' => 22500.00,
                'brutto' => 27675.00,
                'marza' => 5625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 30,
                'ilosc' => 240
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 30000.00,
                'brutto' => 36900.00,
                'marza' => 7500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 37,
                'ilosc' => 295
            ],
            'TOMK' => [
                'netto' => 21000.00,
                'brutto' => 25830.00,
                'marza' => 5250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 28,
                'ilosc' => 225
            ],
            'TOMM' => [
                'netto' => 16000.00,
                'brutto' => 19680.00,
                'marza' => 4000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 22,
                'ilosc' => 175
            ],
            'TOMS' => [
                'netto' => 10500.00,
                'brutto' => 12915.00,
                'marza' => 2625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 15,
                'ilosc' => 120
            ],
            'TOMT' => [
                'netto' => 8000.00,
                'brutto' => 9840.00,
                'marza' => 2000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 11,
                'ilosc' => 90
            ]
        ]
    ],
    'MM' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 105000.00,
                'brutto' => 129150.00,
                'marza' => 26250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 125,
                'ilosc' => 1000
            ],
            'TOMK' => [
                'netto' => 74000.00,
                'brutto' => 91020.00,
                'marza' => 18500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 98,
                'ilosc' => 780
            ],
            'TOMM' => [
                'netto' => 56000.00,
                'brutto' => 68880.00,
                'marza' => 14000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 74,
                'ilosc' => 590
            ],
            'TOMS' => [
                'netto' => 37500.00,
                'brutto' => 46125.00,
                'marza' => 9375.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 50,
                'ilosc' => 400
            ],
            'TOMT' => [
                'netto' => 28000.00,
                'brutto' => 34440.00,
                'marza' => 7000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 37,
                'ilosc' => 295
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 38000.00,
                'brutto' => 46740.00,
                'marza' => 9500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 46,
                'ilosc' => 370
            ],
            'TOMK' => [
                'netto' => 27000.00,
                'brutto' => 33210.00,
                'marza' => 6750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 36,
                'ilosc' => 290
            ],
            'TOMM' => [
                'netto' => 20000.00,
                'brutto' => 24600.00,
                'marza' => 5000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 27,
                'ilosc' => 220
            ],
            'TOMS' => [
                'netto' => 13500.00,
                'brutto' => 16605.00,
                'marza' => 3375.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 18,
                'ilosc' => 145
            ],
            'TOMT' => [
                'netto' => 10000.00,
                'brutto' => 12300.00,
                'marza' => 2500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 13,
                'ilosc' => 110
            ]
        ]
    ],
    'GS' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 72000.00,
                'brutto' => 88560.00,
                'marza' => 18000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 85,
                'ilosc' => 680
            ],
            'TOMK' => [
                'netto' => 51000.00,
                'brutto' => 62730.00,
                'marza' => 12750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 68,
                'ilosc' => 540
            ],
            'TOMM' => [
                'netto' => 38500.00,
                'brutto' => 47355.00,
                'marza' => 9625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 51,
                'ilosc' => 410
            ],
            'TOMS' => [
                'netto' => 25500.00,
                'brutto' => 31365.00,
                'marza' => 6375.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 34,
                'ilosc' => 270
            ],
            'TOMT' => [
                'netto' => 19500.00,
                'brutto' => 23985.00,
                'marza' => 4875.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 26,
                'ilosc' => 210
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 26000.00,
                'brutto' => 31980.00,
                'marza' => 6500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 31,
                'ilosc' => 250
            ],
            'TOMK' => [
                'netto' => 18500.00,
                'brutto' => 22755.00,
                'marza' => 4625.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 25,
                'ilosc' => 200
            ],
            'TOMM' => [
                'netto' => 14000.00,
                'brutto' => 17220.00,
                'marza' => 3500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 19,
                'ilosc' => 150
            ],
            'TOMS' => [
                'netto' => 9000.00,
                'brutto' => 11070.00,
                'marza' => 2250.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 12,
                'ilosc' => 100
            ],
            'TOMT' => [
                'netto' => 7000.00,
                'brutto' => 8610.00,
                'marza' => 1750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 9,
                'ilosc' => 75
            ]
        ]
    ],
    'WEB' => [
        'zeszly_rok' => [
            'TOMC' => [
                'netto' => 155000.00,
                'brutto' => 190650.00,
                'marza' => 38750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 180,
                'ilosc' => 1440
            ],
            'TOMK' => [
                'netto' => 110000.00,
                'brutto' => 135300.00,
                'marza' => 27500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 145,
                'ilosc' => 1160
            ],
            'TOMM' => [
                'netto' => 83000.00,
                'brutto' => 102090.00,
                'marza' => 20750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 110,
                'ilosc' => 880
            ],
            'TOMS' => [
                'netto' => 55500.00,
                'brutto' => 68265.00,
                'marza' => 13875.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 74,
                'ilosc' => 590
            ],
            'TOMT' => [
                'netto' => 42000.00,
                'brutto' => 51660.00,
                'marza' => 10500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 56,
                'ilosc' => 450
            ]
        ],
        'ostatnie_4_miesiace' => [
            'TOMC' => [
                'netto' => 56000.00,
                'brutto' => 68880.00,
                'marza' => 14000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 67,
                'ilosc' => 535
            ],
            'TOMK' => [
                'netto' => 40000.00,
                'brutto' => 49200.00,
                'marza' => 10000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 53,
                'ilosc' => 425
            ],
            'TOMM' => [
                'netto' => 30000.00,
                'brutto' => 36900.00,
                'marza' => 7500.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 40,
                'ilosc' => 320
            ],
            'TOMS' => [
                'netto' => 20000.00,
                'brutto' => 24600.00,
                'marza' => 5000.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 27,
                'ilosc' => 215
            ],
            'TOMT' => [
                'netto' => 15000.00,
                'brutto' => 18450.00,
                'marza' => 3750.00,
                'procent_marzy' => 25.0,
                'liczba_pozycji' => 20,
                'ilosc' => 165
            ]
        ]
    ]
];

// Create new Spreadsheet
$spreadsheet = new Spreadsheet();

// Function to format header row
function formatHeader($sheet, $row, $startCol, $endCol)
{
    $range = $startCol . $row . ':' . $endCol . $row;
    $sheet->getStyle($range)->applyFromArray([
        'font' => [
            'bold' => true,
            'color' => ['rgb' => 'FFFFFF']
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['rgb' => '4472C4']
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => '000000']
            ]
        ]
    ]);
}

// Function to format data rows
function formatDataRow($sheet, $row, $startCol, $endCol)
{
    $range = $startCol . $row . ':' . $endCol . $row;
    $sheet->getStyle($range)->applyFromArray([
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => 'CCCCCC']
            ]
        ]
    ]);
}

// Function to add a section title
function addSectionTitle($sheet, $row, $title)
{
    $sheet->setCellValue('A' . $row, $title);
    $sheet->getStyle('A' . $row . ':G' . $row)->applyFromArray([
        'font' => [
            'bold' => true,
            'size' => 12,
            'color' => ['rgb' => '000000']
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'startColor' => ['rgb' => 'E7E6E6']
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_LEFT,
            'vertical' => Alignment::VERTICAL_CENTER
        ]
    ]);
    $sheet->mergeCells('A' . $row . ':G' . $row);
}

// Function to add data table
function addDataTable($sheet, &$row, $title, $data)
{
    // Add section title
    addSectionTitle($sheet, $row, $title);
    $row++;
    
    // Add header
    $headers = ['Grupa', 'Wartość netto', 'Wartość brutto', 'Wartość marży', 'Procent marży', 'Liczba pozycji', 'Ilości'];
    $sheet->fromArray($headers, null, 'A' . $row);
    formatHeader($sheet, $row, 'A', 'G');
    $row++;
    
    // Add data rows
    $totals = [
        'netto' => 0,
        'brutto' => 0,
        'marza' => 0,
        'liczba_pozycji' => 0,
        'ilosc' => 0
    ];
    
    foreach ($data as $grupa => $values) {
        $sheet->setCellValue('A' . $row, $grupa);
        $sheet->setCellValue('B' . $row, $values['netto']);
        $sheet->setCellValue('C' . $row, $values['brutto']);
        $sheet->setCellValue('D' . $row, $values['marza']);
        $sheet->setCellValue('E' . $row, $values['procent_marzy'] . '%');
        $sheet->setCellValue('F' . $row, $values['liczba_pozycji']);
        $sheet->setCellValue('G' . $row, $values['ilosc']);
        
        // Format monetary values
        $sheet->getStyle('B' . $row . ':D' . $row)->getNumberFormat()
            ->setFormatCode('#,##0.00 zł');
        
        formatDataRow($sheet, $row, 'A', 'G');
        
        $totals['netto'] += $values['netto'];
        $totals['brutto'] += $values['brutto'];
        $totals['marza'] += $values['marza'];
        $totals['liczba_pozycji'] += $values['liczba_pozycji'];
        $totals['ilosc'] += $values['ilosc'];
        
        $row++;
    }
    
    // Add totals row
    $sheet->setCellValue('A' . $row, 'SUMA');
    $sheet->setCellValue('B' . $row, $totals['netto']);
    $sheet->setCellValue('C' . $row, $totals['brutto']);
    $sheet->setCellValue('D' . $row, $totals['marza']);
    $avgMargin = $totals['netto'] > 0 ? ($totals['marza'] / $totals['netto']) * 100 : 0;
    $sheet->setCellValue('E' . $row, number_format($avgMargin, 2) . '%');
    $sheet->setCellValue('F' . $row, $totals['liczba_pozycji']);
    $sheet->setCellValue('G' . $row, $totals['ilosc']);
    
    $sheet->getStyle('B' . $row . ':D' . $row)->getNumberFormat()
        ->setFormatCode('#,##0.00 zł');
    
    $sheet->getStyle('A' . $row . ':G' . $row)->applyFromArray([
        'font' => ['bold' => true],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_MEDIUM,
                'color' => ['rgb' => '000000']
            ]
        ]
    ]);
    
    $row += 2; // Add spacing
}

// Create Summary Sheet
$summarySheet = $spreadsheet->getActiveSheet();
$summarySheet->setTitle('Podsumowanie');

$row = 1;
$summarySheet->setCellValue('A' . $row, 'PODSUMOWANIE SPRZEDAŻY - PORÓWNANIE OPERATORÓW');
$summarySheet->getStyle('A' . $row)->applyFromArray([
    'font' => [
        'bold' => true,
        'size' => 14,
        'color' => ['rgb' => '000000']
    ]
]);
$row += 2;

// Yearly comparison
addSectionTitle($summarySheet, $row, 'Sprzedaż - Zeszły Rok');
$row++;

$headers = ['Operator', 'Wartość netto', 'Wartość brutto', 'Wartość marży', 'Procent marży'];
$summarySheet->fromArray($headers, null, 'A' . $row);
formatHeader($summarySheet, $row, 'A', 'E');
$row++;

$yearlyTotals = ['netto' => 0, 'brutto' => 0, 'marza' => 0];
foreach ($salesData as $operator => $periods) {
    $netto = array_sum(array_column($periods['zeszly_rok'], 'netto'));
    $brutto = array_sum(array_column($periods['zeszly_rok'], 'brutto'));
    $marza = array_sum(array_column($periods['zeszly_rok'], 'marza'));
    $procent = $netto > 0 ? ($marza / $netto) * 100 : 0;
    
    $summarySheet->setCellValue('A' . $row, $operator);
    $summarySheet->setCellValue('B' . $row, $netto);
    $summarySheet->setCellValue('C' . $row, $brutto);
    $summarySheet->setCellValue('D' . $row, $marza);
    $summarySheet->setCellValue('E' . $row, number_format($procent, 2) . '%');
    
    $summarySheet->getStyle('B' . $row . ':D' . $row)->getNumberFormat()
        ->setFormatCode('#,##0.00 zł');
    
    formatDataRow($summarySheet, $row, 'A', 'E');
    
    $yearlyTotals['netto'] += $netto;
    $yearlyTotals['brutto'] += $brutto;
    $yearlyTotals['marza'] += $marza;
    
    $row++;
}

// Add yearly totals
$summarySheet->setCellValue('A' . $row, 'SUMA');
$summarySheet->setCellValue('B' . $row, $yearlyTotals['netto']);
$summarySheet->setCellValue('C' . $row, $yearlyTotals['brutto']);
$summarySheet->setCellValue('D' . $row, $yearlyTotals['marza']);
$avgMargin = $yearlyTotals['netto'] > 0 ? ($yearlyTotals['marza'] / $yearlyTotals['netto']) * 100 : 0;
$summarySheet->setCellValue('E' . $row, number_format($avgMargin, 2) . '%');

$summarySheet->getStyle('B' . $row . ':D' . $row)->getNumberFormat()
    ->setFormatCode('#,##0.00 zł');

$summarySheet->getStyle('A' . $row . ':E' . $row)->applyFromArray([
    'font' => ['bold' => true],
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_MEDIUM,
            'color' => ['rgb' => '000000']
        ]
    ]
]);

$row += 3;

// Last 4 months comparison
addSectionTitle($summarySheet, $row, 'Sprzedaż - Ostatnie 4 Miesiące');
$row++;

$summarySheet->fromArray($headers, null, 'A' . $row);
formatHeader($summarySheet, $row, 'A', 'E');
$row++;

$last4Totals = ['netto' => 0, 'brutto' => 0, 'marza' => 0];
foreach ($salesData as $operator => $periods) {
    $netto = array_sum(array_column($periods['ostatnie_4_miesiace'], 'netto'));
    $brutto = array_sum(array_column($periods['ostatnie_4_miesiace'], 'brutto'));
    $marza = array_sum(array_column($periods['ostatnie_4_miesiace'], 'marza'));
    $procent = $netto > 0 ? ($marza / $netto) * 100 : 0;
    
    $summarySheet->setCellValue('A' . $row, $operator);
    $summarySheet->setCellValue('B' . $row, $netto);
    $summarySheet->setCellValue('C' . $row, $brutto);
    $summarySheet->setCellValue('D' . $row, $marza);
    $summarySheet->setCellValue('E' . $row, number_format($procent, 2) . '%');
    
    $summarySheet->getStyle('B' . $row . ':D' . $row)->getNumberFormat()
        ->setFormatCode('#,##0.00 zł');
    
    formatDataRow($summarySheet, $row, 'A', 'E');
    
    $last4Totals['netto'] += $netto;
    $last4Totals['brutto'] += $brutto;
    $last4Totals['marza'] += $marza;
    
    $row++;
}

// Add last 4 months totals
$summarySheet->setCellValue('A' . $row, 'SUMA');
$summarySheet->setCellValue('B' . $row, $last4Totals['netto']);
$summarySheet->setCellValue('C' . $row, $last4Totals['brutto']);
$summarySheet->setCellValue('D' . $row, $last4Totals['marza']);
$avgMargin = $last4Totals['netto'] > 0 ? ($last4Totals['marza'] / $last4Totals['netto']) * 100 : 0;
$summarySheet->setCellValue('E' . $row, number_format($avgMargin, 2) . '%');

$summarySheet->getStyle('B' . $row . ':D' . $row)->getNumberFormat()
    ->setFormatCode('#,##0.00 zł');

$summarySheet->getStyle('A' . $row . ':E' . $row)->applyFromArray([
    'font' => ['bold' => true],
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_MEDIUM,
            'color' => ['rgb' => '000000']
        ]
    ]
]);

// Auto-size columns
foreach (range('A', 'G') as $col) {
    $summarySheet->getColumnDimension($col)->setAutoSize(true);
}

// Create sheets for each operator
foreach ($salesData as $operator => $periods) {
    $sheet = $spreadsheet->createSheet();
    $sheet->setTitle($operator);
    
    $row = 1;
    $sheet->setCellValue('A' . $row, 'RAPORT SPRZEDAŻY - ' . $operator);
    $sheet->getStyle('A' . $row)->applyFromArray([
        'font' => [
            'bold' => true,
            'size' => 14,
            'color' => ['rgb' => '000000']
        ]
    ]);
    $row += 2;
    
    // Add yearly data table
    addDataTable($sheet, $row, 'Sprzedaż - Zeszły Rok', $periods['zeszly_rok']);
    
    // Add last 4 months data table
    addDataTable($sheet, $row, 'Sprzedaż - Ostatnie 4 Miesiące', $periods['ostatnie_4_miesiace']);
    
    // Auto-size columns
    foreach (range('A', 'G') as $col) {
        $sheet->getColumnDimension($col)->setAutoSize(true);
    }
}

// Save the file
// Create the exports directory if it doesn't exist (permissions: 0755 = rwxr-xr-x)
$outputDir = __DIR__ . '/exports';
if (!is_dir($outputDir)) {
    mkdir($outputDir, 0755, true);
}

$writer = new Xlsx($spreadsheet);
$outputPath = $outputDir . '/sales_data.xlsx';
$writer->save($outputPath);

// Output success message with file location and size for verification
echo "Excel file generated successfully: " . $outputPath . "\n";
echo "File size: " . filesize($outputPath) . " bytes\n";
