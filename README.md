# Dataapp - Excel Sales Report Generator

This application generates comprehensive Excel reports for TOMEX sales data across multiple operators and product groups.

## Features

- **Individual Operator Sheets**: Detailed breakdown for each operator (AR, TR, MŁ, GJ, MM, GS, WEB)
- **Consolidated Summary**: TOMEX-wide view with aggregated data across all operators
- **Time Period Analysis**: Data for "Zeszły rok" (Last Year) and "Ostatnie 4 miesiące" (Last 4 Months)
- **Product Group Breakdown**: Performance metrics for TOMC, TOMK, TOMM, TOMS, TOMT
- **Key Metrics**:
  - Wartość netto (Net Sales Value)
  - Zwroty netto (Net Returns)
  - Marża (Margin in currency)
  - Marża % (Margin Percentage)
- **Professional Formatting**: Color-coded headers, borders, proper number formatting

## Requirements

- PHP 7.4 or higher
- Composer (PHP dependency manager)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/MarekMejsner/dataapp.git
cd dataapp
```

2. Install dependencies using Composer:
```bash
composer install
```

This will install PHPSpreadsheet library required for Excel generation.

## Usage

Run the script from the command line:

```bash
php generate_sales_report.php
```

The script will generate an Excel file named `raport_sprzedazy_TOMEX_YYYY-MM-DD_HH-MM-SS.xlsx` in the current directory.

## Output Structure

The generated Excel file contains:

### 1. Podsumowanie TOMEX (Summary Sheet)
- Consolidated data across all operators
- Breakdown by product groups (TOMC, TOMK, TOMM, TOMS, TOMT)
- Summary by time period (Last Year, Last 4 Months)
- Operator performance comparison

### 2. Individual Operator Sheets (AR, TR, MŁ, GJ, MM, GS, WEB)
Each operator sheet includes:
- Detailed breakdown by product group
- Metrics for both time periods
- Subtotals and totals with margin calculations

## Data Customization

The sample data is defined in the `$salesData` array within `generate_sales_report.php`. To use your own data:

1. Replace the sample data with your actual data
2. Maintain the same structure:
```php
$salesData = [
    'OPERATOR_CODE' => [
        'zeszly_rok' => [
            'PRODUCT_GROUP' => [
                'wartosc_netto' => 0,
                'wartosc_netto_zwroty' => 0,
                'marza' => 0,
                'marza_procent' => 0
            ],
            // ... more product groups
        ],
        'ostatnie_4_miesiace' => [
            // ... same structure
        ]
    ],
    // ... more operators
];
```

3. You can also integrate this with a database by replacing the sample data array with database queries.

## Extending the Application

### Adding New Operators
Add the operator code to the `$operators` array and include its data in `$salesData`.

### Adding New Product Groups
Add the product group code to the `$productGroups` array and include its data for each operator.

### Adding New Time Periods
Add the period key to the `$periods` array and update the data structure accordingly.

## License

This project is open source and available for use and modification.
 
