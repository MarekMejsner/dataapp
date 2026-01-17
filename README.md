"# Dataapp

Sales data reporting tool for generating detailed Excel reports.

## Features

- Generates comprehensive Excel sales reports with multiple sheets
- Separate sheet for each operator (AR, TR, MŁ, GJ, MM, GS, WEB)
- Consolidated summary comparing all operators
- Detailed breakdowns by product groups (TOMC, TOMK, TOMM, TOMS, TOMT)
- Yearly and last 4 months data comparison
- Professional formatting with bold headers and proper monetary formatting

## Installation

1. Install dependencies using Composer:
```bash
composer install
```

## Usage

Run the script to generate the Excel report:
```bash
php generate_sales_report.php
```

The generated Excel file will be saved to: `exports/sales_data.xlsx`

## Report Structure

### Summary Sheet (Podsumowanie)
- Comparison of all operators for yearly sales
- Comparison of all operators for last 4 months sales
- Aggregated totals and margins

### Individual Operator Sheets
Each operator has a dedicated sheet containing:
- **Yearly Data (Zeszły Rok)**: Complete breakdown by product groups
- **Last 4 Months Data (Ostatnie 4 Miesiące)**: Recent performance by product groups

### Data Columns
- Grupa (Group name)
- Wartość netto (Net value)
- Wartość brutto (Gross value)
- Wartość marży (Margin value)
- Procent marży (Margin percentage)
- Liczba pozycji (Number of positions)
- Ilości (Quantities)

## Requirements

- PHP 8.0 or higher
- Composer
- PhpSpreadsheet library (automatically installed via Composer)" 
