"# Dataapp

## Sales Data CSV Generator

This repository contains a PHP script that generates a CSV file with summarized sales data for operators and their product groups.

### Usage

To generate the sales data CSV file, run:

```bash
php generate_sales_data.php
```

This will create a file `exports/sales_data.csv` containing sales data with the following structure:

### CSV Structure

**Columns:**
- `operator` - The operator name
- `period` - Time period ('zeszly_rok' or 'ostatnie_4_miesiace')
- `product_group` - Product group code (TOMC, TOMK, TOMM, TOMS, TOMT)
- `net_value` - Net sales value
- `number_of_positions` - Number of sales positions
- `returns_net_value` - Net value of returns
- `margin` - Sales margin
- `margin_percentage` - Margin as a percentage of net value
- `quantity` - Quantity sold

### Requirements

- PHP 7.0 or higher

### Directory Structure

```
dataapp/
├── generate_sales_data.php  # Main script
├── exports/                  # Output directory
│   └── sales_data.csv       # Generated CSV file
└── README.md                # This file
```" 
