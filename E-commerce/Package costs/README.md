# Shipping Cost Processor Script
## Overview

This script is designed to process Excel or CSV files containing shipment data.
It integrates with an external API (IdoSell API) to fetch shipping costs for packages and
updates the provided files with this information for analytical reasons.

## Requirements

- **Python 3.x**
- **requests** library
- **openpyxl** library for handling Excel files
- **csv** module for handling CSV files
- Access to a specific API with shipping cost information

## Usage

The script is executed from the command line with the following arguments:

- **-f** or **--file** : The path to Excel or CSV file to be processed.
- **-c** or **--courier** : The courier type, which can be one of "DPD", "InPost" or "GLS".

Example: 

`
python package_costs.py -f path/to/file.xlsx -c DPD
`