# Shipping Cost Processor Script
## Overview

This script is designed to process Excel or CSV files containing shipment data.
It integrates with an external API (IdoSell API) to fetch shipping costs for packages and
updates the provided files with this information for analytical reasons.

## Requirements
- Python 3.x
- **requests** library
- **openpyxl** library for handling Excel files
- **csv** module for handling CSV files
- Access to a specific API with shipping cost information