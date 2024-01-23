# Allegro XLSM File Processing

## Overview

This Python script is designed to update an XLSM file from Allegro with product parameters fetched from an e-shop's API. It automates the task of updating the Excel file with specific product details, ensuring data accuracy and saving significant manual effort.

## Configuration

Before running the script, ensure the following variables are correctly set:

- **MAX_ROWS_TO_PROCESS**: The maximum number of rows to be processed in the Excel file.
- **PRODUCT_ID_COLUMN**: The Excel column that contains the product IDs.
- **PARAMETER_COLUMNS**: A dictionary mapping product parameter IDs to their corresponding Excel column letters.
- **STARTING_ROW**:  The row number in the Excel file where processing should begin.
- **API_TIMEOUT**: Timeout in seconds for the API requests.
- **API_URL**: URL of the e-shop's API endpoint.
- **API_HEADERS**: Headers for the API request, including the API key.