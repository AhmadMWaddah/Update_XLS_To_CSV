# Inventory Update Script

This script processes updated product inventory data and merges it with existing exported Shopify product data.

## Key Functionalities

- **Reads updated inventory data**: 
  - Loads updated product data from an Excel file (.xls or .xlsx)
  - Extracts and processes SKU codes ('Code') and quantities ('Qty')
  - Standardizes SKU format (5-digit numeric codes)

- **Reads exported Shopify data**:
  - Loads exported product data from a CSV file
  - Cleans and standardizes 'Variant SKU' to match updated data format
  - Prepares inventory quantities for updating

- **Data processing**:
  - Merges updated quantities with exported data based on matching SKUs
  - Identifies and logs unmatched SKUs for troubleshooting
  - Updates inventory quantities while preserving other product data

- **Publishing logic**:
  - Automatically updates 'Published' column based on inventory levels
  - Products with stock > 0 are set to 'TRUE'
  - Products with 0 stock are set to 'FALSE'

- **Output generation**:
  - Creates audit trail of processed data
  - Preserves all original columns from exported data
  - Maintains data integrity throughout the process

## File Structure

The script expects the following files in the same directory:

- Input Files:
  - `Updated_Data.xls` or `Updated_Data.xlsx`: Contains updated inventory data
  - `Exported_Data.csv`: Contains exported Shopify product data

- Output Files:
  - `Exported_Data_Processed.csv`: Final output with updated inventory and publishing status
  - `Updated_Data_Processed.xlsx`: Preprocessed version of the updated data for verification

## Features

- Comprehensive error handling and logging
- Data validation and cleaning
- Audit trail generation
- Non-destructive processing (preserves original files)
- Detailed logging for troubleshooting

## Dependencies

- Python 3.x
- Required packages:
  - pandas
  - openpyxl
  - os (standard library)
  - logging (standard library)

## Usage

1. Place input files in the same directory as the script
2. Run the script: `python Update_XLS_To_CSV.py`
3. Check logs for processing details
4. Use output files as needed

> Note: The script includes detailed logging at each processing step for troubleshooting and verification purposes.
