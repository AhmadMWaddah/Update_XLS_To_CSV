import os
import pandas as pd
import logging
from openpyxl import load_workbook


# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def preprocess_updated_data(file_path):
    """
    Preprocesses the updated data file (Excel format).

    Args:
        file_path (str): Path to the updated data file.

    Returns:
        pandas.DataFrame: Preprocessed data as a DataFrame.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        if file_path.endswith(('.xls', '.xlsx')):
            # Use pandas to read Excel file
            df = pd.read_excel(file_path, usecols=[0, 2], engine='openpyxl')  # Adjust engine for compatibility
        else:
            raise ValueError("Unsupported file format. Please use .xls or .xlsx")
    except Exception as e:
        logging.error(f"Error reading updated data file: {e}")
        raise

    # Rename and clean columns
    df.columns = ['Code', 'Qty']

    # Clean the 'Code' column: remove non-digits, keep first 5 digits
    df['Code'] = df['Code'].fillna('').astype(str).str.replace(r'\D+', '', regex=True).str[:5]
    df = df[df['Code'].str.len() == 5]  # Keep only valid 5-character codes

    logging.info(f"Processed Updated_Data.xlsx SKUs (first 10):\n{df['Code'].head(10)}")

    if df.empty:
        logging.warning("Updated data file is empty after preprocessing.")
    return df


def preprocess_exported_data(file_path):
    """
    Preprocesses the exported data file (CSV format).

    Args:
        file_path (str): Path to the exported data file.

    Returns:
        pandas.DataFrame: Preprocessed data as a DataFrame.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        logging.error(f"Error reading exported data file: {e}")
        raise

    # Clean SKU codes:
    df['Variant SKU'] = df['Variant SKU'].fillna('').astype(str).str.replace(r"^\'", '', regex=True)

    # Remove non-digit characters, keep first 5 digits like Excel processing
    df['Variant SKU'] = df['Variant SKU'].str.replace(r'\D+', '', regex=True).str[:5]

    logging.info(f"Processed Exported_Data.csv SKUs (first 10):\n{df['Variant SKU'].head(10)}")

    # Reset inventory quantities
    df['Variant Inventory Qty'] = pd.to_numeric(df['Variant Inventory Qty'], errors='coerce').fillna(0).astype(int)

    if df.empty:
        logging.warning("Exported data file is empty after preprocessing.")
    return df


def update_quantities(updated_data, exported_data):
    logging.info("Merging updated data with exported data...")
    merged_df = exported_data.merge(updated_data, left_on='Variant SKU', right_on='Code', how='left')

    # Log unmatched SKUs for debugging
    unmatched = merged_df[merged_df['Qty'].isna() & (merged_df['Variant SKU'] != '')]
    if not unmatched.empty:
        logging.warning(f"Unmatched SKUs (not found in updated data):\n{unmatched[['Variant SKU']]}")

    # Replace NaN in Qty with 0 and convert to int
    merged_df['Qty'] = merged_df['Qty'].fillna(0).astype(int)

    # Update Variant Inventory Qty for matching SKUs
    merged_df.loc[merged_df['Qty'] > 0, 'Variant Inventory Qty'] = merged_df['Qty']

    # Drop unnecessary columns
    merged_df.drop(['Code', 'Qty'], axis=1, inplace=True)

    logging.info("Inventory quantities updated successfully.")
    return merged_df
    

def update_published_column(df):
    """
    Updates the 'Published' column based on 'Variant Inventory Qty'.

    Args:
        df (pandas.DataFrame): DataFrame containing inventory data.

    Returns:
        pandas.DataFrame: Updated DataFrame with the 'Published' column updated.
    """
    df['Published'] = df['Variant Inventory Qty'].apply(
        lambda x: 'TRUE' if x > 0 else ('FALSE' if x == 0 else '')
    )
    return df


def main():
    """
    Main function to execute the script.
    """
    logging.info("Script started.")

    # Define file paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    updated_data_file = os.path.join(base_dir, 'Updated_Data.xlsx')
    exported_data_file = os.path.join(base_dir, 'Exported_Data.csv')
    output_file = os.path.join(base_dir, 'Exported_Data_Processed.csv')
    processed_data_file = os.path.join(base_dir, 'Updated_Data_Processed.xlsx')

    try:
        # Preprocess updated data
        logging.info("Preprocessing updated data file...")
        updated_df = preprocess_updated_data(updated_data_file)

        # Save the preprocessed updated data for auditing
        updated_df.to_excel(processed_data_file, index=False)
        logging.info(f"Processed updated data saved to: {processed_data_file}")

        # Preprocess exported data
        logging.info("Preprocessing exported data file...")
        exported_df = preprocess_exported_data(exported_data_file)

        # Update quantities and published column
        logging.info("Updating inventory quantities...")
        updated_exported_df = update_quantities(updated_df, exported_df)

        logging.info("Updating 'Published' column...")
        updated_exported_df = update_published_column(updated_exported_df)

        # Save the final processed data to CSV
        updated_exported_df.to_csv(output_file, index=False)
        logging.info(f"Final processed data saved to: {output_file}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        raise

    logging.info("Script completed successfully.")


if __name__ == '__main__':
    main()
