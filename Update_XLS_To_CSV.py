import os
import openpyxl
import pandas as pd
import re
import xlrd


def preprocess_updated_data(file_path):
    """
    Preprocesses the updated data file.

    Args:
        file_path (str): Path to the updated data file.

    Returns:
        pandas.DataFrame: Preprocessed data as a DataFrame.
    """

    # Determine the file type based on extension
    if file_path.endswith('.xls'):
        workbook = xlrd.open_workbook(file_path)
        # Convert xlrd sheet to pandas DataFrame
        sheet = workbook.sheet_by_index(0)
        data = [[sheet.cell_value(rowx, colx) for colx in [0, 2]] for rowx in
                range(1, sheet.nrows)]  # Select columns 0 and 2
        df = pd.DataFrame(data)
    elif file_path.endswith('.xlsx'):
        workbook = openpyxl.load_workbook(file_path)
        worksheet = workbook.active
        data = [[cell.value for cell in row[::2]] for row in
                worksheet.iter_rows(min_row=2, values_only=True)]  # Select every other column
        df = pd.DataFrame(data)
    else:
        raise ValueError("Unsupported file format. Please use .xls or .xlsx")

    # Rename columns
    df.columns = ['Code', 'Qty']

    # Clean SKU codes
    df['Code'] = df['Code'].astype(str).str.replace('\D+', '', regex=True)
    df['Code'] = df['Code'].str[:5]
    df = df[df['Code'].str.len() == 5]

    return df


def preprocess_exported_data(file_path):
    # Load the CSV file
    df = pd.read_csv(file_path)

    # Clean SKU codes
    df['Variant SKU'] = df['Variant SKU'].astype(str).str.replace(r"^\'", '', regex=True)  # Remove leading single quote

    # Convert non-zero values in 'Variant Inventory Qty' to 0
    df['Variant Inventory Qty'] = df['Variant Inventory Qty'].where(df['Variant Inventory Qty'].isnull(), 0)

    return df


def update_quantities(updated_data, exported_data):
    # Merge dataframes based on SKU codes
    merged_df = exported_data.merge(updated_data, left_on='Variant SKU', right_on='Code', how='left')

    # Update quantities only for SKUs with matching data
    merged_df.loc[merged_df['Qty'].notnull(), 'Variant Inventory Qty'] = merged_df['Qty']

    # Drop unnecessary columns
    merged_df.drop(['Code', 'Qty'], axis=1, inplace=True)

    return merged_df


def update_published_column(df):
    # Update 'Published' column based on 'Variant Inventory Qty'
    df['Published'] = df['Variant Inventory Qty'].apply(lambda x: 'TRUE' if x > 0 else ('FALSE' if x == 0 else ''))
    return df


def main():
    # Get the absolute path of the script
    direction_location = os.path.dirname(os.path.abspath(__file__))

    # Files Paths.
    updated_data_file = os.path.join(direction_location, 'Updated_Data.xls')
    exported_data_file = os.path.join(direction_location, 'Exported_Data.csv')
    output_file = os.path.join(direction_location, 'Exported_Data_Processed.csv')
    processed_data_file = os.path.join(direction_location, 'Updated_Data_Processed.xls')

    updated_df = preprocess_updated_data(updated_data_file)

    # Save processed data to a new file
    updated_df.to_excel(processed_data_file, index=False)

    exported_df = preprocess_exported_data(exported_data_file)
    updated_exported_df = update_quantities(updated_df, exported_df)
    updated_exported_df.to_csv(output_file, index=False)


if __name__ == '__main__':
    main()
