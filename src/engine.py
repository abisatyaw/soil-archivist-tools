import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook


def extract_column_from_sheet(file_path: str, sheet_name: str, column_names: list):
    """
    Extracts a column from a specific sheet in an Excel file.

    Args:
        file_path (str): Path to the Excel file.
        sheet_name (str): Sheet name to read.
        column_name (str): Column name to extract.

    Returns:
        dict: Dictionary of column name to list of values.
    
    Raises:
        ValueError: If the sheet or column is not found.
        FileNotFoundError: If the file does not exist.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

        missing_columns = [col for col in column_names if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Columns not found in sheet '{sheet_name}': {missing_columns}")
        
        extracted_data = {
            col: df[col].dropna().tolist()
            for col in column_names
        }

        return extracted_data

    except FileNotFoundError:
        raise FileNotFoundError(f"File not found: {file_path}")
    
    except ValueError as e:
        raise e

    except Exception as e:
        raise RuntimeError(f"Unexpected error: {e}")
    
def process_excel(input_file_path: str, column_map: dict, reference: dict):
    """
    Write data to a new Excel file.
    lithology
    ~ Bore column (kiri) => Name column (kanan)
    ~ Depth1 column (kiri) => Depth top column (kanan)
    ~ Depth2 column (kiri) => Depth bottom column (kanan)
    ~ Keyword column (kiri) => Lithology column (kanan)
    ~ comment column (kiri) => Remarks lithology (kanan)

    Args:
    - output_file (str): Path of the output Excel file.
    - data (list): Data to write to the Excel sheet.
    - sheet_name (str): Name of the sheet to write to.
    - column_name (str): Name of the column in the sheet.
    """
    try:
        print(column_map.keys())
        data_dictionary = extract_column_from_sheet(input_file_path, sheet_name="Lithology", column_names=column_map.keys())
        df = pd.DataFrame(data_dictionary)
        df.rename(columns=column_map, inplace=True)

        write_excel_table(df, output_path=reference['output_path'], sheet_name=reference['sheet_name'],)

    except Exception as e:
        raise RuntimeError(f"Unexpected error: {e}")
    
def write_excel_table(df, output_path="output1.xlsx", sheet_name="Sheet1", table_name="Table1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    max_row = ws.max_row
    max_col = ws.max_column
    start_cell = "A1"
    end_col = ws.cell(row=1, column=max_col).column_letter
    table_range = f"A1:{end_col}{max_row}"
    table = Table(displayName=table_name, ref=table_range)
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    wb.save(output_path)
    print("Written to REAL")

def lithology_extract():
    input_file = "src\\203935 Grimsby Riverside - Soil data 1.xlsx"
    input_columns = ["Bore", "Depth1", "Depth2", "Keyword", "Comment"]
    output_columns = ["* Name [-]", "Depth top [m]", "Depth bottom [m]", "Lithology [-]", "Remarks lithology [-]"]
    column_map = dict(zip(input_columns, output_columns))

    reference = {
        "sheet_name":"INPUT Lithology",
        "output_path":"Lithology_Input_Template_rev2.xlsx"
    }
    try:
        process_excel(input_file, column_map, reference)
    except Exception as e:
        print(f"Error: {e}")

def borehole_extract():
    input_file = "src\\203935 Grimsby Riverside - Soil data 1.xlsx"
    input_columns = []
    output_columns = []
    column_map = dict(zip(input_columns, output_columns))

    reference = {
        "sheet_name":"INPUT Borehole",
        "output_path":"Borehole_Base_input_Template_rev2.xlsx"
    }
    try:
        process_excel(input_file, column_map, reference)
    except Exception as e:
        print(f"Error: {e}")

#############################
if __name__ == "__main__":
    test()
    print("complete")