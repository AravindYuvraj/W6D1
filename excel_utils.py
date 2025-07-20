
# excel_utils.py

import pandas as pd
from typing import Dict, Tuple
import os

def read_excel_file(file_path: str, chunk_size: int = 10000) -> Dict[str, pd.DataFrame]:
    """
    Reads an Excel file with multiple sheets, handling large files in chunks.
    Returns a dict of {sheet_name: DataFrame}.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        excel_data = {}
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        for sheet in xls.sheet_names:
            try:
                # Try to read in chunks if possible (for CSV, not Excel, but keep for future)
                df = pd.read_excel(xls, sheet_name=sheet, engine='openpyxl')
                if df.shape[0] > chunk_size:
                    print(f"Warning: Sheet '{sheet}' is large ({df.shape[0]} rows). Consider processing in chunks.")
                excel_data[sheet] = df
            except Exception as e:
                print(f"Error reading sheet '{sheet}': {e}")
                excel_data[sheet] = pd.DataFrame()  # Empty DataFrame for failed sheets
        return excel_data
    except Exception as e:
        raise RuntimeError(f"Error reading Excel file: {e}")


def preview_excel_sheets(dataframes: Dict[str, pd.DataFrame], rows: int = 5) -> str:
    """
    Efficiently preview the first few rows from each sheet (markdown format).
    """
    previews = []
    for sheet_name, df in dataframes.items():
        preview_str = f"### Sheet: {sheet_name}\n"
        if df.empty or df.shape[0] == 0:
            preview_str += "_(Empty or unreadable sheet)_\n"
        else:
            try:
                preview_str += df.head(rows).to_markdown(index=False) + "\n"
            except Exception as e:
                preview_str += f"_(Error previewing: {e})_\n"
        previews.append(preview_str)
    return "\n\n".join(previews)


def get_sheet_names(file_path: str) -> list:
    """
    Return sheet names in the Excel file, with error handling.
    """
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        return xls.sheet_names
    except Exception as e:
        print(f"Unable to get sheet names: {e}")
        return []

def filter_data(df: pd.DataFrame, column: str, operator: str, value) -> pd.DataFrame:
    """
    Filter a DataFrame by column, operator, and value.
    Supported operators: ==, !=, >, <, >=, <=, in, not in
    Returns filtered DataFrame or empty DataFrame if error.
    """
    if column not in df.columns:
        print(f"Column '{column}' not found in DataFrame.")
        return pd.DataFrame()
    try:
        if operator == '==':
            return df[df[column] == value]
        elif operator == '!=':
            return df[df[column] != value]
        elif operator == '>':
            return df[df[column] > value]
        elif operator == '<':
            return df[df[column] < value]
        elif operator == '>=':
            return df[df[column] >= value]
        elif operator == '<=':
            return df[df[column] <= value]
        elif operator == 'in':
            return df[df[column].isin(value if isinstance(value, list) else [value])]
        elif operator == 'not in':
            return df[~df[column].isin(value if isinstance(value, list) else [value])]
        else:
            print(f"Unsupported operator: {operator}")
            return pd.DataFrame()
    except Exception as e:
        print(f"Error filtering data: {e}")
        return pd.DataFrame()