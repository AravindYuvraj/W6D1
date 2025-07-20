# excel_utils.py

import pandas as pd
from typing import Dict, Tuple
import os

def read_excel_file(file_path: str) -> Dict[str, pd.DataFrame]:
    """
    Reads an Excel file with multiple sheets.
    
    Returns:
        A dictionary where keys are sheet names and values are DataFrames.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    except Exception as e:
        raise RuntimeError(f"Error reading Excel file: {e}")

    return excel_data


def preview_excel_sheets(dataframes: Dict[str, pd.DataFrame], rows: int = 5) -> str:
    """
    Creates a markdown-style preview of the first few rows from each sheet.
    Useful for feeding into an LLM prompt or displaying in UI.

    Args:
        dataframes: Dict of {sheet_name: DataFrame}
        rows: Number of rows to preview

    Returns:
        A string with markdown tables
    """
    previews = []
    for sheet_name, df in dataframes.items():
        preview_str = f"### Sheet: {sheet_name}\n"
        if df.empty:
            preview_str += "_(Empty sheet)_\n"
        else:
            preview_str += df.head(rows).to_markdown(index=False) + "\n"
        previews.append(preview_str)

    return "\n\n".join(previews)


def get_sheet_names(file_path: str) -> list:
    """
    Return sheet names in the Excel file.
    """
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        return xls.sheet_names
    except Exception as e:
        raise RuntimeError(f"Unable to get sheet names: {e}")
