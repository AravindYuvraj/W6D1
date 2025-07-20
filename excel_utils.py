# excel_utils.py

import pandas as pd
from typing import Dict
from rapidfuzz import fuzz

# Synonym dictionary
COLUMN_SYNONYMS = {
    "amt": "amount",
    "revenue_amt": "revenue",
    "qty": "quantity",
    "product_id": "product",
    "cust id": "customer_id",
    "customerid": "customer_id",
    "rev": "revenue",
    "reg" : "region"
}

def normalize_column_name(name: str) -> str:
    name = name.lower().strip().replace("-", "_").replace(" ", "_")

    # Replace synonyms
    for key, val in COLUMN_SYNONYMS.items():
        if fuzz.ratio(name, key) > 85:
            return val

    return name


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_column_name(col) for col in df.columns]
    return df


def read_excel_file(file_path: str) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(file_path)
    sheets = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)
        sheets[sheet_name] = clean_dataframe(df)
    return sheets


def preview_excel_sheets(sheets: Dict[str, pd.DataFrame], rows: int = 5) -> str:
    preview = ""
    for name, df in sheets.items():
        preview += f"### Sheet: {name}\n"
        preview += df.head(rows).to_markdown(index=False)
        preview += "\n\n"
    return preview
