from typing import List, Optional
import pandas as pd
from langchain.tools import tool


# Pivot table tool for advanced summarization
@tool("pivot_table", return_direct=False)
def pivot_table(
    df: pd.DataFrame,
    index: list,
    values: str,
    aggfunc: str = "sum",
    columns: list = None,
    fill_value: float = 0.0
) -> pd.DataFrame:
    """
    Create a pivot table from a DataFrame.
    Args:
        df: DataFrame to pivot
        index: List of columns to use as index (rows)
        values: Column to aggregate
        aggfunc: Aggregation function (sum, mean, count, min, max)
        columns: List of columns to use as columns (optional)
        fill_value: Value to replace missing values (default 0.0)
    Returns:
        Pivoted DataFrame
    """
    try:
        pt = pd.pivot_table(
            df,
            index=index,
            values=values,
            aggfunc=aggfunc,
            columns=columns,
            fill_value=fill_value
        )
        return pt.reset_index()
    except Exception as e:
        print(f"Error creating pivot table: {e}")
        return pd.DataFrame()


# Aggregation tool for groupby operations
@tool("aggregate_data", return_direct=False)
def aggregate_data(
    df: pd.DataFrame,
    groupby_columns: List[str],
    agg_column: str,
    agg_func: str = "sum"
) -> pd.DataFrame:
    """
    Aggregate a DataFrame by groupby columns and aggregation function.
    Args:
        df: DataFrame to aggregate
        groupby_columns: List of columns to group by
        agg_column: Column to aggregate
        agg_func: Aggregation function (sum, mean, count, min, max)
    Returns:
        Aggregated DataFrame
    """
    if not all(col in df.columns for col in groupby_columns + [agg_column]):
        print("One or more columns not found in DataFrame.")
        return pd.DataFrame()
    try:
        grouped = df.groupby(groupby_columns)[agg_column]
        if agg_func == "sum":
            result = grouped.sum().reset_index()
        elif agg_func == "mean":
            result = grouped.mean().reset_index()
        elif agg_func == "count":
            result = grouped.count().reset_index()
        elif agg_func == "min":
            result = grouped.min().reset_index()
        elif agg_func == "max":
            result = grouped.max().reset_index()
        else:
            print(f"Unsupported aggregation function: {agg_func}")
            return pd.DataFrame()
        return result
    except Exception as e:
        print(f"Error in aggregation: {e}")
        return pd.DataFrame()


# Custom Excel tools for LangChain agent
@tool("filter_data", return_direct=False)
def filter_data(df: pd.DataFrame, column: str, operator: str, value) -> pd.DataFrame:
    """
    Filter a DataFrame by column, operator, and value.
    Supported operators: ==, !=, >, <, >=, <=, in, not in
    Returns filtered DataFrame or empty DataFrame if error.
    Args:
        df: The DataFrame to filter.
        column: The column name to filter on.
        operator: The comparison operator (==, !=, >, <, >=, <=, in, not in).
        value: The value to compare against.
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


