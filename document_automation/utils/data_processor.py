"""
DataProcessor - Transform and analyze tabular data for report generation.
"""

import pandas as pd


class DataProcessor:
    """Provides data transformation and statistical summary utilities."""

    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()

    def get_summary_statistics(self) -> pd.DataFrame:
        """Return descriptive statistics for numeric columns."""
        numeric_df = self.df.select_dtypes(include="number")
        if numeric_df.empty:
            return pd.DataFrame()
        return numeric_df.describe().round(2)

    def get_column_info(self) -> pd.DataFrame:
        """Return column name, dtype, non-null count, and null count."""
        info = pd.DataFrame({
            "Column": self.df.columns,
            "Type": [str(dt) for dt in self.df.dtypes],
            "Non-Null": self.df.notnull().sum().values,
            "Null": self.df.isnull().sum().values,
            "Unique": [self.df[col].nunique() for col in self.df.columns],
        })
        return info

    def filter_rows(self, column: str, operator: str, value) -> "DataProcessor":
        """
        Filter rows by a condition. Returns a new DataProcessor.
        Operators: ==, !=, >, <, >=, <=, contains, startswith
        """
        ops = {
            "==": lambda s, v: s == v,
            "!=": lambda s, v: s != v,
            ">": lambda s, v: s > v,
            "<": lambda s, v: s < v,
            ">=": lambda s, v: s >= v,
            "<=": lambda s, v: s <= v,
            "contains": lambda s, v: s.astype(str).str.contains(str(v), case=False, na=False),
            "startswith": lambda s, v: s.astype(str).str.startswith(str(v), na=False),
        }
        if operator not in ops:
            raise ValueError(f"Unknown operator '{operator}'. Use: {', '.join(ops.keys())}")
        mask = ops[operator](self.df[column], value)
        return DataProcessor(self.df[mask])

    def group_summary(self, group_col: str, agg_col: str, agg_func: str = "mean") -> pd.DataFrame:
        """Group by a column and aggregate another column."""
        result = self.df.groupby(group_col)[agg_col].agg(agg_func).reset_index()
        result.columns = [group_col, f"{agg_col}_{agg_func}"]
        return result

    def add_computed_column(self, new_col: str, expression: str) -> "DataProcessor":
        """
        Add a column using a pandas eval expression.
        Example: processor.add_computed_column('total', 'price * quantity')
        """
        self.df[new_col] = self.df.eval(expression)
        return self

    def top_n(self, column: str, n: int = 10, ascending: bool = False) -> pd.DataFrame:
        """Return the top N rows sorted by a column."""
        return self.df.nlargest(n, column) if not ascending else self.df.nsmallest(n, column)

    def to_dataframe(self) -> pd.DataFrame:
        """Return the underlying DataFrame."""
        return self.df
