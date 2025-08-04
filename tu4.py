import pandas as pd

def build_concat(df: pd.DataFrame, cols: list[str] | None = None) -> pd.DataFrame:
    """
    Turn every cell into a clean text string (dropping .0 on whole floats,
    converting NaN → ""), then concatenate the desired columns into a new
    'Concatenated' column.
    
    Parameters
    ----------
    df : pandas.DataFrame
        The raw DataFrame returned from your database query.
    cols : list of str, optional
        The names of the columns to include in the concatenation, in order.
        If None, all columns in `df` will be used.
    
    Returns
    -------
    pandas.DataFrame
        A new DataFrame containing only:
          • the specified `cols` (converted to text per-cell)
          • plus a 'Concatenated' column where each row’s values are joined.
    """
    # 1) Decide which columns to concatenate
    use_cols = cols if cols is not None else df.columns.tolist()

    # 2) Make a working copy of just those columns
    df2 = df[use_cols].copy()

    # 3) Define per‐cell formatter
    def to_text(x):
        # NaN / None → empty string
        if pd.isna(x):
            return ""
        # Floats: drop the ".0" if it’s a whole number
        if isinstance(x, float):
            return str(int(x)) if x.is_integer() else str(x)
        # Ints: just convert to string
        if isinstance(x, int):
            return str(x)
        # Everything else → string
        return str(x)

    # 4) Apply formatter to every cell
    df2 = df2.applymap(to_text)

    # 5) Build the concatenation key
    df2["Concatenated"] = df2.agg("".join, axis=1)

    return df2