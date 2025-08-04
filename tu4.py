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
    
    
    
    
    
    
for sheet in sheets:
    print(f"\n▶ Comparing DB → Sheet '{sheet}'")

    # 1) Load the Excel sheet
    df_sheet = pd.read_excel(
        master_xl,
        sheet_name=sheet,
        engine="openpyxl",
        dtype=str
    )

    # 2) (Optional) Drop A–C / keep from 'Concatenated' onward for ABC
    if sheet == "ABC":
        start = df_sheet.columns.get_loc("Concatenated")
        df_sheet = df_sheet.iloc[:, start:].copy()

    # 3) Filter out any '…-deleted' versions
    if "Version" in df_sheet.columns:
        df_sheet = df_sheet[~df_sheet["Version"].str.endswith("-deleted", na=False)]

    # 4) Sanity check
    if "Concatenated" not in df_sheet.columns:
        raise KeyError(f"'{sheet}' missing required 'Concatenated' column.")

    # 5) Fetch DB data & build its concatenation
    sql    = SQL_QUERIES[sheet]
    df_db  = query_to_df(conn, sql)
    cfg    = concat_map.get(sheet)
    cols   = cfg(df_db) if callable(cfg) else cfg
    df_db  = build_concat(df_db, cols)

    # 6) Find mismatches on the 'Concatenated' key
    only_db, only_sheet = compare_mismatches(df_sheet, df_db)

    # 7) Prepare full‐row mismatch details with added columns
    #    – determine which original columns were concatenated
    used_cols = cols or [c for c in df_db.columns if c != "Concatenated"]

    # 7a) DB-only rows
    mask_db  = df_db["Concatenated"].isin(only_db)
    db_rows  = df_db.loc[mask_db, used_cols].copy()
    db_rows["DB_Concatenation"]    = df_db.loc[mask_db, "Concatenated"].values
    db_rows["Sheet_Concatenation"] = ""
    db_rows.insert(0, "MismatchType", "DB only")

    # 7b) Sheet-only rows
    mask_sh     = df_sheet["Concatenated"].isin(only_sheet)
    sheet_rows  = df_sheet.loc[mask_sh, used_cols].copy()
    sheet_rows["Sheet_Concatenation"] = df_sheet.loc[mask_sh, "Concatenated"].values
    sheet_rows["DB_Concatenation"]    = ""
    sheet_rows.insert(0, "MismatchType", "Sheet only")

    # 7c) Combine and write
    mismatch_details = pd.concat([db_rows, sheet_rows], ignore_index=True)
    mismatch_details.to_excel(
        writer,
        sheet_name=f"{sheet}_MismatchDetails",
        index=False
    )

    # 8) Write duplicate‐row sheets as before
    dup_s = find_duplicates(df_sheet)
    if not dup_s.empty:
        dup_s.to_excel(writer,
                       sheet_name=f"{sheet}_SheetDupes",
                       index=False)

    dup_d = find_duplicates(df_db)
    if not dup_d.empty:
        dup_d.to_excel(writer,
                       sheet_name=f"{sheet}_DBDupes",
                       index=False)