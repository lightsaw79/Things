for sheet in sheets:
    print(f"\n▶ Comparing DB → Sheet '{sheet}'")

    # 1) Load the Excel sheet
    df_sheet = pd.read_excel(
        master_xl,
        sheet_name=sheet,
        engine="openpyxl",
        dtype=str
    )

    # 2) If this is your special 'ABC' sheet, drop columns A–C and keep from 'Concatenated' onward
    if sheet == "ABC":
        start = df_sheet.columns.get_loc("Concatenated")
        df_sheet = df_sheet.iloc[:, start:].copy()

    # 3) Remove any '…-deleted' versions before comparing
    if "Version" in df_sheet.columns:
        df_sheet = df_sheet[~df_sheet["Version"].str.endswith("-deleted", na=False)]

    # 4) Sanity check
    if "Concatenated" not in df_sheet.columns:
        raise KeyError(f"'{sheet}' missing required 'Concatenated' column.")

    # 5) Fetch DB data and build its concatenated key
    sql = SQL_QUERIES[sheet]
    df_db = query_to_df(conn, sql)

    # 5a) pick per‐sheet cols (None => all cols)
    cols = concat_map.get(sheet, None)
    cols = cols(df_db) if callable(cols) else cols

    df_db = build_concat(df_db, cols)

    # 6) Determine mismatches on the 'Concatenated' key
    only_db, only_sheet = compare_mismatches(df_sheet, df_db)

    # 7) Export full‐row mismatch details
    #    – use the same columns you concatenated
    used_cols = cols or [c for c in df_db.columns if c != "Concatenated"]

    db_rows    = df_db.loc[df_db["Concatenated"].isin(only_db), used_cols].copy()
    sheet_rows = df_sheet.loc[df_sheet["Concatenated"].isin(only_sheet), used_cols].copy()

    db_rows.insert(0, "MismatchType", "DB only")
    sheet_rows.insert(0, "MismatchType", "Sheet only")

    details = pd.concat([db_rows, sheet_rows], ignore_index=True)
    details.to_excel(
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