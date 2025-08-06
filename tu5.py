import os

# ─── Before you open the two connections, ask for output path ────────────────

export_dir  = input("Enter folder to save the comparison report: ").strip()
base_name   = input("Enter base filename (without .xlsx): ").strip()
os.makedirs(export_dir, exist_ok=True)
out_xl      = os.path.join(export_dir, f"{base_name}.xlsx")
writer      = pd.ExcelWriter(out_xl, engine="openpyxl")

# ─── Then open your two connections as cfg1/cfg2, usr1/..., pw2/..., sql1/sql2 ─────────

conn1 = connect_to_oracle(cfg1["host"], cfg1["port"], cfg1["svc"], usr1, pw1)
conn2 = connect_to_oracle(cfg2["host"], cfg2["port"], cfg2["svc"], usr2, pw2)

label1 = cfg1["label"]
label2 = cfg2["label"]

# ─── Loop through sheets_set exactly as before but using label1/label2 ─────────

for sheet in sheets_set:
    print(f"\n▶ Comparing {label1} → {label2} on sheet '{sheet}'")

    # 1) fetch & concat from first DB
    df1 = query_to_df(conn1, queries[sheet])
    cols1 = concat_map.get(sheet)
    cols1 = cols1(df1) if callable(cols1) else cols1
    df1 = build_concat(df1, cols1)

    # 2) fetch & concat from second DB
    df2 = query_to_df(conn2, queries[sheet])
    cols2 = concat_map.get(sheet)
    cols2 = cols2(df2) if callable(cols2) else cols2
    df2 = build_concat(df2, cols2)

    # 3) compare
    only_2, only_1 = compare_mismatches(df1, df2)

    # 4) prepare full-row details
    use1 = cols1 or [c for c in df1.columns if c!="Concatenated"]
    use2 = cols2 or [c for c in df2.columns if c!="Concatenated"]

    # 4a) rows only in second DB
    m2 = df2["Concatenated"].isin(only_2)
    df2_only = df2.loc[m2, use2].copy()
    df2_only.insert(0, "MismatchType", f"{label2} only")
    df2_only["DB1_Concat"] = ""
    df2_only["DB2_Concat"] = df2.loc[m2, "Concatenated"].values

    # 4b) rows only in first DB
    m1 = df1["Concatenated"].isin(only_1)
    df1_only = df1.loc[m1, use1].copy()
    df1_only.insert(0, "MismatchType", f"{label1} only")
    df1_only["DB2_Concat"] = ""
    df1_only["DB1_Concat"] = df1.loc[m1, "Concatenated"].values

    # 4c) combine & write detailed mismatches
    detail = pd.concat([df2_only, df1_only], ignore_index=True)
    detail.to_excel(
        writer,
        sheet_name=f"{sheet}_MismatchDetails",
        index=False
    )

    # 5) duplicates in each DB
    dup1 = find_duplicates(df1)
    if not dup1.empty:
        dup1.to_excel(writer,
                      sheet_name=f"{sheet}_{label1}_Dupes",
                      index=False)

    dup2 = find_duplicates(df2)
    if not dup2.empty:
        dup2.to_excel(writer,
                      sheet_name=f"{sheet}_{label2}_Dupes",
                      index=False)

# ─── Finalize ───────────────────────────────────────────────────────────────────

writer.save()
conn1.close()
conn2.close()
print(f"\n✅ DB-vs-DB report written to:\n   {out_xl}")