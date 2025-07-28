import pandas as pd
import oracledb
import getpass

# ─── Helpers ────────────────────────────────────────────────────────────────────

def connect_to_oracle(host, port, service, user, pw):
    dsn = oracledb.makedsn(host, port, service_name=service)
    return oracledb.connect(user=user, password=pw, dsn=dsn)

def query_to_df(conn, sql):
    with conn.cursor() as cur:
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        rows = cur.fetchall()
    return pd.DataFrame(rows, columns=cols)

def build_concat(df, cols=None):
    df = df.fillna("").astype(str)
    if cols is None:
        cols = df.columns.tolist()
    df["Concatenated"] = df[cols].apply(lambda r: "".join(r.values), axis=1)
    return df

def compare_mismatches(left_df, right_df):
    L = left_df[["Concatenated"]].dropna().drop_duplicates().astype(str)
    R = right_df[["Concatenated"]].dropna().drop_duplicates().astype(str)
    merged = L.merge(R, on="Concatenated", how="outer", indicator=True)
    only_right = merged[merged["_merge"]=="right_only"]["Concatenated"].tolist()
    only_left  = merged[merged["_merge"]=="left_only" ]["Concatenated"].tolist()
    return only_right, only_left

def find_duplicates(df, col="Concatenated"):
    vc = df[col].value_counts()
    keys = vc[vc>1].index.tolist()
    return df[df[col].isin(keys)].copy()

# ─── 1) Select comparison mode ───────────────────────────────────────────────────

print("Select comparison mode:")
print("  1: Database vs Datasheet")
print("  2: Database vs Database")
mode = input("Enter 1 or 2: ").strip()

# ─── Common environment definitions ──────────────────────────────────────────────

envs = {
    1: {"label":"SIT_STG", "host":"NYKDSR000007912.intranet.barcapint.com", "port":1523, "svc":"TTMUS02P"},
    2: {"label":"SIT_CDS", "host":"NYKDSR000007912.intranet.barcapint.com", "port":1523, "svc":"TTMUS02P"},
    3: {"label":"UAT_STG","host":"isamusatdb.barcapint.com",        "port":1523, "svc":"TTMUS01P"},
    4: {"label":"UAT_CDS","host":"isamusatdb.barcapint.com",        "port":1523, "svc":"TTMUS01P"},
    5: {"label":"PROD",   "host":"your.prod.host.company.com",      "port":1521, "svc":"PROD_SVC"}
}

# ─── Mode 1: DB vs Datasheet ─────────────────────────────────────────────────────

if mode == "1":
    # 1a) pick environment
    print("\nSelect database environment:")
    for i,e in envs.items():
        print(f"  {i}: {e['label']}")
    cfg = envs[int(input("Enter 1–5: ").strip())]
    usr = input(f"{cfg['label']} username: ")
    pw  = getpass.getpass(f"{cfg['label']} password: ")

    # 1b) load datasheet and pick sheets
    master_xl = input("\nPath to master datasheet (.xlsx): ").strip()
    xls = pd.ExcelFile(master_xl, engine="openpyxl")
    # prompt user which sheets, 0=all
    print("\nWhich sheets to compare?")
    print("  0: All sheets")
    for idx,name in enumerate(xls.sheet_names, start=1):
        print(f"  {idx}: {name}")
    sel = input("Enter 0 or comma-separated numbers: ").strip()
    if sel == "0":
        sheets = xls.sheet_names
    else:
        nums = [int(x) for x in sel.split(",") if x.strip().isdigit()]
        sheets = [xls.sheet_names[n-1] for n in nums]

    # 1c) hard-coded SQL per sheet
    SQL_QUERIES = {
        "Scales":     "SELECT * FROM SAM_GLOBAL_CLASS.SAM_ML_RISK_SCORE_MAP",
        "Thresholds": "SELECT * FROM SAM_GLOBAL_CLASS.SAM_TRANSACTION_SYSTEM_CD_MAP",
        # add more sheets & their SQL here...
    }

    # 1d) concatenation rules per sheet
    concat_map = {
        # skip first two DB cols for Thresholds
        "Thresholds": lambda df: df.columns[2:].tolist(),
        # define for others or leave out to use all cols
    }

    # 1e) run comparisons
    conn = connect_to_oracle(cfg["host"], cfg["port"], cfg["svc"], usr, pw)
    out_xl = master_xl.replace(".xlsx", "_db_vs_sheet.xlsx")
    writer = pd.ExcelWriter(out_xl, engine="openpyxl")

    for sheet in sheets:
        if sheet not in SQL_QUERIES:
            raise KeyError(f"No SQL defined for '{sheet}'.")
        sql = SQL_QUERIES[sheet]

        print(f"\n▶ Comparing DB → Sheet '{sheet}'")
        df_sheet = pd.read_excel(master_xl, sheet_name=sheet, engine="openpyxl", dtype=str)
        if "Concatenated" not in df_sheet.columns:
            raise KeyError(f"'{sheet}' missing 'Concatenated' column.")

        df_db = query_to_df(conn, sql)
        # pick cols for concat
        cols = (concat_map[sheet](df_db) if sheet in concat_map else None)
        df_db = build_concat(df_db, cols)

        only_db, only_sheet = compare_mismatches(df_sheet, df_db)
        # write mismatches
        if not only_db and not only_sheet:
            pd.DataFrame([{"Result":"All rows match"}]).to_excel(
                writer, sheet_name=f"{sheet}_Mismatches", index=False)
            print("  ✔️ No mismatches")
        else:
            rows = ([{"Source":"DB only",    "Concatenated":v} for v in only_db] +
                    [{"Source":"Sheet only", "Concatenated":v} for v in only_sheet])
            pd.DataFrame(rows).to_excel(writer, sheet_name=f"{sheet}_Mismatches", index=False)
            print(f"  ⚠️ {len(only_db)} only in DB, {len(only_sheet)} only in Sheet")

        # duplicates
        ds = find_duplicates(df_sheet)
        if not ds.empty:
            ds.to_excel(writer, sheet_name=f"{sheet}_SheetDupes", index=False)
        dd = find_duplicates(df_db)
        if not dd.empty:
            dd.to_excel(writer, sheet_name=f"{sheet}_DBDupes", index=False)

    writer.save()
    conn.close()
    print(f"\n✅ Report: {out_xl}")

# ─── Mode 2: DB vs DB ─────────────────────────────────────────────────────────────

elif mode == "2":
    # 2a) first DB
    print("\n--- First database ---")
    for i,e in envs.items():
        print(f"  {i}: {e['label']}")
    cfg1 = envs[int(input("Enter 1–5: ").strip())]
    usr1 = input(f"{cfg1['label']} username: ")
    pw1  = getpass.getpass(f"{cfg1['label']} password: ")
    sql1 = input("\nEnter SQL query for first DB:\n").strip()

    # 2b) second DB
    print("\n--- Second database ---")
    for i,e in envs.items():
        print(f"  {i}: {e['label']}")
    cfg2 = envs[int(input("Enter 1–5: ").strip())]
    usr2 = input(f"{cfg2['label']} username: ")
    pw2  = getpass.getpass(f"{cfg2['label']} password: ")
    sql2 = input("\nEnter SQL query for second DB:\n").strip()

    # 2c) connect & fetch
    conn1 = connect_to_oracle(cfg1["host"], cfg1["port"], cfg1["svc"], usr1, pw1)
    conn2 = connect_to_oracle(cfg2["host"], cfg2["port"], cfg2["svc"], usr2, pw2)
    df1 = build_concat(query_to_df(conn1, sql1))
    df2 = build_concat(query_to_df(conn2, sql2))
    conn1.close()
    conn2.close()

    # 2d) compare
    only_2, only_1 = compare_mismatches(df1, df2)

    # 2e) output
    out_xl = "db_vs_db_comparison.xlsx"
    writer = pd.ExcelWriter(out_xl, engine="openpyxl")

    if not only_2 and not only_1:
        pd.DataFrame([{"Result":"All rows match"}]).to_excel(
            writer, sheet_name="Mismatches", index=False)
        print("\n✔️ No mismatches between DBs")
    else:
        rows = ([{"Source":f"{cfg2['label']} only", "Concatenated":v} for v in only_2] +
                [{"Source":f"{cfg1['label']} only", "Concatenated":v} for v in only_1])
        pd.DataFrame(rows).to_excel(writer, sheet_name="Mismatches", index=False)
        print(f"\n⚠️ {len(only_2)} rows only in {cfg2['label']}, {len(only_1)} only in {cfg1['label']}")

    # duplicates in each DB
    dup1 = find_duplicates(df1)
    if not dup1.empty:
        dup1.to_excel(writer, sheet_name=f"{cfg1['label']}_Dupes", index=False)
    dup2 = find_duplicates(df2)
    if not dup2.empty:
        dup2.to_excel(writer, sheet_name=f"{cfg2['label']}_Dupes", index=False)

    writer.save()
    print(f"\n✅ DB vs DB report: {out_xl}")

else:
    print("Invalid mode selected. Exiting.")
