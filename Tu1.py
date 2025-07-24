import pandas as pd
import oracledb
import getpass

def connect_to_oracle(host, port, service, user, pw):
    dsn = oracledb.makedsn(host, port, service_name=service)
    return oracledb.connect(user=user, password=pw, dsn=dsn)

def query_to_df(conn, sql):
    with conn.cursor() as cur:
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        rows = cur.fetchall()
    return pd.DataFrame(rows, columns=cols)

def build_db_concat(df_db, sheet_name, concat_map):
    """
    Build 'Concatenated' based on concat_map:
     - if concat_map[sheet_name] is a list: use those cols
     - if it's a callable: call it with df_db to get list
     - else default to all df_db.columns
    """
    df = df_db.fillna("").astype(str)
    if sheet_name in concat_map:
        sel = concat_map[sheet_name]
        if callable(sel):
            cols = sel(df)
        else:
            cols = sel
    else:
        cols = df.columns.tolist()
    df["Concatenated"] = df[cols].apply(lambda r: "".join(r.values), axis=1)
    return df

def compare_mismatches(df_sheet, df_db):
    left  = df_sheet[["Concatenated"]].dropna().astype(str).drop_duplicates()
    right = df_db[["Concatenated"]].drop_duplicates()
    merged = left.merge(right, on="Concatenated", how="outer", indicator=True)
    only_db    = merged[merged["_merge"]=="right_only"]["Concatenated"].tolist()
    only_sheet = merged[merged["_merge"]=="left_only"]["Concatenated"].tolist()
    return only_db, only_sheet

def find_duplicates(df, col="Concatenated"):
    vc = df[col].value_counts()
    keys = vc[vc>1].index.tolist()
    return df[df[col].isin(keys)].copy()

# ─── 1) Environment selection ────────────────────────────────────────────────────
envs = {
    1: {"label":"SIT_STG", "host":"NYKDSR000007912.intranet.barcapint.com", "port":1523, "svc":"TTMUS02P"},
    2: {"label":"SIT_CDS", "host":"NYKDSR000007912.intranet.barcapint.com", "port":1523, "svc":"TTMUS02P"},
    3: {"label":"UAT_STG","host":"isamusatdb.barcapint.com",        "port":1523, "svc":"TTMUS01P"},
    4: {"label":"UAT_CDS","host":"isamusatdb.barcapint.com",        "port":1523, "svc":"TTMUS01P"},
    5: {"label":"PROD",   "host":"your.prod.host.company.com",      "port":1521, "svc":"PROD_SVC"}
}
print("Select environment:")
for i, e in envs.items():
    print(f"  {i}: {e['label']}")
cfg = envs[int(input("Enter 1–5: ").strip())]
user = input(f"{cfg['label']} user: ")
pw   = getpass.getpass(f"{cfg['label']} password: ")

# ─── 2) Excel & sheet choice ─────────────────────────────────────────────────────
master_xl = input("\nPath to master datasheet (.xlsx): ").strip()
xls       = pd.ExcelFile(master_xl, engine="openpyxl")

print("\nWhich sheets to compare?")
print("  0: All sheets")
for idx, name in enumerate(xls.sheet_names, start=1):
    print(f"  {idx}: {name}")
sel = input("Enter 0 or comma-separated numbers (e.g. 1,3): ").strip()
if sel=="0":
    sheets = xls.sheet_names
else:
    nums = [int(x) for x in sel.split(",") if x.strip().isdigit()]
    sheets = [xls.sheet_names[n-1] for n in nums]

# ─── 3) Per-sheet SQL prompts ─────────────────────────────────────────────────────
queries = {}
for sh in sheets:
    queries[sh] = input(f"\nSQL for '{sh}':\n").strip()

# ─── 4) Define concatenation rules ───────────────────────────────────────────────
concat_map = {
    # For Thresholds, skip the first two DB columns, use the rest:
    "Thresholds": lambda df: df.columns[2:].tolist(),
    # For other sheets, replace the placeholder list below with your actual column names:
    # "Scales":     ["enter", "your", "col", "names", "..."],
    # "AnotherTab": ["colA", "colB", "colC"],
}

# ─── 5) Run comparisons ───────────────────────────────────────────────────────────
conn   = connect_to_oracle(cfg["host"], cfg["port"], cfg["svc"], user, pw)
out_xl = master_xl.replace(".xlsx", "_comparison.xlsx")
writer = pd.ExcelWriter(out_xl, engine="openpyxl")

for sheet_name, sql in queries.items():
    print(f"\n▶ Processing '{sheet_name}'")
    df_sheet = pd.read_excel(master_xl, sheet_name=sheet_name,
                             engine="openpyxl", dtype=str)
    if "Concatenated" not in df_sheet.columns:
        raise KeyError(f"'{sheet_name}' missing 'Concatenated' column")

    df_db = query_to_df(conn, sql)
    df_db = build_db_concat(df_db, sheet_name, concat_map)

    only_db, only_sheet = compare_mismatches(df_sheet, df_db)
    if not only_db and not only_sheet:
        print("   ✔️ No mismatches")
        mismatch_df = pd.DataFrame([{"Result":"All rows match"}])
    else:
        print(f"   ⚠️ {len(only_db)} only in DB, {len(only_sheet)} only in Sheet")
        rows = ([{"Source":"DB only",    "Concatenated":v} for v in only_db]
              + [{"Source":"Sheet only", "Concatenated":v} for v in only_sheet])
        mismatch_df = pd.DataFrame(rows)
    mismatch_df.to_excel(writer, sheet_name=f"{sheet_name}_Mismatches", index=False)

    ds = find_duplicates(df_sheet, "Concatenated")
    if not ds.empty:
        ds.to_excel(writer, sheet_name=f"{sheet_name}_SheetDupes", index=False)

    dd = find_duplicates(df_db, "Concatenated")
    if not dd.empty:
        dd.to_excel(writer, sheet_name=f"{sheet_name}_DBDupes", index=False)

writer.save()
conn.close()
print(f"\n✅ Report written to: {out_xl}")
