import pandas as pd
import oracledb
import getpass

def connect_to_oracle(host, port, service, user, pw):
    """
    Create and return an Oracle connection using oracledb.
    """
    dsn = oracledb.makedsn(host, port, service_name=service)
    return oracledb.connect(user=user, password=pw, dsn=dsn)

def query_to_df(conn, sql):
    """
    Execute the given SQL on conn and return the results as a DataFrame.
    """
    with conn.cursor() as cur:
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        rows = cur.fetchall()
    return pd.DataFrame(rows, columns=cols)

def build_db_concat(df_db):
    """
    On df_db, fill NaNs with "" and cast to str, then
    concatenate all columns into a single 'Concatenated' column.
    """
    df_db = df_db.fillna("").astype(str)
    df_db["Concatenated"] = df_db.apply(lambda r: "".join(r.values), axis=1)
    return df_db

def compare_mismatches(df_sheet, df_db):
    """
    Compare the two DataFrames by their 'Concatenated' column.
    Return (only_in_db, only_in_sheet) as lists of strings.
    """
    left  = df_sheet[["Concatenated"]].dropna().astype(str).drop_duplicates()
    right = df_db[["Concatenated"]].drop_duplicates()
    merged = left.merge(right, on="Concatenated", how="outer", indicator=True)
    only_db    = merged[merged["_merge"] == "right_only" ]["Concatenated"].tolist()
    only_sheet = merged[merged["_merge"] == "left_only"  ]["Concatenated"].tolist()
    return only_db, only_sheet

def find_duplicates(df, col="Concatenated"):
    """
    Return the subset of df where 'col' appears more than once.
    """
    vc = df[col].value_counts()
    dup_keys = vc[vc > 1].index.tolist()
    return df[df[col].isin(dup_keys)].copy()

# ──── 1) Select environment ─────────────────────────────────────────────────────

envs = {
    1: {"label":"SIT_STG", "host":"NYKDSR000007912.intranet.barcapint.com", "port":1523, "svc":"TTMUS02P"},
    2: {"label":"SIT_CDS", "host":"NYKDSR000007912.intranet.barcapint.com", "port":1523, "svc":"TTMUS02P"},
    3: {"label":"UAT_STG","host":"isamusatdb.barcapint.com",        "port":1523, "svc":"TTMUS01P"},
    4: {"label":"UAT_CDS","host":"isamusatdb.barcapint.com",        "port":1523, "svc":"TTMUS01P"},
    5: {"label":"PROD",   "host":"your.prod.host.company.com",      "port":1521, "svc":"PROD_SVC"}
}

print("Select environment:")
for num, cfg in envs.items():
    print(f"  {num}: {cfg['label']}")
choice = int(input("Enter environment number (1–5): ").strip())
cfg    = envs[choice]

username = input(f"{cfg['label']} username: ")
password = getpass.getpass(f"{cfg['label']} password: ")

# ──── 2) Load Excel and choose sheets ────────────────────────────────────────────

master_xl = input("\nPath to your master datasheet (.xlsx): ").strip()
xls       = pd.ExcelFile(master_xl, engine="openpyxl")

print("\nWhich sheets do you want to compare?")
print("  0: All sheets")
for i, name in enumerate(xls.sheet_names, start=1):
    print(f"  {i}: {name}")
selection = input("Enter 0 or comma-separated sheet numbers (e.g. 1,3,5): ").strip()

if selection == "0":
    sheets_to_process = xls.sheet_names
else:
    nums = [int(x) for x in selection.split(",") if x.strip().isdigit()]
    sheets_to_process = [xls.sheet_names[n-1] for n in nums]

# ──── 3) Prompt for each sheet’s SQL ─────────────────────────────────────────────

queries = {}
for sheet in sheets_to_process:
    sql = input(f"\nPaste SQL query for sheet '{sheet}':\n").strip()
    queries[sheet] = sql

# ──── 4) Connect to DB and prepare output writer ─────────────────────────────────

conn   = connect_to_oracle(cfg["host"], cfg["port"], cfg["svc"], username, password)
out_xl = master_xl.replace(".xlsx", "_comparison.xlsx")
writer = pd.ExcelWriter(out_xl, engine="openpyxl")

# ──── 5) Process each selected sheet ─────────────────────────────────────────────

for sheet_name, sql in queries.items():
    print(f"\n>>> Processing sheet '{sheet_name}'...")

    # 5.1) Read the Excel sheet
    df_sheet = pd.read_excel(
        master_xl,
        sheet_name=sheet_name,
        engine="openpyxl",
        dtype=str
    )
    if "Concatenated" not in df_sheet.columns:
        raise KeyError(f"Sheet '{sheet_name}' missing a 'Concatenated' column.")

    # 5.2) Run the SQL and build the DB concatenated column
    df_db = query_to_df(conn, sql)
    df_db = build_db_concat(df_db)

    # 5.3) Find mismatches
    only_db, only_sheet = compare_mismatches(df_sheet, df_db)
    if not only_db and not only_sheet:
        print(f"    ✔️  No mismatches in '{sheet_name}'.")
        mismatch_df = pd.DataFrame([{"Result":"All rows match"}])
    else:
        print(f"    ⚠️  {len(only_db)} only in DB, {len(only_sheet)} only in Sheet.")
        mismatch_rows = (
            [{"Source":"DB only",    "Concatenated":v} for v in only_db]
          + [{"Source":"Sheet only", "Concatenated":v} for v in only_sheet]
        )
        mismatch_df = pd.DataFrame(mismatch_rows)

    mismatch_df.to_excel(
        writer,
        sheet_name=f"{sheet_name}_Mismatches",
        index=False
    )

    # 5.4) Detect and write sheet-side duplicates
    df_sheet_dupes = find_duplicates(df_sheet, "Concatenated")
    if not df_sheet_dupes.empty:
        df_sheet_dupes.to_excel(
            writer,
            sheet_name=f"{sheet_name}_SheetDupes",
            index=False
        )

    # 5.5) Detect and write DB-side duplicates
    df_db_dupes = find_duplicates(df_db, "Concatenated")
    if not df_db_dupes.empty:
        df_db_dupes.to_excel(
            writer,
            sheet_name=f"{sheet_name}_DBDupes",
            index=False
        )

# ──── 6) Finalize ────────────────────────────────────────────────────────────────

writer.save()
conn.close()
print(f"\n✅ Comparison report written to:\n   {out_xl}")
