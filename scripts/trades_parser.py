"""
trades_parser.py
───────────────────
Reads all export files (.xls / .xlsx) from the imports folder,
wrangles them and writes a single standardised trades CSV to the
portfolio dashboard final/trades output folder.
Source files are deleted after a successful conversion.

Broker detection (by filename)
───────────────────────────────
• Fineco  → filename matches "file", "file(1)", "file(2)", … (i.e. basename
            starts with "file" and has no "Lista Movimenti" prefix)
            Sheet: "Movimenti Dossier Titoli"
            account → Gaetano

• CA      → filename starts with "Lista Movimenti Deposito Titoli_CAI_"
            Sheet: "Lista Movimenti Deposito Titoli"
            account → CA Giovanna

Output columns:
    date, account, broker, symbol, exchange, asset_class,
    action, qty, price, commission, currency, notes

Output filename: trades_YYYYMMDD_HHMM.csv

Notes
─────
• Bonds (BND): Fineco's "Ctv in Eur" for bonds includes accrued interest,
  so commission cannot be cleanly back-calculated from price × qty alone.
  Commission is set to 0.0 for bonds; verify manually if needed.
• "Portafoglio remunerato" rows (money-market rollovers) are skipped.
• "CEDOLA" / "DIVIDENDO" rows are skipped.
"""

import os
import re
import glob
import sqlite3
import pandas as pd
from datetime import datetime
from pathlib import Path


# ── Configuration ─────────────────────────────────────────────────────────────

BASE          = Path(__file__).resolve().parent.parent
IMPORT_FOLDER = str(BASE / "imports")
OUTPUT_FOLDER = str(BASE / "data" / "trades csv")
DB_PATH       = str(BASE / "data" / "portfolio.db")
DEFAULT_CURR  = "EUR"

# Account labels — rename here to genericize before making the repo public
# (e.g. "Portfolio A" / "Portfolio B") without touching parser logic below.
ACCOUNT_1 = "Gaetano"
ACCOUNT_2 = "CA Giovanna"

# Regex patterns used to identify the broker from the filename (basename, no ext)
# Fineco exports are named "file", "file(1)", "file(2)", … by the browser
FINECO_PATTERN = re.compile(r"^file\s*(\(\d+\))?$", re.IGNORECASE)
CA_PATTERN     = re.compile(r"^Lista Movimenti Deposito Titoli_CAI_", re.IGNORECASE)


# ── Broker detection ──────────────────────────────────────────────────────────

def detect_broker(filepath: str) -> str | None:
    """
    Return 'Fineco', 'CA', or None (unrecognised) based on the filename alone.
    Detection happens before opening the file, so bad files are skipped early.
    """
    basename = os.path.splitext(os.path.basename(filepath))[0]
    if FINECO_PATTERN.match(basename):
        return "Fineco"
    if CA_PATTERN.match(basename):
        return "CA"
    return None


# ── Shared helpers ─────────────────────────────────────────────────────────────

def parse_it_number(val) -> float:
    """Parse an Italian-formatted number string: '15.111,66' → 15111.66"""
    if pd.isna(val):
        return float("nan")
    s = str(val).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def infer_from_isin(isin: str, name: str):
    """Exchange + asset class from ISIN prefix (Fineco file)."""
    country = isin[:2].upper() if isin else ""
    exchange_map = {
        "IE": "ETFplus", "LU": "ETFplus", "NL": "ETFplus",
        "DE": "XETRA",   "FR": "Euronext", "IT": "MTA",
    }
    name_up = name.upper()
    if "CW" in name_up or "**" in name_up:
        return "SeDex", "WAR"
    is_etf = any(kw in name_up for kw in [
        "ETF", "UCITS", "ACC", "DIST", "SOURCE", "XTR", "ISHS",
        "LIF ST", "GLB", "SILVER", "GLDFIXPM", "PHYS", "INVESCO",
    ])
    return exchange_map.get(country, "ETFplus"), ("ETF" if is_etf else "STK")


def infer_from_name(name: str):
    """Exchange + asset class from security name (CA file, no ISIN)."""
    n = name.upper()
    if any(kw in n for kw in ["BTP", "CCT", "BOT", "BTP€I"]):
        return "MTA", "BND"
    if any(kw in n for kw in [
        "PHYS GOLD", "AMUNDI", "INVESCO", "XTRACKERS", "ISHS",
        "VANGUARD", "SPDR", "LYXOR", "SOURCE", "ETF",
    ]):
        return "ETFplus", "ETF"
    if any(kw in n for kw in ["ENEL", "ENI", "INTESA", "UNICREDIT", "STELLANTIS"]):
        return "MTA", "STK"
    return "ETFplus", "ETF"   # safe default for this account


# ── Parser: Fineco — Movimenti Dossier Titoli (.xls) ─────────────────────────

def parse_fineco(filepath: str) -> pd.DataFrame:
    """Gaetano's personal Fineco account — .xls export."""
    try:
        xf = pd.ExcelFile(filepath, engine="xlrd")
    except Exception as exc:
        print(f"  [ERROR] Cannot open {os.path.basename(filepath)}: {exc}")
        return pd.DataFrame()

    if "Movimenti Dossier Titoli" not in xf.sheet_names:
        print(f"  [WARN] Expected sheet not found in {os.path.basename(filepath)}")
        return pd.DataFrame()

    raw = pd.read_excel(filepath, engine="xlrd", header=None)

    header_row = next(
        (i for i, row in raw.iterrows() if str(row.iloc[0]).strip() == "Operazione"),
        None
    )
    if header_row is None:
        print(f"  [WARN] Header row not found: {filepath}")
        return pd.DataFrame()

    df = raw.iloc[header_row + 1:, :11].copy()
    df.columns = [
        "operazione", "data_valuta", "descrizione",
        "titolo", "isin", "segno",
        "quantita", "divisa", "prezzo", "cambio", "controvalore"
    ]
    df = df.dropna(subset=["operazione", "titolo"]).reset_index(drop=True)

    # Skip money-market rollovers
    df = df[df["descrizione"].str.strip() != "Portafoglio remunerato"].copy()
    if df.empty:
        return pd.DataFrame()

    df["operazione"]   = pd.to_datetime(df["operazione"],   dayfirst=True, errors="coerce")
    df["quantita"]     = pd.to_numeric(df["quantita"],      errors="coerce")
    df["prezzo"]       = pd.to_numeric(df["prezzo"],        errors="coerce")
    df["controvalore"] = pd.to_numeric(df["controvalore"],  errors="coerce")
    df = df.dropna(subset=["operazione", "quantita", "prezzo"])

    ec = df.apply(
        lambda r: infer_from_isin(str(r["isin"]).strip(), str(r["titolo"]).strip()), axis=1
    )
    df["exchange"]    = ec.apply(lambda x: x[0])
    df["asset_class"] = ec.apply(lambda x: x[1])

    df["date"]       = df["operazione"].dt.strftime("%Y-%m-%d")
    df["account"]    = ACCOUNT_1
    df["broker"]     = "Fineco"
    df["symbol"]     = df["titolo"].str.strip()
    df["action"]     = df["segno"].str.strip().map({"A": "BUY", "V": "SELL"}).fillna("UNKNOWN")
    df["qty"]        = df["quantita"].abs()
    df["price"]      = df["prezzo"]
    df["currency"]   = df["divisa"].str.strip().fillna(DEFAULT_CURR)
    df["commission"] = (df["controvalore"] - df["qty"] * df["price"]).abs().round(4)
    df["notes"]      = "ISIN:" + df["isin"].str.strip()

    return df[[
        "date", "account", "broker", "symbol", "exchange",
        "asset_class", "action", "qty", "price",
        "commission", "currency", "notes"
    ]]


# ── Parser: CA — Lista Movimenti Deposito Titoli CAI (.xlsx) ─────────────────

SKIP_CAUSALI = {"CEDOLA", "DIVIDENDO", "GIROCONTO", "RIMBORSO"}
BUY_CAUSALI  = {"ACQ.CONT.SU MERC.", "ACQUISTO"}
SELL_CAUSALI = {"VEN.CONT.SU MERC.", "VENDITA"}

def parse_ca(filepath: str) -> pd.DataFrame:
    """CA Giovanna joint account — .xlsx CAI export."""
    try:
        xf = pd.ExcelFile(filepath, engine="openpyxl")
    except Exception as exc:
        print(f"  [ERROR] Cannot open {os.path.basename(filepath)}: {exc}")
        return pd.DataFrame()

    if "Lista Movimenti Deposito Titoli" not in xf.sheet_names:
        print(f"  [WARN] Expected sheet not found in {os.path.basename(filepath)}")
        return pd.DataFrame()

    raw = pd.read_excel(filepath, engine="openpyxl", header=None)

    header_row = next(
        (i for i, row in raw.iterrows() if str(row.iloc[2]).strip() == "Data operazione"),
        None
    )
    if header_row is None:
        print(f"  [WARN] Header row not found: {filepath}")
        return pd.DataFrame()

    df = raw.iloc[header_row + 1:].copy()
    df.columns = [
        "_a", "_b", "data_operazione", "nome",
        "divisa", "causale", "prezzo", "divisa_prezzo",
        "cambio", "quantita", "ctv_eur", "data_valuta"
    ]
    df = df.dropna(subset=["data_operazione", "nome"]).reset_index(drop=True)

    df["causale_up"] = df["causale"].str.strip().str.upper()
    df = df[~df["causale_up"].isin(SKIP_CAUSALI)].copy()
    if df.empty:
        return pd.DataFrame()

    df["prezzo"]          = df["prezzo"].apply(parse_it_number)
    df["quantita"]        = df["quantita"].apply(parse_it_number)
    df["ctv_eur"]         = df["ctv_eur"].apply(parse_it_number)
    df["data_operazione"] = pd.to_datetime(
        df["data_operazione"], dayfirst=True, errors="coerce"
    )
    df = df.dropna(subset=["data_operazione", "quantita", "prezzo"])

    ec = df["nome"].apply(lambda n: infer_from_name(str(n).strip()))
    df["exchange"]    = ec.apply(lambda x: x[0])
    df["asset_class"] = ec.apply(lambda x: x[1])

    df["date"]    = df["data_operazione"].dt.strftime("%Y-%m-%d")
    df["account"] = ACCOUNT_2
    df["broker"]  = "CA"
    df["symbol"]  = df["nome"].str.strip()
    df["action"]  = df["causale_up"].apply(
        lambda c: "BUY" if c in BUY_CAUSALI else ("SELL" if c in SELL_CAUSALI else "UNKNOWN")
    )
    df["qty"]      = df["quantita"].abs()
    df["price"]    = df["prezzo"]
    df["currency"] = DEFAULT_CURR

    # Bonds: Ctv includes accrued interest → commission not reliably derivable
    def calc_commission(row):
        if row["asset_class"] == "BND":
            return 0.0
        return round(abs(row["ctv_eur"] - row["qty"] * row["price"]), 4)

    df["commission"] = df.apply(calc_commission, axis=1)
    df["notes"]      = df.apply(
        lambda r: (
            r["causale"].strip() + " - commission=0 (Ctv includes accrued interest)"
            if r["asset_class"] == "BND"
            else r["causale"].strip()
        ), axis=1
    )

    return df[[
        "date", "account", "broker", "symbol", "exchange",
        "asset_class", "action", "qty", "price",
        "commission", "currency", "notes"
    ]]


# ── Dispatcher ────────────────────────────────────────────────────────────────

def detect_and_parse(filepath: str):
    """
    Identify broker from filename, then route to the correct parser.
    Returns (DataFrame, broker_label) or (empty DataFrame, None).
    """
    broker = detect_broker(filepath)

    if broker is None:
        print(f"  [WARN] Unrecognised filename, skipped: {os.path.basename(filepath)}")
        return pd.DataFrame(), None

    print(f"  -> {os.path.basename(filepath)}  [broker: {broker}]")

    if broker == "Fineco":
        frame = parse_fineco(filepath)
    elif broker == "CA":
        frame = parse_ca(filepath)
    else:
        frame = pd.DataFrame()

    return frame, broker


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # Validate folders
    if not os.path.isdir(IMPORT_FOLDER):
        print(f"[ERROR] Import folder not found:\n  {IMPORT_FOLDER}")
        return

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

    # Ensure DB and table exist
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS trades (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            date        TEXT NOT NULL,
            account     TEXT NOT NULL,
            broker      TEXT NOT NULL,
            symbol      TEXT NOT NULL,
            exchange    TEXT,
            asset_class TEXT,
            action      TEXT NOT NULL,
            qty         REAL NOT NULL,
            price       REAL NOT NULL,
            commission  REAL DEFAULT 0.0,
            currency    TEXT DEFAULT 'EUR',
            notes       TEXT,
            UNIQUE(date, account, broker, symbol, action, qty, price)
        )
    """)
    conn.commit()

    seen, files = set(), []
    for ext in ("*.xls", "*.XLS", "*.xlsx", "*.XLSX"):
        for f in glob.glob(os.path.join(IMPORT_FOLDER, ext)):
            key = os.path.normcase(f)
            if key not in seen:
                seen.add(key)
                files.append(f)

    if not files:
        print(f"[ERROR] No .xls/.xlsx files found in:\n  {IMPORT_FOLDER}")
        return

    print(f"Found {len(files)} file(s) to process:\n")

    all_frames, parsed_files = [], []

    for f in sorted(files):
        frame, broker = detect_and_parse(f)
        if frame is not None and not frame.empty:
            all_frames.append(frame)
            parsed_files.append(f)
            print(f"     {len(frame)} trade row(s) extracted.")
        else:
            print(f"     No usable rows found.")

    if not all_frames:
        print("\nNo trades found. Nothing written, no files deleted.")
        return

    new_trades = pd.concat(all_frames, ignore_index=True)
    print(f"\n  {len(new_trades)} new trade row(s) from imports.")

    # ── Write to SQLite ───────────────────────────────────────────────────────
    rows_attempted = len(new_trades)
    rows_before = conn.execute("SELECT COUNT(*) FROM trades").fetchone()[0]

    insert_sql = """
        INSERT OR IGNORE INTO trades
            (date, account, broker, symbol, exchange, asset_class,
             action, qty, price, commission, currency, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    conn.executemany(
        insert_sql,
        new_trades[[
            "date", "account", "broker", "symbol", "exchange", "asset_class",
            "action", "qty", "price", "commission", "currency", "notes"
        ]].itertuples(index=False, name=None)
    )
    conn.commit()

    rows_after    = conn.execute("SELECT COUNT(*) FROM trades").fetchone()[0]
    rows_inserted = rows_after - rows_before
    rows_ignored  = rows_attempted - rows_inserted

    print(f"\nDB insert summary:")
    print(f"  Rows attempted : {rows_attempted}")
    print(f"  Rows inserted  : {rows_inserted}  (new)")
    print(f"  Rows ignored   : {rows_ignored}  (duplicates)")
    print(f"  Total in DB    : {rows_after}")

    conn.close()

    # Load the latest existing trades file and carry it forward
    existing_csvs = sorted(glob.glob(os.path.join(OUTPUT_FOLDER, "trades_*.csv")))
    if existing_csvs:
        latest_csv = existing_csvs[-1]
        print(f"  Loading previous trades from: {os.path.basename(latest_csv)}")
        prev = pd.read_csv(latest_csv)
        combined = pd.concat([prev, new_trades], ignore_index=True)
    else:
        print("  No previous trades file found - starting fresh.")
        combined = new_trades

    combined = (
        combined
        .drop_duplicates(subset=["date", "account", "broker", "symbol", "action", "qty", "price"])
        .sort_values(["date", "broker", "account", "symbol"])
        .reset_index(drop=True)
    )

    # Summary by broker
    print("\nBroker summary (full combined file):")
    for broker, grp in combined.groupby("broker"):
        print(f"  {broker}: {len(grp)} rows")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_path  = os.path.join(OUTPUT_FOLDER, f"trades_{timestamp}.csv")
    combined.to_csv(out_path, index=False, float_format="%.4f")
    print(f"\n[OK] {len(combined)} total rows written to:\n  {out_path}")

    print("\nDeleting source files from imports:")
    for f in parsed_files:
        try:
            os.remove(f)
            print(f"  [deleted] {os.path.basename(f)}")
        except Exception as exc:
            print(f"  [WARN] Could not delete {os.path.basename(f)}: {exc}")

    print("\nPreview (first 8 rows):")
    print(combined.head(8).to_string(index=False))


def query_trades(sql: str) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql(sql, conn)
    conn.close()
    return df


if __name__ == "__main__":
    main()
    print("\n--- DB preview ---")
    print(query_trades(
        "SELECT broker, account, COUNT(*) as n_trades, "
        "ROUND(SUM(qty*price),2) as total_invested "
        "FROM trades GROUP BY broker, account"
    ))
