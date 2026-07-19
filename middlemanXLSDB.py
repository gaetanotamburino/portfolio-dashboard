import re
import xlwings as xw
import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "data" / "portfolio.db"

ISIN_TICKER_MAP: dict[str, str] = {
    "NLBNPIT34AB0": "NLBNPIT34AB0",
    "IE00B579F325": "SGLD.AS",
    "IE00BG0SKF03": "5MVL.FRA",
    "LU1834988864": "UTI.MI",
    "IE0003Z9E2Y3": "4COP.FRA",
    "IE00B43VDT70": "8PSB.FRA",
    "LU0290358497": "XEON.DE",
}

SYMBOL_TICKER_MAP: dict[str, str] = {
    "AMUNDI STOXX EUROPE": "UTI.PA",
    "INVESCO PHYS GOLD":   "SGLE.MI",
    "ENEL":                "ENEL.MI",
    "BTP 1.8.39":          "IT0004286966",
}

_ISIN_RE = re.compile(r"ISIN:([A-Z]{2}[A-Z0-9]{10})")

def _resolve_ticker(symbol, notes):
    if notes:
        m = _ISIN_RE.search(notes)
        if m:
            return ISIN_TICKER_MAP.get(m.group(1))
    for fragment, ticker in SYMBOL_TICKER_MAP.items():
        if fragment in symbol:
            return ticker
    return None

def main():
    wb = xw.Book.caller()
    ws = wb.sheets["data"]

    try:
        conn = sqlite3.connect(DB_PATH)
        db_rows = conn.execute("SELECT symbol, price FROM live_prices").fetchall()
        trades_rows = conn.execute(
            "SELECT DISTINCT symbol, notes FROM trades"
        ).fetchall()
        conn.close()
    except Exception as e:
        ws.range("G1").value = f"DB Error: {e}"
        return

    # ticker → instrument name
    ticker_to_inst: dict[str, str] = {}
    for sym, notes in trades_rows:
        ticker = _resolve_ticker(sym, notes)
        if ticker and (ticker not in ticker_to_inst or (notes and "ISIN:" in notes)):
            ticker_to_inst[ticker] = sym

    # instrument name → live price
    inst_to_price: dict[str, float] = {}
    for ticker, price in db_rows:
        if ticker in ticker_to_inst:
            inst_to_price[ticker_to_inst[ticker]] = price

    # Read instrument names from col B (row 2 downward)
    last_row = ws.range("B2").end("down").row
    inst_col = ws.range(f"B2:B{last_row}").value
    if not isinstance(inst_col, list):
        inst_col = [inst_col]

    # Write live prices to col G only
    ws.range(f"G2:G{last_row}").value = [
        [inst_to_price.get(name, "")] for name in inst_col
    ]
