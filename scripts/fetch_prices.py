"""
fetch_prices.py
────────────────
Fetches daily closing prices for all instruments in portfolio.db
and writes them to the `prices` table.  Idempotent — safe to re-run
at any time; duplicates are silently ignored.

Routing (same logic as live_price_scraper.py):
  Italian ISIN  (IT*)  → Euronext MOT  (AES-encrypted AJAX)
  .FRA ticker          → Yahoo .DE first; Deutsche Börse page as fallback
  All other tickers    → Yahoo Finance v8 API

.FRA FALLBACK NOTE
───────────────────
Deutsche Börse's live page returns only the current price (no history).
On the first run the backfill will be limited to one data point for any
instrument whose .DE ticker is also not found on Yahoo.  Subsequent daily
runs will add one row per day.  If you find a working Yahoo ticker for one
of these instruments, add it to ISIN_TICKER_SUPPLEMENT.
"""

import base64
import binascii
import hashlib
import json
import os
import re
import shutil
import sqlite3
import tempfile
import time
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from pathlib import Path
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

# ── Paths ──────────────────────────────────────────────────────────────────────

BASE      = Path(__file__).resolve().parent.parent
DB_PATH   = BASE / "data" / "portfolio.db"
XLSM_PATH = BASE / "portfolio-dashboard.xlsm"

REQUEST_DELAY = 0.5   # seconds between HTTP requests
TIMEOUT       = 15    # request timeout

# ── URL overrides ──────────────────────────────────────────────────────────────
# Paste a live.deutsche-boerse.com product URL here for any .FRA ticker that
# returns N/A.  Takes priority over all automatic routing.
URL_OVERRIDES: dict[str, str] = {
    "8PSB.FRA": "https://live.deutsche-boerse.com/en/etf/invesco-physical-silver-etc?currency=EUR",
}

# ── ISIN → canonical ticker ────────────────────────────────────────────────────
# For ISINs whose List-sheet proxy is a plain ticker (not ISIN.EXCHANGE),
# auto-extraction won't work — supply them here.
ISIN_TICKER_SUPPLEMENT: dict[str, str] = {
    "IE00B579F325": "SGLD.AS",   # GLDFIXPM/SOURCE 00   Invesco Physical Gold ETC
    "IE00BG0SKF03": "5MVL.FRA",  # ISHS IV EM VAL USD   iShares MSCI EM Value Factor
    "LU1834988864": "UTI.MI",    # LIF ST EU 600 U UC   Amundi STOXX EU600 Utilities
    "IE0003Z9E2Y3": "4COP.FRA",  # GLB X CP USD-ACC     Global X Copper Miners UCITS
    "IE00B43VDT70": "8PSB.FRA",  # SILVER/SOURCE 00     Invesco Physical Silver ETC
    "LU0290358497": "XEON.DE",   # XTR2 EUR OR SW 1CC   Xtrackers EUR Overnight Rate
}

# Bond symbol (as it appears in trades) → (ISIN, MIC) for Euronext
BTP_EURONEXT: dict[str, tuple[str, str]] = {
    "BTP 1.8.39 5%": ("IT0004286966", "MOTX"),
}

# CA name fragment → Yahoo ticker (CA trades have no ISIN in notes)
SYMBOL_TICKER_CA: dict[str, str] = {
    "AMUNDI STOXX EUROPE": "UTI.PA",
    "INVESCO PHYS GOLD":   "SGLE.MI",
    "ENEL":                "ENEL.MI",  # Italian equity, CA account
}

# Price history for these yahoo_tickers starts at the first BTP trade date,
# not at the instrument's own first trade date.
# Reason: ENEL has a long trade history; we only need prices from when the
# BTP position was opened to keep the data scope consistent.
PRICE_START_AT_BTP: frozenset[str] = frozenset({"ENEL.MI"})

# Deutsche Börse product categories, tried in order
_DB_CATEGORIES = ("etf", "etp", "etc", "equity", "bond", "fund", "certificate", "warrant")

# JSON price keys used by live.deutsche-boerse.com (same as live_price_scraper.py)
_DB_PRICE_KEYS = {
    "lastPrice", "last", "price", "currentPrice", "tradePrice",
    "tradedPrice", "close", "closingPrice", "referencePrice",
    "lastTradedPrice", "netAssetValue", "nav",
}

# Runtime ISIN → ticker (populated from Excel + supplement at startup)
_isin_to_ticker: dict[str, str] = {}


# ── HTTP session ───────────────────────────────────────────────────────────────

_session = requests.Session()
_session.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
})


# ── Ticker-map builder ─────────────────────────────────────────────────────────

def _build_ticker_map() -> dict[str, str]:
    """
    Reads the List sheet from the portfolio Excel file.

    Returns name_map {canonical_ticker → instrument_name} (used by the
    Deutsche Börse slug builder) and populates the global _isin_to_ticker dict.

    Auto-extractable entries: rows whose proxy column matches the pattern
    ISIN.EXCHANGE (e.g. IE0003Z9E2Y3.SG → 4COP.FRA).
    The remainder are supplied via ISIN_TICKER_SUPPLEMENT.
    """
    name_map: dict[str, str] = {}
    try:
        import openpyxl
        src = Path(XLSM_PATH)
        tmp = Path(tempfile.gettempdir()) / src.name
        shutil.copy2(src, tmp)
        wb = openpyxl.load_workbook(str(tmp), read_only=True, data_only=True, keep_vba=False)
        ws = wb["List"]
        isin_re = re.compile(r"^([A-Z]{2}[0-9A-Z]{10})\.[A-Z]+$")
        for row in ws.iter_rows(values_only=True):
            if not row or row[0] == "Symbol/Proxy":
                continue
            proxy  = str(row[0]).strip() if row[0] else ""
            symbol = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            name   = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            m = isin_re.match(proxy)
            if m and symbol:
                _isin_to_ticker[m.group(1)] = symbol
            if symbol and name:
                name_map[symbol] = name
        print(f"  {len(_isin_to_ticker)} ISIN→ticker entries from List sheet.")
    except Exception as e:
        print(f"  [WARN] Could not read List sheet: {e}")
    _isin_to_ticker.update(ISIN_TICKER_SUPPLEMENT)
    return name_map


def _extract_isin(notes: str) -> Optional[str]:
    """'ISIN:IE00B579F325' → 'IE00B579F325'"""
    m = re.search(r"ISIN:([A-Z]{2}[0-9A-Z]{10})", str(notes))
    return m.group(1) if m else None


def _resolve(
    symbol: str, asset_class: str, notes: str, broker: str
) -> tuple[object, str, str]:
    """
    Returns (ticker_info, source, price_symbol).

    source values:
      'yahoo'    → ticker_info is a Yahoo ticker string
      'fra'      → ticker_info is the .FRA ticker string
      'euronext' → ticker_info is (isin, mic) tuple
      None       → skip this instrument
    """
    if asset_class == "WAR":
        return None, None, None

    if asset_class == "BND":
        for key, (isin, mic) in BTP_EURONEXT.items():
            if key.upper() in symbol.upper():
                return (isin, mic), "euronext", isin
        return None, None, None

    isin = _extract_isin(notes)
    if isin:
        ticker = _isin_to_ticker.get(isin)
        if ticker:
            source = "fra" if ticker.upper().endswith(".FRA") else "yahoo"
            return ticker, source, ticker
        print(f"  [WARN] ISIN {isin} not in ticker map (symbol={symbol!r})")
        return None, None, None

    sym_up = symbol.upper()
    for fragment, ticker in SYMBOL_TICKER_CA.items():
        if fragment.upper() in sym_up:
            return ticker, "yahoo", ticker

    print(f"  [WARN] Cannot resolve ticker for {symbol!r} ({broker})")
    return None, None, None


# ── Source 1: Yahoo Finance v8 ─────────────────────────────────────────────────

def _fetch_yahoo(ticker: str, start_date: str) -> list[tuple[str, float]]:
    """Daily closing prices from Yahoo Finance from start_date to today."""
    start_ts = int(
        datetime.strptime(start_date, "%Y-%m-%d")
        .replace(tzinfo=timezone.utc).timestamp()
    )
    end_ts = int(datetime.now(timezone.utc).timestamp())

    url    = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
    params = {"interval": "1d", "period1": start_ts, "period2": end_ts, "events": "history"}
    r = _session.get(url, params=params, timeout=TIMEOUT)
    r.raise_for_status()

    data   = r.json()
    result = data.get("chart", {}).get("result")
    if not result:
        err = data.get("chart", {}).get("error", {})
        raise ValueError(f"Yahoo API error: {err}")

    res        = result[0]
    timestamps = res.get("timestamp", [])
    closes     = res["indicators"]["quote"][0].get("close", [])

    rows: list[tuple[str, float]] = []
    for ts, close in zip(timestamps, closes):
        if close is None:
            continue
        d = datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
        rows.append((d, round(close, 6)))
    return rows


def _fetch_yahoo_live(ticker: str) -> tuple[float, str]:
    """
    Fetch the current market price from Yahoo Finance v8.
    Returns (regularMarketPrice, marketState).
    """
    url    = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
    params = {"interval": "1d", "range": "1d"}
    r = _session.get(url, params=params, timeout=TIMEOUT)
    r.raise_for_status()
    data   = r.json()
    result = data.get("chart", {}).get("result")
    if not result:
        err = data.get("chart", {}).get("error", {})
        raise ValueError(f"Yahoo API error: {err}")
    meta         = result[0].get("meta", {})
    price        = meta.get("regularMarketPrice")
    market_state = meta.get("marketState", "UNKNOWN")
    if price is None or price <= 0:
        raise ValueError(f"No regularMarketPrice in meta for {ticker}")
    return round(float(price), 6), market_state


# ── Source 2: Deutsche Börse (for .FRA tickers) ────────────────────────────────
# Same strategy as live_price_scraper.py: build a URL slug from the instrument
# name and scrape the live page.  Returns only today's price (no history).

def _name_to_db_slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def _walk_json_for_price(node, depth: int = 0) -> Optional[float]:
    """Recursively find the first positive numeric price field in a JSON tree."""
    if depth > 12:
        return None
    if isinstance(node, dict):
        for key, val in node.items():
            if key in _DB_PRICE_KEYS and isinstance(val, (int, float)) and val > 0:
                return float(val)
        for val in node.values():
            found = _walk_json_for_price(val, depth + 1)
            if found is not None:
                return found
    elif isinstance(node, list):
        for item in node:
            found = _walk_json_for_price(item, depth + 1)
            if found is not None:
                return found
    return None


def _scrape_db_page(url: str) -> Optional[float]:
    """
    Scrape the current price from a live.deutsche-boerse.com product page.
    Mirrors the three-strategy approach in live_price_scraper.py.
    """
    try:
        r = _session.get(url, timeout=TIMEOUT)
        r.raise_for_status()
        html = r.text

        # Strategy 1: __NEXT_DATA__ (Next.js server-side props)
        m = re.search(
            r'<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>',
            html, re.DOTALL,
        )
        if m:
            try:
                price = _walk_json_for_price(json.loads(m.group(1)))
                if price is not None:
                    return round(price, 6)
            except json.JSONDecodeError:
                pass

        # Strategy 2: JSON-LD structured data
        for ld in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.DOTALL,
        ):
            try:
                price = _walk_json_for_price(json.loads(ld))
                if price is not None:
                    return round(price, 6)
            except json.JSONDecodeError:
                continue

        # Strategy 3: plain-text pattern near a price label
        m2 = re.search(
            r'(?:Last Price|Letzter Kurs|Current Price|Price)[^\d]{0,30}([\d]+[.,][\d]+)',
            html, re.IGNORECASE,
        )
        if m2:
            try:
                return float(m2.group(1).replace(",", "."))
            except ValueError:
                pass

    except Exception as e:
        print(f"  [WARN] Deutsche Börse page error: {e}")
    return None


def _fetch_fra_by_name(name: str) -> Optional[float]:
    """Try all Deutsche Börse product-category slugs derived from the name."""
    candidates = [name]
    if re.search(r"\bET$", name):          # truncated "...UCITS ET" → add "F"
        candidates.append(name + "F")

    for candidate in candidates:
        slug = _name_to_db_slug(candidate)
        for category in _DB_CATEGORIES:
            url   = f"https://live.deutsche-boerse.com/en/{category}/{slug}?currency=EUR"
            price = _scrape_db_page(url)
            if price is not None:
                return price
    return None


def _fetch_fra_history(
    ticker: str, name: Optional[str], start_date: str
) -> tuple[list[tuple[str, float]], str]:
    """
    Fetch price history for a .FRA-listed instrument.

    Tries in order:
      1. Yahoo Finance with .DE suffix (Xetra, works for most dual-listed ETFs).
      2. Deutsche Börse live page via URL override.
      3. Deutsche Börse live page via name slug.

    Returns (rows, note) where note describes which source was used.
    Deutsche Börse fallback yields at most one row (today's close).
    """
    de_ticker = re.sub(r"\.FRA$", ".DE", ticker, flags=re.IGNORECASE)
    try:
        rows = _fetch_yahoo(de_ticker, start_date)
        if rows:
            return rows, f"Yahoo({de_ticker})"
    except Exception:
        pass

    # Deutsche Börse fallback ─ today's price only
    url_override = URL_OVERRIDES.get(ticker, "").strip()
    if url_override:
        price = _scrape_db_page(url_override)
    elif name:
        price = _fetch_fra_by_name(name)
    else:
        price = None

    if price is not None:
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        return [(today, price)], "DeutscheBoerse(live only)"

    return [], "no source found"


def _fetch_fra_live(ticker: str, name: Optional[str]) -> tuple[float, str]:
    """
    Live quote for a .FRA-listed instrument.
    Tries the .DE Yahoo ticker first; falls back to Deutsche Börse scraping.
    Returns (price, market_state).
    """
    de_ticker = re.sub(r"\.FRA$", ".DE", ticker, flags=re.IGNORECASE)
    try:
        return _fetch_yahoo_live(de_ticker)
    except Exception:
        pass

    url_override = URL_OVERRIDES.get(ticker, "").strip()
    if url_override:
        price = _scrape_db_page(url_override)
    elif name:
        price = _fetch_fra_by_name(name)
    else:
        price = None

    if price is not None:
        return price, "CLOSED"

    raise ValueError(f"No live price found for {ticker}")


# ── Source 3: Euronext MOT (BTP / bonds) ───────────────────────────────────────
# The old /en/ajax/getHistoricalPriceData endpoint returned 404.
# We use the intraday quote endpoint instead, which returns today's last-traded
# price (same pattern as the Deutsche Börse live fallback: one row per daily run).

_EN_BASE         = "https://live.euronext.com"
_EN_PAGE_TPL     = _EN_BASE + "/it/product/bonds/{dna}"
_EN_CHART_TPL    = _EN_BASE + "/en/intraday_chart/getChartData/{dna}/max"
_EN_QUOTE_TPL    = _EN_BASE + "/en/intraday_chart/getDetailedQuoteAjax/{dna}/full"
_EN_FALLBACK_KEY = "24ayqVo7yJma"


def _evp_bytes_to_key(password: bytes, salt: bytes, key_len: int = 32) -> tuple[bytes, bytes]:
    """OpenSSL EVP_BytesToKey (MD5, 1 iteration) — mirrors CryptoJS EvpKDF."""
    d, chunk = b"", b""
    while len(d) < key_len + 16:
        chunk = hashlib.md5(chunk + password + salt).digest()
        d += chunk
    return d[:key_len], d[key_len : key_len + 16]


def _decrypt_cryptojs(json_str: str, passphrase: str) -> str:
    data   = json.loads(json_str)
    ct     = base64.b64decode(data["ct"])
    iv     = binascii.unhexlify(data["iv"])
    salt   = binascii.unhexlify(data["s"])
    key, _ = _evp_bytes_to_key(passphrase.encode(), salt)
    cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
    dec    = cipher.decryptor()
    padded = dec.update(ct) + dec.finalize()
    return padded[: -padded[-1]].decode("utf-8")


def _get_euronext_key(dna: str) -> str:
    try:
        r = _session.get(_EN_PAGE_TPL.format(dna=dna), timeout=TIMEOUT)
        m = re.search(r'"ajax_secure"\s*:\s*\{"kye"\s*:\s*"([^"]+)"', r.text)
        if m:
            return m.group(1)
    except Exception:
        pass
    return _EN_FALLBACK_KEY


def _parse_euronext_chart(payload) -> list[tuple[str, float]]:
    """Parse {time, price, volume} rows from the Euronext chart endpoint."""
    rows: list[tuple[str, float]] = []
    if not isinstance(payload, list):
        return rows
    for item in payload:
        if not isinstance(item, dict):
            continue
        t = item.get("time", "")
        p = item.get("price")
        if not t or p is None:
            continue
        try:
            rows.append((str(t)[:10], round(float(p), 6)))  # "2024-05-31 02:00" → "2024-05-31"
        except (ValueError, TypeError):
            continue
    return rows


def _fetch_euronext_live(
    isin: str, mic: str, key: Optional[str] = None
) -> tuple[float, str]:
    """
    Fetch the last-traded price for a Euronext bond from the intraday quote endpoint.
    Returns (price, market_state).  Pass key to reuse an already-fetched AES key.
    """
    dna = f"{isin}-{mic}"
    if key is None:
        key = _get_euronext_key(dna)
    r = _session.get(
        _EN_QUOTE_TPL.format(dna=dna),
        headers={
            "Referer":            _EN_PAGE_TPL.format(dna=dna),
            "X-Requested-With":   "XMLHttpRequest",
            "Accept":             "*/*",
        },
        timeout=TIMEOUT,
    )
    r.raise_for_status()
    raw = r.text.strip()
    if raw.startswith("{") and '"ct"' in raw and '"iv"' in raw:
        raw = _decrypt_cryptojs(raw, key)
    try:
        html = json.loads(raw)
    except (json.JSONDecodeError, ValueError):
        html = raw
    soup = BeautifulSoup(html, "html.parser")
    for row in soup.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 2:
            continue
        if cells[0].get_text(strip=True) in ("Last Traded", "Last"):
            raw_val = cells[1].get_text(strip=True).replace(",", ".")
            price = float(re.sub(r"[^\d.]", "", raw_val))
            if price > 0:
                return round(price, 6), "REGULAR"
    raise ValueError(f"No live price in Euronext intraday quote for {dna}")


def _fetch_euronext_history(
    isin: str, mic: str, _: str
) -> list[tuple[str, float]]:
    """
    Fetch daily closing prices for a bond from Euronext.

    Primary: chart endpoint (/max) — returns ~2 years of daily OHLCV data
    including today, same AES-encrypted format.  Replaces the dead
    /ajax/getHistoricalPriceData endpoint (which now returns 404).

    Fallback: intraday quote endpoint — returns today's last-traded price
    only, parsed from an HTML table.
    """
    dna = f"{isin}-{mic}"
    key = _get_euronext_key(dna)
    time.sleep(REQUEST_DELAY)

    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    # ── Primary: chart data (full history) ────────────────────────────────────
    rows: list[tuple[str, float]] = []
    try:
        r = _session.get(
            _EN_CHART_TPL.format(dna=dna),
            headers={"Referer": _EN_PAGE_TPL.format(dna=dna),
                     "X-Requested-With": "XMLHttpRequest"},
            timeout=TIMEOUT,
        )
        r.raise_for_status()
        raw = r.text.strip()
        if raw.startswith("{") and '"ct"' in raw and '"iv"' in raw:
            raw = _decrypt_cryptojs(raw, key)
        rows = _parse_euronext_chart(json.loads(raw))
    except Exception as e:
        print(f"  [WARN] Euronext chart fetch failed for {dna}: {e}")

    # ── Supplement: intraday quote when today's price is missing ──────────────
    today_missing = not any(r[0] == today for r in rows)
    is_weekday    = datetime.now(timezone.utc).weekday() < 5
    if today_missing and is_weekday:
        try:
            price, _ = _fetch_euronext_live(isin, mic, key=key)
            rows.append((today, price))
        except Exception as e:
            print(f"  [WARN] Euronext intraday supplement failed for {dna}: {e}")

    return rows


# ── Database helpers ───────────────────────────────────────────────────────────

def _setup_db(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS prices (
            date   TEXT NOT NULL,
            symbol TEXT NOT NULL,
            close  REAL NOT NULL,
            PRIMARY KEY (date, symbol)
        )
    """)
    conn.commit()


def _setup_live_prices_table(conn: sqlite3.Connection) -> None:
    conn.execute("""
        CREATE TABLE IF NOT EXISTS live_prices (
            symbol       TEXT PRIMARY KEY,
            price        REAL NOT NULL,
            market_state TEXT,
            fetched_at   TEXT NOT NULL
        )
    """)
    conn.commit()


def _last_price_date(conn: sqlite3.Connection, symbol: str) -> Optional[str]:
    row = conn.execute(
        "SELECT MAX(date) FROM prices WHERE symbol = ?", (symbol,)
    ).fetchone()
    return row[0] if row and row[0] else None


def _insert_prices(
    conn: sqlite3.Connection,
    symbol: str,
    rows: list[tuple[str, float]],
) -> tuple[int, int]:
    """INSERT OR IGNORE rows.  Returns (inserted, ignored)."""
    before = conn.execute(
        "SELECT COUNT(*) FROM prices WHERE symbol = ?", (symbol,)
    ).fetchone()[0]
    conn.executemany(
        "INSERT OR IGNORE INTO prices (date, symbol, close) VALUES (?, ?, ?)",
        [(d, symbol, c) for d, c in rows],
    )
    conn.commit()
    after    = conn.execute(
        "SELECT COUNT(*) FROM prices WHERE symbol = ?", (symbol,)
    ).fetchone()[0]
    inserted = after - before
    return inserted, len(rows) - inserted


def _upsert_live_price(
    conn: sqlite3.Connection,
    symbol: str,
    price: float,
    market_state: str,
    fetched_at: str,
) -> None:
    conn.execute(
        "INSERT OR REPLACE INTO live_prices (symbol, price, market_state, fetched_at)"
        " VALUES (?, ?, ?, ?)",
        (symbol, price, market_state, fetched_at),
    )
    conn.commit()


def _run_live_mode(conn: sqlite3.Connection, name_map: dict) -> None:
    """Fetch current intraday prices for all open positions into live_prices."""
    _setup_live_prices_table(conn)

    open_symbols = set(
        row[0]
        for row in conn.execute("""
            SELECT symbol
            FROM   trades
            GROUP  BY symbol
            HAVING SUM(CASE WHEN action = 'BUY' THEN qty ELSE -qty END) > 0.001
        """).fetchall()
    )

    instruments = pd.read_sql(
        "SELECT symbol, asset_class, broker, notes FROM trades GROUP BY symbol",
        conn,
    )
    instruments = instruments[instruments["symbol"].isin(open_symbols)].reset_index(drop=True)

    print(f"Live price fetch for {len(instruments)} open instrument(s)...\n")
    print(f"  {'Symbol':<22}  {'Price':>10}  {'State':<10}  Fetched")
    print(f"  {'-'*22}  {'-'*10}  {'-'*10}  -------")

    for _, row in instruments.iterrows():
        symbol      = row["symbol"]
        asset_class = row["asset_class"]
        broker      = row["broker"]
        notes       = row["notes"]

        ticker, source, price_symbol = _resolve(symbol, asset_class, notes, broker)
        if ticker is None:
            label = "WAR" if asset_class == "WAR" else "no ticker"
            print(f"  {symbol:<22}  [SKIP] {label}")
            continue

        fetched_at = datetime.now(ZoneInfo("Europe/Rome")).strftime("%Y-%m-%dT%H:%M:%S")
        hhmm       = fetched_at[11:16]

        try:
            if source == "euronext":
                isin, mic = ticker
                price, market_state = _fetch_euronext_live(isin, mic)
            elif source == "fra":
                price, market_state = _fetch_fra_live(ticker, name_map.get(ticker))
            else:
                price, market_state = _fetch_yahoo_live(ticker)

            _upsert_live_price(conn, price_symbol, price, market_state, fetched_at)
            print(f"  {price_symbol:<22}  {price:>10.4f}  {market_state:<10}  @ {hhmm} CET")

        except Exception as e:
            print(f"  {price_symbol:<22}  [WARN] {e}")

        time.sleep(REQUEST_DELAY)

    print()


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--backfill-from",
        metavar="YYYY-MM-DD",
        default=None,
        help="Force fetch history from this date for every instrument, "
             "bypassing the last-price short-circuit. Use once to anchor YTD baselines.",
    )
    parser.add_argument(
        "--live",
        action="store_true",
        help="Fetch current intraday prices into live_prices table. "
             "Does NOT write to prices table and does NOT run analytics.",
    )
    args = parser.parse_args()
    backfill_from: str | None = args.backfill_from

    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    _setup_db(conn)

    print("Building ticker map…")
    name_map = _build_ticker_map()
    print(f"  {len(_isin_to_ticker)} total ISIN->ticker entries after supplement.\n")

    if args.live:
        _run_live_mode(conn, name_map)
        conn.close()
        return

    if backfill_from:
        print(f"[backfill mode] fetch_from forced to {backfill_from} for all instruments.\n")

    instruments = pd.read_sql(
        "SELECT symbol, asset_class, broker, notes, MIN(date) AS first_date "
        "FROM trades GROUP BY symbol",
        conn,
    )

    # Only fetch prices for instruments with a currently open position.
    # Net qty is computed directly from trades so this works even before
    # build_analytics has run.
    open_symbols = set(
        row[0]
        for row in conn.execute("""
            SELECT symbol
            FROM   trades
            GROUP  BY symbol
            HAVING SUM(CASE WHEN action = 'BUY' THEN qty ELSE -qty END) > 0.001
        """).fetchall()
    )
    closed = instruments[~instruments["symbol"].isin(open_symbols)]
    for sym in closed["symbol"]:
        print(f"  {sym:<30}  [SKIP] position closed")
    instruments = instruments[instruments["symbol"].isin(open_symbols)].reset_index(drop=True)

    print(f"Found {len(instruments)} open instrument(s) in trades (closed positions skipped).\n")

    total_fetched = total_inserted = total_ignored = 0

    for _, row in instruments.iterrows():
        symbol      = row["symbol"]
        asset_class = row["asset_class"]
        broker      = row["broker"]
        notes       = row["notes"]
        first_date  = row["first_date"]

        ticker, source, price_symbol = _resolve(symbol, asset_class, notes, broker)

        if ticker is None:
            label = f"WAR" if asset_class == "WAR" else "no ticker"
            print(f"  {symbol:<30}  [SKIP] {label}")
            continue

        # Determine the start date for price fetching.
        # In backfill mode the caller-supplied date overrides everything so that
        # earlier history (e.g. a 2025 year-end close) can be added to existing series.
        # In normal mode: resume from last stored price, or fall back to first trade date.
        if backfill_from:
            fetch_from = backfill_from
        else:
            last_price = _last_price_date(conn, price_symbol)
            if last_price:
                fetch_from = last_price
            elif price_symbol in PRICE_START_AT_BTP:
                btp_date = conn.execute(
                    "SELECT MIN(date) FROM trades WHERE asset_class = 'BND'"
                ).fetchone()[0]
                fetch_from = btp_date or first_date
            else:
                fetch_from = first_date

        note = ""
        try:
            if source == "euronext":
                isin, mic = ticker
                rows_data = _fetch_euronext_history(isin, mic, fetch_from)
            elif source == "fra":
                instrument_name = name_map.get(ticker)
                rows_data, note = _fetch_fra_history(ticker, instrument_name, fetch_from)
            else:
                rows_data = _fetch_yahoo(ticker, fetch_from)
        except Exception as e:
            print(f"  {price_symbol:<20}  [ERROR] {e}")
            time.sleep(REQUEST_DELAY)
            continue

        n_fetched         = len(rows_data)
        inserted, ignored = _insert_prices(conn, price_symbol, rows_data)

        note_str = f"  ({note})" if note else ""
        print(
            f"  {price_symbol:<20}"
            f"  -> {n_fetched:4d} fetched,"
            f" {inserted:4d} inserted,"
            f" {ignored:4d} ignored"
            f"{note_str}"
        )

        total_fetched  += n_fetched
        total_inserted += inserted
        total_ignored  += ignored

        time.sleep(REQUEST_DELAY)

    print(
        f"\nTotal:  {total_fetched} fetched,"
        f"  {total_inserted} inserted,"
        f"  {total_ignored} ignored."
    )
    conn.close()


if __name__ == "__main__":
    main()
