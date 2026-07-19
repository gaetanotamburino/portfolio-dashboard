"""
build_analytics.py
──────────────────
Builds three analytics layers on top of the trades and prices tables
already in portfolio.db.

  Step 1 — instruments     : master registry of all traded instruments
  Step 2 — positions       : daily positions with cost basis and P&L
  Step 3 — portfolio_summary : daily equity curve with returns
"""

import re
import sqlite3
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from accrued import daily_accrual_series

DB_PATH = Path(__file__).resolve().parent.parent / "data" / "portfolio.db"

# ── Ticker maps ────────────────────────────────────────────────────────────────
# ISIN → Yahoo ticker (matches fetch_prices.py ISIN_TICKER_SUPPLEMENT +
#                       the auto-extracted ISIN.EXCHANGE entries from the List sheet)
ISIN_TICKER_MAP: dict[str, str] = {
    "IE00B579F325": "SGLD.AS",   # GLDFIXPM/SOURCE 00   Invesco Physical Gold ETC
    "IE00BG0SKF03": "5MVL.FRA",  # ISHS IV EM VAL USD   iShares MSCI EM Value Factor
    "LU1834988864": "UTI.MI",    # LIF ST EU 600 U UC   Amundi STOXX EU600 Utilities
    "IE0003Z9E2Y3": "4COP.FRA",  # GLB X CP USD-ACC     Global X Copper Miners UCITS
    "IE00B43VDT70": "8PSB.FRA",  # SILVER/SOURCE 00     Invesco Physical Silver ETC
    "LU0290358497": "XEON.DE",   # XTR2 EUR OR SW 1CC   Xtrackers EUR Overnight Rate
}

# CA instrument name fragment → Yahoo ticker (CA trades have no ISIN in notes)
# BTP: no ISIN in notes either; use the ISIN itself as the prices-table key
# (fetch_prices.py stores BTP prices keyed by ISIN, not by a Yahoo ticker)
SYMBOL_TICKER_MAP: dict[str, str] = {
    "AMUNDI STOXX EUROPE": "UTI.PA",
    "INVESCO PHYS GOLD":   "SGLE.MI",
    "ENEL":                "ENEL.MI",   # Italian equity, CA account
    "BTP 1.8.39":          "IT0004286966",  # BTP prices stored by ISIN in prices table
}

# Yahoo tickers whose instruments are stored in prices but excluded from
# positions and portfolio_summary (analytics).
EXCLUDE_FROM_POSITIONS: frozenset[str] = frozenset({"ENEL.MI"})


# ── Shared helpers ─────────────────────────────────────────────────────────────

def _extract_isin(notes: str) -> str | None:
    m = re.search(r"ISIN:([A-Z]{2}[0-9A-Z]{10})", str(notes or ""))
    return m.group(1) if m else None


def _resolve_ticker(isin: str | None, symbol: str) -> str | None:
    if isin and isin in ISIN_TICKER_MAP:
        return ISIN_TICKER_MAP[isin]
    sym_up = symbol.upper()
    for fragment, ticker in SYMBOL_TICKER_MAP.items():
        if fragment.upper() in sym_up:
            return ticker
    return None


def _nan_to_none(v):
    """Convert NaN / NaT to None and numpy scalars to Python natives for sqlite3."""
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return v.item()          # numpy scalar → Python int / float / bool
    except AttributeError:
        return v


def _rows(df: pd.DataFrame, cols: list[str]):
    """Yield sqlite3-safe value tuples from a DataFrame."""
    for row in df[cols].itertuples(index=False, name=None):
        yield tuple(_nan_to_none(v) for v in row)


# ── Table creation ─────────────────────────────────────────────────────────────

def _create_tables(conn: sqlite3.Connection) -> None:
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS instruments (
            symbol       TEXT PRIMARY KEY,
            isin         TEXT,
            name         TEXT,
            asset_class  TEXT,
            yahoo_ticker TEXT,
            broker       TEXT
        );

        CREATE TABLE IF NOT EXISTS positions (
            date          TEXT NOT NULL,
            account       TEXT NOT NULL,
            symbol        TEXT NOT NULL,
            qty           REAL NOT NULL,
            avg_buy_price REAL NOT NULL,
            close_price   REAL,
            market_value  REAL,
            cost_basis    REAL,
            pnl_eur       REAL,
            pnl_pct       REAL,
            PRIMARY KEY (date, account, symbol)
        );

        CREATE TABLE IF NOT EXISTS portfolio_summary (
            date            TEXT NOT NULL,
            account         TEXT NOT NULL,
            total_value     REAL,
            cost_basis      REAL,
            pnl_eur         REAL,
            pnl_pct_overall REAL,
            ret_1d          REAL,
            ret_1w          REAL,
            ret_1m          REAL,
            ret_ytd         REAL,
            ret_all         REAL,
            PRIMARY KEY (date, account)
        );
    """)
    conn.commit()


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 1 — instruments
# ══════════════════════════════════════════════════════════════════════════════

def build_instruments(conn: sqlite3.Connection) -> None:
    raw = pd.read_sql(
        "SELECT symbol, asset_class, broker, notes FROM trades GROUP BY symbol",
        conn,
    )
    raw["isin"]         = raw["notes"].apply(_extract_isin)
    raw["yahoo_ticker"] = raw.apply(
        lambda r: _resolve_ticker(r["isin"], r["symbol"]), axis=1
    )
    raw["name"] = raw["symbol"]

    cols = ["symbol", "isin", "name", "asset_class", "yahoo_ticker", "broker"]
    conn.executemany(
        """
        INSERT OR REPLACE INTO instruments
            (symbol, isin, name, asset_class, yahoo_ticker, broker)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        list(_rows(raw, cols)),
    )
    conn.commit()
    print(f"[Step 1] instruments table:   {len(raw)} rows upserted")


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 2 — positions
# ══════════════════════════════════════════════════════════════════════════════

def _position_series(group: pd.DataFrame, today: str) -> pd.DataFrame:
    """
    For a single (account, symbol) trade group (sorted by date, id), returns
    a daily DataFrame with qty and avg_buy_price forward-filled from the first
    trade date to today.  Rows where qty == 0 (closed position) are dropped.

    BUY  → new_avg = (old_qty * old_avg + qty * price) / new_qty
    SELL → qty decreases; avg_buy_price unchanged
    """
    qty  = 0.0
    avg  = 0.0
    events: list[tuple[str, float, float]] = []

    for _, t in group.iterrows():
        if t["action"] == "BUY":
            new_qty = qty + t["qty"]
            if new_qty > 0:
                avg = (qty * avg + t["qty"] * t["price"]) / new_qty
            qty = new_qty
        elif t["action"] == "SELL":
            qty = max(0.0, qty - t["qty"])
        events.append((t["date"], qty, avg))

    ev = (
        pd.DataFrame(events, columns=["date", "qty", "avg_buy_price"])
        .groupby("date")
        .last()          # keep final state per date (handles multiple same-day trades)
        .reset_index()
    )

    date_idx = pd.date_range(group["date"].min(), today, freq="D")
    daily    = pd.DataFrame({"date": date_idx.strftime("%Y-%m-%d")})
    daily    = daily.merge(ev, on="date", how="left")
    daily[["qty", "avg_buy_price"]] = daily[["qty", "avg_buy_price"]].ffill()

    daily = daily.dropna(subset=["qty"])
    return daily[daily["qty"] > 0].copy()


def build_positions(conn: sqlite3.Connection, live: bool = False) -> None:
    today  = datetime.now().strftime("%Y-%m-%d")
    trades = pd.read_sql("SELECT * FROM trades ORDER BY date, id", conn)
    prices = pd.read_sql(
        "SELECT date, symbol AS price_key, close AS close_price FROM prices", conn
    )
    ticker_map: dict[str, str | None] = (
        pd.read_sql("SELECT symbol, yahoo_ticker FROM instruments", conn)
        .set_index("symbol")["yahoo_ticker"]
        .to_dict()
    )

    frames: list[pd.DataFrame] = []

    for (account, symbol), grp in trades.groupby(["account", "symbol"]):
        yahoo = ticker_map.get(symbol)
        if yahoo in EXCLUDE_FROM_POSITIONS:
            print(f"  {symbol:<30}  [SKIP] excluded from positions (prices still tracked)")
            continue
        daily = _position_series(grp, today)
        if daily.empty:
            continue
        daily["account"]     = account
        daily["symbol"]      = symbol
        daily["price_key"]   = yahoo         # None → NaN → no price match
        daily["asset_class"] = grp["asset_class"].iloc[0]
        frames.append(daily)

    if not frames:
        print("[Step 2] positions table:     0 rows upserted")
        return

    pos = pd.concat(frames, ignore_index=True)

    # Join prices (price_key = yahoo_ticker; NaN keys never match → NULL close_price)
    pos = pos.merge(prices, on=["date", "price_key"], how="left")

    if live:
        live_p = pd.read_sql(
            "SELECT symbol AS price_key, price AS live_price FROM live_prices", conn
        )
        if not live_p.empty:
            pos = pos.merge(live_p, on="price_key", how="left")
            today_mask = (pos["date"] == today) & pos["live_price"].notna()
            pos.loc[today_mask, "close_price"] = pos.loc[today_mask, "live_price"]
            pos.drop(columns=["live_price"], inplace=True)

    pos.drop(columns=["price_key"], inplace=True)

    # Bond prices are quoted as % of par value (e.g. 113.83 = 113.83% of face).
    # Divide by 100 to convert to EUR per unit of face value.
    price_scale = pos["asset_class"].map(lambda ac: 0.01 if ac == "BND" else 1.0)
    pos.drop(columns=["asset_class"], inplace=True)

    pos["cost_basis"]   = pos["qty"] * pos["avg_buy_price"] * price_scale
    pos["market_value"] = pos["qty"] * pos["close_price"]   * price_scale  # NaN on non-trading days
    pos["pnl_eur"]      = pos["market_value"] - pos["cost_basis"]
    pos["pnl_pct"]      = pos["pnl_eur"] / pos["cost_basis"]

    cols = [
        "date", "account", "symbol", "qty", "avg_buy_price",
        "close_price", "market_value", "cost_basis", "pnl_eur", "pnl_pct",
    ]
    conn.executemany(
        """
        INSERT OR REPLACE INTO positions
            (date, account, symbol, qty, avg_buy_price,
             close_price, market_value, cost_basis, pnl_eur, pnl_pct)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        list(_rows(pos, cols)),
    )
    conn.commit()
    print(f"[Step 2] positions table:     {len(pos):,} rows upserted")


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 2b — accrued_cash
# ══════════════════════════════════════════════════════════════════════════════

def build_accrued_cash(conn: sqlite3.Connection) -> None:
    """
    For every bond position with a coupon schedule in the `bonds` table, adds
    a synthetic CASH position 'ACCRUED_<isin>' per (date, account) whose value
    is the cumulative interest earned since the position was opened: past
    coupons received + the current period's live accrued interest. A paid
    coupon doesn't reset this to zero — it just converts accrued value into
    real cash, so the running total (and total_value) stays continuous.
    """
    bond_isins = {row[0] for row in conn.execute("SELECT isin FROM bonds")}
    if not bond_isins:
        print("[Step 2b] accrued_cash:      0 bonds registered, skipped")
        return

    bonds = pd.read_sql(
        "SELECT symbol, yahoo_ticker, broker FROM instruments WHERE asset_class = 'BND'",
        conn,
    )
    bonds = bonds[bonds["yahoo_ticker"].isin(bond_isins)]
    if bonds.empty:
        print("[Step 2b] accrued_cash:      0 matching bond positions, skipped")
        return

    positions = pd.read_sql(
        "SELECT date, account, symbol, qty FROM positions WHERE symbol IN ({})".format(
            ",".join("?" * len(bonds))
        ),
        conn,
        params=list(bonds["symbol"]),
    )

    instrument_rows = []
    position_rows = []

    for _, b in bonds.iterrows():
        isin, symbol = b["yahoo_ticker"], b["symbol"]
        accrued_symbol = f"ACCRUED_{isin}"
        instrument_rows.append(
            (accrued_symbol, isin, f"Accrued interest — {symbol}", "CASH", None, b["broker"])
        )

        sym_pos = positions[positions["symbol"] == symbol]
        for account, grp in sym_pos.groupby("account"):
            grp = grp.sort_values("date")
            dates = [date.fromisoformat(d) for d in grp["date"]]
            rates = daily_accrual_series(isin, dates)

            running = 0.0
            for d, qty, rate in zip(grp["date"], grp["qty"], rates):
                running += qty * rate / 100.0
                position_rows.append(
                    (d, account, accrued_symbol, 1.0, 0.0, running, running, 0.0, running, None)
                )

    conn.executemany(
        """
        INSERT OR REPLACE INTO instruments
            (symbol, isin, name, asset_class, yahoo_ticker, broker)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        instrument_rows,
    )
    conn.executemany(
        """
        INSERT OR REPLACE INTO positions
            (date, account, symbol, qty, avg_buy_price,
             close_price, market_value, cost_basis, pnl_eur, pnl_pct)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        position_rows,
    )
    conn.commit()
    print(
        f"[Step 2b] accrued_cash:      {len(position_rows):,} rows upserted "
        f"across {len(instrument_rows)} bond(s)"
    )


# ══════════════════════════════════════════════════════════════════════════════
#  STEP 3 — portfolio_summary
# ══════════════════════════════════════════════════════════════════════════════

def _equity_curve(grp: pd.DataFrame, today: str) -> pd.DataFrame:
    """
    Given a per-account positions slice (date, symbol, qty, close_price,
    market_value, cost_basis), returns a daily equity curve with price-only
    returns that are unaffected by capital flows (buys / sells).

    Daily return = weighted average of per-instrument price changes, where
    the weights are the *previous* day's qty.  Because the new units bought
    on trade day T enter the weight only from T+1 onwards, the capital
    injection never inflates the return figure for day T.  Multi-period
    returns (1W, 1M, YTD) are built by chaining these daily price returns.
    """
    # ── Forward-fill price and market_value per symbol across non-trading days ──
    work = grp.sort_values(["symbol", "date"]).copy()
    work["price_ff"] = work.groupby("symbol")["close_price"].ffill()
    work["mv_ff"]    = work.groupby("symbol")["market_value"].ffill()

    # ── Pivot to date × symbol matrices ────────────────────────────────────────
    kw = {"aggfunc": "last"}
    price_piv = work.pivot_table(index="date", columns="symbol", values="price_ff",  **kw)
    qty_piv   = work.pivot_table(index="date", columns="symbol", values="qty",       **kw)
    mv_piv    = work.pivot_table(index="date", columns="symbol", values="mv_ff",     **kw)
    cb_piv    = work.pivot_table(index="date", columns="symbol", values="cost_basis",**kw)

    # ── Price-only daily return ─────────────────────────────────────────────────
    # prev_qty × today_price   →  what yesterday's portfolio is worth today
    # prev_qty × yest_price    →  what yesterday's portfolio was worth yesterday
    # Ratio - 1  =  pure price return, capital flows excluded.
    # Only count (date, symbol) pairs where price exists on both consecutive days.
    price_prev = price_piv.shift(1)
    qty_prev   = qty_piv.shift(1)

    valid    = price_piv.notna() & price_prev.notna() & qty_prev.notna()
    hyp_cur  = (qty_prev * price_piv).where(valid)
    hyp_prev = (qty_prev * price_prev).where(valid)

    hyp_cur_sum  = hyp_cur.sum(axis=1)
    hyp_prev_sum = hyp_prev.sum(axis=1)
    daily_ret = (hyp_cur_sum / hyp_prev_sum - 1).where(hyp_prev_sum > 0)

    # ── Aggregate total_value and cost_basis ────────────────────────────────────
    priced_cb   = cb_piv.where(mv_piv.notna())
    total_value = mv_piv.sum(axis=1, min_count=1)
    cost_basis  = priced_cb.sum(axis=1, min_count=1)

    # ── Expand to all calendar days ─────────────────────────────────────────────
    full_dates = pd.date_range(grp["date"].min(), today, freq="D")
    full = pd.DataFrame({"date": full_dates.strftime("%Y-%m-%d")})
    base = pd.DataFrame({
        "date":        price_piv.index.astype(str),
        "total_value": total_value.values,
        "cost_basis":  cost_basis.values,
        "daily_ret":   daily_ret.values,
    })
    full = full.merge(base, on="date", how="left")
    full["total_value"] = full["total_value"].ffill()
    full["cost_basis"]  = full["cost_basis"].ffill()
    full["daily_ret"]   = full["daily_ret"].fillna(0)   # weekends / holidays → 0

    # ── P&L ─────────────────────────────────────────────────────────────────────
    full["pnl_eur"]         = full["total_value"] - full["cost_basis"]
    full["pnl_pct_overall"] = full["pnl_eur"] / full["cost_basis"]

    # ── Period returns via chained daily price returns ───────────────────────────
    growth = 1 + full["daily_ret"]
    full["ret_1d"] = full["daily_ret"]
    full["ret_1w"] = growth.rolling(5,  min_periods=5).apply(lambda x: x.prod(), raw=True) - 1
    full["ret_1m"] = growth.rolling(21, min_periods=21).apply(lambda x: x.prod(), raw=True) - 1

    # ── ret_ytd: chain from last trading day of the previous calendar year ───────
    full["_dt"]   = pd.to_datetime(full["date"])
    full["_year"] = full["_dt"].dt.year
    full["_cum"]  = growth.cumprod()

    prev_year_cum = (
        full.groupby("_year")["_cum"]
        .last()
        .rename("_prev_cum")
        .reset_index()
        .assign(_year=lambda df: df["_year"] + 1)
    )
    full = full.merge(prev_year_cum, on="_year", how="left")
    full["ret_ytd"] = (full["_cum"] / full["_prev_cum"] - 1).where(full["_prev_cum"].notna())

    # ── ret_all: vs very first non-null total_value ──────────────────────────────
    tv        = full["total_value"]
    first_val = tv.dropna().iloc[0] if tv.notna().any() else None
    full["ret_all"] = (tv / first_val - 1) if (first_val and first_val != 0) else None

    full.drop(columns=["daily_ret", "_dt", "_year", "_cum", "_prev_cum"],
              inplace=True, errors="ignore")
    return full


def build_portfolio_summary(conn: sqlite3.Connection) -> None:
    today     = datetime.now().strftime("%Y-%m-%d")
    positions = pd.read_sql(
        "SELECT date, account, symbol, qty, close_price, market_value, cost_basis FROM positions",
        conn,
    )

    if positions.empty:
        print("[Step 3] portfolio_summary:   0 rows upserted")
        return

    frames: list[pd.DataFrame] = []

    for account, grp in positions.groupby("account"):
        curve            = _equity_curve(grp.copy(), today)
        curve["account"] = account
        frames.append(curve)

    result = pd.concat(frames, ignore_index=True)

    cols = [
        "date", "account", "total_value", "cost_basis", "pnl_eur",
        "pnl_pct_overall", "ret_1d", "ret_1w", "ret_1m", "ret_ytd", "ret_all",
    ]
    conn.executemany(
        """
        INSERT OR REPLACE INTO portfolio_summary
            (date, account, total_value, cost_basis, pnl_eur,
             pnl_pct_overall, ret_1d, ret_1w, ret_1m, ret_ytd, ret_all)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        list(_rows(result, cols)),
    )
    conn.commit()
    print(f"[Step 3] portfolio_summary:   {len(result):,} rows upserted")


# ── Preview ────────────────────────────────────────────────────────────────────

def _fmt_val(v) -> str:
    return f"{v:>12,.2f}" if pd.notna(v) else f"{'N/A':>12}"


def _fmt_pct(v) -> str:
    if pd.isna(v):
        return f"{'N/A':>8}"
    sign = "+" if v >= 0 else ""
    return f"{sign}{v * 100:.2f}%"


def print_preview(conn: sqlite3.Connection) -> None:
    df = pd.read_sql(
        """
        SELECT account, date, total_value, pnl_pct_overall, ret_1d, ret_ytd
        FROM   portfolio_summary
        WHERE  date = (SELECT MAX(date) FROM portfolio_summary)
        ORDER  BY account
        """,
        conn,
    )
    if df.empty:
        print("\n(no portfolio_summary data to preview)")
        return

    sep = "-" * 72
    print(f"\n{sep}")
    print(
        f"  {'account':<16}  {'date':<12}"
        f"  {'total_value':>12}  {'pnl_pct':>8}"
        f"  {'ret_1d':>8}  {'ret_ytd':>8}"
    )
    print(sep)
    for _, row in df.iterrows():
        print(
            f"  {row['account']:<16}  {row['date']:<12}"
            f"  {_fmt_val(row['total_value'])}"
            f"  {_fmt_pct(row['pnl_pct_overall']):>8}"
            f"  {_fmt_pct(row['ret_1d']):>8}"
            f"  {_fmt_pct(row['ret_ytd']):>8}"
        )
    print(sep)


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--live", action="store_true")
    args = parser.parse_args()

    conn = sqlite3.connect(DB_PATH)
    _create_tables(conn)
    build_instruments(conn)
    build_positions(conn, live=args.live)
    build_accrued_cash(conn)
    build_portfolio_summary(conn)
    print_preview(conn)
    conn.close()


if __name__ == "__main__":
    main()
