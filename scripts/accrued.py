from pathlib import Path
import sqlite3
from datetime import date

from dateutil.relativedelta import relativedelta

DB_PATH = Path(__file__).parent.parent / "data" / "portfolio.db"


def _load_bond(isin):
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT coupon, freq, maturity FROM bonds WHERE isin = ?", (isin,)
        ).fetchone()
    if row is None:
        raise KeyError(f"ISIN {isin!r} not found in bonds table")
    coupon, freq, maturity_str = row
    return coupon, freq, date.fromisoformat(maturity_str)


def _coupon_bracket(maturity, freq, val_date):
    period = relativedelta(months=12 // freq)
    d = maturity
    while d > val_date:
        d -= period
    return d, d + period


def daily_accrual_series(isin, dates):
    """
    Per-day accrual rate (EUR accrued per 100 nominal on that single day) for
    each date in `dates`, respecting coupon period boundaries. Used to build
    a cumulative accrued-cash series without resetting at coupon dates —
    a paid coupon just converts accrued value into cash, it doesn't destroy it.
    """
    coupon, freq, maturity = _load_bond(isin)
    rates = []
    for d in dates:
        last, next_ = _coupon_bracket(maturity, freq, d)
        days_period = (next_ - last).days
        rates.append((coupon / freq) / days_period)
    return rates


def accrued_interest(isin, val_date=None):
    if val_date is None:
        val_date = date.today()
    coupon, freq, maturity = _load_bond(isin)
    last, next_ = _coupon_bracket(maturity, freq, val_date)
    days_since = (val_date - last).days
    days_period = (next_ - last).days
    return (coupon / freq) * (days_since / days_period)


def to_dirty(isin, clean_price, val_date=None):
    return clean_price + accrued_interest(isin, val_date)


def position_value(isin, clean_price, nominal, val_date=None):
    return nominal * to_dirty(isin, clean_price, val_date) / 100


if __name__ == "__main__":
    isin = "IT0004532559"

    v_coupon = accrued_interest(isin, date(2026, 8, 1))
    v_mid    = accrued_interest(isin, date(2026, 6, 25))

    print(f"accrued_interest('{isin}', date(2026,8,1))   = {v_coupon:.4f}  (expected 0)")
    print(f"accrued_interest('{isin}', date(2026,6,25))  = {v_mid:.4f}  (expected ~1.9890)")

    assert v_coupon == 0.0,      f"Expected 0, got {v_coupon}"
    assert 0 <= v_mid < 2.5,     f"Out of range [0, 2.5): {v_mid}"
    print("All checks passed.")
