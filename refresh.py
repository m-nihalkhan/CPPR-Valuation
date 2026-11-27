"""
Daily refresh for the ETF valuation model.

  1. Reads the tickers on the Holdings sheet of CPPR_Valuation_Model.xlsx
  2. Pulls the latest official close on each stock's home exchange, plus shares, EPS (trailing + consensus), balance sheet and EBITDA from Yahoo Finance (yfinance)
  3. Writes the raw numbers into the Data sheet (the Valuation/Dashboard formulas do the rest in Excel)
  4. Writes data/valuation.json with the same metrics, for the Lovable dashboard

Run:  pip install -r requirements.txt
      python refresh.py
"""
import datetime as dt
import json
import math
import pathlib
import time
from zoneinfo import ZoneInfo

import openpyxl
import yfinance as yf

import copper

ROOT = pathlib.Path(__file__).resolve().parent
WORKBOOK = ROOT / "CPPR_Valuation_Model.xlsx"
JSON_OUT = ROOT / "data" / "valuation.json"
FIRST = 5            # first data row on Holdings / Data
RETRIES = 3
PAUSE = 1.0          # seconds between tickers – be gentle with Yahoo

# Yahoo quotes some markets in sub-units (London in pence etc.)
SUBUNITS = {"GBp": ("GBP", 100), "GBX": ("GBP", 100), "ZAc": ("ZAR", 100), "ILA": ("ILS", 100)}

DATA_COLS = ["ticker", "status", "updated", "price", "ccy", "fin_ccy", "fx_fin_trade", "fx_trade_aud",
             "shares", "eps_ttm", "eps_fy0", "eps_fy1", "debt", "cash", "minority", "ebitda_ttm",
             "assets", "liabilities", "equity", "y_pe", "y_fpe", "y_ev_ebitda", "y_name",
             "stmt_date", "next_earnings", "close_date", "exchange"]
DATA_HEADERS = {"stmt_date": "Latest statement period end", "next_earnings": "Next earnings date (Yahoo)",
                "close_date": "Close date (home exchange)", "exchange": "Exchange"}

# Fallback local closing hour if Yahoo doesn't say whether the market is still open
DEFAULT_CLOSE_HOUR = 17.5


def last_close(t, info, h=None):
    """Official close of the most recent *completed* session on the stock's home exchange.

    Uses the daily price history (unadjusted). If today's bar exists but that exchange is
    still open (or hasn't reached its close), today's bar is dropped so we never use a
    mid-session price. Returns (close, 'YYYY-MM-DD') or (None, None)."""
    if h is None:
        try:
            h = t.history(period="10d", interval="1d", auto_adjust=False)
        except Exception:
            return None, None
    if h is None or h.empty or "Close" not in h:
        return None, None
    h = h.dropna(subset=["Close"])
    if h.empty:
        return None, None
    tzname = info.get("exchangeTimezoneName")
    state = str(info.get("marketState") or "").upper()
    if tzname:
        now_local = dt.datetime.now(ZoneInfo(tzname))
        last_day = h.index[-1].date()
        if last_day >= now_local.date():
            if state in ("POST", "POSTPOST", "CLOSED"):
                pass                                    # today's session finished – keep it
            elif state in ("REGULAR", "PRE", "PREPRE"):
                h = h.iloc[:-1]                         # still trading / not started – drop it
            elif now_local.hour + now_local.minute / 60 < DEFAULT_CLOSE_HOUR:
                h = h.iloc[:-1]
    if h.empty:
        return None, None
    return num(h["Close"].iloc[-1]), h.index[-1].date().isoformat()


SISTER_SUFFIXES = [(".TO", ".V"), (".V", ".TO"), (".V", ".CN"), (".TO", ".NE"), (".AX", ".NZ")]


def alternates(ticker):
    """Other exchange suffixes to try if a ticker returns nothing."""
    return [ticker[: -len(a)] + b for a, b in SISTER_SUFFIXES if ticker.endswith(a)]


def period_end(*dfs):
    """Most recent period-end date (YYYY-MM-DD) across yfinance statements, or None."""
    best = None
    for df in dfs:
        if df is None or getattr(df, "empty", True):
            continue
        try:
            d = max(df.columns).date().isoformat()
        except Exception:
            continue
        best = max(best, d) if best else d
    return best


def next_earnings_date(t):
    """Next scheduled earnings date from Yahoo's calendar, or None."""
    try:
        cal = t.calendar or {}
        dates = cal.get("Earnings Date") if isinstance(cal, dict) else None
        if dates:
            today = dt.date.today()
            fut = sorted(d for d in dates if hasattr(d, "isoformat") and d >= today)
            return fut[0].isoformat() if fut else None
    except Exception:
        pass
    return None

_fx_cache = {}


# --------------------------------------------------------------------------- helpers
def num(x):
    """float or None (drops NaN / inf / non-numbers)."""
    try:
        x = float(x)
    except (TypeError, ValueError):
        return None
    return x if math.isfinite(x) else None


def fx(a, b):
    """Units of currency b per 1 unit of currency a."""
    if not a or not b or a == b:
        return 1.0
    key = (a, b)
    if key not in _fx_cache:
        rate = None
        for sym, inv in ((f"{a}{b}=X", False), (f"{b}{a}=X", True)):
            try:
                p = num(yf.Ticker(sym).fast_info["last_price"])
                if p:
                    rate = 1 / p if inv else p
                    break
            except Exception:
                continue
        if rate is None:
            raise RuntimeError(f"No FX rate for {a}->{b}")
        _fx_cache[key] = rate
    return _fx_cache[key]


def latest(df, *names):
    """Most recent non-empty value of the first matching row label in a yfinance statement."""
    if df is None or getattr(df, "empty", True):
        return None
    for n in names:
        if n in df.index:
            s = df.loc[n].dropna()
            if len(s):
                return num(s.iloc[0])
    return None


def ttm(qdf, adf, *names):
    """Sum of the last 4 quarters; falls back to the latest annual figure."""
    if qdf is not None and not getattr(qdf, "empty", True):
        for n in names:
            if n in qdf.index:
                s = qdf.loc[n].dropna()
                if len(s) >= 4:
                    return float(s.iloc[:4].sum())
    return latest(adf, *names)


def eps_to_trading(eps, fx_ft, price, ref_pe):
    """Yahoo is inconsistent about whether EPS is in trading or financial currency for
    cross-listed stocks. Use Yahoo's own P/E (when available) to decide; otherwise assume
    the EPS is in the financial currency."""
    if eps is None:
        return None
    if fx_ft == 1:
        return eps
    if ref_pe and price and eps > 0:
        as_is, converted = price / eps, price / (eps * fx_ft)
        if abs(math.log(as_is / ref_pe)) <= abs(math.log(converted / ref_pe)):
            return eps
    return eps * fx_ft


# --------------------------------------------------------------------------- fetch
def fetch(ticker):
    t = yf.Ticker(ticker)
    info = t.info or {}
    try:   # 2 years of daily prices: latest close + weekly closes for the copper beta
        hist = t.history(period="2y", interval="1d", auto_adjust=False)
    except Exception:
        hist = None
    price, close_date = last_close(t, info, hist)
    if price is None:   # no history – fall back to Yahoo's regular-session price
        price = num(info.get("regularMarketPreviousClose")) if str(info.get("marketState") or "").upper() == "REGULAR" \
            else (num(info.get("regularMarketPrice")) or num(info.get("currentPrice")))
        close_date = None
    if price is None:
        raise RuntimeError("no price")
    ccy = info.get("currency") or t.fast_info.get("currency")
    if ccy in SUBUNITS:
        ccy, div = SUBUNITS[ccy]
        price /= div
    fin_ccy = info.get("financialCurrency") or ccy
    if fin_ccy in SUBUNITS:
        fin_ccy = SUBUNITS[fin_ccy][0]
    fx_ft = fx(fin_ccy, ccy)
    fx_aud = fx(ccy, "AUD")

    bs = t.quarterly_balance_sheet
    if bs is None or bs.empty:
        bs = t.balance_sheet
    shares = (num(info.get("sharesOutstanding")) or num(info.get("impliedSharesOutstanding"))
              or latest(bs, "Ordinary Shares Number", "Share Issued"))

    y_pe, y_fpe = num(info.get("trailingPE")), num(info.get("forwardPE"))
    eps_ttm = eps_to_trading(num(info.get("trailingEps")), fx_ft, price, y_pe)

    eps_fy0 = eps_fy1 = None
    try:
        est = t.earnings_estimate
        if est is not None and not est.empty and "avg" in est.columns:
            raw0 = num(est.loc["0y", "avg"]) if "0y" in est.index else None
            raw1 = num(est.loc["+1y", "avg"]) if "+1y" in est.index else None
            ref = raw1 if raw1 else raw0
            conv = eps_to_trading(ref, fx_ft, price, y_fpe)
            factor = (conv / ref) if (ref and conv) else fx_ft
            eps_fy0 = raw0 * factor if raw0 is not None else None
            eps_fy1 = raw1 * factor if raw1 is not None else None
    except Exception:
        pass

    q_is, a_is = t.quarterly_income_stmt, t.income_stmt
    ebitda = ttm(q_is, a_is, "EBITDA", "Normalized EBITDA")
    if ebitda is None:
        ebitda = num(info.get("ebitda"))

    return {
        "price": price, "ccy": ccy, "fin_ccy": fin_ccy, "fx_fin_trade": fx_ft, "fx_trade_aud": fx_aud,
        "shares": shares, "eps_ttm": eps_ttm, "eps_fy0": eps_fy0, "eps_fy1": eps_fy1,
        "debt": latest(bs, "Total Debt"),
        "cash": latest(bs, "Cash Cash Equivalents And Short Term Investments", "Cash And Cash Equivalents"),
        "minority": latest(bs, "Minority Interest"),
        "ebitda_ttm": ebitda,
        "assets": latest(bs, "Total Assets"),
        "liabilities": latest(bs, "Total Liabilities Net Minority Interest"),
        "equity": latest(bs, "Stockholders Equity", "Common Stock Equity"),
        "y_pe": y_pe, "y_fpe": y_fpe, "y_ev_ebitda": num(info.get("enterpriseToEbitda")),
        "y_name": info.get("longName") or info.get("shortName"),
        "stmt_date": period_end(bs, q_is, a_is),
        "next_earnings": next_earnings_date(t),
        "close_date": close_date,
        "exchange": info.get("fullExchangeName") or info.get("exchange"),
        "weekly": copper.weekly_closes(hist),
    }


def fetch_copper():
    """Latest completed copper futures close (US$/lb) and weekly closes."""
    for attempt in range(RETRIES):
        try:
            t = yf.Ticker(copper.COPPER_TICKER)
            info = t.info or {}
            hist = t.history(period="2y", interval="1d", auto_adjust=False)
            close, date = last_close(t, info, hist)
            return close, date, copper.weekly_closes(hist)
        except Exception as e:
            print(f"Copper fetch failed ({e}); retrying")
            time.sleep(PAUSE * (attempt + 2) * 2)
    return None, None, None


# --------------------------------------------------------------------------- metrics (mirror of the Valuation sheet formulas)
def metrics(d, weight, broker_nav=None):
    def ratio(a, b):
        if a is None or b is None:
            return None
        return a / b if b != 0 else "NM"   # negative = loss-making, shown as a real number

    out = {"weight": weight}
    if not d or d.get("price") is None:
        return out
    p, f, aud, sh = d["price"], d["fx_fin_trade"], d["fx_trade_aud"], d.get("shares")
    eq = p * sh / 1e6 if sh else None
    nd = ((d.get("debt") or 0) - (d.get("cash") or 0)) * f / 1e6
    mi = (d.get("minority") or 0) * f / 1e6
    ev = eq + nd + mi if eq is not None else None
    ebitda = d["ebitda_ttm"] * f / 1e6 if d.get("ebitda_ttm") is not None else None
    nav = d["equity"] * f / sh if (d.get("equity") is not None and sh) else None
    out.update({
        "price": p, "currency": d["ccy"], "price_aud": p * aud if aud else None, "fx_to_aud": aud,
        "equity_value_m": eq, "net_debt_m": nd, "minority_m": mi, "enterprise_value_m": ev,
        "ebitda_ttm_m": ebitda, "ev_ebitda": ratio(ev, ebitda) if ev is not None else None,
        "pe_trailing": ratio(p, d.get("eps_ttm")), "pe_current": ratio(p, d.get("eps_fy0")),
        "pe_forward": ratio(p, d.get("eps_fy1")),
        "nav_per_share": nav, "p_nav": ratio(p, nav),
        "broker_nav": broker_nav, "p_broker_nav": ratio(p, broker_nav) if broker_nav else None,
        "equity_value_aud_m": eq * aud if eq is not None else None,
        "enterprise_value_aud_m": ev * aud if ev is not None else None,
        "yahoo_pe": d.get("y_pe"), "yahoo_forward_pe": d.get("y_fpe"), "yahoo_ev_ebitda": d.get("y_ev_ebitda"),
    })
    return out


def harmonic(rows, key):
    w = [(r["weight"], r.get(key)) for r in rows if isinstance(r.get(key), (int, float)) and r.get(key) > 0]
    tw = sum(a for a, _ in w)
    return (tw / sum(a / m for a, m in w) if w else None), tw


# --------------------------------------------------------------------------- Data sheet formats
MIL = '#,##0.00,,"m";(#,##0.00,,"m");"-"'          # raw units shown in millions
NUM2 = '#,##0.00;(#,##0.00);"-"'
MULT = '0.00"x";[Red]-0.00"x";"-"'
FX4 = "0.0000"
DATA_FORMATS = {"price": NUM2, "fx_fin_trade": FX4, "fx_trade_aud": FX4, "shares": MIL,
                "eps_ttm": NUM2, "eps_fy0": NUM2, "eps_fy1": NUM2,
                "debt": MIL, "cash": MIL, "minority": MIL, "ebitda_ttm": MIL,
                "assets": MIL, "liabilities": MIL, "equity": MIL,
                "y_pe": MULT, "y_fpe": MULT, "y_ev_ebitda": MULT}


def style_data_sheet(ws, n):
    """Re-apply number formats every run, so the Data sheet always reads cleanly."""
    from openpyxl.styles import Alignment
    for i in range(n + 5):
        r = FIRST + i
        for c, k in enumerate(DATA_COLS, 1):
            cell = ws.cell(r, c)
            if k in DATA_FORMATS:
                cell.number_format = DATA_FORMATS[k]
                cell.alignment = Alignment(horizontal="right")


# --------------------------------------------------------------------------- main
def main():
    wb = openpyxl.load_workbook(WORKBOOK)
    ho, da, li = wb["Holdings"], wb["Data"], wb["Lists"]

    holdings = []
    for r in range(FIRST, ho.max_row + 1):
        etf, name, tk = ho.cell(r, 1).value, ho.cell(r, 2).value, ho.cell(r, 3).value
        if name and tk:
            holdings.append({"etf": str(etf).strip(), "company": name, "ticker": str(tk).strip(),
                             "country": ho.cell(r, 4).value, "mv": num(ho.cell(r, 6).value) or 0,
                             "broker_nav": num(ho.cell(r, 8).value), "broker_src": ho.cell(r, 9).value})

    # previous values, so a failed ticker keeps its last good data
    prev = {}
    for r in range(FIRST, da.max_row + 1):
        tk = da.cell(r, 1).value
        if tk:
            prev[tk] = {k: da.cell(r, i + 1).value for i, k in enumerate(DATA_COLS)}

    now = dt.datetime.now(ZoneInfo("Australia/Sydney"))
    stamp = now.strftime("%Y-%m-%d %H:%M %Z")
    rows, seen = {}, set()
    for h in holdings:
        tk = h["ticker"]
        if tk in seen:
            continue
        seen.add(tk)
        rec, err, used = None, None, tk
        for attempt in range(RETRIES):
            try:
                rec = fetch(tk)
                break
            except Exception as e:  # rate limit, bad ticker, missing data…
                err = e
                time.sleep(PAUSE * (attempt + 2) * 2)
        if not rec:   # wrong exchange suffix? try the sister listing (e.g. .TO <-> .V)
            for alt in alternates(tk):
                try:
                    rec, used = fetch(alt), alt
                    break
                except Exception as e:
                    err = e
        if rec:
            rec.update(ticker=tk, status="OK" if used == tk else f"OK (via {used} – update Holdings)",
                       updated=stamp)
            old = (prev.get(tk) or {}).get("stmt_date")
            # New financial statements picked up since the last run?
            rec["new_results"] = bool(old and rec.get("stmt_date") and str(rec["stmt_date"]) > str(old))
        elif tk in prev and isinstance(prev[tk].get("price"), (int, float)):
            rec = dict(prev[tk])
            rec["status"] = f"Stale – last good data {prev[tk].get('updated')}"
        else:
            rec = {"ticker": tk, "status": f"No data – check ticker ({type(err).__name__})", "updated": stamp}
        rows[tk] = rec
        print(f"{tk:12s} {rec['status']}")
        time.sleep(PAUSE)

    # ---- write Data sheet
    for r in range(FIRST, max(da.max_row, FIRST) + 1):
        for c in range(1, len(DATA_COLS) + 1):
            da.cell(r, c).value = None
    for i, tk in enumerate(rows):
        for c, k in enumerate(DATA_COLS, 1):
            da.cell(FIRST + i, c, rows[tk].get(k))
    da["C2"] = stamp
    style_data_sheet(da, len(rows))
    for c, k in enumerate(DATA_COLS, 1):
        if k in DATA_HEADERS and not da.cell(FIRST - 1, c).value:
            da.cell(FIRST - 1, c, DATA_HEADERS[k])

    # ---- ETF list for the dropdown
    etfs = sorted({h["etf"] for h in holdings})
    for r in range(5, 55):
        li.cell(r, 1).value = None
    for i, e in enumerate(etfs):
        li.cell(5 + i, 1, e)

    # ---- copper beta: weekly history, regression, Copper History / Returns sheets
    prev_weekly = copper.read_previous_weekly(wb)
    prev_close = wb[copper.HIST_SHEET]["C2"].value if copper.HIST_SHEET in wb.sheetnames else None
    cu_close, cu_date, cu_weekly = fetch_copper()
    if cu_weekly is None:
        cu_weekly = prev_weekly.get(copper.COPPER_TICKER)
    if cu_close is None and isinstance(prev_close, (int, float)):
        cu_close = prev_close
        cu_date = wb[copper.HIST_SHEET]["E2"].value
    stock_weekly = {tk: rec.get("weekly") if rec.get("weekly") is not None else prev_weekly.get(tk)
                    for tk, rec in rows.items()}
    prices = copper.price_table(cu_weekly, stock_weekly, list(rows), dt.datetime.now(dt.timezone.utc))
    rets = copper.returns_table(prices)
    betas = {tk: copper.beta_stats(rets[tk], rets[copper.COPPER_TICKER]) for tk in rows}
    copper.write_sheets(wb, prices, cu_close, cu_date, stamp)
    copper.init_scenario(wb, cu_close)
    print(f"Copper: {cu_close} US$/lb ({cu_date}); betas for "
          f"{sum(1 for b in betas.values() if b['copper_beta'] is not None)}/{len(betas)} tickers")

    wb.calculation.fullCalcOnLoad = True   # Excel recalculates everything on open
    wb.save(WORKBOOK)

    # ---- JSON for the Lovable dashboard
    out = {"updated": stamp, "source": "Yahoo Finance via yfinance", "etfs": {}}
    for e in etfs:
        hs = [h for h in holdings if h["etf"] == e]
        tot = sum(h["mv"] for h in hs) or 1
        items = []
        for h in hs:
            d = rows.get(h["ticker"])
            m = metrics(d if d and isinstance(d.get("price"), (int, float)) else None, h["mv"] / tot, h["broker_nav"])
            m.update(company=h["company"], ticker=h["ticker"], country=h["country"],
                     broker_nav_source=h["broker_src"], status=(d or {}).get("status"),
                     last_updated=(d or {}).get("updated"),
                     statement_date=(d or {}).get("stmt_date"),
                     next_earnings_date=(d or {}).get("next_earnings"),
                     **betas.get(h["ticker"], {}),
                     close_date=(d or {}).get("close_date"),
                     exchange=(d or {}).get("exchange"),
                     new_results=bool((d or {}).get("new_results")))
            items.append(m)
        summary = {"holdings": len(items),
                   "with_data": sum(1 for x in items if x.get("price") is not None),
                   "loss_making": sum(1 for x in items if isinstance(x.get("pe_trailing"), (int, float)) and x["pe_trailing"] < 0),
                   "market_value_aud": sum(h["mv"] for h in hs)}
        for key in ("pe_trailing", "pe_current", "pe_forward", "ev_ebitda", "p_nav"):
            summary[f"{key}_weighted"], summary[f"{key}_coverage"] = harmonic(items, key)
        horizon = (now.date() + dt.timedelta(days=14)).isoformat()
        summary["new_results"] = [x["company"] for x in items if x.get("new_results")]
        summary["upcoming_earnings"] = sorted(
            ({"company": x["company"], "date": str(x["next_earnings_date"])} for x in items
             if x.get("next_earnings_date") and str(x["next_earnings_date"]) <= horizon),
            key=lambda r: r["date"])
        summary.update(copper.etf_summary(items, cu_close, cu_date, prices))
        out["etfs"][e] = {"summary": summary, "holdings": items}
    JSON_OUT.parent.mkdir(exist_ok=True)
    JSON_OUT.write_text(json.dumps(out, indent=2, default=str))
    ok = sum(1 for r in rows.values() if r["status"] == "OK")
    print(f"\nDone: {ok}/{len(rows)} tickers refreshed at {stamp}")


if __name__ == "__main__":
    main()
