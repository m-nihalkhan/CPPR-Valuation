"""
Copper price sensitivity (copper beta) for the ETF valuation model.

Method (also documented on the Copper Beta sheet and in README):
  * Weekly closes (Friday) over the last 2 years: 105 prices -> 104 weekly returns.
  * Stock returns are in each stock's own trading currency (dividend-adjusted closes);
    copper is COMEX copper futures (Yahoo: HG=F) in US$/lb. Using local-currency returns
    keeps exchange-rate moves out of the copper beta.
  * Copper beta = OLS slope of stock returns on copper returns (= Excel SLOPE), using only
    weeks where both returns exist. R², correlation, observations, standard error and a
    95% confidence interval are reported. Fewer than 52 paired weeks -> no beta.
  * Scenario: estimated share-price move = beta x copper price change. Everything else
    (EPS, EBITDA, net debt, book NAV, FX, shares) is held constant: this is a market-price
    sensitivity, not a fundamental earnings forecast.

refresh.py calls these functions each morning; the Excel sheets use the same maths in
live formulas (SLOPE, RSQ, CORREL, TINV), so the workbook and the email always agree.
"""
import datetime as dt
import math

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

COPPER_TICKER = "HG=F"
WEEKS = 104                  # weekly returns used (≈ 2 years)
MIN_OBS = 52                 # fewer paired weeks than this -> no beta
WEAK_R2 = 0.10               # below this the relationship is flagged as weak
HIST_SHEET, RET_SHEET = "Copper History", "Copper Returns"
SENS_SHEET = "Copper Sensitivity"
FIRST_ROW = 5                # first data row on the history / returns sheets

NAVY = "1F3864"
HDR_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor=NAVY)
BODY = Font(name="Arial", size=10)


# --------------------------------------------------------------------------- data
def weekly_closes(hist):
    """Friday-ending weekly closes from a daily yfinance history (dividend-adjusted)."""
    if hist is None or getattr(hist, "empty", True):
        return None
    col = "Adj Close" if "Adj Close" in hist and hist["Adj Close"].notna().any() else "Close"
    s = hist[col].dropna()
    if s.empty:
        return None
    idx = s.index
    if getattr(idx, "tz", None) is not None:
        idx = idx.tz_localize(None)          # keep each exchange's own calendar date
    s = pd.Series(s.values, index=pd.DatetimeIndex(idx).normalize())
    return s.resample("W-FRI").last().dropna()


def last_complete_friday(now_utc):
    """Most recent Friday whose trading week has fully closed (US close ≈ 21:00 UTC)."""
    d = now_utc.date()
    if d.weekday() == 4 and now_utc.hour >= 21:
        return d
    back = (d.weekday() - 4) % 7 or 7
    return d - dt.timedelta(days=back)


def price_table(copper_weekly, stock_weekly, tickers, now_utc):
    """Regular grid of the last WEEKS+1 completed Fridays; columns copper + tickers."""
    end = pd.Timestamp(last_complete_friday(now_utc))
    grid = pd.DatetimeIndex([end - pd.Timedelta(weeks=k) for k in range(WEEKS, -1, -1)])
    cols = {COPPER_TICKER: copper_weekly}
    cols.update({tk: stock_weekly.get(tk) for tk in tickers})
    table = pd.DataFrame(index=grid)
    for name, s in cols.items():
        table[name] = s.reindex(grid) if s is not None else float("nan")
    return table


def returns_table(prices):
    """Week-on-week returns; blank where either week is missing (same as the Excel sheet)."""
    prev = prices.shift(1)
    r = prices / prev - 1
    r[(prev <= 0) | prev.isna() | prices.isna()] = float("nan")
    return r.iloc[1:]


def t_inv_two_tailed(df, alpha=0.05):
    """Two-tailed Student-t critical value (≈ Excel TINV), Cornish-Fisher expansion."""
    if df <= 0:
        return None
    z = 1.959963984540054 if alpha == 0.05 else None
    g1 = (z ** 3 + z) / 4
    g2 = (5 * z ** 5 + 16 * z ** 3 + 3 * z) / 96
    g3 = (3 * z ** 7 + 19 * z ** 5 + 17 * z ** 3 - 15 * z) / 384
    g4 = (79 * z ** 9 + 776 * z ** 7 + 1482 * z ** 5 - 1920 * z ** 3 - 945 * z) / 92160
    return z + g1 / df + g2 / df ** 2 + g3 / df ** 3 + g4 / df ** 4


def beta_stats(stock_ret, copper_ret):
    """OLS copper beta and diagnostics for one stock (mirrors the Copper Beta sheet)."""
    m = stock_ret.notna() & copper_ret.notna()
    y, x = stock_ret[m].astype(float), copper_ret[m].astype(float)
    n = int(m.sum())
    out = {"copper_obs": n, "copper_beta": None, "copper_r2": None, "copper_corr": None,
           "copper_beta_se": None, "copper_beta_low": None, "copper_beta_high": None}
    if n < 3 or x.var() == 0 or y.var() == 0:
        out["copper_flag"] = f"Insufficient data (<{MIN_OBS} weeks)" if n < MIN_OBS else "No variation"
        return out
    beta = float(((x - x.mean()) * (y - y.mean())).sum() / ((x - x.mean()) ** 2).sum())
    corr = float(x.corr(y))
    r2 = corr ** 2
    if n < MIN_OBS:
        out.update(copper_r2=r2, copper_corr=corr, copper_flag=f"Insufficient data (<{MIN_OBS} weeks)")
        return out
    se = abs(beta) * math.sqrt((1 / r2 - 1) / (n - 2)) if r2 > 0 else None
    t = t_inv_two_tailed(n - 2)
    out.update(copper_beta=beta, copper_r2=r2, copper_corr=corr, copper_beta_se=se,
               copper_beta_low=beta - t * se if se is not None else None,
               copper_beta_high=beta + t * se if se is not None else None,
               copper_flag="Weak (R² < 0.10)" if r2 < WEAK_R2 else "OK")
    return out


def scenario_factor(beta, change):
    """Price multiplier for a copper move; a share price can't fall below zero."""
    return max(0.0, 1 + beta * change)


# --------------------------------------------------------------------------- previous data
def read_previous_weekly(wb):
    """Weekly prices already stored on the Copper History sheet (fallback for failed tickers)."""
    if HIST_SHEET not in wb.sheetnames:
        return {}
    ws = wb[HIST_SHEET]
    heads = {c: ws.cell(FIRST_ROW - 1, c).value for c in range(2, ws.max_column + 1)}
    dates = []
    for r in range(FIRST_ROW, ws.max_row + 1):
        v = ws.cell(r, 1).value
        dates.append((r, pd.Timestamp(v) if v else None))
    out = {}
    for c, name in heads.items():
        if not name:
            continue
        vals = {d: ws.cell(r, c).value for r, d in dates if d is not None}
        s = pd.Series({d: v for d, v in vals.items() if isinstance(v, (int, float))}, dtype=float)
        if not s.empty:
            key = COPPER_TICKER if str(name).startswith("Copper") else str(name)
            out[key] = s.sort_index()
    return out


# --------------------------------------------------------------------------- workbook
def _sheet(wb, name, after):
    if name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for c in row:
                c.value = None
        return ws
    ws = wb.create_sheet(name, index=wb.sheetnames.index(after) + 1)
    return ws


def _header(ws, row, values):
    ws.row_dimensions[row].height = 32
    for c, v in enumerate(values, 1):
        cell = ws.cell(row, c, v)
        cell.font, cell.fill = HDR_FONT, HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_sheets(wb, prices, copper_close, copper_date, stamp):
    """Rewrite the Copper History (values) and Copper Returns (formulas) sheets."""
    anchor = "Copper Beta" if "Copper Beta" in wb.sheetnames else "Valuation"
    hs = _sheet(wb, HIST_SHEET, anchor)
    rs = _sheet(wb, RET_SHEET, HIST_SHEET)
    names = list(prices.columns)

    hs["A1"] = "Copper History – weekly closes (written by refresh.py – do not edit)"
    hs["A1"].font = Font(name="Arial", size=16, bold=True, color=NAVY)
    hs["A2"] = "Latest copper close (US$/lb):"
    hs["C2"] = copper_close
    hs["D2"] = "Close date:"
    hs["E2"] = copper_date
    hs["G2"] = (f"Friday closes, last {WEEKS + 1} weeks. Copper = COMEX futures (HG=F, US$/lb); "
                f"stocks in their own trading currency, dividend-adjusted. Refreshed {stamp}.")
    for ref in ("A2", "D2"):
        hs[ref].font = Font(name="Arial", size=10, bold=True)
    hs["C2"].number_format = '"US$"0.0000'
    hs["G2"].font = Font(name="Arial", size=9, italic=True, color="595959")
    _header(hs, FIRST_ROW - 1, ["Week ending", "Copper (HG=F)"] + names[1:])
    for i, (d, row) in enumerate(prices.iterrows()):
        r = FIRST_ROW + i
        hs.cell(r, 1, d.date()).number_format = "dd-mmm-yy"
        for c, name in enumerate(names, 2):
            v = row[name]
            cell = hs.cell(r, c, None if pd.isna(v) else float(v))
            cell.number_format = "#,##0.0000"
            cell.font = BODY
    hs.freeze_panes = "B5"
    hs.column_dimensions["A"].width = 12
    for c in range(2, len(names) + 2):
        hs.column_dimensions[hs.cell(4, c).column_letter].width = 11

    rs["A1"] = "Copper Returns – week-on-week % change (formulas on Copper History)"
    rs["A1"].font = Font(name="Arial", size=16, bold=True, color=NAVY)
    rs["A2"] = "Blank where either week's price is missing. Column B is copper; the Copper Beta sheet regresses each stock column on it."
    rs["A2"].font = Font(name="Arial", size=9, italic=True, color="595959")
    _header(rs, FIRST_ROW - 1, ["Week ending", "Copper (HG=F)"] + names[1:])
    q = "'" + HIST_SHEET + "'!"
    for i in range(len(prices) - 1):
        r, src, prv = FIRST_ROW + i, FIRST_ROW + i + 1, FIRST_ROW + i
        rs.cell(r, 1, f"={q}A{src}").number_format = "dd-mmm-yy"
        for c in range(2, len(names) + 2):
            L = rs.cell(r, c).column_letter
            cell = rs.cell(r, c, f'=IF(AND(ISNUMBER({q}{L}{src}),ISNUMBER({q}{L}{prv}),N({q}{L}{prv})>0),'
                                 f'{q}{L}{src}/{q}{L}{prv}-1,"")')
            cell.number_format = '0.00%;[Red]-0.00%'
            cell.font = BODY
    rs.freeze_panes = "B5"
    rs.column_dimensions["A"].width = 12
    for c in range(2, len(names) + 2):
        rs.column_dimensions[rs.cell(4, c).column_letter].width = 11
    for ws in (hs, rs):
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = "B8692E"


def init_scenario(wb, copper_close):
    """Give the what-if copper price a starting value once; never overwrite the user's input."""
    if SENS_SHEET in wb.sheetnames and copper_close:
        cell = wb[SENS_SHEET]["C7"]
        if not isinstance(cell.value, (int, float)):
            cell.value = round(copper_close * 1.10, 2)


# --------------------------------------------------------------------------- ETF summary
def etf_summary(items, copper_close, copper_date, prices):
    """Weighted copper beta (arithmetic, covered holdings re-weighted) and ETF scenarios."""
    cov = [(x["weight"], x["copper_beta"]) for x in items if isinstance(x.get("copper_beta"), (int, float))]
    covered = sum(w for w, _ in cov)
    beta = sum(w * b for w, b in cov) / covered if covered else None
    cu = prices[COPPER_TICKER].dropna()
    wk = float(cu.iloc[-1] / cu.iloc[-2] - 1) if len(cu) >= 2 else None
    scen = {}
    for shock in (-0.20, -0.10, 0.10, 0.20):
        scen[f"{shock:+.0%}"] = beta * shock if beta is not None else None
    ranked = sorted((x for x in items if isinstance(x.get("copper_beta"), (int, float))),
                    key=lambda x: x["copper_beta"], reverse=True)
    return {"copper_price": copper_close, "copper_price_date": copper_date, "copper_week_change": wk,
            "copper_beta_weighted": beta, "copper_beta_coverage": covered,
            "copper_scenarios": scen,
            "copper_top_beta": [{"company": x["company"], "beta": x["copper_beta"], "r2": x.get("copper_r2"),
                                 "flag": x.get("copper_flag")} for x in ranked[:5]]}
