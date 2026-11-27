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
SHOCKS = (-0.30, -0.20, -0.10, -0.05, 0.05, 0.10, 0.20, 0.30)   # copper moves shown in the email
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


def set_whatif(wb, value):
    """Apply the WHATIF_COPPER_PRICE setting (GitHub variable) to the yellow cell, if given.
    Returns the what-if price now in the workbook (or None)."""
    if SENS_SHEET not in wb.sheetnames:
        return None
    cell = wb[SENS_SHEET]["C7"]
    try:
        v = float(str(value).replace("$", "").replace("US", "").strip()) if value not in (None, "") else None
    except ValueError:
        print(f"WHATIF_COPPER_PRICE '{value}' is not a number – ignored")
        v = None
    if v is not None and v > 0:
        cell.value = round(v, 4)
    return cell.value if isinstance(cell.value, (int, float)) else None


def whatif_summary(items, etf_beta, copper_close, whatif_price):
    """ETF and holding moves at the what-if copper price (same maths as the Copper Sensitivity sheet)."""
    if not (isinstance(whatif_price, (int, float)) and isinstance(copper_close, (int, float)) and copper_close > 0):
        return None
    change = whatif_price / copper_close - 1
    moves = []
    for x in items:
        b = x.get("copper_beta")
        if isinstance(b, (int, float)):
            f = scenario_factor(b, change)
            x["copper_whatif_move"] = f - 1
            p = x.get("price_aud")
            x["copper_whatif_price_aud"] = p * f if isinstance(p, (int, float)) else None
            moves.append(x)
        else:
            x["copper_whatif_move"] = None
            x["copper_whatif_price_aud"] = None
    ranked = sorted(moves, key=lambda x: abs(x["copper_whatif_move"]), reverse=True)[:5]
    return {"price": whatif_price, "change": change,
            "etf_move": etf_beta * change if isinstance(etf_beta, (int, float)) else None,
            "most_affected": [{"company": x["company"], "beta": x["copper_beta"], "move": x["copper_whatif_move"],
                               "price_aud": x.get("price_aud"), "whatif_price_aud": x["copper_whatif_price_aud"]}
                              for x in ranked]}


# --------------------------------------------------------------------------- ETF summary
def etf_summary(items, copper_close, copper_date, prices):
    """Weighted copper beta (arithmetic, covered holdings re-weighted) and ETF scenarios."""
    cov = [(x["weight"], x["copper_beta"]) for x in items if isinstance(x.get("copper_beta"), (int, float))]
    covered = sum(w for w, _ in cov)
    beta = sum(w * b for w, b in cov) / covered if covered else None
    cu = prices[COPPER_TICKER].dropna()
    wk = float(cu.iloc[-1] / cu.iloc[-2] - 1) if len(cu) >= 2 else None
    scen = {}
    for shock in SHOCKS:
        scen[f"{shock:+.0%}"] = beta * shock if beta is not None else None
    ranked = sorted((x for x in items if isinstance(x.get("copper_beta"), (int, float))),
                    key=lambda x: x["copper_beta"], reverse=True)
    return {"copper_price": copper_close, "copper_price_date": copper_date, "copper_week_change": wk,
            "copper_beta_weighted": beta, "copper_beta_coverage": covered,
            "copper_scenarios": scen,
            "copper_top_beta": [{"company": x["company"], "beta": x["copper_beta"], "r2": x.get("copper_r2"),
                                 "flag": x.get("copper_flag")} for x in ranked[:5]]}


# --------------------------------------------------------------------------- charts & sparklines
CHART_SHEET = "Copper Charts"
SPARK_COL = "G"            # sparkline column on the Copper Charts table
TABLE_ROW = 22             # header row of the holdings table on Copper Charts


def _holdings_rows(wb):
    """Holdings sheet rows: (sheet row, ETF, company, ticker, weight within its ETF)."""
    ho = wb["Holdings"]
    rows, tot = [], {}
    for r in range(5, 155):
        name, tk = ho.cell(r, 2).value, ho.cell(r, 3).value
        if name and tk:
            etf, mv = str(ho.cell(r, 1).value).strip(), ho.cell(r, 6).value
            mv = float(mv) if isinstance(mv, (int, float)) else 0.0
            rows.append([r, etf, name, str(tk).strip(), mv])
            tot[etf] = tot.get(etf, 0) + mv
    return [(r, e, n, t, (mv / tot[e]) if tot[e] else 0) for r, e, n, t, mv in rows]


def write_charts(wb, prices):
    """(Re)build the Copper Charts sheet: copper price chart, copper vs holdings index,
    copper-beta bar chart and a holdings table with 2-year sparklines (added after saving)."""
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.styles import Border, Side
    from openpyxl.utils import get_column_letter

    if CHART_SHEET in wb.sheetnames:
        del wb[CHART_SHEET]
    anchor = next((a for a in ("Copper Scenarios", "Copper Beta", "Valuation") if a in wb.sheetnames))
    ws = wb.create_sheet(CHART_SHEET, index=wb.sheetnames.index(anchor) + 1)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "B8692E"
    ws["A1"] = "Copper Charts – two years of weekly prices at a glance"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color=NAVY)
    ws["A2"] = ("Rebuilt by the daily refresh from the Copper History sheet (weekly Friday closes, last 2 years). "
                "Sparklines show each holding's own-currency price trend; the orange dot is the latest week.")
    ws["A2"].font = Font(name="Arial", size=9, italic=True, color="595959")

    hold = _holdings_rows(wb)
    names = list(prices.columns)
    hist_col = {name: get_column_letter(i + 2) for i, name in enumerate(names)}     # on Copper History
    n = len(prices)
    last_row = FIRST_ROW + n - 1

    # ---- hidden helper block (AA:AC): dates, copper rebased, weighted holdings index rebased
    cu = prices[COPPER_TICKER]
    first_cu = cu.dropna().iloc[0] if cu.notna().any() else None
    base = {t: prices[t].dropna().iloc[0] for _, _, _, t, _ in hold if t in prices and prices[t].notna().any()}
    ws["AA4"], ws["AB4"], ws["AC4"] = "Week ending", "Copper (rebased 100)", "Holdings index (rebased 100)"
    for i, (d, row) in enumerate(prices.iterrows()):
        r = 5 + i
        ws[f"AA{r}"] = d.date()
        ws[f"AA{r}"].number_format = "mmm-yy"
        ws[f"AB{r}"] = float(row[COPPER_TICKER] / first_cu * 100) if first_cu and pd.notna(row[COPPER_TICKER]) else None
        num = den = 0.0
        for _, etf, _, t, w in hold:
            v = row.get(t)
            if t in base and pd.notna(v) and base[t]:
                num += w * v / base[t] * 100
                den += w
        ws[f"AC{r}"] = num / den if den else None
    for c in ("AA", "AB", "AC"):
        ws.column_dimensions[c].hidden = True

    def style(ch, title, ytitle):
        ch.title = title
        ch.y_axis.title = ytitle
        ch.roundedCorners = False
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.visible_cells_only = False          # helper columns are hidden
        if ch.legend is not None:
            ch.legend.position = "b"
            ch.legend.overlay = False
        return ch

    hs = wb[HIST_SHEET]
    # Chart 1: copper price
    c1 = style(LineChart(), "Copper price (US$/lb, weekly)", "US$/lb")
    c1.add_data(Reference(hs, min_col=2, min_row=FIRST_ROW - 1, max_row=last_row), titles_from_data=True)
    c1.set_categories(Reference(hs, min_col=1, min_row=FIRST_ROW, max_row=last_row))
    c1.legend = None
    c1.series[0].graphicalProperties.line.solidFill = "B8692E"
    c1.series[0].graphicalProperties.line.width = 28000
    c1.series[0].smooth = False
    c1.x_axis.number_format = "mmm-yy"
    c1.y_axis.number_format = '"$"0.00'
    c1.y_axis.numFmt.sourceLinked = False if c1.y_axis.numFmt is not None else None
    c1.x_axis.tickLblSkip = 13
    c1.height, c1.width = 7.5, 16
    ws.add_chart(c1, "A4")

    # Chart 2: copper vs holdings (both rebased to 100)
    c2 = style(LineChart(), "Copper vs holdings (rebased to 100)", "Index")
    c2.add_data(Reference(ws, min_col=28, max_col=29, min_row=4, max_row=4 + n), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=27, min_row=5, max_row=4 + n))
    for s_, colour in zip(c2.series, ("B8692E", "1F3864")):
        s_.graphicalProperties.line.solidFill = colour
        s_.graphicalProperties.line.width = 22000
        s_.smooth = False
    c2.x_axis.number_format = "mmm-yy"
    c2.y_axis.number_format = "0"
    c2.x_axis.tickLblSkip = 13
    c2.height, c2.width = 7.5, 16
    ws.add_chart(c2, "H4")

    # ---- holdings table with sparkline column
    thin = Side(style="thin", color="D5DBE5")
    heads = ["Company", "Ticker", "Weight", "Copper beta", "R²", "Reliability", "2-year price trend",
             "1-year change", "2-year change"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(TABLE_ROW, j, h)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[TABLE_ROW].height = 30

    def change(series, weeks):
        s = series.dropna()
        if len(s) < 2:
            return None
        end = s.iloc[-1]
        past = series.iloc[-1 - weeks] if len(series) > weeks else None
        if past is None or pd.isna(past):
            past = s.iloc[0] if weeks >= len(series) - 1 else None
        return float(end / past - 1) if past else None

    spark = []
    table = [("Copper (HG=F)", COPPER_TICKER, None, None)] + [(nm, t, w, r) for r, _, nm, t, w in hold]
    for i, (nm, t, w, hr) in enumerate(table):
        r = TABLE_ROW + 1 + i
        ws.cell(r, 1, nm)
        ws.cell(r, 2, "HG=F" if t == COPPER_TICKER else t)
        if w is not None:
            ws.cell(r, 3, w).number_format = "0.00%"
            ws.cell(r, 4, f"='Copper Beta'!G{hr}").number_format = '0.00;[Red]-0.00;0.00'
            ws.cell(r, 5, f"='Copper Beta'!H{hr}").number_format = "0.000"
            ws.cell(r, 6, f"='Copper Beta'!M{hr}")
        series = prices[t] if t in prices else None
        if series is not None:
            ws.cell(r, 8, change(series, 52))
            ws.cell(r, 9, change(series, n - 1))
            col = hist_col[t]
            spark.append((f"{SPARK_COL}{r}", f"'{HIST_SHEET}'!{col}{FIRST_ROW}:{col}{last_row}"))
        for j in (8, 9):
            ws.cell(r, j).number_format = '+0.0%;[Red]-0.0%;0.0%'
        for j in range(1, 10):
            cell = ws.cell(r, j)
            cell.font = Font(name="Arial", size=10, bold=(i == 0))
            cell.border = Border(bottom=thin)
            if j in (3, 4, 5, 8, 9):
                cell.alignment = Alignment(horizontal="right")
            elif j == 2:
                cell.alignment = Alignment(horizontal="center")
            elif j == 6:
                cell.alignment = Alignment(horizontal="left", indent=1)
        ws.row_dimensions[r].height = 20
    for c, wdt in {"A": 34, "B": 11, "C": 8, "D": 10, "E": 8, "F": 26, "G": 26, "H": 10, "I": 10}.items():
        ws.column_dimensions[c].width = wdt
    ws.freeze_panes = None

    # Chart 3: copper beta by holding (sorted), next to the table
    beta_rows = TABLE_ROW + 1
    ws["AE4"], ws["AF4"] = "Company", "Copper beta"
    # values for the bar chart are filled by refresh (betas known only after regression)
    ws.column_dimensions["AE"].hidden = True
    ws.column_dimensions["AF"].hidden = True
    wb._copper_spark = spark                          # picked up by add_sparklines after saving
    return ws


def fill_beta_chart(wb, items_by_ticker):
    """Sorted copper betas for the bar chart (Python values, same as the Copper Beta sheet)."""
    from openpyxl.chart import BarChart, Reference
    ws = wb[CHART_SHEET]
    data = sorted(((x["company"], x["copper_beta"]) for x in items_by_ticker.values()
                   if isinstance(x.get("copper_beta"), (int, float))), key=lambda z: z[1])
    for i, (nm, b) in enumerate(data):
        ws[f"AE{5 + i}"], ws[f"AF{5 + i}"] = nm, b
    if not data:
        return
    ch = BarChart()
    ch.type = "bar"
    ch.title = "Copper beta by holding (2-year weekly)"
    ch.roundedCorners = False
    ch.legend = None
    ch.visible_cells_only = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.add_data(Reference(ws, min_col=32, min_row=4, max_row=4 + len(data)), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=31, min_row=5, max_row=4 + len(data)))
    ch.series[0].graphicalProperties.solidFill = "1F3864"
    ch.series[0].graphicalProperties.line.solidFill = "1F3864"
    ch.gapWidth = 40
    ch.y_axis.number_format = "0.0"
    ch.y_axis.majorGridlines = None
    ch.height, ch.width = max(10, 0.55 * len(data)), 14
    ws.add_chart(ch, f"K{TABLE_ROW}")


def add_sparklines(path, wb_spark):
    """openpyxl can't write sparklines, so inject them into the saved file's sheet XML."""
    import re
    import shutil
    import tempfile
    import zipfile
    from xml.sax.saxutils import escape as xesc

    if not wb_spark:
        return
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    wbxml = files["xl/workbook.xml"].decode()
    m = re.search(r'<sheet[^>]*name="' + re.escape(CHART_SHEET) + r'"[^>]*r:id="(rId\d+)"', wbxml) or \
        re.search(r'<sheet[^>]*r:id="(rId\d+)"[^>]*name="' + re.escape(CHART_SHEET) + r'"', wbxml)
    if not m:
        return
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    t = re.search(r'<Relationship[^>]*Id="' + m.group(1) + r'"[^>]*Target="([^"]+)"', rels) or \
        re.search(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="' + m.group(1) + r'"', rels)
    target = t.group(1).lstrip("/")
    sheet_path = target if target.startswith("xl/") else "xl/" + target
    xml = files[sheet_path].decode()
    xml = re.sub(r"<extLst>.*?</extLst>", "", xml, flags=re.S)
    lines = "".join(f"<x14:sparkline><xm:f>{xesc(src)}</xm:f><xm:sqref>{cell}</xm:sqref></x14:sparkline>"
                    for cell, src in wb_spark)
    ext = ('<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" '
           'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
           '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
           '<x14:sparklineGroup displayEmptyCellsAs="gap" lineWeight="1.25" last="1">'
           '<x14:colorSeries rgb="FF1F3864"/><x14:colorNegative rgb="FFD00000"/><x14:colorAxis rgb="FF000000"/>'
           '<x14:colorMarkers rgb="FFB8692E"/><x14:colorFirst rgb="FF9AA3B0"/><x14:colorLast rgb="FFB8692E"/>'
           '<x14:colorHigh rgb="FF1A7F37"/><x14:colorLow rgb="FFD00000"/>'
           f'<x14:sparklines>{lines}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>')
    xml = xml.replace("</worksheet>", ext + "</worksheet>")
    files[sheet_path] = xml.encode()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    shutil.move(tmp, path)
