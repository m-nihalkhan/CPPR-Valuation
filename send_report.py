"""
Morning valuation email.

Reads data/valuation.json (written by refresh.py), waits until the send time in
Sydney (default 08:00), then emails an HTML summary with the workbook attached.

Settings come from environment variables (GitHub → Settings → Secrets and variables → Actions):
  SMTP_USER       sender address, e.g. yourname@gmail.com            (secret)
  SMTP_PASSWORD   Gmail *app password* (not your normal password)     (secret)
  REPORT_TO       comma-separated recipients                          (secret or variable)
  DASHBOARD_URL   link to the published Lovable site (optional)       (variable)
  SMTP_HOST / SMTP_PORT   default smtp.gmail.com / 465
  SEND_AT         local send time HH:MM, default 08:00
  SEND_NOW=1      skip the wait (for testing)

If SMTP_USER / SMTP_PASSWORD / REPORT_TO are missing it prints a note and exits
without failing, so the daily refresh still works before email is set up.
"""
import datetime as dt
import json
import os
import pathlib
import smtplib
import ssl
import sys
import time
from email.message import EmailMessage
from html import escape
from zoneinfo import ZoneInfo

ROOT = pathlib.Path(__file__).resolve().parent
JSON_IN = ROOT / "data" / "valuation.json"
WORKBOOK = ROOT / "CPPR_Valuation_Model.xlsx"
TZ = ZoneInfo("Australia/Sydney")
MAX_WAIT_HOURS = 4     # never sit waiting longer than this


def wait_until_send_time():
    if os.environ.get("SEND_NOW") == "1":
        return
    hh, mm = (int(x) for x in os.environ.get("SEND_AT", "08:00").split(":"))
    now = dt.datetime.now(TZ)
    target = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
    secs = (target - now).total_seconds()
    if 0 < secs <= MAX_WAIT_HOURS * 3600:
        print(f"Waiting {secs / 60:.0f} min until {target:%H:%M %Z} to send…", flush=True)
        time.sleep(secs)
    else:
        print("Send time already passed (or too far away) – sending now.")


# ---------------------------------------------------------------- email design
# Email clients ignore most CSS, so everything is tables + inline styles.
NAVY, COPPER, INK, MUTED, LINE, ZEBRA, BG = "#13294b", "#b8692e", "#1b2430", "#6b7686", "#e3e7ee", "#f7f9fc", "#eef1f6"
FONT = "font-family:'Segoe UI',Helvetica,Arial,sans-serif"


def cell_x(v):
    """Multiple, with NM / n/a greyed out."""
    if v == "NM":
        return f'<span style="color:#9aa3b0;font-style:italic">NM</span>'
    if isinstance(v, (int, float)):
        if v < 0:
            return f'<span style="color:#b42318">{v:.1f}x</span>'
        return f"{v:.1f}x"
    return '<span style="color:#c3c9d2">n/a</span>'


def aud(v, dp=2, prefix="A$"):
    if not isinstance(v, (int, float)):
        return '<span style="color:#c3c9d2">n/a</span>'
    if dp == 2 and abs(v) < 1:
        dp = 3
    return f"{prefix}{v:,.{dp}f}"


def nice_date(iso):
    try:
        return dt.date.fromisoformat(str(iso)[:10]).strftime("%a %d %b")
    except ValueError:
        return str(iso)


def tile(label, value, sub):
    return (f'<td width="33%" style="padding:6px"><table width="100%" cellpadding="0" cellspacing="0" '
            f'style="background:#ffffff;border:1px solid {LINE};border-top:3px solid {COPPER};border-radius:6px">'
            f'<tr><td style="padding:12px 14px 2px;{FONT};font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:{MUTED}">{label}</td></tr>'
            f'<tr><td style="padding:0 14px;{FONT};font-size:26px;font-weight:600;color:{NAVY}">{value}</td></tr>'
            f'<tr><td style="padding:2px 14px 12px;{FONT};font-size:11px;color:{MUTED}">{sub}</td></tr></table></td>')


def section(title, note=""):
    n = f'<span style="font-weight:400;color:{MUTED};font-size:12px">&nbsp;&nbsp;{note}</span>' if note else ""
    return (f'<tr><td style="padding:26px 28px 8px;{FONT};font-size:16px;font-weight:600;color:{NAVY}">'
            f'<span style="display:inline-block;width:4px;height:14px;background:{COPPER};margin-right:8px;vertical-align:-1px"></span>'
            f'{escape(title)}{n}</td></tr>')


def callout(colour, bg, title, body):
    return (f'<tr><td style="padding:6px 28px"><table width="100%" cellpadding="0" cellspacing="0" '
            f'style="background:{bg};border-left:4px solid {colour};border-radius:4px"><tr><td style="padding:12px 16px;{FONT};font-size:13px;color:{INK}">'
            f'<b style="color:{colour}">{title}</b><br>{body}</td></tr></table></td></tr>')


def mini_rank(title, rows, key):
    body = "".join(
        f'<tr><td style="padding:5px 0;{FONT};font-size:13px;color:{INK};border-bottom:1px solid {LINE}">{escape(x["company"])}</td>'
        f'<td align="right" style="padding:5px 0;{FONT};font-size:13px;font-weight:600;color:{NAVY};border-bottom:1px solid {LINE}">{cell_x(x[key])}</td></tr>'
        for x in rows) or f'<tr><td style="{FONT};font-size:12px;color:{MUTED}">No data</td></tr>'
    return (f'<td width="50%" valign="top" style="padding:0 10px"><div style="{FONT};font-size:11px;letter-spacing:.06em;'
            f'text-transform:uppercase;color:{MUTED};padding-bottom:4px">{title}</div>'
            f'<table width="100%" cellpadding="0" cellspacing="0">{body}</table></td>')


def pct(v, dp=1, colour=True):
    if not isinstance(v, (int, float)):
        return '<span style="color:#c3c9d2">n/a</span>'
    txt = f"{v * 100:+.{dp}f}%"
    if colour and v < 0:
        return f'<span style="color:#b42318">{txt}</span>'
    if colour and v > 0:
        return f'<span style="color:#1a7f37">{txt}</span>'
    return txt


def beta_txt(v):
    return f"{v:.2f}" if isinstance(v, (int, float)) else '<span style="color:#c3c9d2">n/a</span>'


def copper_section(s):
    """Copper price, ETF copper beta and bear/bull scenarios."""
    if s.get("copper_beta_weighted") is None and s.get("copper_price") is None:
        return []
    out = [section("Copper sensitivity", "2-year weekly copper beta · market-price sensitivity, not an earnings forecast")]
    cu, wk = s.get("copper_price"), s.get("copper_week_change")
    cu_txt = f"US${cu:,.2f}" if isinstance(cu, (int, float)) else "n/a"
    cov = s.get("copper_beta_coverage")
    sc = s.get("copper_scenarios") or {}
    out.append('<tr><td style="padding:0 22px"><table width="100%" cellpadding="0" cellspacing="0"><tr>'
               + tile("Copper (US$/lb)", cu_txt, f'{pct(wk)} on the week &middot; settle {nice_date(s.get("copper_price_date"))}')
               + tile("ETF copper beta", beta_txt(s.get("copper_beta_weighted")),
                      f'{cov * 100:.0f}% of fund covered' if isinstance(cov, (int, float)) else "")
               + tile("If copper moves ±10%", f'{pct(sc.get("+10%"), 1, False)}',
                      f'estimated ETF move (&minus;10% copper: {pct(sc.get("-10%"), 1, False)})')
               + '</tr></table></td></tr>')
    th = f'style="padding:7px 10px;{FONT};font-size:11px;color:#ffffff;background:{NAVY};text-align:right"'
    td = f'style="padding:7px 10px;{FONT};font-size:12px;color:{INK};border-bottom:1px solid {LINE};text-align:right"'
    tdl = f'style="padding:7px 10px;{FONT};font-size:12px;color:{INK};border-bottom:1px solid {LINE};text-align:left;white-space:nowrap"'
    heads = ["", "Bear<br>&minus;20%", "&minus;10%", "+10%", "Bull<br>+20%"]
    keys, shocks = ["-20%", "-10%", "+10%", "+20%"], [-0.2, -0.1, 0.1, 0.2]
    price_row = [f"US${cu * (1 + k):,.2f}" if isinstance(cu, (int, float)) else "n/a" for k in shocks]
    move_row = [f"<b>{pct(sc.get(k))}</b>" for k in keys]
    table = (f'<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid {LINE};margin-top:6px">'
             + "<tr>" + "".join(f"<th {th}>{h}</th>" for h in heads) + "</tr>"
             + f'<tr><td {tdl}>Copper price</td>' + "".join(f"<td {td}>{v}</td>" for v in price_row) + "</tr>"
             + f'<tr><td {tdl}>Estimated ETF move</td>' + "".join(f"<td {td}>{v}</td>" for v in move_row) + "</tr></table>")
    top = s.get("copper_top_beta") or []
    top_rows = ""
    for t in top:
        flag = "" if t.get("flag") == "OK" else (f' <span style="font-size:10px;color:#9c5700">'
                                               f'({escape(str(t.get("flag")))})</span>')
        r2 = f'R&sup2; {t["r2"]:.2f}' if isinstance(t.get("r2"), (int, float)) else ""
        top_rows += (f'<tr><td style="padding:5px 0;{FONT};font-size:13px;color:{INK};border-bottom:1px solid {LINE}">'
                     f'{escape(t["company"])}{flag}</td>'
                     f'<td align="right" style="padding:5px 0;{FONT};font-size:13px;font-weight:600;color:{NAVY};'
                     f'border-bottom:1px solid {LINE};white-space:nowrap">&beta; {beta_txt(t.get("beta"))} '
                     f'<span style="font-weight:400;color:{MUTED};font-size:11px">{r2}</span></td></tr>')
    out.append(f'<tr><td style="padding:4px 28px 0"><table width="100%" cellpadding="0" cellspacing="0"><tr>'
               f'<td width="52%" valign="top" style="padding-right:18px;{FONT};font-size:11px;letter-spacing:.06em;'
               f'text-transform:uppercase;color:{MUTED}">ETF scenarios{table}</td>'
               f'<td width="48%" valign="top"><div style="{FONT};font-size:11px;letter-spacing:.06em;text-transform:uppercase;'
               f'color:{MUTED};padding-bottom:4px">Most copper-sensitive holdings</div>'
               f'<table width="100%" cellpadding="0" cellspacing="0">{top_rows}</table></td></tr></table></td></tr>')
    return out


def holdings_source():
    """Holdings source note from the model's Holdings sheet (cell A2)."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(WORKBOOK, read_only=True, data_only=True)
        v = wb["Holdings"]["A2"].value
        wb.close()
        return str(v) if v else None
    except Exception:
        return None


def sources_block(data, s, hs):
    """'Sources' footer listing where every figure in the email comes from."""
    li = f'style="padding:3px 0;{FONT};font-size:11px;color:{MUTED};line-height:1.5"'
    a = f'style="color:{MUTED}"'
    items = [
        f'<b>Share prices, financial statements, analyst consensus EPS and FX rates:</b> '
        f'<a href="https://finance.yahoo.com" {a}>Yahoo Finance</a> via the yfinance Python library. Each price is the '
        f'latest official close on the stock&rsquo;s home exchange; data refreshed {escape(str(data.get("updated")))}.',
    ]
    if s.get("copper_price") is not None:
        items.append(
            f'<b>Copper price:</b> COMEX copper futures, continuous front-month contract (Yahoo ticker HG=F), US$/lb, '
            f'at the 1:00pm New York settlement time on {nice_date(s.get("copper_price_date"))} &mdash; '
            f'<a href="https://www.cmegroup.com/trading/metals/base/copper-futures-and-options.html" {a}>CME Group</a>.')
        items.append('<b>Copper beta:</b> model calculation &ndash; ordinary least squares regression of each holding&rsquo;s '
                     'weekly return on the COMEX copper futures return over the last 2 years (104 weeks).')
    src = holdings_source()
    items.append(f'<b>Holdings and weights:</b> {escape(src)}' if src else
                 '<b>Holdings and weights:</b> ETF Shares CPPR full holdings table (Holdings sheet of the attached model).')
    brokers = sorted({f'{x["company"]} &ndash; {escape(str(x["broker_nav_source"]))}' for x in hs
                      if x.get("broker_nav") and x.get("broker_nav_source")})
    if brokers:
        items.append("<b>Broker NAV:</b> " + "; ".join(brokers) + ".")
    items.append('<b>Valuation multiples:</b> model calculations from the above (P/E, EV/EBITDA and P/NAV per holding; '
                 'ETF figures are weight-based harmonic averages of positive multiples).')
    body = "".join(f"<tr><td {li}>{i}</td></tr>" for i in items)
    return [section("Sources"),
            f'<tr><td style="padding:0 28px 6px"><table width="100%" cellpadding="0" cellspacing="0">{body}</table></td></tr>']


def holdings_table(hs):
    cols = [("#", "center"), ("Company", "left"), ("Weight", "right"), ("Price", "right"), ("Mkt cap A$m", "right"),
            ("EV A$m", "right"), ("Trail P/E", "right"), ("Curr P/E", "right"), ("Fwd P/E", "right"),
            ("EV/EBITDA", "right"), ("P/NAV", "right"), ("Cu &beta;", "right")]
    head = "".join(f'<th align="{a}" style="padding:9px 7px;{FONT};font-size:11px;font-weight:600;color:#ffffff;'
                   f'background:{NAVY};white-space:nowrap">{c}</th>' for c, a in cols)
    rows = []
    for i, x in enumerate(sorted(hs, key=lambda r: r.get("weight") or 0, reverse=True), 1):
        bg = ZEBRA if i % 2 == 0 else "#ffffff"
        td = f'style="padding:7px;{FONT};font-size:12px;color:{INK};border-bottom:1px solid {LINE};white-space:nowrap"'
        badge = (f' <span style="background:{COPPER};color:#fff;font-size:9px;font-weight:600;padding:1px 5px;'
                 f'border-radius:8px;vertical-align:1px">NEW RESULTS</span>') if x.get("new_results") else ""
        name = (f'<span style="font-weight:600">{escape(x["company"])}</span>{badge}<br>'
                f'<span style="font-size:10px;color:{MUTED}">{escape(str(x.get("ticker") or ""))} · {escape(str(x.get("country") or ""))}</span>')
        w = f'{x["weight"] * 100:.1f}%' if isinstance(x.get("weight"), (int, float)) else "n/a"
        cd = x.get("close_date")
        price = aud(x.get("price_aud")) + (f'<br><span style="font-size:10px;color:{MUTED}">close {nice_date(cd)}</span>' if cd else "")
        vals = [str(i), name, w, price, aud(x.get("equity_value_aud_m"), 0, ""),
                aud(x.get("enterprise_value_aud_m"), 0, ""), cell_x(x.get("pe_trailing")), cell_x(x.get("pe_current")),
                cell_x(x.get("pe_forward")), cell_x(x.get("ev_ebitda")), cell_x(x.get("p_nav")),
                beta_txt(x.get("copper_beta"))]
        rows.append(f'<tr style="background:{bg}">' + "".join(
            f'<td align="{a}" {td}>{v}</td>' for v, (_, a) in zip(vals, cols)) + "</tr>")
    return (f'<tr><td style="padding:0 20px"><table width="100%" cellpadding="0" cellspacing="0" '
            f'style="border-collapse:collapse;border:1px solid {LINE}"><tr>{head}</tr>{"".join(rows)}</table></td></tr>')


def build_html(data):
    dash = os.environ.get("DASHBOARD_URL", "").strip()
    today = dt.datetime.now(TZ).strftime("%A %d %B %Y")
    out = [f'<div style="background:{BG};padding:24px 0;margin:0"><table align="center" width="900" cellpadding="0" cellspacing="0" '
           f'style="background:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.08)">']

    for code, etf in data.get("etfs", {}).items():
        s, hs = etf["summary"], etf["holdings"]
        # -- header banner
        button = (f'<a href="{escape(dash)}" style="display:inline-block;background:{COPPER};color:#ffffff;{FONT};'
                  f'font-size:13px;font-weight:600;text-decoration:none;padding:10px 18px;border-radius:5px">'
                  f'Open live dashboard &rarr;</a>') if dash else ""
        out.append(
            f'<tr><td style="background:{NAVY};padding:26px 28px">'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr><td>'
            f'<div style="{FONT};font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:{COPPER};font-weight:600">'
            f'Morning valuation update</div>'
            f'<div style="{FONT};font-size:26px;font-weight:600;color:#ffffff;padding-top:4px">{escape(code)} &middot; Global Copper Miners</div>'
            f'<div style="{FONT};font-size:13px;color:#b9c4d6;padding-top:6px">{today} &nbsp;|&nbsp; '
            f'Latest close on each stock\'s home exchange &nbsp;|&nbsp; all values in AUD &nbsp;|&nbsp; '
            f'<span style="white-space:nowrap">refreshed {escape(str(data.get("updated")))}</span></div>'
            f'</td><td align="right" valign="middle">{button}</td></tr></table></td></tr>')
        out.append(f'<tr><td style="background:{ZEBRA};padding:10px 28px;{FONT};font-size:12px;color:{MUTED};border-bottom:1px solid {LINE}">'
                   f'&#128206; Full Excel valuation model attached &nbsp;&middot;&nbsp; '
                   f'<b style="color:{INK}">{s.get("with_data")}/{s.get("holdings")}</b> holdings with live data &nbsp;&middot;&nbsp; '
                   f'<b style="color:{INK}">{s.get("loss_making")}</b> loss-making &nbsp;&middot;&nbsp; '
                   f'<b style="color:{INK}">{len(s.get("new_results") or [])}</b> with new results</td></tr>')

        # -- KPI tiles
        cov = lambda k: f'{s.get(k + "_coverage", 0) * 100:.0f}% of fund covered' if isinstance(s.get(k + "_coverage"), (int, float)) else ""
        fund = s.get("market_value_aud")
        out.append(section("ETF valuation snapshot", "weighted harmonic averages"))
        out.append('<tr><td style="padding:0 22px"><table width="100%" cellpadding="0" cellspacing="0"><tr>'
                   + tile("Trailing P/E", cell_x(s.get("pe_trailing_weighted")), cov("pe_trailing"))
                   + tile("Current P/E (FY0)", cell_x(s.get("pe_current_weighted")), cov("pe_current"))
                   + tile("Forward P/E (FY1)", cell_x(s.get("pe_forward_weighted")), cov("pe_forward"))
                   + '</tr><tr>'
                   + tile("EV / EBITDA", cell_x(s.get("ev_ebitda_weighted")), cov("ev_ebitda"))
                   + tile("P / NAV (book)", cell_x(s.get("p_nav_weighted")), cov("p_nav"))
                   + tile("Fund value", f'A${fund / 1e6:,.1f}m' if isinstance(fund, (int, float)) else "n/a",
                          f'{s.get("holdings")} holdings')
                   + '</tr></table></td></tr>')

        out.extend(copper_section(s))

        # -- alerts
        if s.get("new_results"):
            out.append(section("Company news"))
            out.append(callout(COPPER, "#fbf1e8", "New financial statements picked up overnight",
                               escape(", ".join(s["new_results"])) + " &mdash; multiples now reflect the latest reported figures."))
        if s.get("upcoming_earnings"):
            if not s.get("new_results"):
                out.append(section("Company news"))
            items = "<br>".join(f'<b>{nice_date(u["date"])}</b> &nbsp;{escape(u["company"])}' for u in s["upcoming_earnings"])
            out.append(callout("#2f6fb5", "#eaf2fb", "Reporting in the next 14 days", items))

        # -- highlights
        def lowest(key, n=5):
            v = [x for x in hs if isinstance(x.get(key), (int, float)) and x[key] > 0]
            return sorted(v, key=lambda x: x[key])[:n]
        out.append(section("Valuation highlights", "lowest positive multiples in the fund"))
        out.append('<tr><td style="padding:4px 18px 0"><table width="100%" cellpadding="0" cellspacing="0"><tr>'
                   + mini_rank("Cheapest on forward P/E", lowest("pe_forward"), "pe_forward")
                   + mini_rank("Cheapest on EV / EBITDA", lowest("ev_ebitda"), "ev_ebitda")
                   + '</tr></table></td></tr>')

        # -- all holdings
        out.append(section(f"All {len(hs)} holdings", "ranked by weight · prices and values in A$ · red = negative (loss-making, excluded from fund averages) · Cu &beta; = copper beta · n/a = no data"))
        out.append(holdings_table(hs))

        problems = [x for x in hs if (x.get("status") or "") != "OK"]
        if problems:
            out.append(f'<tr><td style="padding:14px 28px 0;{FONT};font-size:11px;color:#8a5a00">'
                       f'<b>Data notes:</b> ' + escape("; ".join(f'{x["company"]} – {x.get("status")}' for x in problems))
                       + '</td></tr>')

        out.extend(sources_block(data, s, hs))

    out.append('<tr><td style="padding:14px"></td></tr></table></div>')
    return "".join(out)


def main():
    user, pwd = os.environ.get("SMTP_USER"), os.environ.get("SMTP_PASSWORD")
    to = [a.strip() for a in os.environ.get("REPORT_TO", "").split(",") if a.strip()]
    if not (user and pwd and to):
        print("Email not configured (SMTP_USER / SMTP_PASSWORD / REPORT_TO missing) – skipping send.")
        return 0

    data = json.loads(JSON_IN.read_text())
    wait_until_send_time()

    msg = EmailMessage()
    today = dt.datetime.now(TZ).strftime("%a %d %b %Y")
    codes = ", ".join(data.get("etfs", {}).keys())
    msg["Subject"] = f"{codes} morning valuation update – {today}"
    msg["From"], msg["To"] = user, ", ".join(to)
    msg.set_content("This email is HTML – please view it in an HTML-capable mail client.")
    msg.add_alternative(build_html(data), subtype="html")
    if WORKBOOK.exists():
        msg.add_attachment(WORKBOOK.read_bytes(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=WORKBOOK.name)

    host, port = os.environ.get("SMTP_HOST", "smtp.gmail.com"), int(os.environ.get("SMTP_PORT", "465"))
    with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context()) as s:
        s.login(user, pwd)
        s.send_message(msg)
    print(f"Sent to {len(to)} recipient(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
