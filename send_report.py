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


def fmt_x(v):
    if v == "NM":
        return "NM"
    return f"{v:.1f}x" if isinstance(v, (int, float)) else "n/a"


def fmt_pct(v):
    return f"{v * 100:.1f}%" if isinstance(v, (int, float)) else "n/a"


def fmt_m(v):
    return f"{v:,.1f}" if isinstance(v, (int, float)) else "n/a"


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


TH = 'style="text-align:left;padding:6px 10px;border-bottom:2px solid #1f3a5f;font-size:12px;color:#1f3a5f"'
TD = 'style="padding:5px 10px;border-bottom:1px solid #e3e7ee;font-size:13px"'
TDR = 'style="padding:5px 10px;border-bottom:1px solid #e3e7ee;font-size:13px;text-align:right"'


def table(headers, rows, right_from=1):
    h = "".join(f"<th {TH}>{escape(x)}</th>" for x in headers)
    body = ""
    for r in rows:
        body += "<tr>" + "".join(
            f"<td {TDR if i >= right_from else TD}>{escape(str(c))}</td>" for i, c in enumerate(r)) + "</tr>"
    return f'<table style="border-collapse:collapse;margin:6px 0 18px 0">{"<tr>" + h + "</tr>"}{body}</table>'


def build_html(data):
    dash = os.environ.get("DASHBOARD_URL", "").strip()
    parts = [
        '<div style="font-family:Segoe UI,Arial,sans-serif;color:#1b2430;max-width:760px">',
        '<h2 style="color:#1f3a5f;margin-bottom:2px">ETF Valuation – morning update</h2>',
        f'<div style="color:#6b7686;font-size:13px">Data as at {escape(str(data.get("updated")))} '
        f'(closing prices; source: {escape(str(data.get("source", "Yahoo Finance")))})</div>',
    ]
    if dash:
        parts.append(f'<p><a href="{escape(dash)}" style="color:#1f6feb">Open the live dashboard →</a></p>')

    for code, etf in data.get("etfs", {}).items():
        s, hs = etf["summary"], etf["holdings"]
        parts.append(f'<h3 style="color:#1f3a5f;margin:22px 0 4px">{escape(code)}</h3>')
        parts.append(table(
            ["Metric (weighted harmonic)", "Value", "Coverage"],
            [["Trailing P/E", fmt_x(s.get("pe_trailing_weighted")), fmt_pct(s.get("pe_trailing_coverage"))],
             ["Forward P/E", fmt_x(s.get("pe_forward_weighted")), fmt_pct(s.get("pe_forward_coverage"))],
             ["EV / EBITDA", fmt_x(s.get("ev_ebitda_weighted")), fmt_pct(s.get("ev_ebitda_coverage"))],
             ["P / NAV (book)", fmt_x(s.get("p_nav_weighted")), fmt_pct(s.get("p_nav_coverage"))]]))
        parts.append(f'<div style="font-size:13px">{s.get("with_data")}/{s.get("holdings")} holdings with data · '
                     f'{s.get("loss_making")} loss-making</div>')

        if s.get("new_results"):
            parts.append('<p style="background:#fff6d6;padding:8px 10px;border-radius:6px;font-size:13px">'
                         '<b>New financial statements picked up:</b> '
                         + escape(", ".join(s["new_results"])) + '</p>')
        if s.get("upcoming_earnings"):
            parts.append('<div style="font-size:13px;margin-top:8px"><b>Reporting in the next 14 days:</b> '
                         + escape("; ".join(f'{u["company"]} ({u["date"]})' for u in s["upcoming_earnings"]))
                         + '</div>')

        top = sorted(hs, key=lambda x: x.get("weight") or 0, reverse=True)[:10]
        parts.append('<div style="font-size:13px;margin-top:14px"><b>Top 10 holdings</b></div>')
        parts.append(table(
            ["Company", "Weight", "Price", "Trail P/E", "Fwd P/E", "EV/EBITDA", "P/NAV"],
            [[x["company"], fmt_pct(x.get("weight")),
              f'{x["price"]:,.2f} {x.get("currency", "")}' if isinstance(x.get("price"), (int, float)) else "n/a",
              fmt_x(x.get("pe_trailing")), fmt_x(x.get("pe_forward")),
              fmt_x(x.get("ev_ebitda")), fmt_x(x.get("p_nav"))] for x in top]))

        problems = [x for x in hs if (x.get("status") or "") != "OK"]
        if problems:
            parts.append('<div style="font-size:12px;color:#8a5a00"><b>Data issues:</b> '
                         + escape("; ".join(f'{x["company"]} – {x.get("status")}' for x in problems)) + '</div>')

    parts.append('<p style="font-size:11px;color:#6b7686;margin-top:24px">NM = not meaningful (negative/zero '
                 'denominator); n/a = no data from Yahoo. Generated automatically; for educational use. '
                 'Full model attached.</p></div>')
    return "".join(parts)


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
    msg["Subject"] = f"{codes} valuation update – {today}"
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
