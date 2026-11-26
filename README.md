# ETF Valuation Model – CPPR (ETFS Global Pure Play Copper Miners ETF)

This is a simplified multiples model covering all 50 CPPR holdings. It refreshes every weekday from Yahoo Finance and feeds a Lovable dashboard.

```
Yahoo Finance ──(yfinance, daily via GitHub Actions ~6am Sydney)──► refresh.py
                                                         ├─► CPPR_Valuation_Model.xlsx  (Data sheet → formulas → Dashboard)
                                                         └─► data/valuation.json        (read live by the Lovable website)
                                                                    │
                                              send_report.py ◄──────┘  waits until 08:00 Sydney, emails summary + workbook
```

## What's in the repo

| File | Purpose |
|---|---|
| `CPPR_Valuation_Model.xlsx` | The model. Sheets: Cover, Dashboard, Holdings, Data, Valuation, Lists |
| `refresh.py` | Pulls Yahoo data, fills the Data sheet, and writes `data/valuation.json` |
| `send_report.py` | Builds the 8am HTML email (ETF summary, top 10, new results, upcoming reports) and attaches the workbook |
| `.github/workflows/daily-refresh.yml` | Runs `refresh.py` at about 6:20am Sydney (5:20am in winter), Tue–Sat, publishes the data, then sends the email at 8:00 |
| `requirements.txt` | Python packages |

## Outputs per holding

Trailing / current / forward P/E, equity value, enterprise value, EV/EBITDA, book NAV per share, P/NAV, an optional broker NAV with P/broker NAV, and A$ values. You also get ETF-level weighted (harmonic) multiples.

## Setup (about 10 minutes)

1. **Create a public GitHub repo**, e.g. `cppr-valuation`, and upload everything in this folder. That includes the `.github` folder; in the GitHub web uploader, drag the whole folder in.
2. **Run it once by hand.** Go to the **Actions** tab, click **Daily valuation refresh**, then **Run workflow**. After about 3–5 minutes the workbook and `data/valuation.json` are committed back with live data.
3. **Check the result.** Download the workbook and open **Dashboard**. Pick the ETF, then a company. Look at the **Data** sheet: the Status column flags any ticker that returned no data.
4. **Set up the 8am email** (optional – the refresh works without it). In the repo go to **Settings → Secrets and variables → Actions**:
   - *Secrets* tab → add `SMTP_USER` (your Gmail address), `SMTP_PASSWORD` (a Gmail **app password**: Google Account → Security → 2-Step Verification → App passwords) and `REPORT_TO` (recipients, comma-separated).
   - *Variables* tab → add `DASHBOARD_URL` (your published Lovable link) so the email links to it.
   - Test it: **Actions → Daily valuation refresh → Run workflow**, tick *Also send the email*. It sends immediately.
5. **Copy the JSON link** for Lovable:
   `https://raw.githubusercontent.com/<your-username>/cppr-valuation/main/data/valuation.json`

To run it on your own computer instead: `pip install -r requirements.txt`, then `python refresh.py`.

## How the automatic updates work

- **Daily close prices.** GitHub runs the workflow every weekday at about 6:20am Sydney (AEDT; 5:20am in winter). By then the ASX, Asian, European and US/Canadian markets have all closed, so every price is the previous session's close. The JSON and workbook are committed straight away, so the dashboard is current from about 6:30am.
- **8am email.** The same job then waits until 08:00 Sydney time (daylight saving is handled automatically) and sends the email. GitHub's scheduler can start jobs late at busy times; the early start leaves room for that. If the job ever starts after 8am, the email goes out as soon as the refresh finishes.
- **New financial statements.** Every run re-downloads each company's latest balance sheet, income statement and consensus estimates, so new results flow into the multiples the morning after Yahoo publishes them. You don't need a separate trigger. The script records each company's latest statement period end. When it moves forward, the company is flagged as **new results** in the JSON, the dashboard and the email. The email also lists companies due to report in the next 14 days. To pick up a big result immediately, use **Run workflow**.
- **Weekends and holidays.** Saturday's run picks up Friday's US close. On market holidays the prices simply don't change.
- **Keep-alive.** GitHub pauses scheduled workflows in repos with no activity for 60 days. The daily commit normally counts as activity, but if you see a "scheduled workflow disabled" banner, re-enable it in the Actions tab.

## Maintaining it

- **Holdings change** when the index rebalances. Update the Holdings sheet from the ETF Shares CPPR page using the "Full Holdings" table. Weights recalculate from the market values.
- **Adding another ETF:** add rows on Holdings with the new ETF code in column A. It appears in both dropdowns and in the JSON automatically.
- **Broker NAV:** type a value into Holdings column H (yellow cells). BHP is pre-filled with A$66.73/sh, the FY26 NAV from the Goldman Sachs model.
- **Tickers marked "Check"** on Holdings column K are the ones most likely to need correcting if Yahoo returns nothing.

## Lovable prompt (paste this into Lovable)

> Build an internal "ETF Valuation Dashboard" web app. No login and no backend needed.
>
> **Data:** on load, fetch JSON from `https://raw.githubusercontent.com/<your-username>/cppr-valuation/main/data/valuation.json` and add a "Refresh" button that re-fetches. Shape:
> `{ updated: string, etfs: { [etfCode]: { summary: {...}, holdings: [ {...} ] } } }`.
> Summary fields: `holdings, with_data, loss_making, market_value_aud, pe_trailing_weighted, pe_trailing_coverage, pe_current_weighted, pe_current_coverage, pe_forward_weighted, pe_forward_coverage, ev_ebitda_weighted, ev_ebitda_coverage, p_nav_weighted, p_nav_coverage`.
> Holding fields: `company, ticker, country, weight, price, currency, equity_value_m, equity_value_aud_m, net_debt_m, minority_m, enterprise_value_m, enterprise_value_aud_m, ebitda_ttm_m, ev_ebitda, pe_trailing, pe_current, pe_forward, nav_per_share, p_nav, broker_nav, broker_nav_source, p_broker_nav, yahoo_pe, yahoo_forward_pe, yahoo_ev_ebitda, status, last_updated, statement_date, next_earnings_date, new_results`.
> Summary also has `new_results` (array of company names) and `upcoming_earnings` (array of `{company, date}`).
> Any metric can be a number, the string "NM" (not meaningful, e.g. negative earnings), or null (no data). Show null as "n/a" and "NM" as "NM" in grey.
>
> **Layout:**
> 1. Header: title, "Data as at {updated}", Refresh button. If `updated` is more than 30 hours old, show an amber "data may be stale" badge.
> 1b. An alert strip under the header when the selected ETF's `summary.new_results` is non-empty ("New results: …"), and a small "Reporting soon" list from `summary.upcoming_earnings`.
> 2. Two dropdowns side by side: **ETF** (keys of `etfs`), then **Company** (holdings of the selected ETF, sorted by weight, descending). Changing the ETF resets the company to the largest holding.
> 3. **Company panel** (cards): Price + currency, Weight, Equity value (local m and A$m), Enterprise value (local m and A$m), Trailing P/E, Current P/E, Forward P/E, EV/EBITDA, Book NAV/share, P/NAV, Broker NAV and P/Broker NAV (only if `broker_nav` is set, with the source as a tooltip). Below that, a small "Yahoo cross-check" row with Yahoo's own P/E, forward P/E and EV/EBITDA, plus the data status, last-updated time, latest statement date and next earnings date. Show a "New results" badge when `new_results` is true.
> 4. **ETF summary** strip: number of holdings, holdings with data, loss-making count, fund value (A$), and weighted Trailing P/E, Current P/E, Forward P/E, EV/EBITDA and P/NAV, each with its coverage %.
> 5. **Holdings table** for the selected ETF: Company, Ticker, Country, Weight, Price (A$), Trailing P/E, Current P/E, Forward P/E, EV/EBITDA, P/NAV, Equity value (A$m). Sortable columns, a search box, and clicking a row selects that company in the dropdown.
> 6. A bar chart of Forward P/E by company (skip NM/n/a) for the selected ETF.
>
> **Currency:** show every price and value in Australian dollars (`price_aud`, `equity_value_aud_m`, `enterprise_value_aud_m`), prefixed `A$`. Show the local price only as small grey secondary text.
>
> **Formatting:** multiples as `12.3x`, money in millions with 1 decimal, weights as `%` with 1 decimal. Clean, professional finance styling, dark navy accents, responsive.

## Limitations (worth stating in the write-up)

- yfinance is an unofficial Yahoo wrapper for personal and educational use. It can rate-limit or return gaps, so failed tickers keep their last good data and are marked "Stale".
- Consensus estimates (current and forward P/E) are often missing for junior explorers, which is common in an equal-weight copper index.
- EV uses the latest balance sheet and trailing-12-month EBITDA. Book NAV understates value for miners, which is why there's a broker NAV column.
