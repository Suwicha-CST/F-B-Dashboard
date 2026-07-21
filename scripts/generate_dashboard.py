"""
Regenerates index.html from F_B_Master_Database.xlsx.

This reproduces the exact data pipeline used throughout the dashboard's
development: parses Revenue_Daily / Budget / Outlet / FX sheets, converts
GBP/EUR revenue into THB using the flat annual FX rates, and embeds the
result into template/template.html + template/app.js to produce a single
self-contained index.html file.

Run this whenever F_B_Master_Database.xlsx has been updated in the repo.
"""

import json
import sys
import pandas as pd

EXCEL_PATH = "F_B_Master_Database.xlsx"
TEMPLATE_HTML_PATH = "template/template.html"
TEMPLATE_JS_PATH = "template/app.js"
OUTPUT_PATH = "index.html"

# Country -> local currency. Add entries here if new outlets open in new countries.
CURRENCY_MAP = {"JLD": "GBP", "ZPM": "EUR"}


def build_dashboard_data():
    xl = pd.ExcelFile(EXCEL_PATH)
    rev = xl.parse("Revenue_Daily")
    bud = xl.parse("Budget")
    outlet = xl.parse("Outlet")
    fx = xl.parse("FX")

    rates = dict(zip(fx["Currency"], fx["FX to THB"]))

    rev_actual = rev.dropna(subset=["Revenue"]).copy()
    rev_actual["Date"] = pd.to_datetime(rev_actual["Date"]).dt.strftime("%Y-%m-%d")
    rev_actual["Month"] = rev_actual["Date"].str.slice(0, 7)
    bud["Date"] = pd.to_datetime(bud["Date"]).dt.strftime("%Y-%m-%d")
    bud["Month"] = bud["Date"].str.slice(0, 7)

    outlet_map = dict(zip(outlet["Outlet"], outlet["Brand"]))
    rev_actual["Brand"] = rev_actual["Outlet"].map(outlet_map)
    bud["Brand"] = bud["Outlet"].map(outlet_map)

    def fx_rate(o):
        cur = CURRENCY_MAP.get(o)
        return rates.get(cur, 1)

    actual_records = []
    for _, r in rev_actual.iterrows():
        rate = fx_rate(r["Outlet"])
        actual_records.append({
            "Date": r["Date"], "Month": r["Month"], "Outlet": r["Outlet"], "Brand": r["Brand"],
            "MealPeriod": r["Meal Period"],
            "RevenueLocal": float(r["Revenue"]), "Currency": CURRENCY_MAP.get(r["Outlet"], "THB"),
            "FXRate": rate, "Revenue": float(r["Revenue"]) * rate,
            "Covers": float(r["Covers"]),
        })

    if not actual_records:
        print("ERROR: no recorded (non-blank) Revenue rows found in Revenue_Daily.", file=sys.stderr)
        sys.exit(1)

    budget_records = []
    for _, r in bud.iterrows():
        rate = fx_rate(r["Outlet"])
        budget_records.append({
            "Date": r["Date"], "Month": r["Month"], "Outlet": r["Outlet"], "Brand": r["Brand"],
            "MealPeriod": r["Meal Period"],
            "BudgetRevenueLocal": float(r["Revenue Budget"]), "Currency": CURRENCY_MAP.get(r["Outlet"], "THB"),
            "FXRate": rate, "BudgetRevenue": float(r["Revenue Budget"]) * rate,
            "BudgetCovers": float(r["Covers Budget"]),
        })

    outlet_records = outlet.copy()
    outlet_records["Opening Date"] = outlet_records["Opening Date"].dt.strftime("%Y-%m-%d")
    outlet_records["Currency"] = outlet_records["Outlet"].map(CURRENCY_MAP)

    dates = sorted(r["Date"] for r in actual_records)

    return {
        "actual": actual_records,
        "budget": budget_records,
        "outlets": outlet_records.to_dict("records"),
        "fxRates": rates,
        "reportingCurrency": "THB",
        "firstActualDate": dates[0],
        "lastActualDate": dates[-1],
    }


def main():
    data = build_dashboard_data()

    with open(TEMPLATE_HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()
    with open(TEMPLATE_JS_PATH, "r", encoding="utf-8") as f:
        app_js = f.read()

    html = html.replace("__DATA_JSON__", json.dumps(data))
    html = html.replace("__APP_JS__", app_js)

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Regenerated {OUTPUT_PATH}")
    print(f"Data range: {data['firstActualDate']} -> {data['lastActualDate']}")
    print(f"Total actual rows: {len(data['actual'])}")


if __name__ == "__main__":
    main()
