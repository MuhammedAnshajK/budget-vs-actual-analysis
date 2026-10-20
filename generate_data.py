"""Creates the illustrative budget and actual data for the Budget vs Actual Analysis.

Company: an illustrative sportswear brand selling three product lines, financial year
January to December 2025. Units in thousands, prices and unit costs in $ per unit,
expenses in $ thousands. The actuals contain deliberate, explainable deviations
(a marketing campaign, a supplier price rise, discounting, freight costs, hiring delays
and a one-off IT project) so the analysis has real stories to find.

Output: data/budget.csv and data/actual.csv (one row per month and line).
Run:  python generate_data.py
"""
import csv
import random
from pathlib import Path

random.seed(7)
MONTHS = [f"2025-{m:02d}" for m in range(1, 13)]
SEASON = [0.85, 0.85, 0.95, 1.00, 1.05, 1.05, 1.00, 1.00, 0.95, 1.00, 1.15, 1.35]
PRODUCTS = {  # base monthly units (k), price ($), unit cost ($)
    "Running Shoes": (20, 95.0, 42.0),
    "Apparel": (35, 40.0, 17.0),
    "Accessories": (25, 22.0, 8.0),
}
OPEX = ["Marketing", "Salaries and benefits", "Rent and occupancy", "Freight and logistics",
        "IT and software", "Travel", "Other admin"]


def r(x, d=1):
    return round(x, d)


def build():
    budget, actual = [], []
    for i, month in enumerate(MONTHS):
        m = i + 1
        s = SEASON[i]
        b_rev = a_rev = 0.0
        for name, (units, price, cost) in PRODUCTS.items():
            bu = units * s
            # actual units, price and cost, with the planned deviations
            au, ap, ac = bu * random.uniform(0.98, 1.02), price, cost
            if name == "Running Shoes" and m >= 6:
                ac = cost * 1.06                      # supplier price increase from June
            if name == "Apparel":
                au *= 1.18 if 3 <= m <= 7 else (1.08 if m >= 8 else 1.0)   # spring campaign lift
            if name == "Accessories":
                if m >= 10:
                    ap, au = price * 0.85, au * 1.05  # Q4 discounting: price -15%, volume +5%
                else:
                    au *= 0.97
            budget += [(month, f"{name} units", r(bu, 2)), (month, f"{name} price", price), (month, f"{name} unit cost", cost)]
            actual += [(month, f"{name} units", r(au, 2)), (month, f"{name} price", r(ap, 2)), (month, f"{name} unit cost", r(ac, 2))]
            b_rev += bu * price
            a_rev += au * ap
        b = {
            "Marketing": 230 * s,
            "Salaries and benefits": 520,
            "Rent and occupancy": 180,
            "Freight and logistics": 0.04 * b_rev,
            "IT and software": 60,
            "Travel": 25,
            "Other admin": 70,
        }
        a = {
            "Marketing": 230 * s * (1.40 if 3 <= m <= 5 else random.uniform(0.97, 1.03)),
            "Salaries and benefits": 470 if m <= 6 else 515,          # vacancies until mid-year
            "Rent and occupancy": 180,
            "Freight and logistics": (0.04 if m <= 6 else 0.055) * a_rev,  # shipping rates rose in H2
            "IT and software": 60 + (240 if m == 9 else 0),            # one-off ERP upgrade
            "Travel": 25 * random.uniform(0.8, 0.95),
            "Other admin": 70 * random.uniform(0.95, 1.05),
        }
        budget += [(month, k, r(b[k])) for k in OPEX]
        actual += [(month, k, r(a[k])) for k in OPEX]
    out = Path(__file__).parent / "data"
    out.mkdir(exist_ok=True)
    for name, rows in (("budget", budget), ("actual", actual)):
        with open(out / f"{name}.csv", "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["month", "line", "value"])
            w.writerows(rows)
    print("Wrote data/budget.csv and data/actual.csv")


if __name__ == "__main__":
    build()
