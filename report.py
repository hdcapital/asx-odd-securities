"""Writes the Excel report. Uses openpyxl only - never opens or touches Excel itself."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")
NEW_FILL = PatternFill("solid", fgColor="E2EFDA")

COMMON = ["code", "parent", "company", "description", "bid", "ask", "last",
          "volume_today", "avg_daily_value", "spread_pct", "flags", "new_since_last_run"]

SHEETS = {
    "Options": ["code", "parent", "company", "description", "expiry", "days_to_expiry",
                "share_price", "strike", "strike_source", "breakeven_strike", "bid", "ask",
                "last", "intrinsic", "moneyness", "time_value_pct_of_share", "vol_est",
                "model_value", "model_to_ask", "volume_today", "avg_daily_value",
                "spread_pct", "flags", "key_features", "terms_confidence", "terms_source",
                "new_since_last_run"],
    "Hybrids & notes": ["code", "parent", "company", "category", "description", "rate_type",
                        "coupon", "margin", "cumulative", "perpetual", "call_or_maturity",
                        "years_to_call_or_maturity", "bid", "ask", "last", "assumed_face",
                        "price_to_face", "running_yield", "approx_yield_to_call",
                        "approx_trading_margin", "yield_api", "next_ex_date",
                        "volume_today", "avg_daily_value", "spread_pct", "flags",
                        "key_features", "terms_confidence", "terms_source",
                        "new_since_last_run"],
    "Partly paid & deferred": ["code", "parent", "company", "category", "description",
                               "share_price", "unpaid_amount", "fully_paid_equivalent_cost",
                               "bid", "ask", "last", "discount_to_shares", "volume_today",
                               "avg_daily_value", "spread_pct", "flags", "key_features",
                               "terms_confidence", "terms_source", "new_since_last_run"],
    "Other": ["code", "parent", "company", "category", "description", "share_price", "bid",
              "ask", "last", "volume_today", "avg_daily_value", "spread_pct", "flags",
              "new_since_last_run"],
}

PCT = {"spread_pct", "time_value_pct_of_share", "vol_est", "coupon", "margin",
       "price_to_face", "running_yield", "approx_yield_to_call", "approx_trading_margin",
       "discount_to_shares"}
MONEY = {"bid", "ask", "last", "share_price", "strike", "breakeven_strike", "intrinsic",
         "model_value", "unpaid_amount", "fully_paid_equivalent_cost", "assumed_face"}
INT = {"volume_today", "avg_daily_value", "days_to_expiry"}
RATIO = {"moneyness", "model_to_ask", "years_to_call_or_maturity"}

NOTES = [
    ("What this is", "Every quoted ASX security that is not a plain ordinary share, with "
     "simple mispricing checks. Prices are delayed and the last price can be weeks old for "
     "thinly traded lines, so check bid/ask and volume before acting."),
    ("Flags tab", "Only securities that triggered a check, real ones first. Securities whose "
     "only flag is 'No bid or offer' are left off the Flags tab but stay on their own tab."),
    ("Options: strike", "ASX does not publish strikes for quoted options in its code list. "
     "Strikes come from option_terms.csv (researched from announcements, or typed in by you), "
     "then the ASX description, then inferred from an unquoted series with the same expiry. "
     "Anything still blank is listed in option_terms.csv to fill in."),
    ("Options: breakeven strike", "Share price minus option ask. If the real strike is BELOW "
     "this, buying the option and exercising costs less than buying the shares."),
    ("Options: model value", "Black-Scholes with volatility estimated from the share's 52-week "
     "high/low. Very rough: ignores dilution, dividends and thin trading. A screening aid only."),
    ("Hybrids: face value", "Assumed $100 if the price is $50-$200, $1 if $0.50-$1.50, "
     "otherwise left blank (check the terms). Called-but-unredeemed lines use maturity, "
     "or show running yield only if perpetual."),
    ("Hybrids: yields", "Approximate yield to first call/maturity, ignoring franking. Floating "
     "rate coupons use the BBSW_3M setting (default 3.6%) plus the margin."),
    ("Partly paid", "Fully-paid equivalent cost = partly paid ask + amount still unpaid, "
     "compared with the ordinary share bid."),
    ("All terms tab", "Researched key terms for every security: strike, expiry, options on issue, "
     "face value, coupon/margin, call and maturity dates, conversion terms, special features, "
     "with source link and confidence. Stored in security_terms.csv - edit it to correct anything."),
    ("Not covered", "Exchange-traded options, bank warrants, ETFs, wholesale-only debt and "
     "unquoted securities are deliberately excluded."),
]


def _fmt_sheet(ws, cols: list[str]) -> None:
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=1, column=i)
        cell.value = c.replace("_", " ").capitalize()
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        letter = get_column_letter(i)
        width = {"description": 42, "company": 30, "flags": 48, "strike_source": 26,
                 "key_features": 60, "terms_source": 30}.get(c, 12)
        ws.column_dimensions[letter].width = width
        fmt = None
        if c in PCT:
            fmt = "0.0%"
        elif c in MONEY:
            fmt = "#,##0.000"
        elif c in INT:
            fmt = "#,##0"
        elif c in RATIO:
            fmt = "0.00"
        for row in range(2, ws.max_row + 1):
            if fmt:
                ws.cell(row=row, column=i).number_format = fmt
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "B2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions
    if "flags" in cols:
        fi = cols.index("flags") + 1
        ni = cols.index("new_since_last_run") + 1 if "new_since_last_run" in cols else None
        for row in range(2, ws.max_row + 1):
            flag = ws.cell(row=row, column=fi).value
            new = ws.cell(row=row, column=ni).value if ni else False
            fill = NEW_FILL if new else (FLAG_FILL if flag and flag != "No bid or offer" else None)
            if fill:
                for col in range(1, len(cols) + 1):
                    ws.cell(row=row, column=col).fill = fill


def flagged(df: pd.DataFrame) -> pd.DataFrame:
    f = df[df["flags"].fillna("").str.len() > 0].copy()
    f = f[f["flags"] != "No bid or offer"]
    return f.sort_values(["category", "avg_daily_value"], ascending=[True, False])


def write_excel(df: pd.DataFrame, gone: pd.DataFrame, path: Path, isin_updated: str) -> None:
    df = df.copy()
    for c in ("expiry", "call_or_maturity"):
        if c in df:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        fl = flagged(df)
        flag_cols = ["code", "category", "company", "description", "flags", "bid", "ask",
                     "share_price", "strike", "avg_daily_value", "spread_pct",
                     "new_since_last_run"]
        fl.reindex(columns=flag_cols).to_excel(xw, "Flags", index=False)
        _fmt_sheet(xw.sheets["Flags"], flag_cols)

        cats = {
            "Options": df.category == "Option",
            "Hybrids & notes": df.category.isin(["Hybrid / preference", "Convertible", "Note / bond"]),
            "Partly paid & deferred": df.category.isin(["Partly paid", "Deferred settlement"]),
            "Other": df.category.isin(["Rights", "Other"]),
        }
        for name, mask in cats.items():
            cols = SHEETS[name]
            sub = df[mask].reindex(columns=cols)
            sort = "days_to_expiry" if name == "Options" else "code"
            sub.sort_values(sort).to_excel(xw, name, index=False)
            _fmt_sheet(xw.sheets[name], cols)

        new = df[df.new_since_last_run].reindex(columns=["code", "category", "company", "description"])
        new.to_excel(xw, "New listings", index=False)
        _fmt_sheet(xw.sheets["New listings"], list(new.columns))
        gone.to_excel(xw, "Gone since last run", index=False)
        _fmt_sheet(xw.sheets["Gone since last run"], list(gone.columns))

        terms_path = Path(__file__).resolve().parent / "security_terms.csv"
        if terms_path.exists():
            t = pd.read_csv(terms_path, dtype=str).fillna("")
            t = df[["code", "company", "category", "description", "ask"]].merge(
                t.drop(columns=["category"], errors="ignore"), on="code", how="left")
            t.sort_values(["category", "code"]).to_excel(xw, "All terms", index=False)
            _fmt_sheet(xw.sheets["All terms"], list(t.columns))

        notes = pd.DataFrame(NOTES + [("ASX code list date", isin_updated or "unknown"),
                                      ("Securities scanned", str(len(df)))],
                             columns=["Topic", "Explanation"])
        notes.to_excel(xw, "Read me", index=False)
        ws = xw.sheets["Read me"]
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 110
        for row in ws.iter_rows(min_row=2):
            row[1].alignment = Alignment(wrap_text=True, vertical="top")
            row[0].font = Font(bold=True)
