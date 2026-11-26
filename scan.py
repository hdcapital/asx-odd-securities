"""
ASX Odd Securities Scanner
==========================

Finds every quoted ASX security that is NOT a plain ordinary share
(company options, preference shares, hybrids/capital notes, convertible
bonds, listed notes, partly paid shares, deferred-settlement lines,
rights) and flags the ones that look mispriced.

Run:   python scan.py            (writes output/asx_odd_securities_YYYY-MM-DD.xlsx)
       python scan.py --email    (also emails the report, needs SMTP settings)

Data:  ASX ISIN list (every security code and its description) plus the
       delayed quote feed that powers asx.com.au company pages.
"""

from __future__ import annotations

import argparse
import datetime as dt
import math
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd
import requests

try:  # .env support for local runs (GitHub uses repo secrets instead)
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUT_DIR = ROOT / "output"
TERMS_FILE = ROOT / "option_terms.csv"
SEC_TERMS_FILE = ROOT / "security_terms.csv"  # researched terms for every security

ISIN_URLS = [
    "https://www.asx.com.au/content/dam/asx/issuers/ISIN.xls",
    "https://www.asx.com.au/programs/ISIN.xls",
]
ISIN_PAGE = "https://www.asx.com.au/markets/market-resources/asx-codes-and-descriptors"
API = "https://asx.api.markitdigital.com/asx-research/1.0/companies/{code}/{what}"

RISK_FREE = float(os.getenv("RISK_FREE_RATE", "0.04"))  # for the option model
BBSW_3M = float(os.getenv("BBSW_3M", "0.036"))  # used to estimate floating-rate coupons
WORKERS = int(os.getenv("WORKERS", "6"))

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; asx-odd-securities/1.0)",
    "Accept": "application/json, text/plain, */*",
}

ORDINARY_TYPES = (
    "ORDINARY FULLY PAID",
    "ORDINARY UNITS FULLY PAID",
    "FULLY PAID ORDINARY/UNITS STAPLED SECURITIES",
    "FULLY PAID UNITS STAPLED SECURITIES",
    "CHESS DEPOSITARY INTERESTS",
    "CDI ",
)
# Things in the ISIN list that are not interesting / not tradeable by retail
EXCLUDE_WORDS = (
    "MIN CONS",  # wholesale-only debt ($500k minimum)
    "PLACEMENT",
    "INITIAL PUBLIC OFFER",
    "TEST ",
    "BSK",
    "BASKET",
    "STRUCTURED",
    " ETF",
    "RESTRICTED",
    "PERFORMANCE",
    "DEFERRED PURCHASE AGREEMENT",
    "WARRANT",
    "WARR ",
    "INSTAL",
    "CTW ",
    "CITIWARRANTS",
)

MONTHS = {m: i for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], 1)}

session = requests.Session()
session.headers.update(HEADERS)


def log(msg: str) -> None:
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


# --------------------------------------------------------------------------
# 1. Universe: the ASX ISIN list
# --------------------------------------------------------------------------

def download_isin_list() -> tuple[pd.DataFrame, str]:
    urls = list(ISIN_URLS)
    try:  # ASX moves this file occasionally; look for the current link
        html = session.get(ISIN_PAGE, timeout=30).text
        for href in re.findall(r'href="([^"]*ISIN\.xls)"', html):
            full = href if href.startswith("http") else "https://www.asx.com.au" + href
            if full not in urls:
                urls.insert(0, full)
    except requests.RequestException:
        pass

    for url in urls:
        try:
            r = session.get(url, timeout=60)
            if r.status_code == 200 and r.content[:4] == b"\xd0\xcf\x11\xe0":
                path = DATA_DIR / "ISIN.xls"
                path.parent.mkdir(exist_ok=True)
                path.write_bytes(r.content)
                raw = pd.read_excel(path, header=None)
                break
        except requests.RequestException:
            continue
    else:
        sys.exit("Could not download the ASX ISIN list (all known URLs failed).")

    updated = ""
    for v in raw.iloc[:6, 0].fillna("").map(str):
        m = re.search(r"updated\s*:\s*(\S+)", v, re.I)
        if m:
            updated = m.group(1)
    df = raw.iloc[5:, :4].copy()
    df.columns = ["code", "name", "type", "isin"]
    df = df.dropna(subset=["code"])
    for c in df.columns:
        df[c] = df[c].fillna("").map(lambda v: str(v).strip())
    return df.reset_index(drop=True), updated


def classify(t: str) -> str:
    t = t.upper()
    if t.startswith("OPTION"):
        return "Option"
    if "PARTLY PAID" in t:
        return "Partly paid"
    if "CONVERT" in t or t.startswith("CNV ") or "CONV " in t:
        return "Convertible"
    if any(k in t for k in ("PREF", "CAP NOTE", "CAPITAL NOTE", "PERP")):
        return "Hybrid / preference"
    if any(k in t for k in ("NOTE", "BOND", "LOAN")):
        return "Note / bond"
    if t.startswith("RIGHTS") or "ENT TO " in t:
        return "Rights"
    if "DEFERRED" in t and "ORDINARY" in t:
        return "Deferred settlement"
    return "Other"


def build_universe(isin: pd.DataFrame) -> pd.DataFrame:
    is_ord = isin["type"].str.upper().str.startswith(ORDINARY_TYPES) & (isin.code.str.len() <= 4)
    ords = isin[is_ord & (isin.code.str.len() == 3)]
    name_to_parent = dict(zip(ords["name"], ords["code"]))
    ord_codes = set(ords["code"])

    rows = []
    for r in isin.itertuples():
        code, t = r.code, r.type.upper()
        if any(w in t for w in EXCLUDE_WORDS):
            continue
        if len(code) == 3:
            if is_ord[r.Index] or classify(t) == "Other":
                continue  # plain ordinary share
        elif code[3] == "A":
            continue  # 'A' suffix = unquoted (employee options, rights etc.)
        parent = name_to_parent.get(r.name) or (code[:3] if code[:3] in ord_codes else "")
        if not parent:
            continue  # ETFs, bank warrants, govt bonds etc. that have no listed parent share
        rows.append(dict(code=code, parent=parent, company=r.name, description=r.type,
                         category=classify(t), isin=r.isin, source="ASX ISIN list"))
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# 2. Prices
# --------------------------------------------------------------------------

def get_json(code: str, what: str) -> dict | None:
    for attempt in range(3):
        try:
            r = session.get(API.format(code=code, what=what), timeout=20)
            if r.status_code == 200:
                return r.json().get("data")
            if r.status_code in (400, 404):
                return None
        except (requests.RequestException, ValueError):
            pass
        time.sleep(1.5 * (attempt + 1))
    return None


def quote(code: str) -> dict | None:
    h = get_json(code, "header")
    if not h:
        return None
    k = get_json(code, "key-statistics") or {}
    return {
        "code": code,
        "api_name": h.get("displayName"),
        "bid": h.get("priceBid"),
        "ask": h.get("priceAsk"),
        "last": h.get("priceLast"),
        "volume_today": h.get("volume"),
        "status": h.get("statusCode") or "",
        "sec_type": h.get("securityType"),
        "avg_volume": k.get("volumeAverage"),
        "high_52w": k.get("priceFiftyTwoWeekHigh"),
        "low_52w": k.get("priceFiftyTwoWeekLow"),
        "api_description": k.get("shareDescription") or "",
        "yield_api": k.get("yieldAnnual"),
        "next_ex_date": k.get("dateExDate"),
        "dividend": k.get("dividend"),
    }


def quote_many(codes: list[str], label: str) -> dict[str, dict]:
    codes = sorted(set(codes))
    log(f"Fetching quotes for {len(codes)} {label}...")
    with ThreadPoolExecutor(WORKERS) as ex:
        res = list(ex.map(quote, codes))
    return {r["code"]: r for r in res if r}


def probe_new_series(univ: pd.DataFrame, ords: list[str]) -> pd.DataFrame:
    """The ISIN list is only refreshed every few weeks. Look for option and
    hybrid series listed since then by trying the next codes in sequence."""
    tries = set()
    have = set(univ.code)
    for p in ords:
        if p + "O" not in have:
            tries.add(p + "O")
    for code in univ.code:
        m = re.fullmatch(r"([A-Z0-9]{3})(O|P)([A-Z]?)", code)
        if m:
            base, kind, letter = m.groups()
            nxt = "A" if not letter else (chr(ord(letter) + 1) if letter < "Z" else None)
            if nxt and base + kind + nxt not in have:
                tries.add(base + kind + nxt)
    log(f"Probing {len(tries)} possible new codes not yet in the ISIN list...")
    with ThreadPoolExecutor(WORKERS) as ex:
        found = [h for h in ex.map(lambda c: (c, get_json(c, "header")), sorted(tries)) if h[1]]
    rows = []
    for code, h in found:
        desc = h.get("displayName") or ""
        cat = classify(desc)
        if h.get("securityType") == 40:
            cat = "Option"
        elif cat == "Other" and h.get("securityType") == 21:
            cat = "Hybrid / preference"
        rows.append(dict(code=code, parent=code[:3], company="", description=desc,
                         category=cat, isin="", source="Found by probing (new)"))
    log(f"  found {len(rows)} extra securities")
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------
# 3. Parsing terms
# --------------------------------------------------------------------------

def parse_date(s: str) -> dt.date | None:
    s = s.upper()
    m = re.search(r"(\d{1,2})-([A-Z]{3})-(\d{2,4})", s)
    if m and m.group(2) in MONTHS:
        y = int(m.group(3))
        y = y + 2000 if y < 100 else y
        try:
            return dt.date(y, MONTHS[m.group(2)], int(m.group(1)))
        except ValueError:
            return None
    m = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2})\b", s)
    if m:
        try:
            return dt.date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def parse_money(s: str) -> float | None:
    s = s.strip().upper()
    if s == "NIL":
        return 0.0
    m = re.fullmatch(r"\$?([\d.]+)(C?)", s)
    if not m:
        return None
    v = float(m.group(1))
    return v / 100 if m.group(2) == "C" else v


def option_strike_from_text(t: str) -> float | None:
    m = re.search(r"\bEX\s+(NIL|\$?[\d.]+C?)", t.upper())
    return parse_money(m.group(1)) if m else None


def load_manual_terms() -> pd.DataFrame:
    if TERMS_FILE.exists():
        df = pd.read_csv(TERMS_FILE, dtype=str).fillna("")
        return df
    return pd.DataFrame(columns=["code", "strike", "shares_per_option", "note"])


def load_security_terms() -> dict[str, dict]:
    if SEC_TERMS_FILE.exists():
        df = pd.read_csv(SEC_TERMS_FILE, dtype=str).fillna("")
        return {r["code"]: r for r in df.to_dict("records")}
    return {}


def _iso(s: str) -> dt.date | None:
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except ValueError:
        return None


def infer_strikes(isin: pd.DataFrame) -> dict[tuple[str, str], float]:
    """Unquoted option series (code suffix A..) usually show the strike.
    If a company has exactly one strike for a given expiry, the quoted
    series with that expiry very likely has the same strike."""
    out: dict[tuple[str, str], set] = {}
    for r in isin.itertuples():
        t = r.type.upper()
        if not t.startswith("OPTION") or len(r.code) < 4 or r.code[3] != "A":
            continue
        k = option_strike_from_text(t)
        d = parse_date(t)
        if k is None or d is None:
            continue
        out.setdefault((r.code[:3], d.isoformat()), set()).add(k)
    return {key: next(iter(v)) for key, v in out.items() if len(v) == 1}


# --------------------------------------------------------------------------
# 4. Analytics
# --------------------------------------------------------------------------

def norm_cdf(x: float) -> float:
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def bs_call(S, K, T, r, sigma) -> float | None:
    if not all(v is not None and v == v for v in (S, K, T, sigma)) or S <= 0 or T <= 0 or sigma <= 0:
        return None
    if K <= 0:
        return S
    d1 = (math.log(S / K) + (r + sigma ** 2 / 2) * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)
    return max(S * norm_cdf(d1) - K * math.exp(-r * T) * norm_cdf(d2), 0.0)


def vol_from_range(hi, lo) -> float | None:
    """Rough annual volatility from the 52-week high/low (Parkinson range
    estimator). Crude, but needs no price history."""
    if hi and lo and hi > 0 and lo > 0 and hi > lo:
        return min(max(math.log(hi / lo) / math.sqrt(4 * math.log(2)), 0.15), 2.5)
    return None


def num(v):
    try:
        f = float(v)
        return f if f == f and f > 0 else None
    except (TypeError, ValueError):
        return None


def analyse_option(r: dict, P: dict, today: dt.date, inferred, manual) -> dict:
    out = {}
    exp = parse_date(r["description"])
    out["expiry"] = exp
    out["days_to_expiry"] = (exp - today).days if exp else None
    S = num(P.get("last"))
    S_bid = num(P.get("bid")) or S
    out["share_price"] = S

    strike, src, ratio = None, "", 1.0
    man = manual.get(r["code"])
    if man and man.get("strike", "") != "":
        note = str(man.get("note", ""))
        strike = float(man["strike"])
        src = note.split(":")[0] if note.startswith("Researched") else "Manual (option_terms.csv)"
        if man.get("shares_per_option"):
            ratio = float(man["shares_per_option"])
    if strike is None:
        k = option_strike_from_text(r["description"])
        if k is not None:
            strike, src = k, "ASX description"
    if strike is None and exp is not None:
        k = inferred.get((r["parent"], exp.isoformat()))
        if k is not None:
            strike, src = k, "Inferred from unquoted series (check)"
    out["strike"] = strike
    out["strike_source"] = src or "UNKNOWN - add to option_terms.csv"

    ask, bid = num(r.get("ask")), num(r.get("bid"))
    # The highest strike at which buying the option and exercising is no
    # dearer than buying the shares. Useful when the strike is unknown.
    if S and ask is not None:
        out["breakeven_strike"] = round(S - ask / ratio, 6)
    flags = []
    if strike is not None and S:
        intrinsic = max(S_bid * ratio - strike * ratio, 0)
        out["intrinsic"] = intrinsic
        out["moneyness"] = S / strike if strike > 0 else None
        if ask is not None:
            out["time_value_pct_of_share"] = (ask - intrinsic) / (S * ratio)
            if intrinsic > 0 and ask < intrinsic * 0.98:
                flags.append(f"Ask below intrinsic value ({(1 - ask / intrinsic):.0%} discount)")
        T = (out["days_to_expiry"] or 0) / 365
        sigma = vol_from_range(num(P.get("high_52w")), num(P.get("low_52w")))
        out["vol_est"] = sigma
        model = bs_call(S, strike, T, RISK_FREE, sigma) if sigma else None
        if model is not None:
            model *= ratio
            out["model_value"] = model
            if ask:
                out["model_to_ask"] = model / ask
                if out["model_to_ask"] >= 2 and T > 0.5 and (out.get("moneyness") or 0) > 0.6:
                    flags.append(f"Cheap vs rough model ({out['model_to_ask']:.1f}x ask)")
        if intrinsic > 0 and 0 < (out["days_to_expiry"] or 0) <= 45:
            flags.append("In the money and expiring within 45 days")
    out["flags"] = flags
    return out


def parse_hybrid(desc: str, api_desc: str) -> dict:
    t = (desc + " " + api_desc).upper()
    out = {}
    m = re.search(r"BBSW\s*\+\s*([\d.]+)", t)
    if m:
        out["margin"] = float(m.group(1)) / 100
        out["rate_type"] = "Floating"
    else:
        m = re.search(r"([\d.]+)\s*%", t)
        if m:
            out["coupon"] = float(m.group(1)) / 100
            out["rate_type"] = "Fixed"
    out["perpetual"] = "PERP" in t
    out["cumulative"] = "NON-CUM" not in t and "CUM" in t
    maturity = None if out["perpetual"] else parse_date(re.sub(r"\bT-\d{1,2}-\d{2}\b", "", desc))
    m = re.search(r"\bT-(\d{1,2})-(\d{2})\b", t)  # first call / conversion date
    call = dt.date(2000 + int(m.group(2)), int(m.group(1)), 1) if m else None
    if call and call > dt.date.today():
        out["call_or_maturity"] = call
    elif maturity:
        out["call_or_maturity"] = maturity
    return out


def analyse_income(r: dict, today: dt.date, terms: dict | None = None) -> dict:
    out = parse_hybrid(r["description"], r.get("api_description", ""))
    terms = terms or {}
    good = terms.get("confidence") in ("high", "medium")
    px = num(r.get("ask")) or num(r.get("last"))
    face = num(terms.get("face_value")) if good else None
    if face is None and px:
        face = 100.0 if 50 <= px <= 200 else (1.0 if 0.5 <= px <= 1.5 else None)
    if good:
        for key in ("first_call_date", "maturity_date"):
            d = _iso(terms.get(key, ""))
            if d and d > today:
                out["call_or_maturity"] = d
                break
        m = re.search(r"([\d.]+)\s*%", terms.get("coupon_or_margin", ""))
        if m and "BBSW" in terms.get("coupon_or_margin", "").upper():
            out["margin"], out["rate_type"] = float(m.group(1)) / 100, "Floating"
            out.pop("coupon", None)
        elif m:
            out["coupon"], out["rate_type"] = float(m.group(1)) / 100, "Fixed"
            out.pop("margin", None)
    out["assumed_face"] = face
    flags = []
    if px and face:
        out["price_to_face"] = px / face
        cpn = out.get("coupon")
        if cpn is None and out.get("margin") is not None:
            cpn = BBSW_3M + out["margin"]
        out["running_yield"] = cpn * face / px if cpn else None
        when = out.get("call_or_maturity")
        if when and cpn:
            yrs = max((when - today).days / 365, 0.25)
            out["years_to_call_or_maturity"] = round(yrs, 2)
            out["approx_yield_to_call"] = (cpn * face + (face - px) / yrs) / ((face + px) / 2)
            if out.get("margin") is not None:
                out["approx_trading_margin"] = out["approx_yield_to_call"] - BBSW_3M
        if px / face < 0.95:
            flags.append(f"Trading at {px / face:.0%} of assumed face value")
    out["flags"] = flags
    return out


def analyse_partly(r: dict, P: dict) -> dict:
    t = r["description"].upper()
    out, flags = {}, []
    m = re.search(r"UNPAID\s+(\$?[\d.]+C?)", t)
    unpaid = parse_money(m.group(1)) if m else None
    out["unpaid_amount"] = unpaid
    ask, S = num(r.get("ask")), num(P.get("bid")) or num(P.get("last"))
    out["share_price"] = S
    if unpaid is not None and ask and S:
        cost = ask + unpaid
        out["fully_paid_equivalent_cost"] = cost
        out["discount_to_shares"] = 1 - cost / S
        if out["discount_to_shares"] > 0.03:
            flags.append(f"Pay-up cost {out['discount_to_shares']:.0%} below share price")
    out["flags"] = flags
    return out


def analyse_vs_parent(r: dict, P: dict) -> dict:
    out, flags = {}, []
    ask, S = num(r.get("ask")), num(P.get("bid")) or num(P.get("last"))
    out["share_price"] = S
    if ask and S:
        out["discount_to_shares"] = 1 - ask / S
        if out["discount_to_shares"] > 0.03:
            flags.append(f"{out['discount_to_shares']:.0%} below the ordinary shares")
    out["flags"] = flags
    return out


# --------------------------------------------------------------------------
# 5. Main
# --------------------------------------------------------------------------

def run(send_email: bool) -> Path:
    today = dt.date.today()
    DATA_DIR.mkdir(exist_ok=True)
    OUT_DIR.mkdir(exist_ok=True)

    log("Downloading the ASX ISIN list...")
    isin, isin_updated = download_isin_list()
    log(f"  {len(isin)} securities (list last updated {isin_updated or 'unknown'})")

    univ = build_universe(isin)
    is_ord = isin["type"].str.upper().str.startswith(ORDINARY_TYPES) & (isin.code.str.len() == 3)
    ord_codes = sorted(set(isin.loc[is_ord, "code"]))
    extra = probe_new_series(univ, ord_codes)
    if len(extra):
        univ = pd.concat([univ, extra], ignore_index=True)

    quotes = quote_many(list(univ.code), "odd securities")
    parents = quote_many(list(univ.parent), "parent shares")

    univ = univ[univ.code.isin(quotes)].copy()  # keep only quoted ones
    log(f"{len(univ)} quoted odd securities found")

    inferred = infer_strikes(isin)
    manual_df = load_manual_terms()
    manual = {r["code"]: r for r in manual_df.to_dict("records")}
    sec_terms = load_security_terms()

    records = []
    for r in univ.to_dict("records"):
        q = quotes[r["code"]]
        r.update({k: v for k, v in q.items() if k != "code"})
        if not r.get("company"):
            r["company"] = (parents.get(r["parent"]) or {}).get("api_name", "")
        P = parents.get(r["parent"], {})
        cat = r["category"]
        if cat == "Option":
            info = analyse_option(r, P, today, inferred, manual)
        elif cat in ("Hybrid / preference", "Convertible", "Note / bond"):
            info = analyse_income(r, today, sec_terms.get(r["code"]))
            info["share_price"] = num(P.get("last"))
        elif cat == "Partly paid":
            info = analyse_partly(r, P)
        elif cat == "Deferred settlement":
            info = analyse_vs_parent(r, P)
        else:
            info = {"share_price": num(P.get("last")), "flags": []}
        r.update(info)
        t = sec_terms.get(r["code"], {})
        r["key_features"] = t.get("key_features", "")
        r["terms_confidence"] = t.get("confidence", "")
        r["terms_source"] = t.get("source_url", "")

        bid, ask, last = num(r["bid"]), num(r["ask"]), num(r["last"])
        r["spread_pct"] = (ask - bid) / ((ask + bid) / 2) if bid and ask else None
        r["avg_daily_value"] = (r["avg_volume"] or 0) * (last or 0)
        if not bid and not ask:
            r["flags"].append("No bid or offer")
        r["flags"] = "; ".join(r["flags"])
        records.append(r)

    df = pd.DataFrame(records)

    # ---- what's new since the last run --------------------------------
    latest = DATA_DIR / "latest.csv"
    if latest.exists():
        prev = pd.read_csv(latest, dtype=str)
        prev_codes = set(prev.code)
        df["new_since_last_run"] = ~df.code.isin(prev_codes)
        gone = prev[~prev.code.isin(set(df.code))][["code", "company", "description", "category"]]
    else:
        df["new_since_last_run"] = False
        gone = pd.DataFrame(columns=["code", "company", "description", "category"])

    # ---- save history ---------------------------------------------------
    snap_cols = ["code", "parent", "company", "category", "description", "bid", "ask",
                 "last", "volume_today", "avg_volume", "share_price", "strike", "flags"]
    snap = df.reindex(columns=snap_cols)
    snap.to_csv(latest, index=False)
    hist = DATA_DIR / "history"
    hist.mkdir(exist_ok=True)
    snap.to_csv(hist / f"{today.isoformat()}.csv", index=False)

    # options missing a strike -> add blank rows to option_terms.csv
    missing = df[(df.category == "Option") & df.strike.isna()]
    if len(missing):
        add = missing[~missing.code.isin(set(manual_df.code))][["code"]].copy()
        if len(add):
            add["strike"] = ""
            add["shares_per_option"] = ""
            m = missing.set_index("code").loc[add.code]
            add["note"] = (m["company"] + " - " + m["description"]).values
            pd.concat([manual_df, add], ignore_index=True).sort_values("code").to_csv(TERMS_FILE, index=False)

    from report import write_excel
    path = OUT_DIR / f"asx_odd_securities_{today.isoformat()}.xlsx"
    write_excel(df, gone, path, isin_updated)
    log(f"Report written: {path}")

    if send_email:
        from emailer import send_report
        send_report(df, gone, path)
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--email", action="store_true", help="email the report when done")
    args = ap.parse_args()
    run(args.email)
