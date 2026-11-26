# ASX odd securities scanner

Every weekday evening this scans **every quoted ASX security that isn't a plain ordinary share**
and emails you a spreadsheet with anything that looks mispriced. It covers:

- **Company options** (the `XYZO`, `XYZOA` ... codes)
- **Hybrids and preference shares** (bank capital notes, cumulative prefs, etc.)
- **Convertible bonds and listed notes** (excluding $500k-minimum wholesale debt)
- **Partly paid shares**, **deferred settlement** lines and **rights**

ETFs, bank warrants, exchange-traded options and unquoted securities are left out on purpose.

## What it checks

| Security | Check |
|---|---|
| Options | Ask below intrinsic value · in the money and expiring within 45 days · cheap against a rough option model |
| Hybrids / converts / notes | Trading below 95% of face value · approximate yield to call and trading margin |
| Partly paid | Partly paid price + amount still owed, against the ordinary share price |
| Deferred settlement | Discount to the ordinary shares |
| Everything | New listings and delistings since the last run, liquidity, bid/ask spread |

The **Flags** tab in the spreadsheet holds just the hits. Green rows are new since the last run.

## Option strike prices: one thing you need to do

ASX doesn't publish strike prices for quoted options in its code list, so most options start with
the strike blank. The scanner fills strikes in where it can (about 10% so far), and adds every
missing one to **`option_terms.csv`** for you.

To add strikes, open `option_terms.csv` on GitHub, click the pencil icon, type the strike (in dollars,
e.g. `0.05`) next to each code you care about, and click **Commit changes**. You only do this once per
option series. If one option converts into more than one share, put that number in `shares_per_option`.

Until a strike is added, the **Breakeven strike** column still helps. It's the share price minus the
option's ask. If the real strike is below that number, buying the option and exercising costs less
than buying the shares.

## Running on GitHub (automatic)

It runs by itself Monday–Friday at about 6:30pm Sydney time.

**To get it emailed to you**, go to the repo's **Settings → Secrets and variables → Actions →
New repository secret** and add:

| Name | What to put |
|---|---|
| `SMTP_USER` | The Gmail address that sends it |
| `SMTP_PASSWORD` | A Gmail **app password** (Google account → Security → 2-Step Verification → App passwords). Not your normal password. |
| `EMAIL_TO` | Where it goes (separate several with commas) |

Without these it still runs, and you download the spreadsheet from the run instead:
**Actions** tab → click the latest run → **Artifacts** at the bottom.

**To run it now:** Actions tab → *Daily ASX odd securities scan* → **Run workflow**.

## Running on your PC

1. Download the repo (green **Code** button → **Download ZIP**) and unzip it.
2. Double-click **`run.bat`**.
3. The spreadsheet appears in the `output` folder. It's a new file, so your open Excel
   workbooks aren't touched.

To email it from your PC too, copy `.env.example` to `.env`, fill it in, then run
`run.bat --email` from Command Prompt.

## Good to know

- **The data is delayed.** For thinly traded lines, the "last" price can be weeks old, so go by
  bid/ask and volume.
- The option model value is crude. It estimates volatility from the 52-week high and low and
  ignores dilution, so treat it as a screen and nothing more.
- Hybrid yields assume $100 face value (or $1 for small converts), ignore franking, and use a
  3.6% BBSW assumption for floating coupons. You can change that with a `BBSW_3M` setting.
- Data comes from the ASX code list and the delayed quotes behind asx.com.au. **ASX's website terms
  allow private/personal use. Commercial use by a fund manager may need an ASX data licence.**
- **This repo is public.** Anyone can see the code and the daily snapshots in `data/`. To change
  that, go to Settings → General → Danger zone → Change visibility.
- Daily snapshots build up in `data/history/`, which gives you a record of every odd security over time.

## Files

| File | What it is |
|---|---|
| `scan.py` | The scanner |
| `report.py` | Builds the spreadsheet |
| `emailer.py` | Sends the email |
| `option_terms.csv` | Your option strike prices |
| `data/latest.csv`, `data/history/` | Daily snapshots (used to spot new listings) |
| `.github/workflows/scan.yml` | The daily schedule |
| `run.bat` | Double-click to run it on Windows |
