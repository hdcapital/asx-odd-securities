"""Emails the report. Settings come from environment variables
(a .env file locally, repository secrets on GitHub):

    SMTP_USER      e.g. you@gmail.com
    SMTP_PASSWORD  for Gmail this is an "app password", not your normal password
    EMAIL_TO       who gets it (comma-separate several addresses)
    SMTP_HOST      optional, default smtp.gmail.com
    SMTP_PORT      optional, default 587
    EMAIL_FROM     optional, defaults to SMTP_USER
"""

from __future__ import annotations

import datetime as dt
import html
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

import pandas as pd

from report import flagged


def _table(df: pd.DataFrame, cols: list[str]) -> str:
    if df.empty:
        return "<p><i>None.</i></p>"
    head = "".join(f"<th style='text-align:left;padding:4px 8px;background:#1F3864;color:#fff'>"
                   f"{html.escape(c.replace('_', ' ').capitalize())}</th>" for c in cols)
    rows = []
    for _, r in df.iterrows():
        cells = []
        for c in cols:
            v = r.get(c)
            if isinstance(v, float):
                v = "" if v != v else f"{v:,.3f}"
            cells.append(f"<td style='padding:4px 8px;border-bottom:1px solid #ddd'>"
                         f"{html.escape(str(v if v is not None else ''))}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return (f"<table style='border-collapse:collapse;font-family:Arial;font-size:12px'>"
            f"<tr>{head}</tr>{''.join(rows)}</table>")


def send_report(df: pd.DataFrame, gone: pd.DataFrame, xlsx: Path) -> None:
    user = os.getenv("SMTP_USER")
    pwd = os.getenv("SMTP_PASSWORD")
    to = os.getenv("EMAIL_TO")
    if not (user and pwd and to):
        print("Email skipped: SMTP_USER, SMTP_PASSWORD and EMAIL_TO must all be set.")
        return

    fl = flagged(df)
    new = df[df.new_since_last_run]
    today = dt.date.today()

    body = f"""
    <div style='font-family:Arial;font-size:13px'>
    <p><b>ASX odd securities scan - {today:%a %d %b %Y}</b><br>
    {len(df)} quoted securities scanned · {len(fl)} flagged · {len(new)} new · {len(gone)} gone.
    Full detail in the attached spreadsheet.</p>
    <h3>Flags</h3>
    {_table(fl.head(40), ['code', 'category', 'company', 'flags', 'bid', 'ask', 'share_price', 'strike'])}
    <h3>New since last run</h3>
    {_table(new, ['code', 'category', 'company', 'description'])}
    <h3>Gone since last run</h3>
    {_table(gone, ['code', 'category', 'company', 'description'])}
    <p style='color:#888'>Delayed data. Screening aid only - check terms and live prices before trading.</p>
    </div>"""

    msg = EmailMessage()
    msg["Subject"] = f"ASX odd securities: {len(fl)} flagged, {len(new)} new ({today:%d %b})"
    msg["From"] = os.getenv("EMAIL_FROM") or user
    msg["To"] = to
    msg.set_content("Your email program does not show HTML. See the attached spreadsheet.")
    msg.add_alternative(body, subtype="html")
    msg.add_attachment(xlsx.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       filename=xlsx.name)

    host = os.getenv("SMTP_HOST") or "smtp.gmail.com"
    port = int(os.getenv("SMTP_PORT") or 587)
    with smtplib.SMTP(host, port, timeout=60) as s:
        s.starttls()
        s.login(user, pwd)
        s.send_message(msg)
    print(f"Email sent to {to}")
