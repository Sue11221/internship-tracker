"""Email the new postings shortlist.py queued in data/shortlist_new.json.

Order in the email: Manhattan / Jersey City, then remote, then other US
cities; inside each, Quant, Data Science, AI Consulting, Backend.

Config: email_config.json next to this file (addresses and server only).
The Gmail app password is NOT kept in any file. It lives in Windows
Credential Manager, stored once with --set-password (or set_email_password.bat).

    .venv\\Scripts\\python.exe send_email.py                 # send queued new postings
    .venv\\Scripts\\python.exe send_email.py --test          # send the current full shortlist
    .venv\\Scripts\\python.exe send_email.py --set-password  # store the app password in the vault, then test
    .venv\\Scripts\\python.exe send_email.py --clear-password

No config or no stored password: prints a notice and exits 0, and the queue
is left intact so nothing is lost. A failed send also keeps the queue.
"""

import getpass
import json
import os
import smtplib
import sys
from datetime import datetime
from email.message import EmailMessage
from html import escape

try:
    import keyring  # Windows Credential Manager on this machine
except ImportError:  # pragma: no cover
    keyring = None

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(ROOT, "email_config.json")
NEW = os.path.join(ROOT, "data", "shortlist_new.json")
SEEN = os.path.join(ROOT, "data", "shortlist_seen.json")
LATEST = os.path.join(ROOT, "data", "shortlist_latest.json")
VAULT_SERVICE = "internship-tracker-gmail"  # entry name in Windows Credential Manager

TIERS = [("nyc", "Manhattan / Jersey City"), ("remote", "Remote / online"), ("other", "Other US cities")]
GROUPS = ["Quant", "Data Science / ML / AI", "AI & Data Consulting", "Internet / Backend Software"]


def load_config(need_password=True):
    if not os.path.exists(CONFIG):
        print("send_email: no email_config.json, skipping.")
        return None
    cfg = json.load(open(CONFIG, encoding="utf-8"))
    if not cfg.get("username") or not cfg.get("to"):
        print("send_email: email_config.json needs 'username' and 'to'.")
        return None
    if not need_password:
        return cfg
    pw = vault_get(cfg["username"])
    if not pw:
        print("send_email: no app password stored yet, skipping. Double-click set_email_password.bat once.")
        return None
    cfg["app_password"] = pw
    return cfg


def vault_get(username):
    if keyring is None:
        print("send_email: the 'keyring' package is missing (pip install keyring).")
        return None
    try:
        return keyring.get_password(VAULT_SERVICE, username)
    except Exception as exc:
        print(f"send_email: could not read Windows Credential Manager: {exc}")
        return None


def smtp_login(cfg):
    """Open the SMTP connection and log in; the caller closes it."""
    s = smtplib.SMTP(cfg.get("smtp_host", "smtp.gmail.com"), int(cfg.get("smtp_port", 587)), timeout=30)
    s.ehlo()
    s.starttls()
    s.login(cfg["username"], cfg["app_password"])
    return s


def set_password(cfg):
    """Ask for the app password (hidden), check it against Gmail, store it in the vault."""
    if keyring is None:
        print("The 'keyring' package is missing: .venv\\Scripts\\python.exe -m pip install keyring")
        return 1
    print(f"Gmail account: {cfg['username']}")
    print("Get a 16-letter app password at https://myaccount.google.com/apppasswords")
    print("Paste it below. Nothing is shown while you type; spaces are fine.")
    pw = getpass.getpass("App password: ").replace(" ", "").strip()
    if not pw:
        print("Nothing entered, nothing saved.")
        return 1
    cfg["app_password"] = pw
    print("Checking it with Gmail...")
    try:
        smtp_login(cfg).quit()
    except Exception as exc:
        print(f"Gmail rejected it, nothing saved: {exc}")
        return 1
    keyring.set_password(VAULT_SERVICE, cfg["username"], pw)
    print(f"Saved in Windows Credential Manager under '{VAULT_SERVICE}'. No file in this folder contains it.")
    return 0


def clear_password(cfg):
    if keyring is None:
        return 1
    try:
        keyring.delete_password(VAULT_SERVICE, cfg["username"])
        print("Removed from Windows Credential Manager.")
    except keyring.errors.PasswordDeleteError:
        print("Nothing stored, nothing to remove.")
    return 0


def render(records, heading):
    html = [f"<h2 style='font-family:sans-serif'>{escape(heading)}</h2>"]
    text = [heading, ""]
    for tkey, tlabel in TIERS:
        in_tier = [r for r in records if r["tier"] == tkey]
        if not in_tier:
            continue
        html.append(f"<h3 style='font-family:sans-serif;border-bottom:2px solid #333'>"
                    f"{escape(tlabel)} ({len(in_tier)})</h3>")
        text += [f"##### {tlabel} ({len(in_tier)}) #####", ""]
        for g in GROUPS:
            rows = [r for r in in_tier if r["group"] == g]
            if not rows:
                continue
            html.append(f"<h4 style='font-family:sans-serif;margin:12px 0 4px'>{escape(g)} ({len(rows)})</h4>")
            html.append("<table style='font-family:sans-serif;font-size:14px;border-collapse:collapse'>")
            text += [f"== {g} ({len(rows)}) =="]
            for r in rows:
                html.append(
                    "<tr><td style='padding:4px 8px;border-bottom:1px solid #ddd'>"
                    f"<b>{escape(r['company'])}</b><br>"
                    f"<a href='{escape(r['url'])}'>{escape(r['title'])}</a><br>"
                    f"<span style='color:#555'>{escape(r['location'][:90])} · {escape(r['season'])}"
                    f" · posted {escape(r['posted'])}</span></td></tr>")
                text.append(f"- {r['company']}: {r['title']}\n"
                            f"  {r['location'][:90]} | {r['season']} | posted {r['posted']}\n  {r['url']}")
            html.append("</table>")
            text.append("")
    return "\n".join(html), "\n".join(text)


def send(cfg, subject, html, text):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = cfg.get("from") or cfg["username"]
    msg["To"] = cfg["to"]
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")
    with smtp_login(cfg) as s:
        s.send_message(msg)


def summary(records):
    n = {k: sum(1 for r in records if r["tier"] == k) for k, _ in TIERS}
    return f"{len(records)} new ({n['nyc']} NYC, {n['remote']} remote, {n['other']} other US)"


def send_test(cfg):
    if not os.path.exists(LATEST):
        print("send_email: no data/shortlist_latest.json yet, run shortlist.py first.")
        return 1
    rows = json.load(open(LATEST, encoding="utf-8")).get("records", [])
    html, text = render(rows, "Test: current full shortlist")
    send(cfg, f"[Internships] test email, {len(rows)} postings", html, text)
    print(f"send_email: test email sent to {cfg['to']} ({len(rows)} postings).")
    return 0


def main():
    if "--set-password" in sys.argv:
        cfg = load_config(need_password=False)
        if cfg is None:
            return 1
        rc = set_password(cfg)
        return rc if rc else send_test(cfg)
    if "--clear-password" in sys.argv:
        cfg = load_config(need_password=False)
        return clear_password(cfg) if cfg else 1

    cfg = load_config()
    if cfg is None:
        return 0

    if "--test" in sys.argv:
        return send_test(cfg)

    queue = json.load(open(NEW, encoding="utf-8")) if os.path.exists(NEW) else {"records": []}
    records = queue.get("records", [])
    if cfg.get("only_nyc"):
        records = [r for r in records if r["tier"] == "nyc"]
    if not records:
        print("send_email: nothing new to send.")
        return 0

    stamp = datetime.now().strftime("%Y-%m-%d")
    html, text = render(records, f"{summary(records)}, {stamp}")
    try:
        send(cfg, f"[Internships] {summary(records)}, {stamp}", html, text)
    except Exception as exc:  # keep the queue so the next run retries
        print(f"send_email: FAILED, queue kept: {exc}")
        return 1

    # Mark as reported: {id: date}. shortlist.py deletes entries past its age cutoff.
    seen = json.load(open(SEEN, encoding="utf-8")) if os.path.exists(SEEN) else {}
    if isinstance(seen, list):
        seen = {i: stamp for i in seen}
    for r in queue.get("records", []):
        for i in r.get("ids", []):
            seen.setdefault(i, stamp)
    json.dump(seen, open(SEEN, "w", encoding="utf-8"), indent=0, sort_keys=True)
    json.dump({"generated_at": queue.get("generated_at"), "records": []}, open(NEW, "w", encoding="utf-8"))
    print(f"send_email: sent {len(records)} postings to {cfg['to']}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
