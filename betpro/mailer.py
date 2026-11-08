"""Trimite raportul pe Gmail (parola de aplicatie din GitHub Secrets)."""
import os
import smtplib
from email.mime.text import MIMEText


def send(subject: str, html: str) -> str:
    user = os.environ.get("GMAIL_USER", "").strip()
    pw = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "").strip()
    to = [x.strip() for x in os.environ.get("MAIL_TO", user).split(",") if x.strip()]
    if not user or not pw or not to:
        return "lipsesc GMAIL_USER / GMAIL_APP_PASSWORD - emailul nu a plecat"
    msg = MIMEText(html, "html", "utf-8")
    msg["Subject"] = subject
    msg["From"] = f"BET PRO <{user}>"
    msg["To"] = user
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
        s.login(user, pw)
        s.sendmail(user, to, msg.as_string())
    return f"email trimis catre {len(to)} adrese"
