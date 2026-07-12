"""
דוח בוקר אוטומטי: בודק מיילים מהלילה (ואירועי יומן אם הוגדר קישור iCal),
מסכם עם Gemini (או IAC כגיבוי), ושולח מייל סיכום למשתמש.

גרסה חדשה: משתמש ב-IMAP/SMTP עם סיסמת אפליקציה (App Password) במקום OAuth,
כך שהטוקן לא פג לעולם והדוח לא נשבר.

רץ עצמאי (לא תלוי ב-Streamlit) דרך GitHub Actions, גם כשהמחשב כבוי.
"""

import email
import imaplib
import os
import smtplib
import ssl
from datetime import datetime, timedelta
from email.header import decode_header, make_header
from email.mime.text import MIMEText
from email.utils import parsedate_to_datetime

import requests

GEMINI_MODEL = "gemini-2.5-flash"
IAC_BASE_URL = "https://server.iac.ac.il/api/v1/studentapi"

MY_EMAIL = os.environ.get("DIGEST_TO_EMAIL", "elankimi1@gmail.com")
GMAIL_USER = os.environ.get("GMAIL_USER", MY_EMAIL)
# סיסמת אפליקציה של Google (16 אותיות, בלי רווחים)
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "")
# קישור סודי בפורמט iCal ליומן (אופציונלי) - אם ריק, מדלגים על אירועים
ICAL_URL = os.environ.get("ICAL_URL", "").strip()


def _decode(value: str) -> str:
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value or ""


def fetch_recent_emails(hours: int = 12) -> str:
    """קורא מיילים מהשעות האחרונות דרך IMAP עם סיסמת אפליקציה."""
    if not GMAIL_APP_PASSWORD:
        return "(לא הוגדרה סיסמת אפליקציה - אי אפשר לקרוא מיילים)"
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    imap.select("INBOX")

    since_date = (datetime.now() - timedelta(hours=hours)).strftime("%d-%b-%Y")
    status, data = imap.search(None, f'(SINCE "{since_date}")')
    ids = data[0].split() if data and data[0] else []
    if not ids:
        imap.logout()
        return "אין מיילים חדשים מהלילה."

    cutoff = datetime.now().astimezone() - timedelta(hours=hours)
    summaries = []
    for msg_id in ids[-20:]:
        status, msg_data = imap.fetch(msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
        if not msg_data or not msg_data[0]:
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        # מסננים לפי זמן אמיתי (SINCE ב-IMAP הוא ברמת יום)
        try:
            msg_dt = parsedate_to_datetime(msg.get("Date"))
            if msg_dt and msg_dt < cutoff:
                continue
        except Exception:
            pass
        frm = _decode(msg.get("From", "?"))
        subj = _decode(msg.get("Subject", "(ללא נושא)"))
        summaries.append(f"מאת: {frm} | נושא: {subj}")

    imap.logout()
    return "\n".join(summaries) if summaries else "אין מיילים חדשים מהלילה."


def fetch_today_events() -> str:
    """קורא אירועי היום מקישור iCal סודי (אם הוגדר)."""
    if not ICAL_URL:
        return "(לא הוגדר קישור יומן - מדלגים על אירועים)"
    try:
        r = requests.get(ICAL_URL, timeout=30)
        r.raise_for_status()
        today = datetime.now().date()
        events = []
        current = {}
        for line in r.text.splitlines():
            if line.startswith("BEGIN:VEVENT"):
                current = {}
            elif line.startswith("SUMMARY:"):
                current["summary"] = line[len("SUMMARY:"):].strip()
            elif line.startswith("DTSTART"):
                val = line.split(":", 1)[-1].strip()
                current["start"] = val
            elif line.startswith("END:VEVENT"):
                start = current.get("start", "")
                if start[:8] == today.strftime("%Y%m%d"):
                    events.append(f"{start} | {current.get('summary', '(ללא כותרת)')}")
        return "\n".join(events) if events else "אין אירועים מתוכננים להיום."
    except Exception as e:
        return f"(לא הצלחתי לקרוא את היומן: {e})"


def summarize_with_gemini(emails_text: str, events_text: str) -> str:
    from google import genai

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("no GEMINI_API_KEY")

    client = genai.Client(api_key=api_key)
    prompt = (
        "אתה עוזר אישי שמכין דוח בוקר קצר וברור בעברית. "
        "סכם בקצרה את הדברים החשובים מהמיילים הבאים (אם יש דחוף/חשוב - תדגיש), "
        "ואת האירועים של היום מלוח השנה. אל תמציא מידע שלא קיים.\n\n"
        f"מיילים מהלילה:\n{emails_text}\n\nאירועים להיום:\n{events_text}"
    )
    response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
    return response.text


def summarize_with_iac(emails_text: str, events_text: str) -> str:
    token = os.environ.get("IAC_TOKEN")
    if not token:
        raise RuntimeError("no IAC_TOKEN")

    headers = {"Authorization": f"Bearer {token}"}
    prompt = (
        "אתה עוזר אישי שמכין דוח בוקר קצר וברור בעברית. "
        "סכם בקצרה את הדברים החשובים מהמיילים הבאים (אם יש דחוף/חשוב - תדגיש), "
        "ואת האירועים של היום מלוח השנה. אל תמציא מידע שלא קיים.\n\n"
        f"מיילים מהלילה:\n{emails_text}\n\nאירועים להיום:\n{events_text}"
    )
    payload = {"messages": [{"role": "user", "content": prompt}], "max_completion_tokens": 2000}
    r = requests.post(f"{IAC_BASE_URL}/chat/completions", headers=headers, json=payload, timeout=60)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]


def send_digest_email(subject: str, body: str):
    """שולח את הדוח דרך SMTP עם סיסמת אפליקציה."""
    msg = MIMEText(body, _charset="utf-8")
    msg["From"] = GMAIL_USER
    msg["To"] = MY_EMAIL
    msg["Subject"] = subject

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_USER, [MY_EMAIL], msg.as_string())


def main():
    emails_text = fetch_recent_emails()
    events_text = fetch_today_events()

    try:
        summary = summarize_with_gemini(emails_text, events_text)
    except Exception:
        summary = summarize_with_iac(emails_text, events_text)

    today_str = datetime.now().strftime("%d/%m/%Y")
    send_digest_email(f"דוח בוקר - {today_str}", summary)
    print("דוח הבוקר נשלח בהצלחה.")


if __name__ == "__main__":
    main()
