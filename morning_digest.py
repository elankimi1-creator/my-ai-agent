"""
דוח בוקר אוטומטי: בודק מיילים מהלילה ואירועים של היום בלוח השנה,
מסכם עם Gemini (או IAC כגיבוי), ושולח מייל סיכום למשתמש.

רץ עצמאי (לא תלוי ב-Streamlit) דרך GitHub Actions, גם כשהמחשב כבוי.
"""

import base64
import json
import os
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText

import requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
]

GEMINI_MODEL = "gemini-2.5-flash"
IAC_BASE_URL = "https://server.iac.ac.il/api/v1/studentapi"

MY_EMAIL = os.environ.get("DIGEST_TO_EMAIL", "elankimi1@gmail.com")


def get_google_creds():
    token_json = os.environ["GMAIL_TOKEN_JSON"]
    creds = Credentials.from_authorized_user_info(json.loads(token_json), GOOGLE_SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    return creds


def fetch_recent_emails(creds, hours: int = 12) -> str:
    service = build("gmail", "v1", credentials=creds)
    after_ts = int((datetime.now(timezone.utc) - timedelta(hours=hours)).timestamp())
    results = service.users().messages().list(
        userId="me", q=f"after:{after_ts}", maxResults=20
    ).execute()
    messages = results.get("messages", [])
    if not messages:
        return "אין מיילים חדשים מהלילה."

    summaries = []
    for msg in messages:
        full = service.users().messages().get(
            userId="me", id=msg["id"], format="metadata",
            metadataHeaders=["From", "Subject"]
        ).execute()
        headers = {h["name"]: h["value"] for h in full["payload"]["headers"]}
        snippet = full.get("snippet", "")
        summaries.append(f"מאת: {headers.get('From', '?')} | נושא: {headers.get('Subject', '?')} | תקציר: {snippet}")
    return "\n".join(summaries)


def fetch_today_events(creds) -> str:
    service = build("calendar", "v3", credentials=creds)
    now = datetime.now(timezone.utc)
    end_of_day = now.replace(hour=23, minute=59, second=59)
    results = service.events().list(
        calendarId="primary", timeMin=now.isoformat(), timeMax=end_of_day.isoformat(),
        singleEvents=True, orderBy="startTime",
    ).execute()
    events = results.get("items", [])
    if not events:
        return "אין אירועים מתוכננים להיום."
    summaries = []
    for ev in events:
        start = ev["start"].get("dateTime", ev["start"].get("date"))
        summaries.append(f"{start} | {ev.get('summary', '(ללא כותרת)')}")
    return "\n".join(summaries)


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


def send_digest_email(creds, subject: str, body: str):
    service = build("gmail", "v1", credentials=creds)
    message = MIMEText(body)
    message["to"] = MY_EMAIL
    message["subject"] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()


def main():
    creds = get_google_creds()
    emails_text = fetch_recent_emails(creds)
    events_text = fetch_today_events(creds)

    try:
        summary = summarize_with_gemini(emails_text, events_text)
    except Exception:
        summary = summarize_with_iac(emails_text, events_text)

    today_str = datetime.now().strftime("%d/%m/%Y")
    send_digest_email(creds, f"דוח בוקר - {today_str}", summary)
    print("דוח הבוקר נשלח בהצלחה.")


if __name__ == "__main__":
    main()
