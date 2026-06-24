"""
סוכן AI אחד (Agent) עם ממשק Streamlit, מבוסס Google Gemini (ה-API האישי שלך, לא תלוי במכללה).
כולל: קבצים, shell, Gmail, חיפוש אינטרנט, ויצירת קבצי Word/PowerPoint.
כשהמכסה היומית של Gemini מתרוקנת, הסוכן עובר אוטומטית לגיבוי - ה-API של המכללה (IAC).
"""

import base64
import json
import os
import time
from email.mime.text import MIMEText
from pathlib import Path

import requests
import streamlit as st
from dotenv import load_dotenv
from docx import Document
from pptx import Presentation
from pptx.util import Inches
from openpyxl import Workbook
from google import genai
from google.genai import errors as genai_errors
from google.genai import types

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

load_dotenv()

GEMINI_MODEL = "gemini-2.5-flash"

IAC_BASE_URL = "https://server.iac.ac.il/api/v1/studentapi"
IAC_MAX_TOKENS = 10000

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]
CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "gmail_token.json"

# ========================================================
# Page configuration
# ========================================================

st.set_page_config(page_title="הסוכן שלי", page_icon="🤖", layout="centered")
st.title("🤖 הסוכן שלי")
st.caption("סוכן אחד עם כלים: קבצים, Word/PowerPoint, Gmail וחיפוש אינטרנט")

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if "created_files" not in st.session_state:
    st.session_state.created_files = []

if "uploaded_paths" not in st.session_state:
    st.session_state.uploaded_paths = []

UPLOADS_DIR = Path("uploads")
UPLOADS_DIR.mkdir(exist_ok=True)


# ========================================================
# טוקנים
# ========================================================

def load_gemini_key() -> str:
    return os.environ.get("GEMINI_API_KEY")


def load_iac_token() -> str:
    env_path = Path(".env")
    if not env_path.exists():
        return None
    for line in env_path.read_text(encoding="utf-8").splitlines():
        if line.startswith("IAC_TOKEN="):
            val = line.split("=", 1)[1].strip().strip("'\"")
            if val and val != "sk-std-YOUR-KEY":
                return val
    return None


def is_gemini_quota_error(e: Exception) -> bool:
    return isinstance(e, genai_errors.ClientError) and "RESOURCE_EXHAUSTED" in str(e)


# ========================================================
# כלים: קבצים, shell, Gmail
# ========================================================

def get_gmail_service():
    creds = None
    if Path(TOKEN_FILE).exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, GMAIL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        Path(TOKEN_FILE).write_text(creds.to_json(), encoding="utf-8")
    return build("gmail", "v1", credentials=creds)


def tool_read_file(path: str) -> str:
    try:
        return Path(path).read_text(encoding="utf-8")
    except Exception as e:
        return f"שגיאה בקריאת הקובץ: {e}"


def tool_write_file(path: str, content: str) -> str:
    try:
        Path(path).write_text(content, encoding="utf-8")
        return f"הקובץ {path} נשמר בהצלחה."
    except Exception as e:
        return f"שגיאה בכתיבת הקובץ: {e}"


def tool_run_shell(command: str) -> str:
    import subprocess
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=60)
        output = result.stdout + result.stderr
        return output.strip() or "(הפקודה רצה בהצלחה, אין פלט)"
    except Exception as e:
        return f"שגיאה בהרצת הפקודה: {e}"


def tool_send_email(to: str, subject: str, body: str) -> str:
    try:
        service = get_gmail_service()
        message = MIMEText(body)
        message["to"] = to
        message["subject"] = subject
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        return f"המייל נשלח בהצלחה ל-{to}."
    except Exception as e:
        return f"שגיאה בשליחת המייל: {e}"


def tool_list_recent_emails(max_results: int = 5) -> str:
    try:
        service = get_gmail_service()
        results = service.users().messages().list(userId="me", maxResults=max_results).execute()
        messages = results.get("messages", [])
        if not messages:
            return "אין מיילים בתיבה."
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
    except Exception as e:
        return f"שגיאה בקריאת המיילים: {e}"


# ========================================================
# כלים: Word / PowerPoint
# ========================================================

def _register_created_file(path: str):
    if path not in st.session_state.created_files:
        st.session_state.created_files.append(path)


def tool_create_word_doc(path: str, title: str, paragraphs: list) -> str:
    try:
        if not path.endswith(".docx"):
            path += ".docx"
        doc = Document()
        if title:
            doc.add_heading(title, level=1)
        for p in paragraphs:
            doc.add_paragraph(p)
        doc.save(path)
        _register_created_file(path)
        return f"קובץ Word נוצר בהצלחה: {path}"
    except Exception as e:
        return f"שגיאה ביצירת קובץ Word: {e}"


def tool_create_pptx(path: str, slides: list) -> str:
    try:
        if not path.endswith(".pptx"):
            path += ".pptx"
        prs = Presentation()
        title_layout = prs.slide_layouts[1]
        for slide_data in slides:
            slide = prs.slides.add_slide(title_layout)
            slide.shapes.title.text = slide_data.get("title", "")
            body = slide.placeholders[1].text_frame
            content = slide_data.get("content", "")
            lines = content.split("\n") if isinstance(content, str) else content
            for i, line in enumerate(lines):
                if i == 0:
                    body.text = line
                else:
                    body.add_paragraph().text = line
        prs.save(path)
        _register_created_file(path)
        return f"מצגת PowerPoint נוצרה בהצלחה: {path}"
    except Exception as e:
        return f"שגיאה ביצירת מצגת: {e}"


def tool_create_excel(path: str, sheet_name: str, headers: list, rows: list) -> str:
    try:
        if not path.endswith(".xlsx"):
            path += ".xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name or "Sheet1"
        if headers:
            ws.append(headers)
        for row in rows:
            ws.append(row)
        wb.save(path)
        _register_created_file(path)
        return f"קובץ Excel נוצר בהצלחה: {path}"
    except Exception as e:
        return f"שגיאה ביצירת קובץ Excel: {e}"


# ========================================================
# כלי: חיפוש אינטרנט (דרך ה-Agent של IAC ברקע)
# ========================================================

def tool_web_search(query: str) -> str:
    token = load_iac_token()
    if not token:
        return "אין גישה לחיפוש אינטרנט כרגע (אין טוקן גיבוי)."
    try:
        headers = {"Authorization": f"Bearer {token}"}
        payload = {
            "input": query,
            "instructions": "Search the web and answer concisely in Hebrew.",
            "tools": [{"type": "web_search"}],
            "reasoning": {"effort": "low"},
            "max_output_tokens": IAC_MAX_TOKENS,
        }
        r = requests.post(f"{IAC_BASE_URL}/responses", headers=headers, json=payload, timeout=90)
        r.raise_for_status()
        data = r.json()
        if data.get("output_text"):
            return data["output_text"]
        texts = []
        for item in data.get("output", []):
            if not isinstance(item, dict) or item.get("type") == "reasoning":
                continue
            for part in item.get("content") or []:
                text = part.get("text") or part.get("output_text")
                if text:
                    texts.append(text)
        return "\n".join(texts) if texts else "לא נמצאו תוצאות."
    except Exception as e:
        return f"שגיאה בחיפוש אינטרנט: {e}"


def call_tool(name: str, args: dict) -> str:
    func = TOOL_FUNCTIONS.get(name)
    if not func:
        return f"כלי לא מוכר: {name}"
    try:
        return func(**args)
    except TypeError:
        if set(args.keys()) - {"parameters", "arguments", "args"} == set() and len(args) == 1:
            inner = next(iter(args.values()))
            if isinstance(inner, dict):
                return func(**inner)
        raise


TOOL_FUNCTIONS = {
    "read_file": tool_read_file,
    "write_file": tool_write_file,
    "run_shell": tool_run_shell,
    "send_email": tool_send_email,
    "list_recent_emails": tool_list_recent_emails,
    "create_word_doc": tool_create_word_doc,
    "create_pptx": tool_create_pptx,
    "create_excel": tool_create_excel,
    "web_search": tool_web_search,
}

TOOLS_SCHEMA = [
    {
        "name": "read_file",
        "description": "קורא את תוכן הקובץ הנתון ומחזיר אותו כטקסט.",
        "parameters": {
            "type": "object",
            "properties": {"path": {"type": "string", "description": "נתיב לקובץ לקריאה"}},
            "required": ["path"],
        },
    },
    {
        "name": "write_file",
        "description": "כותב תוכן לקובץ נתון (יוצר אותו אם לא קיים, או דורס אותו אם קיים).",
        "parameters": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "נתיב לקובץ לכתיבה"},
                "content": {"type": "string", "description": "התוכן לכתוב לקובץ"},
            },
            "required": ["path", "content"],
        },
    },
    {
        "name": "run_shell",
        "description": "מריץ פקודת shell ומחזיר את הפלט שלה (stdout/stderr).",
        "parameters": {
            "type": "object",
            "properties": {"command": {"type": "string", "description": "פקודת shell להרצה"}},
            "required": ["command"],
        },
    },
    {
        "name": "send_email",
        "description": "שולח מייל מתיבת הדואר של המשתמש (Gmail) לכתובת נתונה.",
        "parameters": {
            "type": "object",
            "properties": {
                "to": {"type": "string", "description": "כתובת המייל של הנמען"},
                "subject": {"type": "string", "description": "נושא המייל"},
                "body": {"type": "string", "description": "תוכן המייל"},
            },
            "required": ["to", "subject", "body"],
        },
    },
    {
        "name": "list_recent_emails",
        "description": "מחזיר רשימה של המיילים האחרונים בתיבת הדואר (נושא, שולח, תקציר).",
        "parameters": {
            "type": "object",
            "properties": {
                "max_results": {"type": "integer", "description": "כמה מיילים להחזיר (ברירת מחדל 5)"}
            },
            "required": [],
        },
    },
    {
        "name": "create_word_doc",
        "description": "יוצר קובץ Word (.docx) עם כותרת ופסקאות טקסט.",
        "parameters": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "נתיב/שם הקובץ לשמירה, למשל 'document.docx'"},
                "title": {"type": "string", "description": "כותרת המסמך"},
                "paragraphs": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "רשימת פסקאות טקסט להוספה למסמך",
                },
            },
            "required": ["path", "paragraphs"],
        },
    },
    {
        "name": "create_pptx",
        "description": "יוצר מצגת PowerPoint (.pptx) עם רשימת שקופיות, כל שקופית עם כותרת ותוכן.",
        "parameters": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "נתיב/שם הקובץ לשמירה, למשל 'presentation.pptx'"},
                "slides": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "title": {"type": "string", "description": "כותרת השקופית"},
                            "content": {"type": "string", "description": "תוכן השקופית (אפשר כמה שורות מופרדות ב-\\n)"},
                        },
                    },
                    "description": "רשימת שקופיות, כל אחת עם title ו-content",
                },
            },
            "required": ["path", "slides"],
        },
    },
    {
        "name": "create_excel",
        "description": "יוצר קובץ Excel (.xlsx) עם כותרות עמודות ושורות נתונים.",
        "parameters": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "נתיב/שם הקובץ לשמירה, למשל 'data.xlsx'"},
                "sheet_name": {"type": "string", "description": "שם הגיליון"},
                "headers": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "כותרות העמודות (שורה ראשונה)",
                },
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {}},
                    "description": "רשימת שורות, כל שורה היא רשימת ערכים",
                },
            },
            "required": ["path", "rows"],
        },
    },
    {
        "name": "web_search",
        "description": "מחפש מידע עדכני באינטרנט ומחזיר תקציר תשובה. שימושי לשאלות על אירועים עדכניים, מחירים, חדשות וכו'.",
        "parameters": {
            "type": "object",
            "properties": {"query": {"type": "string", "description": "שאילתת החיפוש"}},
            "required": ["query"],
        },
    },
]

MAX_TOOL_ROUNDS = 10

SYSTEM_INSTRUCTION = (
    "אתה סוכן AI פעיל עם גישה אמיתית לכלים הבאים, ואתה חייב להשתמש בהם בפועל ולא רק לתאר מה "
    "אפשר לעשות:\n"
    "- send_email: שולח מייל אמיתי מתיבת ה-Gmail המחוברת של המשתמש. כשמתבקש לשלוח מייל, "
    "תקרא לכלי הזה ממש - אל תכתוב שאינך יכול לשלוח מיילים.\n"
    "- list_recent_emails: קורא מיילים אמיתיים מהתיבה.\n"
    "- create_word_doc, create_pptx, create_excel: יוצרים קבצים אמיתיים בדיסק שהמשתמש יכול להוריד.\n"
    "- web_search: מחפש מידע עדכני באינטרנט באמת.\n"
    "- read_file, write_file, run_shell: גישה אמיתית למערכת הקבצים ולמסוף.\n"
    "לעולם אל תגיד שאינך מסוגל לבצע פעולה שיש לך כלי בשבילה - תמיד תקרא לכלי המתאים ותבצע "
    "את הבקשה בפועל."
)


# ========================================================
# Gemini
# ========================================================

def get_gemini_client():
    api_key = load_gemini_key()
    if not api_key:
        raise RuntimeError("לא נמצא GEMINI_API_KEY בקובץ .env")
    return genai.Client(api_key=api_key)


def generate_with_retry(client, **kwargs):
    for attempt in range(3):
        try:
            return client.models.generate_content(**kwargs)
        except genai_errors.ServerError:
            time.sleep(3 * (attempt + 1))
    raise RuntimeError("Gemini לא הגיב (עומס שרת חזק).")


def call_agent_gemini(history: list) -> str:
    client = get_gemini_client()
    tool_config = types.Tool(function_declarations=[
        types.FunctionDeclaration(name=t["name"], description=t["description"], parameters=t["parameters"])
        for t in TOOLS_SCHEMA
    ])
    contents = [
        types.Content(role="model" if m["role"] == "assistant" else "user", parts=[types.Part(text=m["content"])])
        for m in history if m["role"] in ("user", "assistant") and m.get("content")
    ]

    for _ in range(MAX_TOOL_ROUNDS):
        response = generate_with_retry(
            client, model=GEMINI_MODEL, contents=contents,
            config=types.GenerateContentConfig(
                tools=[tool_config], system_instruction=SYSTEM_INSTRUCTION
            ),
        )
        candidate = response.candidates[0]
        contents.append(candidate.content)

        function_calls = [p.function_call for p in candidate.content.parts if p.function_call]
        if not function_calls:
            return "".join(p.text for p in candidate.content.parts if p.text)

        response_parts = []
        for call in function_calls:
            st.toast(f"🔧 קורא לכלי: {call.name}")
            result = call_tool(call.name, dict(call.args))
            response_parts.append(types.Part.from_function_response(name=call.name, response={"result": result}))
        contents.append(types.Content(role="user", parts=response_parts))

    return "הגעתי למספר המקסימלי של שלבים מבלי לסיים את המשימה."


# ========================================================
# IAC (גיבוי)
# ========================================================

def _iac_tools_schema() -> list:
    return [{"type": "function", "function": t} for t in TOOLS_SCHEMA]


def call_agent_iac(history: list) -> str:
    token = load_iac_token()
    if not token:
        raise RuntimeError("אין גם טוקן IAC לגיבוי.")
    headers = {"Authorization": f"Bearer {token}"}
    messages = [{"role": "system", "content": SYSTEM_INSTRUCTION}] + [
        {"role": m["role"], "content": m["content"]} for m in history if m.get("content")
    ]

    for _ in range(MAX_TOOL_ROUNDS):
        payload = {"messages": messages, "tools": _iac_tools_schema(), "max_completion_tokens": IAC_MAX_TOKENS}

        for attempt in range(3):
            try:
                r = requests.post(f"{IAC_BASE_URL}/chat/completions", headers=headers, json=payload, timeout=120)
                break
            except requests.exceptions.ReadTimeout:
                if attempt == 2:
                    raise RuntimeError("השרת של המכללה לא הגיב בזמן (timeout חזרתי). נסה שוב בעוד דקה.")
                time.sleep(3)

        r.raise_for_status()
        data = r.json()
        if "choices" not in data:
            raise RuntimeError(data.get("details") or data.get("error") or str(data))
        message = data["choices"][0]["message"]
        tool_calls = message.get("tool_calls")
        if not tool_calls:
            return message.get("content") or str(data)
        messages.append(message)
        for call in tool_calls:
            name = call["function"]["name"]
            args = json.loads(call["function"]["arguments"] or "{}")
            st.toast(f"🔧 קורא לכלי (גיבוי): {name}")
            result = call_tool(name, args)
            messages.append({"role": "tool", "tool_call_id": call["id"], "content": result})

    return "הגעתי למספר המקסימלי של שלבים מבלי לסיים את המשימה."


def call_agent(history: list) -> str:
    try:
        return call_agent_gemini(history)
    except Exception as e:
        if is_gemini_quota_error(e) and load_iac_token():
            st.toast("⚠️ מכסת Gemini הסתיימה - עובר לגיבוי")
            return call_agent_iac(history)
        raise


# ========================================================
# Sidebar
# ========================================================

with st.sidebar:
    st.header("הגדרות")
    st.caption("סוכן אחד עם כל הכלים: קבצים, Word/PowerPoint/Excel, Gmail, חיפוש אינטרנט")
    if st.button("נקה שיחה"):
        st.session_state.chat_history = []
        st.session_state.created_files = []
        st.session_state.uploaded_paths = []
        st.rerun()

    st.divider()
    st.subheader("📎 צרף קובץ")
    uploaded_file = st.file_uploader(
        "תמונה, סרטון או מסמך", type=None, key="file_upload_widget"
    )
    if uploaded_file is not None:
        save_path = UPLOADS_DIR / uploaded_file.name
        save_path.write_bytes(uploaded_file.getbuffer())
        path_str = str(save_path)
        if path_str not in st.session_state.uploaded_paths:
            st.session_state.uploaded_paths.append(path_str)
            st.success(f"הקובץ הועלה: {uploaded_file.name}")

    if st.session_state.uploaded_paths:
        st.caption("קבצים מצורפים לשיחה:")
        for p in st.session_state.uploaded_paths:
            st.text(Path(p).name)
        if st.button("נקה קבצים מצורפים"):
            st.session_state.uploaded_paths = []
            st.rerun()

    if st.session_state.created_files:
        st.divider()
        st.subheader("⬇️ קבצים שנוצרו")
        for fpath in st.session_state.created_files:
            p = Path(fpath)
            if p.exists():
                st.download_button(
                    f"הורד {p.name}", data=p.read_bytes(), file_name=p.name, key=f"dl_{fpath}"
                )


# ========================================================
# Token check
# ========================================================

if not load_gemini_key():
    st.error("לא נמצא מפתח תקין בקובץ .env")
    st.info("יש להוסיף לקובץ .env שורה כזו:")
    st.code("GEMINI_API_KEY=your_key_here")
    st.stop()


# ========================================================
# Display history
# ========================================================

if not st.session_state.chat_history:
    with st.chat_message("assistant"):
        st.markdown("היי! אני יכול לעזור עם קבצים, Word/PowerPoint, מיילים וחיפוש באינטרנט 🙂")

for msg in st.session_state.chat_history:
    if msg["role"] not in ("user", "assistant"):
        continue
    content = msg.get("content")
    if not content:
        continue
    with st.chat_message(msg["role"]):
        st.markdown(content)


# ========================================================
# Chat input
# ========================================================

user_prompt = st.chat_input("כתוב הודעה...")

if user_prompt:
    full_prompt = user_prompt
    if st.session_state.uploaded_paths:
        attachments = ", ".join(st.session_state.uploaded_paths)
        full_prompt += f"\n\n(קבצים מצורפים שזמינים בדיסק לקריאה/עיבוד: {attachments})"

    st.session_state.chat_history.append({"role": "user", "content": full_prompt})

    with st.chat_message("user"):
        st.markdown(user_prompt)
        if st.session_state.uploaded_paths:
            st.caption("📎 " + ", ".join(Path(p).name for p in st.session_state.uploaded_paths))

    with st.chat_message("assistant"):
        response_placeholder = st.empty()
        response_placeholder.markdown("חושב...")

        try:
            assistant_reply = call_agent(st.session_state.chat_history)
            response_placeholder.markdown(assistant_reply)
        except Exception as e:
            assistant_reply = f"שגיאה: {str(e)}"
            response_placeholder.error(assistant_reply)

    st.session_state.chat_history.append({"role": "assistant", "content": assistant_reply})
