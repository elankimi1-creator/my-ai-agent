# הסוכן שלי

סוכן AI עם ממשק Streamlit. כולל:
- שיחה עם כלים: קבצים, shell, Gmail, חיפוש אינטרנט
- יצירת קבצי Word, PowerPoint, Excel
- צירוף קבצים לשיחה
- מודל ראשי: Google Gemini, עם גיבוי אוטומטי ל-API חיצוני כשהמכסה מתרוקנת

## הרצה מקומית

```bash
pip install -r requirements.txt
streamlit run streamlit_agent.py
```

## הגדרת סודות (Secrets)

יש ליצור קובץ `.env` (לא מועלה ל-git) עם:

```
GEMINI_API_KEY=your_key_here
IAC_TOKEN=optional_backup_token
```

לתמיכה ב-Gmail, יש להוסיף גם `credentials.json` (OAuth client מ-Google Cloud Console).
