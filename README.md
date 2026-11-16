# Transcript to Inventory (Streamlit)

Turns raw inspection-dictation transcripts into a formatted list, then Word / Excel exports.
Two methods: Gemini (AI) or Local rules (no API). Two job types: Normal and Pearl.

## Deploy on Streamlit Community Cloud
1. Upload these files to a GitHub repo (keep it private).
2. share.streamlit.io -> New app -> pick the repo, main file `app.py`.
3. App -> Settings -> Secrets, paste:

```
GEMINI_API_KEY = "your key"
APP_PASSWORD = "choose a password"      # optional, recommended
GEMINI_FALLBACK_MODEL = "gemini-3.6-flash"   # optional
```
Local rules mode needs no key. Never commit your key to GitHub.

## Files
- `app.py` - the web page
- `transcript_tool.py`, `prompts.py`, `local_parser.py` - the engine, prompts and no-API rules
- `replacements.json`, `keywords.json` - default replacement words and Y/N keywords (edit in the sidebar, then download the lists and commit them to keep changes)
