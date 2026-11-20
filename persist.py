"""Keeps replacement words / keywords across browser refreshes (server file) and, optionally,
across app reboots/redeploys (private GitHub Gist, via Streamlit secrets GITHUB_TOKEN [+ GIST_ID])."""
import json
import requests
import streamlit as st

API = "https://api.github.com/gists"


def _secret(name):
    try:
        return st.secrets.get(name) or ""
    except Exception:
        return ""


def _hdr():
    return {"Authorization": f"Bearer {_secret('GITHUB_TOKEN')}", "Accept": "application/vnd.github+json"}


def enabled():
    return bool(_secret("GITHUB_TOKEN"))


def gist_read():
    """-> {filename: text} or None"""
    gid = _secret("GIST_ID")
    if not (enabled() and gid):
        return None
    try:
        r = requests.get(f"{API}/{gid}", headers=_hdr(), timeout=15)
        r.raise_for_status()
        return {k: v.get("content", "") for k, v in r.json().get("files", {}).items()}
    except Exception as e:
        st.session_state["persist_err"] = f"Gist read failed: {e}"
        return None


def gist_write(files):
    """files = {filename: text}. Creates the gist if GIST_ID isn't set (id shown in the sidebar)."""
    if not enabled():
        return
    gid = st.session_state.get("new_gist_id") or _secret("GIST_ID")
    try:
        body = {"files": {k: {"content": v} for k, v in files.items()}}
        if gid:
            r = requests.patch(f"{API}/{gid}", headers=_hdr(), json=body, timeout=15)
        else:
            body.update(public=False, description="Inventory tool settings")
            r = requests.post(API, headers=_hdr(), json=body, timeout=15)
            if r.ok:
                st.session_state["new_gist_id"] = r.json()["id"]
        r.raise_for_status()
        st.session_state.pop("persist_err", None)
    except Exception as e:
        st.session_state["persist_err"] = f"Gist save failed: {e}"
