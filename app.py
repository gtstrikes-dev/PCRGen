"""Streamlit web version (single user). Secrets: GEMINI_API_KEY, optional APP_PASSWORD."""
import io
import json
import os

import streamlit as st

import prompts
import transcript_tool as t

st.set_page_config(page_title="Transcript to Inventory", layout="wide")


def secret(name, default=""):
    try:
        return st.secrets[name]
    except Exception:
        return os.environ.get(name, default)


# Optional second lock, in case the app is ever made public by mistake.
pw = secret("APP_PASSWORD")
if pw and st.session_state.get("ok") is not True:
    st.title("Transcript to Inventory")
    if st.text_input("Password", type="password") == pw:
        st.session_state["ok"] = True
        st.rerun()
    st.stop()

key = secret("GEMINI_API_KEY")
st.title("Transcript to Inventory")
if not key:
    st.error("GEMINI_API_KEY is not set. Add it under the app's Secrets.")
    st.stop()

with st.sidebar:
    job = st.radio("Job type", ["Normal", "Pearl"])
    model = st.text_input("Model", value=secret("GEMINI_MODEL", "gemini-3.1-flash-lite-preview"))
    ref = st.file_uploader("Pearl Property Ref (optional)", type=["txt", "md", "docx", "pdf"]) if job == "Pearl" else None
    reorder = st.checkbox("Reorder items to Pearl template", value=True) if job == "Pearl" else False
    layout_label = st.radio("Layout", ["Standard", "Final report"])
    if st.button("Test key / model"):
        ok, msg = t.test_key(key, model.strip())
        (st.success if ok else st.error)(msg)

# ---- separate replacement-word lists for Normal and Pearl (kept for this session; download to keep a copy)
if "rules" not in st.session_state:
    st.session_state["rules"] = t.default_rules()
with st.sidebar.expander("Replacement words"):
    which = st.radio("Edit list for", ["Normal", "Pearl"], horizontal=True, key="rules_which").lower()
    import pandas as pd
    df = pd.DataFrame(st.session_state["rules"][which], columns=["Find", "Replace"])
    edited = st.data_editor(df, num_rows="dynamic", width="stretch", hide_index=True, key=f"ed_{which}")
    st.session_state["rules"][which] = [[str(a), "" if b is None or b != b else str(b)]
                                        for a, b in edited.itertuples(index=False)
                                        if a is not None and a == a and str(a).strip()]
    st.caption('Whole word, ignores case. Spaces matter: " is " removes the word is. Leave Replace empty to delete a word.')
    st.download_button("Download both lists (.json)", json.dumps(st.session_state["rules"], indent=1),
                       "replacements.json", "application/json")
    up = st.file_uploader("Load lists (.json)", type="json")
    if up is not None and st.session_state.get("rules_loaded") != up.file_id:
        try:
            loaded = json.loads(up.getvalue())
            for k in ("normal", "pearl"):
                if isinstance(loaded.get(k), list):
                    st.session_state["rules"][k] = [[str(a), str(b)] for a, b in loaded[k] if str(a).strip()]
            st.session_state["rules_loaded"] = up.file_id
            st.rerun()
        except Exception as e:
            st.error(f"Could not read that file: {e}")
file_name = st.sidebar.text_input("Download file name", value="Inventory")

raw = st.text_area("Raw transcript", height=300, key="raw")

if st.button("1. Process transcript", type="primary"):
    if not raw.strip():
        st.warning("Paste a transcript first.")
    else:
        sp = prompts.NORMAL if job == "Normal" else prompts.PEARL
        if ref is not None:
            reftxt = t.read_reference(ref.name, ref.getvalue())
            sp += "\n\nPEARL PROPERTY REF (formatting and terminology guide ONLY; never copy facts from it):\n" + reftxt
        sp += prompts.FIDELITY
        try:
            with st.spinner("Processing with Gemini..."):
                text, warn = t.call_gemini(key, model.strip(), sp, raw)
            lst = t.strip_fence(text)
            if reorder:
                lst = t.reorder_pearl(lst)
            st.session_state["list"] = lst
            st.session_state["job_used"] = job.lower()
            st.session_state["check"] = t.compare(raw, lst)
            if warn:
                st.warning(warn)
        except Exception as e:
            st.error(str(e))

if "list" in st.session_state:
    st.subheader("2. Formatted list (editable)")
    lst = st.text_area("List", st.session_state["list"], height=400, key="list_edit", label_visibility="collapsed")
    om, ad, no, na = st.session_state["check"]
    with st.expander(f"Omitted / added check ({no} possibly omitted, {na} possibly added)"):
        st.write(f"**Possibly omitted** (fillers, tags and self-corrections are expected): {om or 'none'}")
        st.write(f"**Possibly added** (AU spelling and brand fixes are expected): {ad or 'none'}")

    st.subheader("3. Export")
    layout = "standard" if layout_label == "Standard" else "final"
    data = t.run_macros(t.parse_list(lst), st.session_state["rules"][st.session_state.get("job_used", "normal")])
    base = t.safe_name(file_name)
    wb, xb = io.BytesIO(), io.BytesIO()
    t.write_word(data, layout, wb)
    t.write_excel(data, layout, xb)
    c1, c2 = st.columns(2)
    c1.download_button("Download Word", wb.getvalue(), base + ".docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    c2.download_button("Download Excel", xb.getvalue(), base + ".xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
