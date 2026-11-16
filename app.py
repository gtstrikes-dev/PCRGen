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
with st.sidebar:
    method = st.radio("Method", ["Gemini (AI)", "Local rules (no API)"])
    job = st.radio("Job type", ["Normal", "Pearl"])
    model = st.text_input("Model", value=secret("GEMINI_MODEL", "gemini-3.1-flash-lite-preview"))
    ref = st.file_uploader("Pearl Property Ref (optional)", type=["txt", "md", "docx", "pdf"]) if job == "Pearl" else None
    reorder = st.checkbox("Reorder items to the reference template (Pearl: PDF, Normal: blank report)", value=True)
    layout_label = st.radio("Layout", ["Standard", "Final report"])
    if method.startswith("Gemini") and st.button("Test key / model"):
        ok, msg = t.test_key(key, model.strip()) if key else (False, "GEMINI_API_KEY is not set. Add it under the app's Secrets.")
        (st.success if ok else st.error)(msg)

# ---- separate replacement-word lists for Normal and Pearl (kept for this session; download to keep a copy)
if "rules" not in st.session_state:
    st.session_state["rules"] = t.load_rules()  # starts from replacements.json in the repo if present
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
    up = st.file_uploader("Load lists (.json) or word_replacements.txt", type=["json", "txt"])
    if up is not None and st.session_state.get("rules_loaded") != up.file_id:
        try:
            if up.name.lower().endswith(".txt"):  # "find => replacement" lines, merged into BOTH lists
                new = t.parse_rules_txt(up.getvalue().decode("utf-8-sig", "ignore"))
                for k in ("normal", "pearl"):
                    t.merge_rules(st.session_state["rules"][k], new)
            else:
                loaded = json.loads(up.getvalue())
                for k in ("normal", "pearl"):
                    if isinstance(loaded.get(k), list):
                        st.session_state["rules"][k] = [[str(a), str(b)] for a, b in loaded[k] if str(a).strip()]
            st.session_state["rules_loaded"] = up.file_id
            st.rerun()
        except Exception as e:
            st.error(f"Could not read that file: {e}")
if "kws" not in st.session_state:
    st.session_state["kws"] = t.load_keywords()  # starts from keywords.json in the repo if present
with st.sidebar.expander("Clean / Undamaged / Working keywords"):
    kj = st.radio("Prompt", ["Normal", "Pearl"], horizontal=True, key="kw_job").lower()
    kc = st.radio("Column", ["Clean", "Undamaged", "Working"], horizontal=True, key="kw_col").lower()
    kdf = pd.DataFrame({"Keyword": st.session_state["kws"][kj][kc]})
    ked = st.data_editor(kdf, num_rows="dynamic", width="stretch", hide_index=True, key=f"kw_{kj}_{kc}")
    st.session_state["kws"][kj][kc] = [str(x).strip() for x in ked["Keyword"] if x is not None and x == x and str(x).strip()]
    st.caption("If ANY keyword appears in a description, that column becomes N, otherwise Y. Matches letters inside text, ignoring case.")
    st.download_button("Download keywords (.json)", json.dumps(st.session_state["kws"], indent=1), "keywords.json", "application/json")
    kup = st.file_uploader("Load keywords (.json)", type="json", key="kw_up")
    if kup is not None and st.session_state.get("kw_loaded") != kup.file_id:
        try:
            loaded = json.loads(kup.getvalue())
            for j in ("normal", "pearl"):
                for c in t.KW_COLS:
                    v = loaded.get(j, {}).get(c)
                    if isinstance(v, list):
                        st.session_state["kws"][j][c] = [str(x).strip() for x in v if str(x).strip()]
            st.session_state["kw_loaded"] = kup.file_id
            st.rerun()
        except Exception as e:
            st.error(f"Could not read that file: {e}")
file_name = st.sidebar.text_input("Download file name", value="Inventory")

tf = st.file_uploader("Or load a transcript file (.txt, .docx, .xlsx, .pdf, .vtt, .srt, .csv)", type=["txt", "md", "vtt", "srt", "csv", "docx", "xlsx", "xlsm", "pdf"], key="tr_up")
if tf is not None and st.session_state.get("_tr_loaded") != (tf.name, tf.size):
    try:
        st.session_state["raw"] = t.read_transcript(tf.name, tf.getvalue())
        st.session_state["_tr_loaded"] = (tf.name, tf.size)
    except Exception as e:
        st.error(f"Could not read that file: {e}")
raw = st.text_area("Raw transcript", height=300, key="raw")

if st.button("1. Process transcript", type="primary"):
    if not raw.strip():
        st.warning("Paste a transcript first.")
    elif method.startswith("Local"):
        lst, review = t.process_local(raw, job.lower(), reorder)
        st.session_state["list"] = lst
        st.session_state["job_used"] = job.lower()
        st.session_state["check"] = t.compare(raw, lst)
        st.session_state["raw_used"] = raw
        st.session_state["review"] = review
        st.session_state["local"] = True
    elif not key:
        st.error("GEMINI_API_KEY is not set. Add it under the app's Secrets, or switch Method to Local rules.")
    else:
        sp = prompts.NORMAL if job == "Normal" else prompts.PEARL
        if ref is not None:
            reftxt = t.read_reference(ref.name, ref.getvalue())
            sp += "\n\nPEARL PROPERTY REF (formatting and terminology guide ONLY; never copy facts from it):\n" + reftxt
        sp += prompts.FIDELITY
        try:
            with st.spinner("Processing with Gemini..."):
                ph = st.empty()
                text, warn = t.call_gemini(key, model.strip(), sp, raw, secret("GEMINI_FALLBACK_MODEL", "gemini-3.6-flash"),
                                           lambda m: ph.info(m))
                ph.empty()
            lst = t.strip_fence(text)
            if reorder:
                lst = t.reorder_for(job.lower(), lst)
            st.session_state["list"] = lst
            st.session_state["job_used"] = job.lower()
            st.session_state["check"] = t.compare(raw, lst)
            st.session_state["raw_used"] = raw
            st.session_state["review"] = ""
            st.session_state["local"] = False
            if warn:
                st.warning(warn)
        except Exception as e:
            st.error(str(e))

if "list" in st.session_state:
    st.subheader("2. Formatted list (editable)")
    lst = st.text_area("List", st.session_state["list"], height=400, key="list_edit", label_visibility="collapsed")
    om, ad, no, na = st.session_state["check"]
    po, pa, flag = t.fidelity(st.session_state.get("raw_used", ""), no, na)
    if st.session_state.get("local"):
        st.info("Local rules (no AI): nothing is invented. Words 'missing' below are mostly spoken headings that became item names, "
                "repeated item names, fillers and corrected words.")
        if st.session_state.get("review"):
            with st.expander("Things to check (corrections, go-backs, unresolved lines)", expanded=True):
                st.text(st.session_state["review"])
    elif flag:
        st.warning(f"Check this output: about {po:.1f}% of your transcript words are missing and {pa:.1f}% were added. "
                   "The model may have changed your transcript. Review the check below or try a stronger model.")
    elif st.session_state.get("raw_used") and not st.session_state.get("local"):
        st.caption(f"Looks faithful: {po:.1f}% of words missing, {pa:.1f}% added.")
    with st.expander(f"Omitted / added check ({no} possibly omitted, {na} possibly added)"):
        st.write(f"**Possibly omitted** (fillers, tags and self-corrections are expected): {om or 'none'}")
        st.write(f"**Possibly added** (AU spelling and brand fixes are expected): {ad or 'none'}")

    st.subheader("3. Export")
    layout = "standard" if layout_label == "Standard" else "final"
    data = t.run_macros(t.parse_list(lst), st.session_state["rules"][st.session_state.get("job_used", "normal")],
                        st.session_state["kws"][st.session_state.get("job_used", "normal")])
    base = t.safe_name(file_name)
    wb, xb = io.BytesIO(), io.BytesIO()
    t.write_word(data, layout, wb)
    t.write_excel(data, layout, xb)
    c1, c2 = st.columns(2)
    c1.download_button("Download Word", wb.getvalue(), base + ".docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    c2.download_button("Download Excel", xb.getvalue(), base + ".xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
