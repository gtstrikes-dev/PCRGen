"""
Transcript -> Inventory tool (local)

Setup:   pip install python-docx openpyxl requests
Run:     python transcript_tool.py
API key: set env var GEMINI_API_KEY, or paste it in the app (optionally saved to config.json beside this file).

Pipeline: raw transcript -> Gemini (Normal / Pearl prompt) -> list review ->
          FormatText -> FormatTable -> ChangeYtoN (ported from your .bas macros) -> Word / Excel
"""
import io
import json
import os
import re
import threading
from collections import Counter
from functools import lru_cache
from pathlib import Path

import requests
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

import prompts

APP_DIR = Path(__file__).parent
CONFIG = APP_DIR / "config.json"
DEFAULT_MODEL = "gemini-3.8-flash"
API_BASE = "https://generativelanguage.googleapis.com/v1beta"

# ------------------------------------------------------------------ macro data (from your .bas files)
# FormatText replacements: whole-word, case-insensitive, applied in this order after commas -> full stops.
TEXT_REPLACEMENTS = list(zip(
    ["points and switches", "slash", "windows and frames", "windowsill", "metre", "window covering",
     "window treatment", "benchtop", "benchtops", "colourbond", "colour bond", "intact in working order",
     "intact and working order", "down light", "glove", "gloves", "tower", " is ", " are ", " a ", " there ",
     "watermarks", "watermarked", "the "],
    ["Points & Switches", "/", "Windows & Frames", "window sill", "meter", "Window Covering",
     "Window Treatment", "bench top", "bench top", "Colorbond", "Colorbond", "intact and in working order",
     "intact and in working order", "downlight", "globe", "globes", "towel", " ", " ", " ", " ",
     "water marks", "water marked", ""],
))
CLEAN_KW = ("unclean,grubby,cobweb,cobwebs,dirty,dirt,dusty,dust,residue,spot marks,black marks,dark marks,debris,"
            "weeds,droppings,faeces,mould,smudge,smudging,smudged,white marks,white mark,spot mark,black mark,"
            "drip mark,drip marks,light marks,washable marks,wash marks,wash mark,bug debris,weed,soap scum,"
            "brown mark,brown marks,grease,greasy,minor marks,build up,water mark,water marks,small marks,"
            "darker markings,tyre mark,tyre marks,marked,markings,food remnants,ring mark,finger mark,"
            "non-washable marking,cleaning mark").split(",")
DAMAGE_KW = ("chip,chipped,chipping,chips,crack,cracked,cracks,cracking,scratch,scratched,scratching,scratches,"
             "rust,rusting,rusted,rusty,damaged,dent,dents,dented,bent,broken,furniture indents,indentation,"
             "dropped,loose,pull marks,pulls,rips to,ripped,lifting,gouge,coming away,peeling,peel,calcium,"
             "one tear,a tear,tears,furniture imprints,poorly sealed,fraying,bending,sagging,gaping,gap,gaps,"
             "corrosion,bubbling,deteriorating,deteriorated,deteriorate,dry patches,leaking,torn,one rip,a rip,"
             "water damage,tearing,denting,frayed,detached,split").split(",")
WORKING_KW = ["globe not working", "not working"]

def read_reference(name, data):
    """Text of a reference file (.txt/.md/.docx/.pdf) given its name and bytes."""
    n = name.lower()
    if n.endswith(".docx"):
        d = Document(io.BytesIO(data))
        parts = [x.text for x in d.paragraphs]
        for tb in d.tables:
            for row in tb.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    if n.endswith(".pdf"):
        from pypdf import PdfReader
        return "\n".join((pg.extract_text() or "") for pg in PdfReader(io.BytesIO(data)).pages)
    return data.decode("utf-8", "ignore")


TRANSCRIPT_TYPES = "*.txt *.md *.vtt *.srt *.csv *.docx *.xlsx *.xlsm *.pdf"


def read_transcript(name, data):
    """Raw transcript text from a loaded file: txt/md/vtt/srt/csv, Word (.docx), Excel (.xlsx/.xlsm) or PDF."""
    n = name.lower()
    if n.endswith((".docx", ".pdf")):
        return read_reference(name, data).replace(" | ", " ")
    if n.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
                if cells:
                    lines.append(" ".join(cells))
        return "\n".join(lines)
    if n.endswith((".doc", ".xls")):
        raise ValueError("Old .doc/.xls files are not supported. Save it as .docx / .xlsx first (File > Save As).")
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return data.decode("utf-8", "ignore")


# ------------------------------------------------------------------ Gemini
def get_key(entry_value=""):
    return (entry_value or os.environ.get("GEMINI_API_KEY", "")).strip()


def test_key(key, model):
    """Returns (ok, message). Uses a cheap models.get call."""
    try:
        r = requests.get(f"{API_BASE}/models/{model}", headers={"x-goog-api-key": key}, timeout=20)
    except requests.RequestException as e:
        return False, f"Could not reach Google: {e}"
    if r.status_code == 200:
        return True, f"Key is valid and model '{model}' is available."
    try:
        msg = r.json()["error"]["message"]
    except Exception:
        msg = r.text[:200]
    return False, f"HTTP {r.status_code}: {msg}"


def list_models(key):
    """-> sorted model ids that support generateContent."""
    r = requests.get(f"{API_BASE}/models", params={"pageSize": 200}, headers={"x-goog-api-key": key}, timeout=30)
    r.raise_for_status()
    return sorted(m["name"].split("/", 1)[1] for m in r.json().get("models", [])
                  if "generateContent" in m.get("supportedGenerationMethods", []))


class GeminiError(RuntimeError):
    def __init__(self, status, msg):
        super().__init__(f"Gemini error {status}: {msg}")
        self.status = status


_RETRY_STATUS = (0, 500, 502, 503, 504)  # 0 = network problem / timeout; 503 = "high demand"
_WAITS = (5, 15, 30)                     # seconds between the 4 attempts on one model


def _gemini_once(key, model, system_prompt, transcript):
    body = {
        "systemInstruction": {"parts": [{"text": system_prompt}]},
        "contents": [{"role": "user", "parts": [{"text": "RAW TRANSCRIPT:\n\n" + transcript}]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 65536},
    }
    try:
        r = requests.post(f"{API_BASE}/models/{model}:generateContent",
                          headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                          json=body, timeout=600)
    except requests.RequestException as e:
        raise GeminiError(0, f"network problem: {e}")
    if r.status_code != 200:
        try:
            msg = r.json()["error"]["message"]
        except Exception:
            msg = r.text[:300]
        raise GeminiError(r.status_code, msg)
    cand = (r.json().get("candidates") or [{}])[0]
    text = "".join(p.get("text", "") for p in cand.get("content", {}).get("parts", []))
    if not text.strip():
        raise GeminiError(200, f"empty response (finishReason={cand.get('finishReason')}).")
    warn = "\nWARNING: output hit the token limit and may be truncated." if cand.get("finishReason") == "MAX_TOKENS" else ""
    return text, warn


def call_gemini(key, model, system_prompt, transcript, fallback="", on_status=None):
    """Calls Gemini. Busy/temporary errors (503, 500, timeouts) are retried with waits, then the fallback model is tried.
    Quota, key and bad-model errors are NOT retried. on_status(msg) is told about each retry."""
    import time
    say = on_status or (lambda m: None)
    models = [model] + ([fallback] if fallback and fallback != model else [])
    last = None
    for mi, m in enumerate(models):
        for attempt in range(len(_WAITS) + 1):
            try:
                text, warn = _gemini_once(key, m, system_prompt, transcript)
                if mi:
                    warn += f"\nNote: {models[0]} was busy, so {m} answered. Check tab 3 carefully."
                return text, warn
            except GeminiError as e:
                if e.status not in _RETRY_STATUS:
                    raise
                last = e
                if attempt < len(_WAITS):
                    say(f"{m} is busy (error {e.status or 'network'}). Retrying in {_WAITS[attempt]}s "
                        f"(attempt {attempt + 2} of {len(_WAITS) + 1})...")
                    time.sleep(_WAITS[attempt])
        if mi + 1 < len(models):
            say(f"{m} still busy. Trying {models[mi + 1]}...")
    raise GeminiError(last.status, f"{last.args[0].split(': ', 1)[-1]} (still failing after retries"
                      f"{' and the fallback model' if len(models) > 1 else ''}. Try again in a few minutes.)")


def strip_fence(text):
    m = re.search(r"```[a-zA-Z]*\n(.*?)```", text, re.S)
    return (m.group(1) if m else text).strip("\n")


# ------------------------------------------------------------------ parsing the list
def parse_list(text):
    """-> list of ('room', name) / ('item', name, desc)."""
    rows = []
    for line in text.replace("\r", "").split("\n"):
        if not line.strip():
            continue
        if "\t" in line:
            name, desc = line.split("\t", 1)
            rows.append(("item", name.strip(), re.sub(r"\t+", " ", desc).strip()))
        else:
            rows.append(("room", line.strip()))
    return rows


# ------------------------------------------------------------------ Pearl item order layer
# Item order taken from the "Pearl Property - Reference format" PDF. Runs AFTER the AI step, Pearl jobs only.
_PEARL_CLASSES = [  # (key, regex on lower-case item name) - first match wins
    ("security_screen", r"security"), ("external_walls", r"external wall"), ("eaves", r"\beaves"),
    ("gutters", r"gutter|downpipe"), ("garden_beds", r"garden"), ("lawn", r"\blawn"), ("driveway", r"driveway"),
    ("letterbox", r"letterbox|street number"), ("fences", r"fence|gate"), ("bins", r"garbage|\bbins?\b"),
    ("elec_meter", r"electric.*meter"), ("gas_meter", r"gas.*meter"), ("water_meter", r"water meter"),
    ("taps_hoses", r"hose"), ("clothesline", r"clothesline"), ("hot_water", r"hot water"),
    ("shower_tapware", r"shower.*(tap|rose|head|mixer)"), ("shower_tiles", r"shower.*tile"), ("shower_screen", r"shower"),
    ("bath_tapware", r"bath.*(tap|spout|mixer)"), ("bath_tiles", r"bath.*tile"), ("bath", r"\bbath\b"),
    ("wm_taps", r"washing machine"), ("tapware", r"\btaps?\b|tapware|mixer|spout"),
    ("toilet_roll", r"toilet roll"), ("toilet_bowl", r"toilet|bowl|cistern"),
    ("towel_rails", r"towel"), ("hand_rail", r"hand ?rail"),
    ("splash_back", r"splash"), ("bench_top", r"bench ?top"), ("sink", r"\bsink"), ("wash_basin", r"basin"), ("mirror", r"mirror"),
    ("aircon", r"air ?con"), ("range_hood", r"range ?hood|\bhood"), ("exhaust_fan", r"exhaust|heating|\bvent"),
    ("stove", r"stove|hot ?plate|cook ?top"), ("oven", r"oven|griller"),
    ("other_appliance", r"dishwasher|fridge|microwave|appliance"),
    ("pantry", r"pantry"), ("linen", r"linen"), ("trough_cabinet", r"trough.*cabinet"), ("trough", r"trough|\btub\b"),
    ("wardrobe", r"wardrobe|\brobe"), ("drawers", r"drawer"), ("underbench", r"underbench|vanity|cabinet|cupboard"),
    ("power", r"power|switch"), ("light", r"light"),
    ("door_treatments", r"door treatment"), ("external_door", r"external door|outside door"), ("sliding_door", r"sliding"),
    ("auto_door", r"automatic|roller|garage door"), ("door", r"door"),
    ("blinds", r"blind|curtain"), ("windows", r"window"),
    ("cornices", r"cornice|ceiling|bulkhead"), ("walls", r"wall|picture hook|doorstop"),
    ("skirting", r"skirting"), ("floor_coverings", r"floor|carpet"), ("additional", r"additional"),
]
_PEARL_RX = [(k, re.compile(rx)) for k, rx in _PEARL_CLASSES]

_GENERIC = ["security_screen", "door", "door_treatments", "external_door", "cornices", "exhaust_fan", "light", "walls",
            "aircon", "windows", "blinds", "sliding_door", "power", "skirting", "floor_coverings", "wash_basin",
            "tapware", "bench_top", "splash_back", "mirror", "underbench", "shower_screen", "shower_tiles",
            "shower_tapware", "towel_rails", "hand_rail", "bath", "bath_tiles", "bath_tapware", "toilet_bowl",
            "toilet_roll", "auto_door", "wardrobe", "additional"]
_TOILET = [k for k in _GENERIC if k != "exhaust_fan"]
_TOILET.insert(_TOILET.index("light") + 1, "exhaust_fan")
_TEMPLATES = {
    "generic": _GENERIC,
    "toilet": _TOILET,
    "kitchen": ["door", "cornices", "light", "walls", "aircon", "windows", "blinds", "sliding_door", "power", "skirting",
                "floor_coverings", "drawers", "underbench", "bench_top", "splash_back", "sink", "tapware", "range_hood",
                "stove", "oven", "other_appliance", "pantry", "additional"],
    "laundry": ["door", "door_treatments", "external_door", "blinds", "cornices", "light", "walls", "aircon", "windows",
                "power", "skirting", "floor_coverings", "wm_taps", "trough", "trough_cabinet", "linen", "additional"],
    "front": ["external_walls", "light", "eaves", "gutters", "floor_coverings", "garden_beds", "lawn", "driveway",
              "letterbox", "fences", "bins", "elec_meter", "gas_meter", "water_meter", "additional"],
    "rear": ["external_walls", "light", "gutters", "eaves", "fences", "taps_hoses", "lawn", "clothesline",
             "floor_coverings", "hot_water", "aircon", "additional"],
}


def _room_template(room):
    r = room.lower().lstrip("!# ").strip()
    if "front" in r and "property" in r:
        return "front"
    if "rear" in r and "property" in r:
        return "rear"
    if "kitchen" in r:
        return "kitchen"
    if "laundry" in r:
        return "laundry"
    if re.fullmatch(r"(toilet|wc|powder room)( \d+)?", r):
        return "toilet"
    return "generic"


def _classify(name):
    n = name.lower()
    for k, rx in _PEARL_RX:
        if rx.search(n):
            return k
    return None


def _order_block(lines, tpl):
    pos = {k: i for i, k in enumerate(tpl)}
    keyed, last = [], -1.0
    for idx, ln in enumerate(lines):
        k = _classify(ln.split("\t", 1)[0])
        if k in pos:
            last = float(pos[k])
            keyed.append((last, idx, ln))
        else:  # item not in the reference template: goes at the end of the room, in the order spoken
            keyed.append((float(len(tpl)) + 1.0, idx, ln))
    return [ln for _, _, ln in sorted(keyed)]


def reorder_pearl(text):
    """Reorder items inside each room to the Pearl PDF template. Room order and all text stay untouched."""
    out, block, room = [], [], None

    def flush():
        out.extend(_order_block(block, _TEMPLATES[_room_template(room or "")]))
        block.clear()

    for ln in text.replace("\r", "").split("\n"):
        if not ln.strip():
            continue
        if "\t" in ln:
            block.append(ln)
        else:
            flush()
            out.append(ln)
            room = ln
    flush()
    return "\n".join(out)


# ------------------------------------------------------------------ Normal jobs: item order from the blank report template
_NORMAL_CLASSES = [  # first match wins, so more specific patterns come first
    ("security", r"security"), ("front_door", r"front door"), ("roller_door", r"roller"),
    ("door_cov", r"door (?:coverings?|treatments?)"), ("ext_door", r"external door|sliding door"),
    ("window_treat", r"window (?:treatments?|coverings?)|blinds?|curtains?|shutters?"),
    ("windows", r"window|fly ?screen"), ("door", r"\bdoors?\b|door (?:&|and) frame"),
    ("driveway", r"driveway"), ("gate", r"\bgates?\b"), ("fence", r"fenc"), ("letterbox", r"letter ?box"),
    ("lawn", r"lawn"), ("garden", r"garden"), ("path", r"\bpath|paving|pavers?|stairwell"),
    ("woodwork", r"woodwork"), ("eaves", r"eaves|fascia"), ("roof", r"\broof"), ("alarm", r"alarm(?! clock)"),
    ("reticulation", r"reticulation|sprinkler"), ("taps_hoses", r"hoses?\b|taps? (?:&|and) hoses?"),
    ("meter", r"meter|metre|gas box"), ("gutters", r"gutter|down ?pipe"),
    ("clothesline", r"clothes ?line"), ("hot_water", r"hot water"), ("storage", r"storage|shed"),
    ("light", r"light (?:fitting|fixture)|^lights?\b"), ("points", r"points? ?(?:&|and)|power point|switch"),
    ("electrical", r"^electrical"), ("smoke", r"smoke"), ("aircon", r"air ?con|split system"),
    ("cornice", r"cornice"), ("ceiling", r"ceiling"), ("skirting", r"skirting"), ("floor", r"floor|carpet"),
    ("walls", r"walls?\b"), ("exhaust", r"exhaust|range ?hood"),
    ("bench", r"bench ?top"), ("pantry", r"pantry"), ("linen", r"linen"), ("robe", r"\brobe|wardrobe"),
    ("sink", r"^sink"), ("basin", r"basin"), ("tapware", r"tapware|^taps?\b"), ("mirror", r"mirror"),
    ("towel", r"towel|tower"), ("shower", r"shower"), ("bath", r"^bath\b"), ("toilet_roll", r"toilet (?:roll )?holder|roll holder"),
    ("toilet", r"toilet|cistern"), ("oven", r"oven|grill"), ("stove", r"stove|hot ?plate|cook ?top"),
    ("dishwasher", r"dishwasher"), ("wm_taps", r"washing machine"), ("trough", r"trough"),
    ("cupboards", r"cupboard|drawer|vanity"), ("additional", r"^additional"),
]
_NORMAL_RX = [(k, re.compile(rx)) for k, rx in _NORMAL_CLASSES]
_STD = ["door", "floor", "skirting", "walls", "cornice", "ceiling", "light", "points"]
_NORMAL_TEMPLATES = {
    "front": ["driveway", "fence", "gate", "letterbox", "lawn", "garden", "path", "walls", "woodwork", "eaves", "roof", "alarm",
              "reticulation", "taps_hoses", "meter", "gutters", "light", "points", "electrical", "additional"],
    "porch": ["walls", "ceiling", "floor", "light", "points", "electrical", "additional"],
    "garage": ["roller_door", "door", "floor", "skirting", "walls", "cornice", "ceiling", "light", "points", "additional"],
    "entry": ["security", "front_door", "floor", "skirting", "walls", "cornice", "ceiling", "light", "points", "alarm", "additional"],
    "study": _STD + ["aircon", "windows", "window_treat", "alarm", "additional"],
    "lounge": _STD + ["aircon", "alarm", "windows", "window_treat", "additional"],
    "kitchen": _STD + ["aircon", "alarm", "windows", "window_treat", "bench", "cupboards", "pantry", "sink", "tapware", "exhaust",
                       "oven", "stove", "dishwasher", "additional"],
    "living": ["door", "door_cov", "floor", "skirting", "walls", "cornice", "ceiling", "light", "points", "aircon", "alarm",
               "windows", "window_treat", "additional"],
    "hallway": _STD + ["aircon", "smoke", "alarm", "linen"],
    "backhall": _STD[:-1] + ["points", "alarm", "smoke", "linen", "additional"],
    "bedroom": _STD + ["aircon", "alarm", "windows", "window_treat", "robe", "additional"],
    "ensuite": _STD + ["exhaust", "windows", "window_treat", "bench", "cupboards", "mirror", "basin", "tapware", "towel", "shower",
                       "toilet", "toilet_roll", "additional"],
    "bathroom": _STD[:7] + ["exhaust", "points", "bench", "cupboards", "mirror", "basin", "tapware", "towel", "shower", "bath",
                            "toilet", "toilet_roll", "windows", "window_treat", "additional"],
    "toilet": _STD + ["toilet", "toilet_roll", "exhaust", "windows", "window_treat"],
    "laundry": ["door", "ext_door", "door_cov", "floor", "skirting", "walls", "cornice", "ceiling", "light", "points", "windows",
                "window_treat", "linen", "trough", "cupboards", "bench", "wm_taps", "additional"],
    "patio": _STD[:7],
    "rear": ["walls", "garden", "lawn", "reticulation", "taps_hoses", "light", "path", "eaves", "gutters", "clothesline", "fence",
             "hot_water", "storage", "additional"],
}


def _room_template_normal(room):
    r = room.lower().lstrip("!# ").strip()
    if "front" in r:
        return "front"
    if "rear" in r or "back yard" in r or "backyard" in r:
        return "rear"
    for key, words in (("porch", ("porch",)), ("garage", ("garage", "carport")), ("entry", ("entry", "entrance")),
                       ("study", ("study", "office")), ("kitchen", ("kitchen",)), ("backhall", ("back hall",)),
                       ("hallway", ("hall",)), ("ensuite", ("ensuite", "en-suite", "en suite")), ("bathroom", ("bathroom",)),
                       ("toilet", ("toilet", "wc", "powder")), ("laundry", ("laundry",)),
                       ("bedroom", ("bedroom", "bed room")), ("lounge", ("lounge",)),
                       ("patio", ("patio", "alfresco", "verandah", "balcony", "pergola")),
                       ("living", ("living", "dining", "family", "rumpus", "meals"))):
        if any(w in r for w in words):
            return key
    return "living"


def _classify_normal(name):
    n = name.lower()
    for k, rx in _NORMAL_RX:
        if rx.search(n):
            return k
    return None


def _order_block_normal(lines, tpl):
    pos = {k: i for i, k in enumerate(tpl)}
    keyed = []
    for idx, ln in enumerate(lines):
        k = _classify_normal(ln.split("\t", 1)[0])
        # items not in the template for this room go at the end of the room, in the order spoken
        keyed.append((float(pos[k]) if k in pos else float(len(tpl)) + 1.0, idx, ln))
    return [ln for _, _, ln in sorted(keyed)]


def reorder_normal(text):
    """Normal jobs: order the items inside each room like the blank report template. Rooms and all text stay untouched."""
    out, block, room = [], [], None

    def flush():
        out.extend(_order_block_normal(block, _NORMAL_TEMPLATES[_room_template_normal(room or "")]))
        block.clear()

    for ln in text.replace("\r", "").split("\n"):
        if not ln.strip():
            continue
        if "\t" in ln:
            block.append(ln)
        else:
            flush()
            out.append(ln)
            room = ln
    flush()
    return "\n".join(out)


def reorder_for(used_job, text):
    return reorder_pearl(text) if used_job == "pearl" else reorder_normal(text)


# ------------------------------------------------------------------ macro 1: FormatText
def _ws_rx(find):
    # a leading/trailing space is already a word boundary, so only add lookarounds on the open ends
    pat = re.escape(find.rstrip()).replace(r"\ ", r"\s+") + (r"\s+" if find.endswith(" ") else "")
    # word-boundary checks only where the edge of the find text is a letter/number (not a space or punctuation like ". ")
    lead = r"(?<![A-Za-z0-9])" if find[:1].isalnum() else ""
    trail = r"(?![A-Za-z0-9])" if find.rstrip()[-1:].isalnum() and not find.endswith(" ") else ""
    return re.compile(lead + pat + trail, re.I)


@lru_cache(maxsize=16)
def _compile_rules(rules):
    return [(_ws_rx(f), r) for f, r in rules if f.strip()]


def format_text(s, rules=None):
    """rules = [(find, replace), ...]; defaults to TEXT_REPLACEMENTS."""
    s = s.replace("\u00a0", " ").replace("\u202f", " ").replace("\u201a", ",").replace(",", ".")
    for rx, rep in _compile_rules(tuple((f, r) for f, r in (TEXT_REPLACEMENTS if rules is None else rules))):
        s = rx.sub(lambda m, rep=rep: rep, s)
    return s



REPL_FILE = APP_DIR / "replacements.json"


def default_rules():
    return {"normal": [list(x) for x in TEXT_REPLACEMENTS], "pearl": [list(x) for x in TEXT_REPLACEMENTS]}


def load_rules():
    """Separate replacement lists for Normal and Pearl jobs (saved in replacements.json)."""
    rules = default_rules()
    try:
        saved = json.loads(REPL_FILE.read_text(encoding="utf-8"))
        for k in rules:
            if isinstance(saved.get(k), list):
                rules[k] = [[str(a), str(b)] for a, b in saved[k] if str(a).strip()]
    except Exception:
        pass
    return rules


def save_rules(rules):
    REPL_FILE.write_text(json.dumps(rules, indent=1, ensure_ascii=False), encoding="utf-8")


def _unquote(x):
    x = x.strip()
    return x[1:-1] if len(x) >= 2 and x[0] == x[-1] == '"' else x


def parse_rules_txt(text):
    """word_replacements.txt format: one 'find => replacement' per line. Lines starting with # are ignored.
    Blank replacement removes the word. Wrap a side in "quotes" to keep leading/trailing spaces (e.g. " is " => " ")."""
    out = []
    for line in text.replace("\r", "").split("\n"):
        if not line.strip() or line.lstrip().startswith("#") or "=>" not in line:
            continue
        a, b = line.split("=>", 1)
        a, b = _unquote(a), _unquote(b)
        if a.strip():
            out.append([a, b])
    return out


def rules_to_txt(rules):
    q = lambda v: f'"{v}"' if v != v.strip() else v
    head = ("# CUSTOM WORD REPLACEMENTS\n# One rule per line: find => replacement\n"
            "# Matching ignores capitalisation and uses whole words / phrases.\n"
            "# Leave the replacement blank to remove the word or phrase.\n"
            '# Wrap a side in "quotes" to keep spaces around it, e.g. " is " => " "\n\n')
    return head + "\n".join(f"{q(a)} => {q(b)}".rstrip() for a, b in rules) + "\n"


def merge_rules(existing, new):
    """Add new [find, replace] pairs to a list; a find that already exists (ignoring case) is updated."""
    added = updated = 0
    for a, b in new:
        for row in existing:
            if row[0].lower() == a.lower():
                if row[1] != b:
                    updated += 1
                row[1] = b
                break
        else:
            existing.append([a, b])
            added += 1
    return added, updated


def safe_name(name, default="Inventory"):
    return re.sub(r'[\\/:*?"<>|]+', "-", (name or "").strip()).strip(" .") or default


def unique_base(folder, name, exts):
    """folder/name, adding ' (2)', ' (3)'... until none of the files exist."""
    base, n = os.path.join(folder, name), 1
    while any(os.path.exists(base + e) for e in exts):
        n += 1
        base = os.path.join(folder, f"{name} ({n})")
    return base


# ------------------------------------------------------------------ macro 2: FormatTable
def start_case(s):
    """Capitalise each word, including after '/', '-', '(' (e.g. Cornices/Ceiling). '!' and '#' markers are kept."""
    s = re.sub(r"\s+", " ", s.strip()).lower()
    return re.sub(r"(^|[\s/\-(!#])([a-z])", lambda m: m.group(1) + m.group(2).upper(), s)


def sentence_case(s):
    out, cap, i = [], True, 0
    while i < len(s):
        ch = s[i]
        if ch == "@":
            out.append("@")
            i += 1
            if i < len(s):
                ch = s[i]
                out.append(ch.upper())
                cap = False
        elif cap and ch.isalpha() and ch.isascii():
            out.append(ch.upper())
            cap = False
        else:
            out.append(ch)
            if ch.isdigit():
                cap = False  # a sentence starting with a number must not capitalise the next word
        # sentence end; a decimal point between digits (2.4 meter) is NOT a sentence end
        if ch in "!?" or (ch == "." and not (i > 0 and s[i - 1].isdigit() and i + 1 < len(s) and s[i + 1].isdigit())):
            cap = True
        i += 1
    return "".join(out)


def strip_bang(s):
    out, skip = [], False
    for ch in s:
        if ch == "!":
            skip = True
        elif skip and ch.isalpha() and ch.isascii():
            out.append(ch.upper())
            skip = False
        else:
            out.append(ch)
    return "".join(out).strip()


# ------------------------------------------------------------------ macro 3: ChangeYtoN
def _hit(text, kws):
    return "N" if any(k.strip() and k.strip().lower() in text for k in kws) else "Y"


KW_FILE = APP_DIR / "keywords.json"
KW_COLS = ("clean", "undamaged", "working")
KW_LABELS = {"clean": "Clean", "undamaged": "Undamaged", "working": "Working"}


def _kw_defaults():
    return {"clean": [k.strip() for k in CLEAN_KW], "undamaged": [k.strip() for k in DAMAGE_KW],
            "working": [k.strip() for k in WORKING_KW]}


def default_keywords():
    return {"normal": _kw_defaults(), "pearl": _kw_defaults()}


def load_keywords():
    """Separate Clean / Undamaged / Working keyword lists for Normal and Pearl (saved in keywords.json)."""
    kws = default_keywords()
    try:
        saved = json.loads(KW_FILE.read_text(encoding="utf-8"))
        for j in kws:
            for c in KW_COLS:
                v = saved.get(j, {}).get(c)
                if isinstance(v, list):
                    kws[j][c] = [str(x).strip() for x in v if str(x).strip()]
    except Exception:
        pass
    return kws


def save_keywords(kws):
    KW_FILE.write_text(json.dumps(kws, indent=1, ensure_ascii=False), encoding="utf-8")


def yn(desc, kws=None):
    """A column becomes N when any of its keywords appears in the description (substring, ignores case)."""
    kws = kws or _kw_defaults()
    d = desc.lower()
    return tuple(_hit(d, kws[c]) for c in KW_COLS)


def run_macros(rows, rules=None, kws=None):
    """Apply FormatText -> FormatTable -> ChangeYtoN. -> list of dicts. rules = replacement list for the job."""
    out = []
    for r in rows:
        if r[0] == "room":
            out.append({"room": strip_bang(start_case(format_text(r[1], rules)))})
        else:
            name = strip_bang(start_case(format_text(r[1], rules)))
            desc = strip_bang(sentence_case(format_text(r[2], rules)))
            c, u, w = yn(desc, kws)
            out.append({"item": name, "desc": desc, "clean": c, "undamaged": u, "working": w})
    return out


# ------------------------------------------------------------------ omitted / added check
_TAGS = re.compile(r"<[^>]*>|\[[^\]]*\]|\b\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?\b")
_WORD = re.compile(r"[a-z0-9]+(?:'[a-z]+)?")
# fillers plus small words your FormatText macro deletes anyway, so they never count as "missing"
_FILLERS = {"um", "uh", "ah", "er", "erm", "hmm", "umm", "the", "is", "are", "a", "an", "there"}


def _tok(s):
    return Counter(w for w in _WORD.findall(s.lower()) if w not in _FILLERS)


def compare(raw, list_text):
    rows = parse_list(list_text)
    raw_t = _tok(_TAGS.sub(" ", raw))
    all_out = _tok(" ".join(" ".join(r[1:]) for r in rows))
    desc_out = _tok(" ".join(r[2] for r in rows if r[0] == "item"))
    omitted = raw_t - all_out
    added = desc_out - raw_t
    fmt = lambda c: ", ".join(f"{w}" + (f" x{n}" if n > 1 else "") for w, n in sorted(c.items()))
    return fmt(omitted), fmt(added), sum(omitted.values()), sum(added.values())


def fidelity(raw, no, na):
    """-> (% of transcript words missing, % added, flagged). Flags when the model has probably altered the transcript."""
    total = sum(_tok(_TAGS.sub(" ", raw)).values()) or 1
    po, pa = 100.0 * no / total, 100.0 * na / total
    return po, pa, (po > 4.0 or pa > 3.0)


# ------------------------------------------------------------------ layout
HEAD = {"blank": "", "item": "Item", "desc": "Description", "clean": "Clean", "undamaged": "Undamaged", "working": "Working"}
LAYOUTS = {
    "standard": (["item", "desc", "clean", "undamaged", "working"], [3.4, 8.3, 2.1, 2.1, 2.1]),
    "final": (["item", "clean", "undamaged", "working", "blank", "desc"], [3.4, 2.1, 2.1, 2.1, 2.2, 6.1]),
}
NAVY, ROOM_FILL, N_FILL, N_FONT, Y_FONT = "1F3864", "D9E2F3", "FDE9E7", "C00000", "2E7D32"
SPACE_PT = 4

# ------------------------------------------------------------------ local rules mode (no API)
def process_local(raw, used_job, do_reorder):
    """Rules-based tidy-up, no AI. -> (list_text, review_text). Same list format as the Gemini step."""
    import local_parser
    lst, review = local_parser.local_parse(raw, used_job)
    if do_reorder:
        lst = reorder_for(used_job, lst)
    lines = []
    for r in review:
        lines.append(f"[{r['kind']}] {('(' + r['room'] + ') ') if r['room'] else ''}{r['text']}\n      -> {r['note']}")
    short = local_parser.short_lines(raw, used_job)
    txt = ""
    if lines:
        txt += f"{len(lines)} THINGS TO CHECK (corrections, go-backs and anything unresolved):\n\n" + "\n\n".join(lines) + "\n\n"
    if short:
        txt += ("SHORT LINES NOT RECOGNISED AS HEADINGS (if any is a heading, add it to the item list - button 'Edit item list'):\n"
                + "\n".join(sorted(set(short))) + "\n\n")
    return lst, txt


# ------------------------------------------------------------------ Word
def _shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _set_lang(rPr, code="en-AU"):
    """Mark text as English (Australia) so Word's spell-check treats colour, grey, centre... as correct."""
    for old in rPr.findall(qn("w:lang")):
        rPr.remove(old)
    lang = OxmlElement("w:lang")
    for k in ("val", "eastAsia", "bidi"):
        lang.set(qn(f"w:{k}"), code)
    rPr.append(lang)


def _put(cell, text, bold=False, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(SPACE_PT)
    p.paragraph_format.space_after = Pt(SPACE_PT)
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.bold = bold
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Calibri")
    _set_lang(run._element.get_or_add_rPr())
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def write_word(data, layout, path):
    cols, widths = LAYOUTS[layout]
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(1.5)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(10)
    st.element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "Calibri")
    _set_lang(st.element.get_or_add_rPr())
    # also the document-wide defaults, so empty paragraphs and anything typed later are en-AU too
    dd = doc.styles.element.find(qn("w:docDefaults"))
    if dd is not None:
        rpr = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr")) if False else None
        rd = dd.find(qn("w:rPrDefault"))
        if rd is not None and rd.find(qn("w:rPr")) is not None:
            _set_lang(rd.find(qn("w:rPr")))

    tbl = doc.add_table(rows=1, cols=len(cols))
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        for k, v in (("val", "single"), ("sz", "8"), ("space", "0"), ("color", "000000")):
            e.set(qn(f"w:{k}"), v)
        borders.append(e)
    tblPr.append(borders)

    def finish_row(row, header=False):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if header:
            trPr.append(OxmlElement("w:tblHeader"))
        for c, w in zip(row.cells, widths):
            c.width = Cm(w)

    hdr = tbl.rows[0]
    finish_row(hdr, header=True)
    for c, key in zip(hdr.cells, cols):
        _put(c, HEAD[key], bold=True, color="FFFFFF",
             align=WD_ALIGN_PARAGRAPH.LEFT if key in ("item", "desc") else WD_ALIGN_PARAGRAPH.CENTER)
        _shade(c, NAVY)

    for d in data:
        row = tbl.add_row()
        finish_row(row)
        if "room" in d:
            m = row.cells[0].merge(row.cells[-1])
            _put(m, d["room"], bold=True, color=NAVY)
            _shade(m, ROOM_FILL)
            continue
        for c, key in zip(row.cells, cols):
            v = d.get(key, "")
            if key == "blank":
                _put(c, "")
            elif key == "item":
                _put(c, v, bold=True)
            elif key == "desc":
                _put(c, v)
            else:
                _put(c, v, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.save(path)


# ------------------------------------------------------------------ Excel
# Flat import-style sheet (matches the reference workbook): one row per item, the room repeated in column A.
XL_HEAD = {"category": "category", "item": "subcategory", "desc": "comments", "clean": "clean",
           "undamaged": "undamaged", "working": "working", "blank": ""}
XL_LAYOUTS = {
    "standard": (["category", "item", "desc", "clean", "undamaged", "working"], [38.9, 23.5, 68.0, 11.0, 13.6, 10.5]),
    "final": (["category", "item", "clean", "undamaged", "working", "blank", "desc"], [38.9, 23.5, 11.0, 13.6, 10.5, 10.5, 68.0]),
}


def write_excel(data, layout, path):
    cols, widths = XL_LAYOUTS[layout]
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    font = Font(name="Calibri", size=11)
    top = Border(top=Side(style="thin"))

    for i, (key, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(1, i, XL_HEAD[key])
        c.font = font
        c.alignment = Alignment(wrap_text=(key == "desc"))
        ws.column_dimensions[c.column_letter].width = w

    r, room = 2, ""
    for d in data:
        if "room" in d:
            room = d["room"].upper()
            continue
        vals = {"category": room, "item": d["item"], "desc": d["desc"], "clean": d["clean"],
                "undamaged": d["undamaged"], "working": d["working"], "blank": ""}
        for i, key in enumerate(cols, 1):
            c = ws.cell(r, i, vals[key] or None)
            c.font = font
            c.border = top
            c.alignment = Alignment(vertical="center", wrap_text=key in ("category", "item", "desc"))
        r += 1

    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    wb.save(path)


# ------------------------------------------------------------------ GUI
def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    cfg = {}
    if CONFIG.exists():
        try:
            cfg = json.loads(CONFIG.read_text())
        except Exception:
            pass

    root = tk.Tk()
    root.title("Transcript to Inventory")

    def _report(exc, val, tb):  # works without a console (pythonw): log to error.log and show a box
        import traceback
        msg = "".join(traceback.format_exception(exc, val, tb))
        try:
            with open(APP_DIR / "error.log", "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except OSError:
            pass
        messagebox.showerror("Unexpected error", msg[-1200:])
    root.report_callback_exception = _report
    root.geometry("1100x780")

    key_var = tk.StringVar(value=cfg.get("api_key", ""))
    model_var = tk.StringVar(value=cfg.get("model", DEFAULT_MODEL))
    fallback_var = tk.StringVar(value=cfg.get("fallback_model", "gemini-3.6-flash"))
    remember = tk.BooleanVar(value=bool(cfg.get("api_key")))
    job = tk.StringVar(value="normal")
    want_word, want_excel = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
    layout = tk.StringVar(value="standard")
    status = tk.StringVar(value="Ready.")
    out_dir = tk.StringVar(value=cfg.get("out_dir", ""))
    file_base = tk.StringVar(value=cfg.get("file_base", "Inventory"))
    rules_state = load_rules()
    kw_state = load_keywords()
    state = {"job": None}

    top = ttk.Frame(root, padding=8)
    top.pack(fill="x")
    ttk.Label(top, text="Gemini API key").grid(row=0, column=0, sticky="w")
    ttk.Entry(top, textvariable=key_var, show="*", width=48).grid(row=0, column=1, padx=6)
    ttk.Label(top, text="Model").grid(row=0, column=2, sticky="w")
    model_box = ttk.Combobox(top, textvariable=model_var, width=26,
                             values=["gemini-3.8-flash", "gemini-3.6-flash", "gemini-3.1-pro-preview"])
    model_box.grid(row=0, column=3, padx=6)
    ttk.Label(top, text="If busy, try").grid(row=1, column=2, sticky="w", pady=(4, 0))
    fallback_box = ttk.Combobox(top, textvariable=fallback_var, width=26,
                                values=["", "gemini-3.6-flash", "gemini-3.8-flash", "gemini-3.1-flash-lite-preview"])
    fallback_box.grid(row=1, column=3, padx=6, pady=(4, 0))
    ttk.Label(top, text="(blank = no fallback)", foreground="#555").grid(row=1, column=4, sticky="w", pady=(4, 0))
    ttk.Checkbutton(top, text="Remember key on this PC", variable=remember).grid(row=0, column=4, padx=6)

    def save_cfg():
        data = {"model": model_var.get().strip(), "fallback_model": fallback_var.get().strip(), "out_dir": out_dir.get().strip(), "file_base": file_base.get().strip(), "method": method.get()}
        if remember.get():
            data["api_key"] = key_var.get().strip()
        try:
            CONFIG.write_text(json.dumps(data))
        except OSError:
            pass

    root.protocol("WM_DELETE_WINDOW", lambda: (save_cfg(), root.destroy()))

    def do_test():
        save_cfg()
        k = get_key(key_var.get())
        if not k:
            return messagebox.showwarning("API key", "Enter a key (or set GEMINI_API_KEY).")
        status.set("Testing key and fetching models...")
        def work():
            ok, msg = test_key(k, model_var.get().strip())
            ms = []
            try:
                ms = list_models(k)
            except Exception:
                pass
            def done():
                if ms:
                    model_box["values"] = ms
                    fallback_box["values"] = [""] + ms
                    m2 = msg + f"\n\n{len(ms)} models are now in the Model dropdown - pick one from the list."
                else:
                    m2 = msg
                status.set(m2.splitlines()[0])
                (messagebox.showinfo if ok else messagebox.showerror)("API key test", m2)
            root.after(0, done)
        threading.Thread(target=work, daemon=True).start()
    ttk.Button(top, text="Test key / list models", command=do_test).grid(row=0, column=5)

    def do_list():
        k = get_key(key_var.get())
        if not k:
            return messagebox.showwarning("API key", "Enter a key first.")
        status.set("Fetching model list...")
        def work():
            try:
                ms = list_models(k)
                def ok():
                    model_box["values"] = ms
                    fallback_box["values"] = [""] + ms
                    status.set(f"{len(ms)} models loaded into the Model dropdown.")
                root.after(0, ok)
            except Exception as e:
                err = str(e)
                root.after(0, lambda: (status.set("Failed."), messagebox.showerror("Models", err)))
        threading.Thread(target=work, daemon=True).start()
    ttk.Button(top, text="List models", command=do_list).grid(row=0, column=6, padx=4)

    jb = ttk.LabelFrame(root, text="Job type", padding=6)
    jb.pack(fill="x", padx=8)
    ttk.Radiobutton(jb, text="Normal", variable=job, value="normal").pack(side="left", padx=8)
    ttk.Radiobutton(jb, text="Pearl", variable=job, value="pearl").pack(side="left", padx=8)
    method = tk.StringVar(value=cfg.get("method", "gemini"))
    ttk.Label(jb, text="  Method:").pack(side="left", padx=(16, 0))
    ttk.Radiobutton(jb, text="Gemini (AI)", variable=method, value="gemini").pack(side="left", padx=4)
    ttk.Radiobutton(jb, text="Local rules (no API)", variable=method, value="local").pack(side="left", padx=4)
    def edit_items():
        import local_parser
        if not os.path.exists(local_parser.VOCAB_FILE):
            local_parser.save_vocab(local_parser.load_vocab())
        try:
            os.startfile(local_parser.VOCAB_FILE)
        except Exception:
            messagebox.showinfo("Item list", "Open this file in a text editor:\n" + local_parser.VOCAB_FILE)
    ttk.Button(jb, text="Edit item list", command=edit_items).pack(side="left", padx=4)
    reorder = tk.BooleanVar(value=True)
    ttk.Checkbutton(jb, text="Reorder items to reference template (Pearl: PDF, Normal: blank report)", variable=reorder).pack(side="left", padx=16)

    def open_rules():
        win = tk.Toplevel(root)
        win.title("Replacement words")
        win.geometry("700x560")
        win.transient(root)
        which = tk.StringVar(value=job.get())
        f_var, r_var = tk.StringVar(), tk.StringVar()

        top = ttk.Frame(win, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Edit list for:").pack(side="left")
        tree = ttk.Treeview(win, columns=("find", "repl"), show="headings", height=14)
        tree.heading("find", text="Find (whole word, ignores case)")
        tree.heading("repl", text="Replace with")
        tree.column("find", width=300)
        tree.column("repl", width=300)

        def refresh():
            tree.delete(*tree.get_children())
            for i, (a, b) in enumerate(rules_state[which.get()]):
                tree.insert("", "end", iid=str(i), values=(f'"{a}"', f'"{b}"'))

        def persist():
            try:
                save_rules(rules_state)
            except OSError as e:
                messagebox.showerror("Replacements", f"Could not save: {e}", parent=win)

        for label, val in (("Normal prompt", "normal"), ("Pearl prompt", "pearl")):
            ttk.Radiobutton(top, text=label, variable=which, value=val, command=refresh).pack(side="left", padx=8)
        tree.pack(fill="both", expand=True, padx=8)

        def on_select(_e=None):
            sel = tree.selection()
            if sel:
                a, b = rules_state[which.get()][int(sel[0])]
                f_var.set(a)
                r_var.set(b)
        tree.bind("<<TreeviewSelect>>", on_select)

        form = ttk.Frame(win, padding=8)
        form.pack(fill="x")
        ttk.Label(form, text="Find").grid(row=0, column=0, sticky="w")
        ttk.Entry(form, textvariable=f_var, width=34).grid(row=0, column=1, padx=6)
        ttk.Label(form, text="Replace with").grid(row=0, column=2, sticky="w")
        ttk.Entry(form, textvariable=r_var, width=34).grid(row=0, column=3, padx=6)

        def add_update():
            a, b = f_var.get(), r_var.get()
            if not a.strip():
                return messagebox.showwarning("Replacements", "Enter the word or phrase to find.", parent=win)
            lst = rules_state[which.get()]
            for row in lst:
                if row[0].lower() == a.lower():
                    row[1] = b
                    break
            else:
                lst.append([a, b])
            persist()
            refresh()
            f_var.set("")
            r_var.set("")

        def delete():
            sel = sorted((int(i) for i in tree.selection()), reverse=True)
            if not sel:
                return messagebox.showinfo("Replacements", "Select one or more rows to delete.", parent=win)
            for i in sel:
                del rules_state[which.get()][i]
            persist()
            refresh()

        def reset():
            if messagebox.askyesno("Replacements", f"Reset the {which.get()} list to the original default words?", parent=win):
                rules_state[which.get()] = default_rules()[which.get()]
                persist()
                refresh()

        def import_txt():
            p = filedialog.askopenfilename(parent=win, initialdir=str(APP_DIR), initialfile="word_replacements.txt",
                                           filetypes=[("Replacement list", "*.txt"), ("All files", "*.*")])
            if not p:
                return
            new = parse_rules_txt(Path(p).read_text(encoding="utf-8-sig", errors="ignore"))
            if not new:
                return messagebox.showinfo("Import", "No active rules found. Rules look like:  find => replacement\n"
                                           "(lines starting with # are ignored).", parent=win)
            both = messagebox.askyesnocancel("Import", f"Found {len(new)} rules.\n\nYes = add to BOTH lists\n"
                                             f"No = add to the {which.get()} list only", parent=win)
            if both is None:
                return
            targets = ["normal", "pearl"] if both else [which.get()]
            for k in targets:
                a, u = merge_rules(rules_state[k], new)
            persist()
            refresh()
            messagebox.showinfo("Import", f"Imported {len(new)} rules into: {', '.join(targets)}.", parent=win)

        def export_txt():
            p = filedialog.asksaveasfilename(parent=win, initialdir=str(APP_DIR), initialfile=f"word_replacements_{which.get()}.txt",
                                             defaultextension=".txt", filetypes=[("Replacement list", "*.txt")])
            if p:
                Path(p).write_text(rules_to_txt(rules_state[which.get()]), encoding="utf-8")
                messagebox.showinfo("Export", f"Saved the {which.get()} list to:\n{p}", parent=win)

        btns = ttk.Frame(win, padding=8)
        btns.pack(fill="x")
        ttk.Button(btns, text="Import .txt...", command=import_txt).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Export .txt...", command=export_txt).pack(side="left", padx=(0, 16))
        ttk.Button(btns, text="Add / update", command=add_update).pack(side="left")
        ttk.Button(btns, text="Delete selected", command=delete).pack(side="left", padx=8)
        ttk.Button(btns, text="Reset to default", command=reset).pack(side="left")
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="right")
        ttk.Label(win, wraplength=660, justify="left", foreground="#555",
                  text='Tip: spaces matter. A find of " is " (a space either side) removes the word "is". '
                       'Leave Replace empty to delete the word. Select a row to edit it, then click Add / update. '
                       'Import .txt reads lines like  find => replacement.  Changes apply to the next export and are saved in replacements.json.').pack(fill="x", padx=8, pady=(0, 8))
        refresh()
    ttk.Button(jb, text="Replacement words...", command=open_rules).pack(side="left", padx=4)

    def open_kw():
        win = tk.Toplevel(root)
        win.title("Clean / Undamaged / Working keywords")
        win.geometry("640x600")
        win.transient(root)
        jobv, colv, entry_v = tk.StringVar(value=job.get()), tk.StringVar(value="clean"), tk.StringVar()
        count_v = tk.StringVar()

        cur = lambda: kw_state[jobv.get()][colv.get()]

        def persist():
            try:
                save_keywords(kw_state)
            except OSError as e:
                messagebox.showerror("Keywords", f"Could not save: {e}", parent=win)

        r1 = ttk.Frame(win, padding=(8, 8, 8, 0))
        r1.pack(fill="x")
        ttk.Label(r1, text="Prompt:").pack(side="left")
        r2 = ttk.Frame(win, padding=(8, 4, 8, 0))
        r2.pack(fill="x")
        ttk.Label(r2, text="Column:").pack(side="left")

        lb_frame = ttk.Frame(win)
        lb_frame.pack(fill="both", expand=True, padx=8, pady=6)
        lb = tk.Listbox(lb_frame, selectmode="extended", font=("Calibri", 10))
        sb = ttk.Scrollbar(lb_frame, orient="vertical", command=lb.yview)
        lb.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)

        def refresh():
            lb.delete(0, "end")
            for k in cur():
                lb.insert("end", k)
            count_v.set(f"{len(cur())} keywords. If ANY appears in a description, "
                        f"the {KW_LABELS[colv.get()]} column becomes N (otherwise Y).")

        for label, val in (("Normal", "normal"), ("Pearl", "pearl")):
            ttk.Radiobutton(r1, text=label, variable=jobv, value=val, command=refresh).pack(side="left", padx=8)
        for val in KW_COLS:
            ttk.Radiobutton(r2, text=KW_LABELS[val], variable=colv, value=val, command=refresh).pack(side="left", padx=8)
        ttk.Label(win, textvariable=count_v, wraplength=600, foreground="#555").pack(fill="x", padx=8)

        def on_select(_e=None):
            sel = lb.curselection()
            if len(sel) == 1:
                entry_v.set(lb.get(sel[0]))
        lb.bind("<<ListboxSelect>>", on_select)

        f = ttk.Frame(win, padding=8)
        f.pack(fill="x")
        ttk.Label(f, text="Keyword(s)").pack(side="left")
        ttk.Entry(f, textvariable=entry_v, width=52).pack(side="left", padx=6)

        def add():
            new = [x.strip() for x in entry_v.get().split(",") if x.strip()]
            if not new:
                return messagebox.showwarning("Keywords", "Type a keyword (several can be separated by commas).", parent=win)
            have = {k.lower() for k in cur()}
            n = 0
            for k in new:
                if k.lower() not in have:
                    cur().append(k)
                    have.add(k.lower())
                    n += 1
            persist()
            refresh()
            lb.see("end")
            entry_v.set("")
            if n < len(new):
                messagebox.showinfo("Keywords", f"Added {n}. {len(new) - n} already in the list.", parent=win)

        def update():
            sel = lb.curselection()
            new = entry_v.get().strip()
            if len(sel) != 1 or not new:
                return messagebox.showinfo("Keywords", "Select ONE keyword, edit the text box, then click Update selected.", parent=win)
            if any(k.lower() == new.lower() for i, k in enumerate(cur()) if i != sel[0]):
                return messagebox.showinfo("Keywords", "That keyword is already in the list.", parent=win)
            cur()[sel[0]] = new
            persist()
            refresh()
            entry_v.set("")

        def delete():
            sel = sorted(lb.curselection(), reverse=True)
            if not sel:
                return messagebox.showinfo("Keywords", "Select one or more keywords to delete.", parent=win)
            for i in sel:
                del cur()[i]
            persist()
            refresh()

        def reset():
            if messagebox.askyesno("Keywords", f"Reset the {KW_LABELS[colv.get()]} list ({jobv.get()}) to the original keywords?", parent=win):
                kw_state[jobv.get()][colv.get()] = _kw_defaults()[colv.get()]
                persist()
                refresh()

        b = ttk.Frame(win, padding=(8, 0, 8, 8))
        b.pack(fill="x")
        ttk.Button(b, text="Add", command=add).pack(side="left")
        ttk.Button(b, text="Update selected", command=update).pack(side="left", padx=8)
        ttk.Button(b, text="Delete selected", command=delete).pack(side="left")
        ttk.Button(b, text="Reset to default", command=reset).pack(side="left", padx=8)
        ttk.Button(b, text="Close", command=win.destroy).pack(side="right")
        ttk.Label(win, wraplength=600, justify="left", foreground="#555",
                  text="Matching is by letters inside the text (ignores case), so \"chip\" also catches \"chipped\" and \"chipboard\". "
                       "Changes apply to the next export and are saved in keywords.json.").pack(fill="x", padx=8, pady=(0, 8))
        refresh()
    ttk.Button(jb, text="Y/N keywords...", command=open_kw).pack(side="left", padx=4)
    ref_lbl.pack(side="left")

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=6)
    t_raw = scrolledtext.ScrolledText(nb, wrap="word", font=("Calibri", 10), undo=True)
    t_list = scrolledtext.ScrolledText(nb, wrap="none", font=("Consolas", 10), undo=True)
    t_rev = scrolledtext.ScrolledText(nb, wrap="word", font=("Calibri", 10))
    nb.add(t_raw, text="1. Raw transcript")
    nb.add(t_list, text="2. Formatted list (editable)")
    nb.add(t_rev, text="3. Omitted / added check")

    rb = ttk.Frame(root)
    rb.pack(fill="x", padx=8)
    ttk.Button(rb, text="Paste", command=lambda: (t_raw.delete("1.0", "end"), t_raw.insert("1.0", root.clipboard_get()))).pack(side="left")
    ttk.Button(rb, text="Clear", command=lambda: t_raw.delete("1.0", "end")).pack(side="left", padx=4)
    def load_txt():
        p = filedialog.askopenfilename(filetypes=[("Transcript files", TRANSCRIPT_TYPES), ("All files", "*.*")])
        if p:
            try:
                txt = read_transcript(p, Path(p).read_bytes())
            except Exception as e:
                return messagebox.showerror("Load file", f"Could not read that file: {e}")
            t_raw.delete("1.0", "end")
            t_raw.insert("1.0", txt)
    ttk.Button(rb, text="Load file...", command=load_txt).pack(side="left")

    bot = ttk.LabelFrame(root, text="Output", padding=6)
    bot.pack(fill="x", padx=8, pady=4)
    ttk.Checkbutton(bot, text="Word (.docx)", variable=want_word).pack(side="left", padx=6)
    ttk.Checkbutton(bot, text="Excel (.xlsx)", variable=want_excel).pack(side="left", padx=6)
    ttk.Radiobutton(bot, text="Standard: Item | Description | Clean | Undamaged | Working", variable=layout, value="standard").pack(side="left", padx=14)
    ttk.Radiobutton(bot, text="Final report: Item | Clean | Undamaged | Working | (blank) | Description", variable=layout, value="final").pack(side="left")
    bot2 = ttk.Frame(bot)
    bot2.pack(fill="x", pady=(6, 0))
    ttk.Label(bot2, text="Save to folder").pack(side="left", padx=6)
    ttk.Entry(bot2, textvariable=out_dir, width=52).pack(side="left")
    def browse_dir():
        d = filedialog.askdirectory(initialdir=out_dir.get() or None)
        if d:
            out_dir.set(os.path.normpath(d))
            save_cfg()
    ttk.Button(bot2, text="Browse...", command=browse_dir).pack(side="left", padx=4)
    ttk.Label(bot2, text="File name").pack(side="left", padx=(14, 4))
    ttk.Entry(bot2, textvariable=file_base, width=24).pack(side="left")
    ttk.Label(bot2, text="(folder blank = ask each time)", foreground="#555").pack(side="left", padx=8)

    act = ttk.Frame(root)
    act.pack(fill="x", padx=8, pady=4)
    b_proc = ttk.Button(act, text="1. Process transcript")
    b_exp = ttk.Button(act, text="2. Export")
    b_proc.pack(side="left")
    b_exp.pack(side="left", padx=8)
    ttk.Label(act, textvariable=status).pack(side="left", padx=12)

    def process():
        raw = t_raw.get("1.0", "end").strip()
        k = get_key(key_var.get())
        if not raw:
            return messagebox.showwarning("Transcript", "Paste a raw transcript first.")
        if method.get() == "local":
            save_cfg()
            used_job = job.get()
            lst, review = process_local(raw, used_job, reorder.get())
            om, ad, no, na = compare(raw, lst)
            rows = parse_list(lst)
            state["job"] = used_job
            t_list.delete("1.0", "end")
            t_list.insert("1.0", lst)
            t_rev.delete("1.0", "end")
            t_rev.insert("1.0", f"LOCAL RULES (no AI). Nothing was invented: {na} words added. "
                         f"{no} transcript words are not in the output - mostly the spoken headings (turned into item names), "
                         f"repeated item names, fillers and corrected words.\n\n" + review +
                         f"POSSIBLY OMITTED from transcript:\n{om or 'none'}\n\nPOSSIBLY ADDED in descriptions:\n{ad or 'none'}\n")
            nb.select(t_rev if review else t_list)
            status.set(f"Done (local rules): {sum(1 for r in rows if r[0]=='item')} items. Check tab 3 for items to review, then Export.")
            return
        if not k:
            return messagebox.showwarning("API key", "Enter a Gemini API key (or set GEMINI_API_KEY).")
        save_cfg()
        sp = prompts.NORMAL if job.get() == "normal" else prompts.PEARL
        sp += prompts.FIDELITY
        used_job = job.get()
        do_reorder = reorder.get()
        b_proc.state(["disabled"])
        status.set("Processing with Gemini... (large transcripts can take a minute or two)")

        def work():
            try:
                text, warn = call_gemini(k, model_var.get().strip(), sp, raw, fallback_var.get().strip(),
                                         lambda m: root.after(0, lambda m=m: status.set(m)))
                lst = strip_fence(text)
                if do_reorder:
                    lst = reorder_for(used_job, lst)
                om, ad, no, na = compare(raw, lst)
                po, pa, flag = fidelity(raw, no, na)
                rows = parse_list(lst)
                bad = [r for r in rows if r[0] == "room" and job.get() == "normal" and not r[1].startswith("!")]
                def done():
                    state["job"] = used_job
                    t_list.delete("1.0", "end")
                    t_list.insert("1.0", lst)
                    t_rev.delete("1.0", "end")
                    t_rev.insert("1.0",
                        (f"CHECK THIS OUTPUT: about {po:.1f}% of your transcript words are missing and {pa:.1f}% were added. "
                         f"The model may have changed your transcript. Compare below, or try a stronger model.\n\n" if flag else
                         f"Looks faithful: {po:.1f}% of words missing, {pa:.1f}% added.\n\n") +
                        f"POSSIBLY OMITTED from transcript ({no} words; fillers/tags/self-corrections are expected):\n{om or 'none'}\n\n"
                        f"POSSIBLY ADDED in descriptions ({na} words; AU spelling and brand fixes are expected):\n{ad or 'none'}\n\n"
                        + (f"LINES WITHOUT A TAB OR '!' (check these):\n" + "\n".join(r[1] for r in bad) if bad else ""))
                    nb.select(t_list)
                    status.set(f"Done: {sum(1 for r in rows if r[0]=='item')} items. Review tabs 2 and 3, then Export.{warn}")
                    b_proc.state(["!disabled"])
                    if flag:
                        status.set(f"CHECK: {po:.1f}% words missing, {pa:.1f}% added. See tab 3.")
                        messagebox.showwarning("Check the output", f"About {po:.1f}% of your transcript words are missing and "
                                               f"{pa:.1f}% were added.\n\nThe model ({model_var.get().strip()}) may have changed "
                                               "your transcript. Review tab 3, or re-run with a stronger model.")
                root.after(0, done)
            except Exception as e:
                err = str(e)  # 'e' is deleted when this block ends, so keep a copy for the callback
                def fail():
                    status.set("Failed.")
                    b_proc.state(["!disabled"])
                    messagebox.showerror("Error", err)
                root.after(0, fail)
        threading.Thread(target=work, daemon=True).start()

    def export():
        lst = t_list.get("1.0", "end").strip("\n")
        if not lst.strip():
            return messagebox.showwarning("Export", "Nothing to export. Process a transcript first.")
        if not (want_word.get() or want_excel.get()):
            return messagebox.showwarning("Export", "Tick Word and/or Excel.")
        folder = os.path.expanduser(out_dir.get().strip())
        exts = [e for e, on in ((".docx", want_word.get()), (".xlsx", want_excel.get())) if on]
        if folder:
            try:
                os.makedirs(folder, exist_ok=True)
            except OSError as e:
                return messagebox.showerror("Export", f"Cannot use that folder: {e}")
            base = unique_base(folder, safe_name(file_base.get()), exts)
        else:
            p = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=safe_name(file_base.get()),
                                             filetypes=[("Base file name", "*.*")], title="Choose a file name (extension is added)")
            if not p:
                return
            base = os.path.splitext(p)[0]
        save_cfg()
        jb_used = state["job"] or job.get()
        data = run_macros(parse_list(lst), rules_state[jb_used], kw_state[jb_used])
        made = []
        try:
            if want_word.get():
                write_word(data, layout.get(), base + ".docx")
                made.append(base + ".docx")
            if want_excel.get():
                write_excel(data, layout.get(), base + ".xlsx")
                made.append(base + ".xlsx")
        except PermissionError:
            return messagebox.showerror("Export", "A file is open in Word/Excel. Close it and try again.")
        status.set("Exported: " + ", ".join(os.path.basename(m) for m in made))
        messagebox.showinfo("Export", "Saved:\n" + "\n".join(made))

    b_proc.config(command=process)
    b_exp.config(command=export)
    root.mainloop()


if __name__ == "__main__":
    run_gui()
