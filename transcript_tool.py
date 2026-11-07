"""
Transcript -> Inventory tool (local)

Setup:   pip install python-docx openpyxl requests pypdf
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
CONFIG_VERSION = 2
API_BASE = "https://generativelanguage.googleapis.com/v1beta"

# ------------------------------------------------------------------ macro data (from your .bas files)
# FormatText replacements: whole-word, case-insensitive, applied in this order after commas -> full stops.
TEXT_REPLACEMENTS = list(zip(
    ["points and switches", "slash", "windows and frames", "windowsill", "window covering",
     "window treatment", "benchtop", "benchtops", "colourbond", "colour bond", "intact in working order",
     "intact and working order", "down light", "glove", "gloves", "tower", "watermarks", "watermarked"],
    ["Points & Switches", "/", "Windows & Frames", "window sill", "Window Covering",
     "Window Treatment", "bench top", "bench top", "Colourbond", "Colourbond", "intact and in working order",
     "intact and in working order", "downlight", "globe", "globes", "towel", "water marks", "water marked"],
))
# Never allow a replacement rule to silently delete common spoken words.
_PROTECTED_WORD_REPLACEMENTS = {"is", "are", "a", "an", "there", "the", "to", "of", "in", "on", "with", "and", "or"}

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


def call_gemini(key, model, system_prompt, transcript):
    body = {
        "systemInstruction": {"parts": [{"text": system_prompt}]},
        "contents": [{"role": "user", "parts": [{"text": "RAW TRANSCRIPT:\n\n" + transcript}]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 65536},
    }
    r = requests.post(f"{API_BASE}/models/{model}:generateContent",
                      headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                      json=body, timeout=600)
    if r.status_code != 200:
        try:
            msg = r.json()["error"]["message"]
        except Exception:
            msg = r.text[:300]
        raise RuntimeError(f"Gemini error {r.status_code}: {msg}")
    data = r.json()
    cand = (data.get("candidates") or [{}])[0]
    text = "".join(p.get("text", "") for p in cand.get("content", {}).get("parts", []))
    if not text.strip():
        raise RuntimeError(f"Empty response (finishReason={cand.get('finishReason')}).")
    warn = "\nWARNING: output hit the token limit and may be truncated." if cand.get("finishReason") == "MAX_TOKENS" else ""
    return text, warn


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
    "bedroom": ["door", "cornices", "light", "walls", "aircon", "windows", "blinds", "power", "skirting", "floor_coverings", "wardrobe", "additional"],
    "garage": ["door", "external_door", "cornices", "light", "walls", "power", "floor_coverings", "auto_door", "additional"],
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
    if "bedroom" in r:
        return "bedroom"
    if "garage" in r:
        return "garage"
    if re.fullmatch(r"(toilet|wc|powder room)( \d+)?", r):
        return "toilet"
    return "generic"


def _classify(name):
    n = name.lower().strip()
    # Composite names must be classified by their principal item, not a secondary word.
    if "wardrobe" in n or re.search(r"\brobe\b", n):
        return "wardrobe"
    if "trough cabinet" in n:
        return "trough_cabinet"
    if "toilet roll" in n:
        return "toilet_roll"
    if "shower" in n and re.search(r"tapware|tap|mixer|showerhead|shower head", n):
        return "shower_tapware"
    if "shower" in n and "tile" in n:
        return "shower_tiles"
    if "shower" in n:
        return "shower_screen"
    if "bath" in n and re.search(r"tapware|tap|mixer|spout", n):
        return "bath_tapware"
    if "bath" in n and "tile" in n:
        return "bath_tiles"
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
        else:  # unknown item: stay straight after the item it followed
            keyed.append((last + 0.5, idx, ln))
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


# ------------------------------------------------------------------ macro 1: FormatText
def _ws_rx(find):
    # a leading/trailing space is already a word boundary, so only add lookarounds on the open ends
    pat = re.escape(find.rstrip()).replace(r"\ ", r"\s+") + (r"\s+" if find.endswith(" ") else "")
    return re.compile(("" if find.startswith(" ") else r"(?<![A-Za-z0-9])") + pat
                      + ("" if find.endswith(" ") else r"(?![A-Za-z0-9])"), re.I)


@lru_cache(maxsize=16)
def _compile_rules(rules):
    return [(_ws_rx(f), r) for f, r in rules if f.strip()]


def format_text(s, rules=None):
    """Apply non-destructive terminology fixes; never delete common words."""
    s = (s.replace("\u00a0", " ").replace("\u202f", " ").replace("\u201a", ",")
           .replace(",", ".").replace(";", "."))
    source_rules = TEXT_REPLACEMENTS if rules is None else rules
    safe_rules = []
    for find, replace in source_rules:
        find, replace = str(find), str(replace)
        key = find.strip().lower()
        if not replace.strip() and key in _PROTECTED_WORD_REPLACEMENTS:
            continue
        # Migrate legacy defaults that contradicted the Australian-English prompt.
        if key == "metre" and replace.strip().lower() == "meter":
            replace = "metre"
        if key in {"colourbond", "colour bond"} and replace.strip().lower() == "colorbond":
            replace = "Colourbond"
        safe_rules.append((find, replace))
    for rx, rep in _compile_rules(tuple(safe_rules)):
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
    """Remove only the leading Normal room marker; preserve punctuation elsewhere."""
    return s.strip()[1:].lstrip() if s.strip().startswith("!") else s.strip()


# ------------------------------------------------------------------ macro 3: ChangeYtoN
def _hit(text, kws):
    return "N" if any(k.strip() and k.strip().lower() in text for k in kws) else "Y"


def yn(desc):
    d = desc.lower()
    return _hit(d, CLEAN_KW), _hit(d, DAMAGE_KW), _hit(d, WORKING_KW)


def run_macros(rows, rules=None):
    """Apply FormatText -> FormatTable -> ChangeYtoN. -> list of dicts. rules = replacement list for the job."""
    out = []
    for r in rows:
        if r[0] == "room":
            out.append({"room": strip_bang(start_case(format_text(r[1], rules)))})
        else:
            name = strip_bang(start_case(format_text(r[1], rules)))
            desc = sentence_case(format_text(r[2], rules))
            c, u, w = yn(desc)
            out.append({"item": name, "desc": desc, "clean": c, "undamaged": u, "working": w})
    return out


# ------------------------------------------------------------------ omitted / added check
_TAGS = re.compile(r"<[^>]*>|\[[^\]]*\]|\b\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?\b")
_WORD = re.compile(r"[a-z0-9]+(?:'[a-z]+)?")
_FILLERS = {"um", "uh", "ah", "er", "erm", "hmm", "umm"}


def _tok(s):
    return Counter(w for w in _WORD.findall(s.lower()) if w not in _FILLERS)


def compare(raw, list_text):
    """Approximate word audit against descriptions only (not room/item labels)."""
    rows = parse_list(list_text)
    raw_t = _tok(_TAGS.sub(" ", raw))
    desc_out = _tok(" ".join(r[2] for r in rows if r[0] == "item"))
    omitted = raw_t - desc_out
    added = desc_out - raw_t
    fmt = lambda c: ", ".join(f"{w}" + (f" x{n}" if n > 1 else "") for w, n in sorted(c.items()))
    return fmt(omitted), fmt(added), sum(omitted.values()), sum(added.values())


def audit_structure(list_text, mode):
    """Return structural warnings without silently rewriting model output."""
    warnings = []
    lines = [ln for ln in list_text.replace("\r", "").split("\n") if ln.strip()]
    item_lines = [ln for ln in lines if "\t" in ln]
    header_lines = [ln for ln in lines if "\t" not in ln]
    if not item_lines:
        warnings.append("No tab-separated item rows were detected.")
    for number, line in enumerate(lines, 1):
        if "\t" in line:
            name, desc = line.split("\t", 1)
            if not name.strip():
                warnings.append(f"Line {number}: item name is blank.")
            if not desc.strip():
                warnings.append(f"Line {number}: description is blank.")
        elif mode == "normal" and not line.lstrip().startswith("!"):
            warnings.append(f"Line {number}: Normal room header is missing its leading ! marker: {line[:90]}")
    if len(header_lines) > max(1, len(item_lines)):
        warnings.append("There are many un-tabbed lines; check that descriptions have not been split into accidental room headers.")
    return warnings


# ------------------------------------------------------------------ layout
HEAD = {"blank": "", "item": "Item", "desc": "Description", "clean": "Clean", "undamaged": "Undamaged", "working": "Working"}
LAYOUTS = {
    "standard": (["item", "desc", "clean", "undamaged", "working"], [3.4, 8.3, 2.1, 2.1, 2.1]),
    "final": (["item", "clean", "undamaged", "working", "blank", "desc"], [3.4, 2.1, 2.1, 2.1, 2.2, 6.1]),
}
NAVY, ROOM_FILL, N_FILL, N_FONT, Y_FONT = "1F3864", "D9E2F3", "FDE9E7", "C00000", "2E7D32"
SPACE_PT = 4

# ------------------------------------------------------------------ Word
def _shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


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
def write_excel(data, layout, path):
    cols, widths = LAYOUTS[layout]
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory"
    thin = Side(style="thin", color="000000")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = lambda h: PatternFill("solid", start_color=h, end_color=h)
    f = lambda **k: Font(name="Calibri", size=10, **k)

    for i, (key, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(1, i, HEAD[key])
        c.font, c.fill, c.border = f(bold=True, color="FFFFFF"), fill(NAVY), box
        c.alignment = Alignment(horizontal="left" if key in ("item", "desc") else "center", vertical="center")
        ws.column_dimensions[c.column_letter].width = w * 4.6 if key == "desc" else w * 5.2

    r = 2
    for d in data:
        if "room" in d:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=len(cols))
            c = ws.cell(r, 1, d["room"])
            c.font, c.fill = f(bold=True, color=NAVY), fill(ROOM_FILL)
            c.alignment = Alignment(vertical="center")
            for i in range(1, len(cols) + 1):
                ws.cell(r, i).border = box
                ws.cell(r, i).fill = fill(ROOM_FILL)
        else:
            for i, key in enumerate(cols, 1):
                v = d.get(key, "")
                c = ws.cell(r, i, v)
                c.border = box
                if key == "blank":
                    pass
                elif key == "item":
                    c.font = f(bold=True)
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                elif key == "desc":
                    c.font = f()
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                else:
                    c.font = f(bold=True)
                    c.alignment = Alignment(horizontal="center", vertical="center")
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
    root.geometry("1100x780")

    key_var = tk.StringVar(value=cfg.get("api_key", ""))
    model_var = tk.StringVar(value=cfg.get("model", DEFAULT_MODEL))
    remember = tk.BooleanVar(value=bool(cfg.get("api_key")))
    job = tk.StringVar(value="normal")
    want_word, want_excel = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
    # Version 2 changes the default back to the five-column Standard layout.
    # Older config files stored "final" automatically because it used to be the default.
    saved_layout = cfg.get("layout", "standard") if cfg.get("config_version", 0) >= CONFIG_VERSION else "standard"
    if saved_layout not in LAYOUTS:
        saved_layout = "standard"
    layout = tk.StringVar(value=saved_layout)
    status = tk.StringVar(value="Ready.")
    ref_text = {"text": "", "name": ""}
    out_dir = tk.StringVar(value=cfg.get("out_dir", ""))
    file_base = tk.StringVar(value=cfg.get("file_base", "Inventory"))
    rules_state = load_rules()
    state = {"job": None}

    top = ttk.Frame(root, padding=8)
    top.pack(fill="x")
    ttk.Label(top, text="Gemini API key").grid(row=0, column=0, sticky="w")
    ttk.Entry(top, textvariable=key_var, show="*", width=48).grid(row=0, column=1, padx=6)
    ttk.Label(top, text="Model").grid(row=0, column=2, sticky="w")
    model_box = ttk.Combobox(top, textvariable=model_var, width=26,
                             values=["gemini-3.8-flash", "gemini-3.6-flash", "gemini-3.1-pro-preview"])
    model_box.grid(row=0, column=3, padx=6)
    ttk.Checkbutton(top, text="Remember key on this PC", variable=remember).grid(row=0, column=4, padx=6)

    def save_cfg():
        data = {"config_version": CONFIG_VERSION, "model": model_var.get().strip(), "out_dir": out_dir.get().strip(), "file_base": file_base.get().strip(),
                "layout": layout.get()}
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
        status.set("Testing key...")
        def work():
            ok, msg = test_key(k, model_var.get().strip())
            root.after(0, lambda: (status.set(msg), (messagebox.showinfo if ok else messagebox.showerror)("API key test", msg)))
        threading.Thread(target=work, daemon=True).start()
    ttk.Button(top, text="Test key", command=do_test).grid(row=0, column=5)

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
    reorder = tk.BooleanVar(value=True)
    ttk.Checkbutton(jb, text="Pearl: reorder items to reference template", variable=reorder).pack(side="left", padx=16)
    ref_lbl = ttk.Label(jb, text="(optional) Pearl Property Ref: none")
    def pick_ref():
        p = filedialog.askopenfilename(filetypes=[("Reference", "*.txt *.md *.docx *.pdf")])
        if not p:
            return
        ref_text["text"] = read_reference(p, Path(p).read_bytes())
        ref_text["name"] = os.path.basename(p)
        ref_lbl.config(text=f"Pearl Property Ref: {ref_text['name']}")
    ttk.Button(jb, text="Load Pearl ref...", command=pick_ref).pack(side="left", padx=16)

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

        btns = ttk.Frame(win, padding=8)
        btns.pack(fill="x")
        ttk.Button(btns, text="Add / update", command=add_update).pack(side="left")
        ttk.Button(btns, text="Delete selected", command=delete).pack(side="left", padx=8)
        ttk.Button(btns, text="Reset to default", command=reset).pack(side="left")
        ttk.Button(btns, text="Close", command=win.destroy).pack(side="right")
        ttk.Label(win, wraplength=660, justify="left", foreground="#555",
                  text='Tip: spaces matter. A find of " is " (a space either side) removes the word "is". '
                       'Leave Replace empty to delete the word. Select a row to edit it, then click Add / update. '
                       'Changes apply to the next export and are saved in replacements.json.').pack(fill="x", padx=8, pady=(0, 8))
        refresh()
    ttk.Button(jb, text="Replacement words...", command=open_rules).pack(side="left", padx=4)
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
        p = filedialog.askopenfilename(filetypes=[("Text", "*.txt *.vtt *.srt *.md")])
        if p:
            t_raw.delete("1.0", "end")
            t_raw.insert("1.0", Path(p).read_text(encoding="utf-8", errors="ignore"))
    ttk.Button(rb, text="Load file...", command=load_txt).pack(side="left")

    bot = ttk.LabelFrame(root, text="Output", padding=6)
    bot.pack(fill="x", padx=8, pady=4)
    ttk.Checkbutton(bot, text="Word (.docx)", variable=want_word).pack(side="left", padx=6)
    ttk.Checkbutton(bot, text="Excel (.xlsx)", variable=want_excel).pack(side="left", padx=6)
    ttk.Radiobutton(bot, text="Standard: Item | Description | Clean | Undamaged | Working", variable=layout, value="standard").pack(side="left", padx=14)
    ttk.Radiobutton(bot, text="Final report: Item | Clean | Undamaged | Working | blank column | Description", variable=layout, value="final").pack(side="left")
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
        if not k:
            return messagebox.showwarning("API key", "Enter a Gemini API key (or set GEMINI_API_KEY).")
        save_cfg()
        sp = prompts.NORMAL if job.get() == "normal" else prompts.PEARL
        if job.get() == "pearl" and ref_text["name"]:
            sp += ("\n\nA Pearl reference PDF has been selected. Use it only as item terminology/order guidance. "
                   "Do not copy any property facts from it. The application's Pearl ordering layer applies the reference item order after generation.")
        sp += prompts.FIDELITY
        used_job = job.get()
        do_reorder = used_job == "pearl" and reorder.get()
        b_proc.state(["disabled"])
        status.set("Processing with Gemini... (large transcripts can take a minute or two)")

        def work():
            try:
                text, warn = call_gemini(k, model_var.get().strip(), sp, raw)
                lst = strip_fence(text)
                if do_reorder:
                    lst = reorder_pearl(lst)
                om, ad, no, na = compare(raw, lst)
                rows = parse_list(lst)
                structure_warnings = audit_structure(lst, used_job)
                def done():
                    state["job"] = used_job
                    t_list.delete("1.0", "end")
                    t_list.insert("1.0", lst)
                    t_rev.delete("1.0", "end")
                    t_rev.insert("1.0",
                        f"POSSIBLY OMITTED from transcript ({no} words; fillers/tags/self-corrections are expected):\n{om or 'none'}\n\n"
                        f"POSSIBLY ADDED in descriptions ({na} words; AU spelling and brand fixes are expected):\n{ad or 'none'}\n\n"
                        + ("\n\nSTRUCTURE WARNINGS (review before export):\n" + "\n".join("- " + w for w in structure_warnings) if structure_warnings else "\n\nStructure check: no obvious formatting issues detected."))
                    nb.select(t_list)
                    status.set(f"Done: {sum(1 for r in rows if r[0]=='item')} items. Review tabs 2 and 3, then Export.{warn}")
                    b_proc.state(["!disabled"])
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
        data = run_macros(parse_list(lst), rules_state[state["job"] or job.get()])
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
