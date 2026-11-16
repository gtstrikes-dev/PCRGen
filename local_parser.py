"""Local rules mode (no AI / no API).

Reorders and organises a raw dictation transcript into the same list format the Gemini step produces:
    !Room
    Item<TAB>Description
It only cuts, relabels and re-orders the speaker's own words. Anything it cannot resolve is left in the text
and listed in the review list that local_parse() returns.

Vocabulary lives in items.json (next to this file) and can be edited: each item has a spoken `match` regex,
the name used for Normal jobs and for Pearl jobs, and a few flags. Delete items.json to restore the defaults.
"""
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
VOCAB_FILE = os.path.join(HERE, "items.json")

NUMS = "one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|\\d+"
NUMW = {w: i for i, w in enumerate("zero one two three four five six seven eight nine ten eleven twelve".split())}

# ------------------------------------------------------------------ default vocabulary
# key, spoken regex, Normal name(s), Pearl name(s), flags
#   normal/pearl: list of names. More than one name = the spoken heading is split into several items (later ones
#   start when the speaker begins describing them, e.g. "Ceiling is painted white").
#   container: swallows every following sentence until the next room/container (Robe, Pantry, Linen cupboard).
#   implicit: the item can also start from a sentence such as "Walls are painted ..." with no heading.
#   tap: tapware that is merged into the previous sink/basin/shower/bath item for Normal jobs.
DEFAULT_ITEMS = [
    {"key": "security", "match": "(?:entry )?security screens?(?: door)?", "normal": ["Security Screen"], "pearl": ["Security Screen/Frame/Handles"]},
    {"key": "ext_door", "match": "(?:external|outside|front|back|rear|personnel) doors?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "sliding", "match": "(?:external )?sliding doors?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "doors", "match": "(?:entry |bathroom |internal )?doors? and (?:the )?(?:door |metal |wooden )?frames?", "normal": ["Doors & Frames"], "pearl": ["Doors & Frames"], "lead": True},
    {"key": "doors", "match": "(?:entry |garage roller |roller )?doors?", "normal": ["Doors & Frames"], "pearl": ["Doors & Frames"]},
    {"key": "door_treat", "match": "door treatments?", "normal": ["Door Treatments"], "pearl": ["Door Treatments"]},
    {"key": "floor", "match": "floor coverings?|flooring", "normal": ["Floor Coverings"], "pearl": ["Floor Coverings"], "implicit": True},
    {"key": "skirting", "match": "skirting(?: boards?)?", "normal": ["Skirting"], "pearl": ["Skirting"], "implicit": True, "lead": True},
    {"key": "walls", "match": "(?:balls|walls?)", "normal": ["Walls"], "pearl": ["Walls"], "implicit": True},
    {"key": "ext_walls", "match": "external walls?", "normal": ["External Walls"], "pearl": ["External Walls"], "implicit": True},
    {"key": "cornice_ceiling", "match": "cornices? (?:and|within|in) ceilings?", "normal": ["Cornice", "Ceiling"], "pearl": ["Cornices/Ceiling"], "parts": ["cornice", "ceiling"]},
    {"key": "cornice", "match": "cornices?", "normal": ["Cornice"], "pearl": ["Cornices/Ceiling"], "implicit": True},
    {"key": "ceiling", "match": "ceilings?", "normal": ["Ceiling"], "pearl": ["Cornices/Ceiling"], "implicit": True},
    {"key": "light_fan", "match": "light fittings? and fans?", "normal": ["Light Fitting & Fan"], "pearl": ["Light Fitting & Fan"], "also": ["light"]},
    {"key": "light", "match": "light (?:fittings?|fixtures?)|lights?", "normal": ["Light Fittings"], "pearl": ["Light Fittings"], "nolead": True},
    {"key": "points", "match": "(?:power )?(?:points?|pointed|pointing) and (?:light )?switch(?:es)?|(?:pointed|pointing) switches", "normal": ["Points & Switches"], "pearl": ["Power Points/Light Switches"]},
    {"key": "smoke", "match": "smoke (?:detectors?|alarms?)", "normal": ["Smoke Detector"], "pearl": ["Smoke Detector"]},
    {"key": "doorbell", "match": "door ?bells?", "normal": ["Doorbell"], "pearl": ["Doorbell"]},
    {"key": "windows", "match": "windows? and (?:metal |wooden )?frames?", "normal": ["Windows & Frames"], "pearl": ["Windows/Screens"], "also": ["windows2"]},
    {"key": "windows2", "match": "windows?(?: and (?:fly ?)?screens?)?", "normal": ["Windows & Frames"], "pearl": ["Windows/Screens"], "also": ["windows"]},
    {"key": "window_treat", "match": "window (?:treatments?|trimmers?)|blinds?(?: and curtains?)?|curtains?", "normal": ["Window Treatments"], "pearl": ["Blinds/Curtains"]},
    {"key": "aircon", "match": "air ?conditioning|air ?conditioner|split system", "normal": ["Air Conditioning"], "pearl": ["Air Conditioning"]},
    {"key": "fan", "match": "ceiling fans?|exhaust fans?|heat lights?", "normal": ["{spoken}"], "pearl": ["{spoken}"], "lead": True},
    {"key": "vent", "match": "vents?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "skylight", "match": "skylights?", "normal": ["Skylight"], "pearl": ["Skylight"]},
    {"key": "bench", "match": "bench ?tops?", "normal": ["Bench Top"], "pearl": ["Bench Top"]},
    {"key": "splash", "match": "splash ?backs?(?: tiles?)?", "normal": ["Splashback"], "pearl": ["Splashback"]},
    {"key": "sink", "match": "sinks?", "normal": ["Sink"], "pearl": ["Sink"], "vessel": "Kitchen"},
    {"key": "basin", "match": "(?:wash )?basins?", "normal": ["Basin"], "pearl": ["Basin"], "vessel": "Basin"},
    {"key": "shower_screen", "match": "shower (?:screens?|curtains?)", "normal": ["Shower Screen"], "pearl": ["Shower Screen/Curtain"], "vessel": "Shower"},
    {"key": "shower_tiles", "match": "shower tiles?", "normal": ["Shower Tiles"], "pearl": ["Shower Tiles"], "vessel": "Shower"},
    {"key": "shower", "match": "showers?(?: screen and taps?| and taps?)?", "normal": ["Shower"], "pearl": ["Shower Screen/Curtain"], "vessel": "Shower"},
    {"key": "bath", "match": "bath(?:tub)?", "normal": ["Bath"], "pearl": ["Bath"], "vessel": "Bath"},
    {"key": "bath_tiles", "match": "bath tiles?", "normal": ["Bath Tiles"], "pearl": ["Bath Tiles"]},
    {"key": "tapware", "match": "(?:bath |basin |shower |kitchen |sink )?(?:tap ?wares?|tap wire)|taps?", "normal": ["Tapware"], "pearl": ["Tapware"], "tap": True},
    {"key": "wm_taps", "match": "wash(?:ing)? machine taps?", "normal": ["Washing Machine Taps"], "pearl": ["Washing Machine Taps"]},
    {"key": "trough", "match": "(?:laundry )?troughs?|laundry tub", "normal": ["Trough"], "pearl": ["Trough"]},
    {"key": "toilet_roll", "match": "(?:toilet )?(?:roll )?holders?|toilet roll", "normal": ["Toilet Roll Holder"], "pearl": ["Toilet Roll Holder"]},
    {"key": "toilet", "match": "toilet(?: suite)?|cistern seat and lid|toilet seat", "normal": ["Toilet"], "pearl": ["Toilet"]},
    {"key": "mirror", "match": "mirrors?", "normal": ["Mirror"], "pearl": ["Mirror"]},
    {"key": "towel", "match": "(?:towel|tower) (?:rails?|rings?|holders?)", "normal": ["Towel Rail"], "pearl": ["Towel Rail"]},
    {"key": "vanity", "match": "vanity(?: unit)?|vanity cupboards?", "normal": ["Vanity"], "pearl": ["Vanity"]},
    {"key": "cupboards", "match": "(?:under-?bench |overhead |wall |base )cupboards?|cupboards?|drawers?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "dishwasher", "match": "dishwashers?", "normal": ["Dishwasher"], "pearl": ["Dishwasher"]},
    {"key": "hood", "match": "range ?hoods?|rangehood", "normal": ["Range Hood"], "pearl": ["Exhaust Fan/Range Hood"]},
    {"key": "stove", "match": "stove ?tops?|hot ?plates?|cook ?tops?", "normal": ["Stove Top"], "pearl": ["Stove Top/Hot Plates"]},
    {"key": "oven", "match": "ovens?(?: and griller)?", "normal": ["Oven"], "pearl": ["Oven/Griller"]},
    {"key": "microwave", "match": "microwave(?: recess)?", "normal": ["Microwave Recess"], "pearl": ["Microwave Recess"]},
    {"key": "fridge", "match": "fridge(?: recess)?|refrigerator", "normal": ["Fridge"], "pearl": ["Fridge"]},
    {"key": "pantry", "match": "pantry", "normal": ["Pantry"], "pearl": ["Pantry"], "container": True},
    {"key": "linen", "match": "linen (?:cupboard|press)(?: (?:one|two|three|four|\\d+))?", "normal": ["{spoken}"], "pearl": ["{spoken}"], "container": True},
    {"key": "robe", "match": "(?:built-?in )?(?:robes?|rogue|rope|wardrobe)(?: (?:one|two|three|\\d+))?", "normal": ["Robe"], "pearl": ["Robe"], "container": True},
    {"key": "desk", "match": "desks?", "normal": ["Desk"], "pearl": ["Desk"]},
    {"key": "roof", "match": "roof(?:ing)?", "normal": ["Roof"], "pearl": ["Roof"], "lead": True},
    {"key": "eaves", "match": "eaves|fascia(?: and eaves)?|fascia", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "gutters", "match": "gutters?(?: and downpipes?)?|downpipes?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "fence", "match": "fenc(?:es|ing|e)|gates?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "paving", "match": "path slash paving|paving|pavers?|pathway|driveway|concrete|slab", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "lawn", "match": "lawns?|artificial lawn", "normal": ["Lawn"], "pearl": ["Lawn"]},
    {"key": "garden", "match": "garden beds?|gardens?", "normal": ["Garden Beds"], "pearl": ["Garden Beds"]},
    {"key": "reticulation", "match": "reticulation|sprinklers?", "normal": ["Reticulation"], "pearl": ["Reticulation"]},
    {"key": "hoses", "match": "taps and hoses|hose taps?", "normal": ["Taps & Hoses"], "pearl": ["Taps & Hoses"]},
    {"key": "clothesline", "match": "clothes ?lines?", "normal": ["Clothesline"], "pearl": ["Clothesline"]},
    {"key": "letterbox", "match": "letter ?box(?:es)?", "normal": ["Letterbox"], "pearl": ["Letterbox"]},
    {"key": "bins", "match": "bins?|garbage bins?", "normal": ["Bins"], "pearl": ["Bins"]},
    {"key": "hotwater", "match": "hot water (?:system|unit|service)", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "meter", "match": "(?:electric(?:ity)?|gas|water) (?:meter|metre)s?(?: box)?", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "cistern", "match": "cisterns?", "normal": ["Cistern"], "pearl": ["Cistern"]},
    {"key": "outdoor", "match": "pergola|barbecue area|garden shed|shed|carport roof", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "stairs", "match": "stairwell(?: and pathway)?|stairs|steps", "normal": ["{spoken}"], "pearl": ["{spoken}"]},
    {"key": "electrical", "match": "electrical", "normal": ["Electrical"], "pearl": ["Electrical"]},
    {"key": "exhaust", "match": "exhausts?", "normal": ["Exhaust Fan"], "pearl": ["Exhaust Fan/Range Hood"]},
    {"key": "flyscreen", "match": "(?:internal )?fly ?screens?", "normal": ["Fly Screens"], "pearl": ["Windows/Screens"]},
    {"key": "extdoors", "match": "external doors?", "normal": ["External Doors"], "pearl": ["External Doors"]},
    {"key": "wheelie", "match": "wheelie bins?", "normal": ["Wheelie Bins"], "pearl": ["Wheelie Bins"]},
    {"key": "addinfo", "match": "additional information", "normal": ["Additional"], "pearl": ["Additional"]},
    {"key": "additional", "match": "additional", "normal": ["Additional"], "pearl": ["Additional"]},
]

ROOM_RX = [  # (regex, canonical key (None = spoken lower-case), display name or None = Title Case of spoken)
    (r"front of (?:the )?property|front property(?: and carport)?|front yard", "front", None),
    (r"rear of (?:the )?property|rear property|back ?yard|rear yard", "rear", None),
    (r"(?:master )?bed ?room(?: (?:number )?(?:one|two|three|four|five|six|seven|eight|nine|ten|\d+))?", None, None),
    (r"en-?suite(?: bathroom)?", "ensuite", None),
    (r"(?:main )?bathroom", None, None), (r"toilet(?: (?:one|two|\d+))?", None, None), (r"powder room", None, None),
    (r"kitchen", None, None), (r"laundry", None, None), (r"hall ?way|hall", None, None), (r"(?:sunken )?lounge(?: room)?", None, None),
    (r"(?:living|family|dining|rumpus|meals?) (?:room|area)", None, None), (r"living", None, None),
    (r"entry(?: hall)?|entrance|study|office|garage|carport|alfresco(?: area)?|patio|balcony|verandah|porch|pool area|storeroom|store room|"
     r"stairs|landing|cellar|theatre room|nursery|sunroom", None, None),
]
_ROOMS = [(re.compile(r"(?:" + rx + r")", re.I), k, d) for rx, k, d in ROOM_RX]
_ROOMANY = "|".join("(?:" + rx + ")" for rx, _, _ in ROOM_RX)

FILLERS = re.compile(r"\b(?:um+|uh+|er+m?|ah+|hmm+|mm+)\b[,.]?\s*", re.I)
LEAD_OK = re.compile(r"(?:painted|unpainted|stained|exposed|clean|dusty|dirty|intact|secured?|frameless|framed|damaged|worn|cracked|"
                     r"white|grey|gray|cream|black|brown|blue|green|red|pink|yellow|bronze|gold|silver|beige|have|has|with|in|"
                     r"minor|heavy|small|large|no|nil|a|an|" + NUMS + r")\b", re.I)
PART_WORDS = {"frame", "frames", "handle", "handles", "stopper", "stoppers", "stop", "stops", "hook", "hooks", "closer",
              "mesh", "screen", "screens", "track", "tracks", "rail", "rails", "latch", "lock", "hinge", "hinges", "panel",
              "panels", "bracket", "surround", "base", "seat", "cover", "light", "lights", "switch", "switches", "point",
              "points", "tile", "tiles", "grout", "drain", "drains", "plug", "tap", "taps", "spout", "door", "doors"}


# ------------------------------------------------------------------ vocabulary i/o
def load_vocab():
    if os.path.exists(VOCAB_FILE):
        try:
            data = json.load(open(VOCAB_FILE, encoding="utf-8"))
            if isinstance(data, list) and data:
                return data
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_ITEMS))


def save_vocab(items):
    with open(VOCAB_FILE, "w", encoding="utf-8") as f:
        json.dump(items, f, indent=1, ensure_ascii=False)


def _compile(items):
    comp = []
    for it in items:
        rx = re.compile(r"(?:" + it["match"] + r")", re.I)
        comp.append((it, rx))
    # longest/most specific first: sort by length of the pattern text (compound headings before simple ones)
    comp.sort(key=lambda x: -len(x[0]["match"]))
    return comp


# ------------------------------------------------------------------ text helpers
def clean_raw(raw):
    t = raw.replace("\r", "")
    t = re.sub(r"</?v[^>]*>", " ", t)
    t = re.sub(r"\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?(?:\s*-->\s*\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?)?", " ", t)
    t = re.sub(r"\[\d+\]", " ", t)
    t = re.sub(r"^WEBVTT.*$", " ", t, flags=re.M)
    t = FILLERS.sub(" ", t)
    t = t.replace("’", "'").replace("“", '"').replace("”", '"')
    return re.sub(r"\s+", " ", t).strip()


def split_sentences(t):
    parts = re.split(r"(?<=[.?!])\s+(?=[A-Z0-9\"'(])", t)
    return [p.strip() for p in parts if p.strip()]


def _stem(w):
    if re.search(r"(?:ch|sh|ss|x)es$", w):
        return w[:-2]
    return w[:-1] if w.endswith("s") and not w.endswith("ss") else w


def cap(s):
    s = s.strip()
    return s[:1].upper() + s[1:] if s else s


def title(s):
    return re.sub(r"(^|[\s/\-(])([a-z])", lambda m: m.group(1) + m.group(2).upper(), s.lower())


def end_dot(s):
    s = s.strip().rstrip(".,;:!? ")
    return (cap(s) + ".") if s else ""


def commas_to_stops(s):
    """Commas / semicolons -> full stops (not inside numbers), capitalise what follows."""
    s = re.sub(r"(?<!\d)\s*[,;]\s*(?!\d)", ". ", s)
    s = re.sub(r"\.\s*\.", ".", s)
    s = re.sub(r"(\.\s+)([a-z])", lambda m: m.group(1) + m.group(2).upper(), s)
    return s


def room_match(text, full=True):
    t = text.strip().rstrip(".,?! ")
    for rx, k, d in _ROOMS:
        m = rx.fullmatch(t) if full else rx.match(t)
        if m:
            spoken = m.group(0)
            key = k or re.sub(r"\s+", " ", spoken.lower())
            key = re.sub(r"^master ", "", key) if key.startswith("master bed") and False else key
            name = d or title(re.sub(r"\s+", " ", spoken))
            if k == "front":
                name = "Front Property" + (" And Carport" if "carport" in spoken.lower() else "")
            if k == "rear":
                name = "Rear Of Property"
            return key, name, m.end()
    return None


class Parser:
    def __init__(self, job="normal", vocab=None):
        self.job = job
        self.items = vocab or load_vocab()
        self.vc = _compile(self.items)
        self.byKey = {i["key"]: i for i in self.items}
        self.rooms = []          # [{key, name, items:[...]}]
        self.review = []
        self.room = None
        self.cur = None          # current item dict
        self.target = None       # add-mode target item (or None)
        self.revisit = False     # inside a room that already existed -> mark with #
        self.mark_next = False
        self.last_vessel = None
        self.area_pending = False   # "keep this area under <room>" heard; next heading becomes a single container item

    # ---- vocabulary lookups
    def find_item(self, text, full=True):
        """-> (entry, spoken_text, end_pos) for a heading phrase at the start of text."""
        t = text.strip()
        t = re.sub(r"^(?:the )", "", t, flags=re.I)
        for it, rx in self.vc:
            m = rx.fullmatch(t.rstrip(".,?! ")) if full else rx.match(t)
            if m and (full or re.match(r"\W|$", t[m.end():m.end() + 1] or " ")):
                return it, m.group(0), m.end()
        return None

    def names_for(self, it, spoken):
        names = it["pearl"] if self.job == "pearl" else it["normal"]
        sp = re.sub(r"\s+slash\s+", "/", spoken, flags=re.I)
        sp = re.sub(r"\s+and\s+", "/" if self.job == "pearl" else " & ", sp, flags=re.I)
        return [title(sp) if n == "{spoken}" else n for n in names]

    def flag(self, kind, text, note=""):
        self.review.append({"kind": kind, "text": text, "note": note, "room": self.room["name"] if self.room else ""})

    # ---- structure
    def open_room(self, key, name):
        for r in self.rooms:
            if r["key"] == key:
                self.room, self.revisit = r, True
                self.cur = r["items"][-1] if r["items"] else None
                self.target = None
                return
        self.room = {"key": key, "name": name, "items": []}
        self.rooms.append(self.room)
        self.revisit = False
        self.cur = self.target = None

    def new_item(self, it, spoken, mark=False):
        if self.room is None:
            self.open_room("unsorted", "Unsorted")
            self.flag("no room", spoken, "Item heard before any room heading; placed under 'Unsorted'")
        names = self.names_for(it, spoken)
        prev = self.cur
        # Normal jobs: taps belong with the sink/shower/bath/basin just described
        if it.get("tap") and self.job != "pearl" and prev and prev.get("vessel"):
            prev["tapped"] = True
            self.cur = prev
            self.target = None
            return prev
        name = names[0]
        if it.get("tap") and self.job == "pearl" and prev and prev.get("vessel"):
            name = prev["vessel"] + " Tapware"
        item = {"key": it["key"], "name": name, "entry": it, "spoken": spoken, "sents": [], "pending": names[1:],
                "pend_keys": it.get("parts", [])[1:], "container": bool(it.get("container")),
                "vessel": it.get("vessel"), "combined": len(names) == 1 and len(it.get("parts", [])) > 1,
                "covers": set(it.get("parts", [it["key"]])) | set(it.get("also", [])),
                "drop": (self.byKey.get(it["parts"][0], it)["match"] if it.get("parts") and self.job != "pearl" else it["match"])}
        if self.job == "pearl" and it["key"] in ("cornice", "ceiling", "cornice_ceiling"):
            item["covers"] = {"cornice", "ceiling", "cornice_ceiling"}
        self.room["items"].append(item)
        self.cur = item
        self.target = None
        return item

    def add_sentence(self, item, s, marked=False):
        s = commas_to_stops(s)
        s = self.drop_name(item, s)
        s = end_dot(s)
        if not s:
            return
        if marked or self.revisit or self.mark_next:
            s = "#" + s
        self.mark_next = False
        item["sents"].append(s)

    def drop_name(self, item, s):
        if item["entry"].get("parts") and self.job == "pearl":
            return s
        if item["key"] in ("ext_door", "doors") and False:
            return s
        rx = re.compile(r"^(?:the )?(?:" + item.get("drop", item["entry"]["match"]) + r")\b(?:\s+(?:is|are|was|were))?[,\s]*", re.I)
        m = rx.match(s)
        if not m:
            return s
        rest = s[m.end():]
        if not rest.strip():
            return s
        nxt = re.match(r"[A-Za-z]+", rest)
        had_verb = re.search(r"\b(?:is|are|was|were)\b", s[:m.end()], re.I)
        if nxt and nxt.group(0).lower() in PART_WORDS and not had_verb:
            return s
        return rest

    # ---- sentence classification
    def heading_of(self, s):
        """If sentence starts with a heading, return (kind, payload)."""
        t = s.strip()
        core = t.rstrip(".?! ")
        rm = room_match(core)
        if rm:
            return ("room", rm, "")
        # room + item prefix ("Kitchen floor coverings", "Hallway, doors and frames")
        for rx, k, d in _ROOMS:
            m = rx.match(core)
            if m and m.end() < len(core):
                rest = core[m.end():].lstrip(" ,")
                if not rest or not re.match(r"[A-Za-z]", rest):
                    continue
                fi = self.find_item(rest)
                desc = ""
                if not fi:
                    for it, rx2 in self.vc:
                        m2 = rx2.match(rest)
                        if m2 and m2.end() < len(rest) and rest[m2.end()] == " " and not it.get("nolead") and not it.get("container") \
                                and (it.get("lead") or LEAD_OK.match(rest[m2.end():].strip())) and not re.match(r"(?:is|are|was)\b", rest[m2.end():].strip(), re.I):
                            fi, desc = (it, m2.group(0), m2.end()), rest[m2.end():].strip()
                            break
                if fi:
                    rm2 = room_match(m.group(0))
                    return ("roomitem", rm2, fi, desc)
        fi = self.find_item(core)
        if fi:
            return ("item", fi, "")
        for it, rx in self.vc:
            if it.get("lead"):
                m = rx.match(core)
                if m and m.end() < len(core) and core[m.end()] == " " and not re.match(r"\s*(?:is|are|was)\b", core[m.end():], re.I):
                    desc = core[m.end():].strip()
                    if it["key"] == "doors":
                        desc = re.sub(r"^doors?\s+and\s+", "", core, flags=re.I)
                    return ("item", (it, m.group(0), m.end()), desc)
        for it, rx in self.vc:  # heading directly followed by a description ("Walls painted grey")
            m = rx.match(core)
            if m and m.end() < len(core) and core[m.end()] == " " and not it.get("container") and not it.get("nolead"):
                rest = core[m.end():].strip()
                if LEAD_OK.match(rest) and not re.match(r"(?:is|are|was)\b", rest, re.I):
                    return ("item", (it, m.group(0), m.end()), re.sub(r"^(?:have|has)\s+", "", rest, flags=re.I), "lead")
        # Heading, description
        if "," in core:
            head, rest = core.split(",", 1)
            fi = self.find_item(head)
            if fi and len(head.split()) <= 5:
                return ("item", fi, rest.strip())
            rm = room_match(head)
            if rm and rest.strip():
                fi2 = self.find_item(rest.strip())
                if fi2:
                    return ("roomitem", rm, fi2, "")
        # Heading + externally/internally
        m = re.match(r"(.+?)\s+(externally|internally)$", core, re.I)
        if m:
            fi = self.find_item(m.group(1))
            if fi:
                return ("item", fi, m.group(2))
        return None

    def implicit_of(self, s):
        m = re.match(r"(?:the )?(.+?)\s+(?:is|are)\b", s.strip(), re.I)
        if not m:
            return None
        for it, rx in self.vc:
            if it.get("implicit") and rx.fullmatch(m.group(1).strip()):
                return it, m.group(1).strip()
        return None

    # ---- corrections
    def inline_fix(self, s):
        """Handle 'X, sorry, Y' style self-corrections inside one sentence. Returns (sentence, changed)."""
        t = s
        m = re.match(r"^(\w+),?\s+no,?\s+((?:%s)\b.*)$" % NUMS, t, re.I)
        if m and re.fullmatch(NUMS, m.group(1), re.I):
            return cap(m.group(2)), True
        rx = re.compile(r"[,\s]*\b(?:(?:no|nope)[,\s]+)?sorry\b[,\s]*(?:(?:no|nope)\b[,\s]*)?", re.I)
        changed = False
        for _ in range(4):
            m = rx.search(t)
            if not m:
                break
            before, after = t[:m.start()], t[m.end():]
            if not before.strip() or not after.strip().rstrip(".?!"):
                t = (before + " " + after).strip()
                changed = True
                continue
            fw = re.match(r"[A-Za-z0-9'-]+", after.strip())
            fw = fw.group(0).lower() if fw else ""
            low = before.lower()
            cut = None
            for mm in re.finditer(r"\b" + re.escape(fw) + r"\b", low):
                cut = mm.start()
            if cut is not None and fw:
                t = before[:cut].rstrip(" ,") + " " + after.strip()
                t = t.strip()
            elif len(before.split()) <= 4:
                t = after.strip()
                self.flag("self-correction", s, "Replaced the words before 'sorry' with what followed — please check")
            else:
                self.flag("self-correction", s, "Could not tell what to cut — left as spoken")
                return s, False
            changed = True
        return t.strip(), changed

    def similar(self, a, b):
        wa = {w.rstrip("s") for w in re.findall(r"[a-z]+", a.lower())} - {"one", "two", "three", "four", "five", "six", "the", "a", "of", "to", "and", "with"}
        wb = {w.rstrip("s") for w in re.findall(r"[a-z]+", b.lower())} - {"one", "two", "three", "four", "five", "six", "the", "a", "of", "to", "and", "with"}
        return bool(wa & wb)

    def replace_similar(self, item, new, why):
        words = [w for w in re.findall(r"[a-z]+", new.lower()) if w not in ("one", "two", "three", "four", "five", "six", "the", "a", "with", "and")]
        key = [_stem(words[-1])] if words else []
        for i in range(len(item["sents"]) - 1, -1, -1):
            old = item["sents"][i].lstrip("#")
            if any(k and k in [_stem(w) for w in re.findall(r"[a-z]+", old.lower())] for k in key):
                s2 = "#" + end_dot(commas_to_stops(new))
                item["sents"][i] = s2
                self.flag("replaced", why, f"'{old}' -> '{s2.lstrip('#')}' (marked #)")
                return True
        return False

    def handle_command(self, s):
        low = s.lower()
        if not re.search(r"\b(?:go back|going back|head back|back to|thinking about|add on|could you|change)\b", low):
            return False
        t = re.sub(r"^(?:(?:sorry|so|oh|okay|ok)[,\s]+)+", "", s.strip(), flags=re.I)
        t = re.sub(r"^if you'?re thinking about\s+", "go back to ", t, flags=re.I)
        t = re.sub(r"^(?:and\s+)?(?:could you\s+)?(?:also\s+)?(?:please\s+)?", "", t, flags=re.I)
        # change that to X
        m = re.match(r"(?:could you\s+)?(?:please\s+)?change\s+(?:that|it)?\s*to\s+(.+)$", t.rstrip(".?! "), re.I)
        if m and not re.search(r"\band\s+with\b|\bchange it to\b", m.group(1), re.I):
            cands = ([self.cur] if self.cur else []) + [i for i in reversed(self.room["items"] if self.room else []) if i is not self.cur]
            for tgt in cands:
                if self.replace_similar(tgt, m.group(1), s):
                    return True
            self.flag("unresolved correction", s, "Could not find the earlier line to change — left in place")
            return False
        if re.match(r"(?:could you\s+)?(?:please\s+)?change\b", t, re.I):
            self.flag("unresolved correction", s, "Complex correction — left in place, please fix by hand")
            return False
        nav = re.match(r"(?:(?:go|head|going) back|back)\s+to\s+(?:the\s+)?(.*)$", t.rstrip("?!. "), re.I)
        rest = nav.group(1) if nav else t.rstrip("?!. ")
        if not nav and not re.match(r"(?:could you\s+)?(?:also\s+)?(?:please\s+)?add\b", rest, re.I) and not re.search(r"\badd on\b", rest, re.I):
            return False
        splitter = re.compile(r"\b(?:and\s+)?(?:could you\s+)?(?:also\s+)?(?:please\s+)?add(?:\s+on)?\b[,\s]*", re.I)
        pieces = [p.strip(" ,") for p in splitter.split(rest)]
        first_is_target = bool(nav)
        if not nav and pieces and pieces[0] == "":
            pieces = pieces[1:]
        elif not nav:
            first_payload, pieces = pieces[0], pieces[1:]
            if self.target is not None and first_payload:
                self.add_sentence(self.target, first_payload, marked=True)
            elif first_payload:
                self.flag("unresolved command", s, "Add-on text with no target")
                return False
        ok = True
        for idx, p in enumerate(pieces):
            if not p:
                continue
            if idx == 0 and first_is_target and "," in p and room_match(p.split(",", 1)[0]):
                head, tail = p.split(",", 1)
                rm = room_match(head)
                self.open_room(rm[0], rm[1])
                self.revisit = True
                self.flag("go back", s, f"Went back to room '{rm[1]}'")
                h = self.heading_of(tail.strip())
                if h and h[0] == "item" and self.resolve_target(h[1][1], s, quiet=True):
                    if h[2]:
                        self.add_sentence(self.target, h[2], marked=True)
                else:
                    self.add_sentence(self.cur or self.new_item(self.byKey["additional"], "additional"), tail.strip(), marked=True)
                    self.flag("unresolved go back", s, "Went back to the room but could not tell which item — text added to the last item there")
                continue
            if idx == 0 and first_is_target:
                p2 = re.sub(r"\s+(?:to|in|of)\s+(?:the\s+)?(?:property|house)$", "", p, flags=re.I)
                rm = room_match(p) or room_match(p2)
                if rm:
                    self.open_room(rm[0], rm[1])
                    self.revisit = True
                    self.flag("go back", s, f"Went back to room '{rm[1]}'")
                    continue
                if self.resolve_target(p2, s):
                    continue
                ok = False
                self.flag("unresolved go back", s, f"Could not find '{p}' — original sentence left in place")
                break
            else:
                if self.resolve_target(p, s, quiet=True):
                    continue
                if self.target is None:
                    self.flag("unresolved command", s, f"No target for '{p}'")
                    ok = False
                    break
                self.add_sentence(self.target, p, marked=True)
        if not ok:
            return False
        if self.target is None:
            self.mark_next = False
        self.flag("go back / add on", s, "Applied; added text marked #")
        return True

    def resolve_target(self, phrase, s, quiet=False):
        p = re.sub(r"\s+(?:to|in|of)\s+(?:the\s+)?(?:property|front of property|house)$", "", phrase.strip(), flags=re.I)
        fi = self.find_item(p)
        if not fi:
            p2 = re.sub(r"\b(?:under|in|to|the|of|at)\b", " ", re.sub(_ROOMANY, " ", p, flags=re.I), flags=re.I)
            p2 = re.sub(r"\s+", " ", p2).strip()
            fi = self.find_item(p2) if p2 else None
        if not fi:
            return False
        it, spoken, _ = fi
        found = None
        if self.room:
            for item in reversed(self.room["items"]):
                if it["key"] in item["covers"] or item["key"] == it["key"]:
                    found = item
                    break
        if found is None:
            found = self.new_item(it, spoken)
            self.flag("new item", s, f"'{found['name']}' did not exist in this room, so it was created")
            if not found["sents"]:
                pass
        self.cur = found
        self.target = found
        return True

    # ---- main loop
    def run(self, raw):
        sents = split_sentences(clean_raw(raw))
        for s in sents:
            self.step(s)
        return self.render()

    def step(self, s):
        orig = s
        if re.search(r"\boh actually\b", s, re.I):
            self.flag("dropped", orig, "Speaker changed their mind mid-sentence — sentence not used; check")
            return
        # "Could you please keep this area under sunken lounge?" -> next heading becomes ONE container item (like Robe)
        m = re.match(r"(?:(?:sorry|so|oh|okay|ok)[,\s]+)*(?:could you\s+)?(?:please\s+)?(?:keep|put|place|group|file)\s+(?:this|that|these|it)"
                     r"(?:\s+(?:area|section|part|bit))?\s+(?:all\s+)?(?:under|in|inside|within)\s+(?:the\s+)?(.+?)[.?!\s]*$", s, re.I)
        if m:
            rm = room_match(m.group(1))
            if rm:
                self.open_room(rm[0], rm[1])
                self.revisit = False
                self.area_pending = True
                self.flag("area under room", orig, f"Next heading will be kept as ONE item inside '{rm[1]}' (like Robe)")
            else:
                self.flag("unresolved", orig, f"Could not tell which room '{m.group(1)}' is — left as spoken")
                self.add_sentence(self.cur, s) if self.cur else None
            return
        if self.area_pending:
            core = s.rstrip(".?! ")
            if room_match(core):
                self.area_pending = False
            elif core and len(core.split(",")[0].split()) <= 6:
                head, _, rest = core.partition(",")
                name = re.sub(r"\s+slash\s+", "/", head.strip(), flags=re.I)
                name = re.sub(r"\s+and\s+", "/" if self.job == "pearl" else " & ", name, flags=re.I)
                entry = {"key": "area:" + head.strip().lower(), "match": re.escape(head.strip()), "normal": [title(name)],
                         "pearl": [title(name)], "container": True}
                item = self.new_item(entry, head.strip())
                if rest.strip():
                    self.add_sentence(item, rest.strip())
                self.area_pending = False
                return
            else:
                self.area_pending = False
                self.flag("unresolved", orig, "Expected the area name after 'keep this under <room>' — left as spoken")
        # standalone "Sorry, no." / "Nope, sorry."
        if re.fullmatch(r"(?:sorry|nope|no)[,.\s]+(?:sorry|nope|no)\.?", s, re.I) and self.cur and self.cur["sents"]:
            dropped = self.cur["sents"].pop()
            self.flag("removed", orig, f"Removed previous line: '{dropped.lstrip('#')}'")
            self.mark_next = True
            return
        m = re.match(r"(?:nope\b[,\s]*|no,\s*)(?:sorry[,\s]+)?(.+)$", s, re.I)
        if m and self.cur and self.cur["sents"]:
            if self.replace_similar(self.cur, m.group(1), orig):
                return
            s = cap(m.group(1))
            self.mark_next = True
            self.flag("correction", orig, "Treated as a replacement but no similar earlier line found — kept as new line marked #")
        # commands (go back / add on / change)
        if re.search(r"\b(?:go back|going back|head back|back to|thinking about|add on|could you|change)\b", s, re.I):
            if self.handle_command(s):
                return
        if re.fullmatch(r"back to (?:the )?(.+?)\.?", s, re.I):
            rm = room_match(re.fullmatch(r"back to (?:the )?(.+?)\.?", s, re.I).group(1))
            if rm:
                self.open_room(rm[0], rm[1])
                self.revisit = True
                self.flag("go back", orig, f"Back to room '{rm[1]}' — items added from here are marked #")
                return
        s, ch = self.inline_fix(s)
        h = self.heading_of(s)
        if h:
            self.target = None
            if h[0] == "room":
                self.open_room(h[1][0], h[1][1])
                self.revisit = False
                return
            if h[0] == "roomitem":
                self.open_room(h[1][0], h[1][1])
                self.revisit = False
                it, spoken, _ = h[2]
                item = self.new_item(it, spoken)
                if h[3]:
                    self.add_sentence(item, h[3])
                return
            it, spoken, _ = h[1]
            if len(h) > 3 and h[3] == "lead" and self.cur is not None and (it["key"] in self.cur["covers"]):
                self.add_sentence(self.cur, s)   # "Blinds have ..." while already in Window Treatments: just a sentence
                return
            # a container (robe/pantry/linen) swallows everything except a new container or another room
            if self.cur and self.cur["container"] and not it.get("container") and not (it["key"] in ("pantry",)):
                self.add_sentence(self.cur, s)
                return
            item = self.new_item(it, spoken)
            if h[2]:
                self.add_sentence(item, h[2])
            return
        # add mode
        if self.target is not None:
            imp = self.implicit_of(s)
            if not imp:
                self.add_sentence(self.target, s, marked=True)
                return
            self.target = None
        # implicit item start
        imp = self.implicit_of(s)
        if imp and self.cur is not None and not self.cur["container"]:
            it, spoken = imp
            if it["key"] in self.cur["covers"] and self.cur["key"] != "cornice_ceiling":
                self.add_sentence(self.cur, s)
                return
            if self.cur["pending"] and it["key"] in self.cur.get("pend_keys", []):
                self.cur["pending"].pop(0)
                self.cur["pend_keys"].pop(0)
                self.new_item(it, spoken)
                self.add_sentence(self.cur, s)
                return
            if it["key"] in self.cur["covers"]:
                self.add_sentence(self.cur, s)
                return
            self.new_item(it, spoken)
            self.add_sentence(self.cur, s)
            return
        if self.cur is None:
            self.flag("no item", orig, "Sentence heard before any item heading")
            if self.room is None:
                self.open_room("unsorted", "Unsorted")
            self.cur = {"key": "unsorted", "name": "Unsorted", "entry": {"match": "$^"}, "sents": [], "pending": [],
                        "pend_keys": [], "container": False, "vessel": None, "combined": False, "covers": set()}
            self.room["items"].append(self.cur)
        self.add_sentence(self.cur, s)

    # ---- output
    def render(self):
        out = []
        for r in self.rooms:
            out.append("!" + r["name"])
            for it in r["items"]:
                out.append(it["name"] + "\t" + " ".join(it["sents"]))
        text = "\n".join(out)
        return text, self.review


def local_parse(raw, job="normal"):
    """-> (list_text, review). job: 'normal' or 'pearl'."""
    return Parser(job).run(raw)


def short_lines(raw, job="normal"):
    """Short sentences that were NOT recognised as headings (possible vocabulary gaps)."""
    p = Parser(job)
    out = []
    for s in split_sentences(clean_raw(raw)):
        core = s.rstrip(".?! ")
        w = core.split()
        if 1 <= len(w) <= 3 and not p.heading_of(s) and not re.match(r"(?:%s)\b" % NUMS, core, re.I) and "," not in core:
            if not re.search(r"\b(?:is|are|noted|visible|present|intact|dusty|clean|worn|secure|secured|damaged|scuffed|stained)\b", core, re.I):
                out.append(core)
    return out


if __name__ == "__main__":
    import sys
    raw = open(sys.argv[1], encoding="utf-8").read()
    job = sys.argv[2] if len(sys.argv) > 2 else "normal"
    txt, rev = local_parse(raw, job)
    print(txt)
    print("\n--- REVIEW ---", file=sys.stderr)
    for r in rev:
        print(f"[{r['kind']}] ({r['room']}) {r['text']}  -> {r['note']}", file=sys.stderr)
