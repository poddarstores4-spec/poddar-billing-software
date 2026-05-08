import json, os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

ITEMS_FILE   = os.path.join(DATA_DIR, "items.json")
PARTIES_FILE = os.path.join(DATA_DIR, "parties.json")
INVOICES_FILE= os.path.join(DATA_DIR, "invoices.json")
COUNTER_FILE = os.path.join(DATA_DIR, "counter.json")

# ─── default data from Excel ────────────────────────────────────────────────

DEFAULT_ITEMS = [
    {"name":"ARHAR DAAL",   "hsn":"713",  "unit":"KGS","rate":97,  "gst":0.05,"category":"PULSE"},
    {"name":"MASOOR DAAL",  "hsn":"713",  "unit":"KGS","rate":87,  "gst":0.05,"category":"PULSE"},
    {"name":"CHANA DAAL",   "hsn":"713",  "unit":"KGS","rate":83,  "gst":0.05,"category":"PULSE"},
    {"name":"MOONG DAAL",   "hsn":"713",  "unit":"KGS","rate":110, "gst":0.05,"category":"PULSE"},
    {"name":"CHANA",        "hsn":"713",  "unit":"KGS","rate":68,  "gst":0.05,"category":"PULSE"},
    {"name":"SOYABEAN",     "hsn":"1201", "unit":"KGS","rate":115, "gst":0.05,"category":"PULSE"},
    {"name":"NAMAK",        "hsn":"2501", "unit":"KGS","rate":28,  "gst":0.05,"category":"SALT"},
    {"name":"SARSO TEL",    "hsn":"1514", "unit":"LTR","rate":187, "gst":0.05,"category":"EDIBLE OIL"},
    {"name":"HALDI POWDER", "hsn":"904",  "unit":"KGS","rate":180, "gst":0.05,"category":"SPICES"},
    {"name":"MIRCH POWDER", "hsn":"904",  "unit":"KGS","rate":290, "gst":0.05,"category":"SPICES"},
    {"name":"DHANIYA POWDER","hsn":"904", "unit":"KGS","rate":264, "gst":0.05,"category":"SPICES"},
    {"name":"PUNCH PHORNA", "hsn":"904",  "unit":"KGS","rate":190, "gst":0.05,"category":"SPICES"},
    {"name":"JEERA",        "hsn":"904",  "unit":"KGS","rate":325, "gst":0.05,"category":"SPICES"},
    {"name":"SUKHA MIRCH",  "hsn":"904",  "unit":"KGS","rate":260, "gst":0.05,"category":"SPICES"},
    {"name":"TEJPATTA",     "hsn":"904",  "unit":"KGS","rate":180, "gst":0.05,"category":"SPICES"},
    {"name":"GARAM MASALA", "hsn":"910",  "unit":"PCS","rate":5,   "gst":0.05,"category":"SPICES"},
    {"name":"AALU",         "hsn":"701",  "unit":"KGS","rate":12,  "gst":0.05,"category":"VEGETABLE"},
    {"name":"PAYAJ (ONION)","hsn":"701",  "unit":"KGS","rate":17,  "gst":0.05,"category":"VEGETABLE"},
    {"name":"GOBHI",        "hsn":"701",  "unit":"KGS","rate":30,  "gst":0,   "category":"VEGETABLE"},
    {"name":"KADDU",        "hsn":"701",  "unit":"KGS","rate":40,  "gst":0,   "category":"VEGETABLE"},
    {"name":"HARA MIRCH",   "hsn":"701",  "unit":"KGS","rate":100, "gst":0,   "category":"VEGETABLE"},
    {"name":"PARWAL",       "hsn":"701",  "unit":"KGS","rate":50,  "gst":0,   "category":"VEGETABLE"},
    {"name":"PALAK",        "hsn":"701",  "unit":"KGS","rate":30,  "gst":0,   "category":"VEGETABLE"},
    {"name":"PATTA GOBHI",  "hsn":"701",  "unit":"KGS","rate":30,  "gst":0,   "category":"VEGETABLE"},
    {"name":"SEASONAL FRUIT","hsn":"803", "unit":"PCS","rate":5,   "gst":0,   "category":"SEASONAL FRUIT"},
    {"name":"JLAWAN/ENDHAN","hsn":"4401", "unit":"KGS","rate":14,  "gst":0.05,"category":"WOODS"},
    {"name":"FRUIT KELA",   "hsn":"803",  "unit":"PCS","rate":6,   "gst":0.05,"category":"SEASONAL FRUIT"},
    {"name":"RAW CHANA",    "hsn":"713",  "unit":"KGS","rate":68,  "gst":0.05,"category":"PULSE"},
    {"name":"EGGS",         "hsn":"407",  "unit":"PCS","rate":5,   "gst":0,   "category":"EGGS"},
]

DEFAULT_PARTIES = [
    {"name":"PS PURA",              "address":"PURA",            "block":"SADAR",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS KOPINANDA",         "address":"KOPINANDA",       "block":"BENIPUR",        "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS PANKIPATTI",        "address":"PANKIPATTI",      "block":"GHANSYAMPUR",    "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS SHANKARLOHAR",      "address":"SHANKARLOHAR",    "block":"BAHERI",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS RAJWARATOL",        "address":"RAJWARATOL",      "block":"MANIGACHHI",     "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS SUSARI",            "address":"SUSARI",          "block":"BAHERI",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS NAVTOL BITHAULI",   "address":"BITHAULI",        "block":"BAHERI",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS KAMTOL",            "address":"KAMTOL",          "block":"JALE",           "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS VISHNUPUR",         "address":"VISHNUPUR",       "block":"HANUMAN NAGAR",  "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS KAKARGHATTI URDU",  "address":"KAKARGHATTI",     "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS MADHOPUR GOPALPUR", "address":"MADHOPUR",        "block":"HANUMAN NAGAR",  "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"NPS BIJULI",           "address":"BIJULI",          "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"UMS MIRZAPUR BOYS",    "address":"MIRZAPUR",        "block":"ALINAGAR",       "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS KOPICHAKLA",        "address":"KOPICHAKLA",      "block":"BAHERI",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS PINDARUCH KANYA",   "address":"PINDARUCH",       "block":"KEOWTI",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS BITHAULI",          "address":"BITHAULI",        "block":"BAHERI",         "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS GODYARI",           "address":"GODYARI",         "block":"MANIGACHHI",     "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS CHAKKA",            "address":"CHAKKA",          "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS LAHO",              "address":"LAHO",            "block":"MAANIGACHHI",    "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS GHOURDAUR",         "address":"GHOURDAUR",       "block":"MANIGACHHI",     "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS MAHARANI NAVTOLIYA","address":"NAVTOLIYA",       "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS KAISTH KABAI",      "address":"KAISTH KABAI",    "block":"MANIGACHHI",     "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS PANTA",             "address":"PANTA",           "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS KUMHAROLI HINDI",   "address":"KUMHAROLI",       "block":"JALE",           "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS MAJHIGAMA",         "address":"MAJHIGAMA",       "block":"KEOTI",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"B S GODAIPATTI",       "address":"GODAIPATTI",      "block":"HANUMAN NAGAR",  "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"PS NAYATOL KANSI",     "address":"NAYATOL KANSI",   "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS GODAIPATTI",        "address":"GODAIPATTI",      "block":"HANUMANNGAR",    "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
    {"name":"MS KHARTHUA",          "address":"KHARTHUA",        "block":"SADAR",          "district":"DARBHANGA","state":"BIHAR","state_code":"10"},
]

DEFAULT_COUNTER = {
    "JANUARY":0,"FEBRUARY":0,"MARCH":0,"APRIL":0,"MAY":0,"JUNE":0,
    "JULY":0,"AUGUST":0,"SEPTEMBER":0,"OCTOBER":0,"NOVEMBER":0,"DECEMBER":0
}

# ─── helpers ────────────────────────────────────────────────────────────────

def _load(path, default):
    if os.path.exists(path):
        with open(path) as f: return json.load(f)
    _save(path, default); return default

def _save(path, data):
    with open(path, "w") as f: json.dump(data, f, indent=2)

# ─── public API ─────────────────────────────────────────────────────────────

def get_items():    return _load(ITEMS_FILE, DEFAULT_ITEMS)
def save_items(d):  _save(ITEMS_FILE, d)

def get_parties():   return _load(PARTIES_FILE, DEFAULT_PARTIES)
def save_parties(d): _save(PARTIES_FILE, d)

def get_invoices():   return _load(INVOICES_FILE, [])
def save_invoices(d): _save(INVOICES_FILE, d)

def get_counter():   return _load(COUNTER_FILE, DEFAULT_COUNTER)
def save_counter(d): _save(COUNTER_FILE, d)

def next_invoice_no(month_name):
    c = get_counter()
    month_upper = month_name.upper()
    c[month_upper] = c.get(month_upper, 0) + 1
    save_counter(c)
    mon_abbr = month_name[:3].upper()
    year = str(datetime.now().year)[2:]
    return f"{mon_abbr}/26/{c[month_upper]:03d}"

def add_invoice(inv):
    invs = get_invoices()
    invs.append(inv)
    save_invoices(invs)

def item_by_name(name):
    for i in get_items():
        if i["name"] == name: return i
    return None

def party_by_name(name):
    for p in get_parties():
        if p["name"] == name: return p
    return None
