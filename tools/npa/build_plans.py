# Сборка ORDER_PLANS для index.html из таблиц, извлечённых из PDF приказов №227 и №183.
# Вход: tools/npa/out/tables_227.json, tables_183.json. Выход: tools/npa/out/order_plans.js + сводка в stdout.
import json, io, re, os, sys
sys.stdout.reconfigure(encoding="utf-8")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
t227 = json.load(io.open(os.path.join(OUT, "tables_227.json"), encoding="utf-8"))
t183 = json.load(io.open(os.path.join(OUT, "tables_183.json"), encoding="utf-8"))

# приложение → язык обучения / образец плана
LANG_227 = {"2": "uz", "3": "ru", "4": "q1", "5": "q2"}
DIR_227 = [  # порядок направлений в каждом приложении и ключи
    ("FILOLOGIYA", "philology"), ("XORIJIY TILLAR", "foreign_langs"), ("MOLIYA", "finance"),
    ("AXBOROT TEXNOLOGIYALARI", "it"), ("MUHANDISLIK", "engineering"), ("TABIIY FANLAR", "natural"),
    ("IJTIMOIY-GUMANITAR", "social_humanities"), ("BOSHLANG", "primary_ed"), ("KIMYO-TEXNOLOGIYA", "chem_tech"),
    ("GEOGRAFIYA VA GEOLOGIYA", "geo"), ("MAKTABGACHA", "preschool_psy"), ("TARIX VA ARXEOLOGIYA", "history_arch"),
    ("TURIZM", "tourism"), ("YURISPRUDENSIYA", "law"), ("IJTIMOIY ISH", "social_work"),
    ("YOʻL HARAKATI", "road_safety"), ("SPORT", "sport"), ("HARBIY VATANPARVARLIK", "military"),
]
MAP_183 = {"2.1": ("mf", "uz"), "2.2": ("kb", "uz"), "2.3": ("oc", "uz"), "2.4": ("mc", "uz"),
           "3": ("ix", "uz"), "4": ("ix", "ru"), "5": ("ix", "q1"), "6": ("ix", "q2"),
           "7": ("nd1", "uz"), "8": ("nd1", "ru"), "9": ("nd1", "q1"), "10": ("nd1", "q2"),
           "11": ("nd2", "uz"), "12": ("nd2", "ru"), "13": ("nd2", "q1"), "14": ("nd2", "q2")}
NAME_FIX = {"Informatika va AT": "Informatika va axborot texnologiyalari",
            "ChQBT": "Chaqiruvga qadar boshlangʻich tayyorgarlik"}

def norm_name(n):
    n = re.sub(r"\s+", " ", n.replace("’", "ʼ").replace("'", "ʼ")).strip()
    n = n.replace("sanʼat", "sanʼat")
    return NAME_FIX.get(n, n)

def pack(t):
    g = t["grades"]
    rows = []
    for r in t["rows"]:
        hrs = [r["vals"].get(str(x), 0) for x in g]
        rows.append([norm_name(r["name"]), "m" if r["section"] == "mandatory" else "e", hrs])
    return {"g": g, "r": rows}

def direction_key(title):
    T = title.upper().replace("’", "ʼ")
    for needle, key in DIR_227:
        if needle.upper() in T: return key
    return None

plans = {"227": {}, "183": {}}
problems = []
for t in t227:
    lang = LANG_227.get(t["annex"]); key = direction_key(t["title"])
    if not lang or not key: problems.append(("227", t["annex"], t["page"], t["title"][-80:])); continue
    if lang in plans["227"].setdefault(key, {}): problems.append(("227 dup", key, lang, t["page"]))
    plans["227"][key][lang] = pack(t)
for t in t183:
    m = MAP_183.get(t["annex"])
    if not m: problems.append(("183", t["annex"], t["page"])); continue
    plans["183"].setdefault(m[0], {})[m[1]] = pack(t)

print("227 directions:", len(plans["227"]), {k: sorted(v.keys()) for k, v in plans["227"].items() if len(v) != 4})
print("183 tracks:", {k: sorted(v.keys()) for k, v in plans["183"].items()})
print("problems:", problems)
names = sorted({r[0] for d in plans.values() for x in d.values() for tb in x.values() for r in tb["r"]})
print("names:", len(names), names)
# компактная форма: названия — по индексу в общем списке, строка = [индекс, 0|1 (1 — вторая часть), часы по классам g]
idx = {n: i for i, n in enumerate(names)}
def compact(tb):
    return {"g": tb["g"], "r": [[idx[r[0]], 1 if r[1] == "e" else 0] + r[2] for r in tb["r"]]}
import datetime
# версия данных: дата приказа и файл-источник — рабочая область запоминает, по какой версии заполнен план
cplans = {"meta": {"183": {"date": "2026-05-25", "source": "НПА/www.idum.uz__tanlov_fanlar_183_2026-2027.pdf"},
                   "227": {"date": "2026-06-26", "source": "НПА/227_buyruq_variativ_reja.pdf"}},
          "n": names}
for o in ("227", "183"):
    cplans[o] = {k: {l: compact(tb) for l, tb in v.items()} for k, v in plans[o].items()}
js = ("/* Учебные планы приказов МНиШО №227 (26.06.2026, вариативные планы 7–11) и №183 (25.05.2026,\n"
      "   предметы по выбору 10–11). Извлечено из PDF скриптами tools/npa — не править вручную.\n"
      "   Язык: uz, ru, q1/q2 — родственные языки (казахский, киргизский, таджикский, туркменский),\n"
      "   образец 1 (с русским языком) / 2 (без). Строка: [индекс названия в n, 0 — обязательная часть /\n"
      "   1 — вторая часть (предметы по выбору, профессия), часы по классам g]. */\n"
      "var ORDER_PLANS = " + json.dumps(cplans, ensure_ascii=False, separators=(",", ":")) + ";\n")
io.open(os.path.join(OUT, "order_plans.js"), "w", encoding="utf-8").write(js)
# --inject <index.html>: заменить блок между маркерами ORDER_PLANS:BEGIN / ORDER_PLANS:END
if "--inject" in sys.argv:
    target = sys.argv[sys.argv.index("--inject") + 1]
    page = io.open(target, encoding="utf-8", newline="").read()
    begin = "/* ORDER_PLANS:BEGIN */\n"
    b0 = page.index(begin) + len(begin)
    b1 = page.index("/* ORDER_PLANS:END */")
    io.open(target, "w", encoding="utf-8", newline="").write(page[:b0] + js + page[b1:])
    print("injected into", target)
print("bytes:", len(js.encode("utf-8")))
# выборочная сверка
for key, lang in (("philology", "uz"), ("law", "ru")):
    tb = plans["227"][key][lang]; print("\n227", key, lang, tb["g"]); [print("  ", r) for r in tb["r"]]
tb = plans["183"]["nd2"]["q2"]; print("\n183 nd2 q2", tb["g"]); [print("  ", r) for r in tb["r"]]
