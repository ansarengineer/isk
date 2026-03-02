import json
import re
import tkinter as tk
from dataclasses import dataclass
from collections import defaultdict, Counter
from difflib import SequenceMatcher
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import pandas as pd

# ===================== НАСТРОЙКИ =====================
COURTS_JSON_PATH = "./courts_merged.json"

TH_TAKE_LOCAL = 2.6
TH_TAKE_GLOBAL = 3.4
TH_REGION_FROM_PLACE = 0.55

# Excel: какой столбец адреса искать
ADDRESS_COL_PRIMARY = "Адрес"
ADDRESS_COL_ALIASES = {
    "адрес", "адрес проживания", "адрес регистрации", "адрес объекта", "место жительства", "местожительства"
}


FORCED_COURT_BY_ANCHOR = {
    "ордабас": "Ордабасинский районный суд Туркестанской области",
    "ордабасы": "Ордабасинский районный суд Туркестанской области",
    "ordabasy": "Ордабасинский районный суд Туркестанской области",

    "текели": "Текелийский городской суд области Жетісу",
    "tekeli": "Текелийский городской суд области Жетісу",

    "талдыкорган": "Талдыкорганский городской суд области Жетісу",
    "taldykorgan": "Талдыкорганский городской суд области Жетісу",

    "тюлькубас": "Тюлькубасский районный суд Туркестанской области",
    "түлкібас": "Тюлькубасский районный суд Туркестанской области",
    "tulkibas": "Тюлькубасский районный суд Туркестанской области",

    "махтаарал": "Мактааральский районный суд Туркестанской области",
    "мактаарал": "Мактааральский районный суд Туркестанской области",
    "мақтаарал": "Мактааральский районный суд Туркестанской области",
    "maktaaral": "Мактааральский районный суд Туркестанской области",
    "mahtaaral": "Мактааральский районный суд Туркестанской области",
}

# ===================== НОРМАЛИЗАЦИЯ / СЛОВАРИ =====================
STOP_TOKENS = {
    "суд", "cуд", "суда",
    "районный", "городской", "межрайонный", "специализированный",
    "по", "делам", "административный", "гражданским", "уголовным",
    "г", "г.", "город", "города", "қала",
    "область", "области", "обл", "обл.",
    "республика", "республики", "казахстан",
    "район", "района", "р-н", "рн",
    "аудан", "ауданы", "аудана", "ауд.",
    "№", "номер", "имени"
}

PLACE_ALIASES = {
    "капчагай": "қонаев",
    "капшагай": "қонаев",
    "қапшағай": "қонаев",
}

# РЕГИОН по якорям (оставляем как было, но Ордабасы тут можно оставить для региона)
REGION_ANCHORS = {
    "зко": "ЗКО",
    "з-казахстан": "ЗКО",
    "западно казахстан": "ЗКО",

    "вко": "ВКО",
    "в-казахстан": "ВКО",
    "восточно казахстан": "ВКО",

    "ско": "СКО",
    "с-казахстан": "СКО",
    "северо казахстан": "СКО",

    "уральск": "ЗКО",

    "ордабас": "Туркестанская область",
    "ордабасы": "Туркестанская область",
    "ordabasy": "Туркестанская область",

    "саркан": "Область Жetісу".replace("Жet", "Жет"),
    "сарқанд": "Область Жetісу".replace("Жet", "Жет"),

    "нур-султан": "Астана",
    "нурсултан": "Астана",
    "астан": "Астана",
    "алматы": "Алматы",
    "шымкент": "Шымкент",

    "жетісу": "Область Жetісу".replace("Жet", "Жет"),
    "жетысу": "Область Жetісу".replace("Жet", "Жет"),
    "абай": "Область Абай",
    "улытау": "Область Ұлытау",

    "акмол": "Акмолинская область",
    "актюб": "Актюбинская область",
    "атырау": "Атырауская область",
    "жамбыл": "Жамбылская область",
    "костанай": "Костанайская область",
    "караган": "Карагандинская область",
    "кызылорд": "Кызылординская область",
    "мангист": "Мангистауская область",
    "павлодар": "Павлодарская область",
    "туркестан": "Туркестанская область",

    "қонаев": "Алматинская область",
    "конаев": "Алматинская область",

    "текели": "Область Жетісу",
    "tekeli": "Область Жетісу",

    "талдыкорган": "Область Жетісу",
    "taldykorgan": "Область Жетісу",
    "тaлдыкорган": "Область Жетісу",    

    "тюлькубас": "Туркестанская область",
    "түлкібас": "Туркестанская область",
    "tulkibas": "Туркестанская область",

    "махтаарал": "Туркестанская область",
    "мактаарал": "Туркестанская область",
    "мақтаарал": "Туркестанская область",
    "maktaaral": "Туркестанская область",
    "mahtaaral": "Туркестанская область",

}

REGION_CANON = {
    "западно-казахстанская область": "ЗКО",
    "западно казахстанская область": "ЗКО",
    "зко": "ЗКО",
    "восточно-казахстанская область": "ВКО",
    "восточно казахстанская область": "ВКО",
    "вко": "ВКО",
    "северо-казахстанская область": "СКО",
    "северо казахстанская область": "СКО",
    "ско": "СКО",
    
    "область жетысу": "Область Жетісу",
    "область жетісу": "Область Жетісу",
    "жетысу": "Область Жетісу",
    "жетісу": "Область Жетісу",
}

DISTRICT_PATTERNS = [
    r"\b([a-zа-яёәіңғқөұүһ0-9\- ]{2,}?)\s+(район|р\-н|рн)\b",
    r"\b([a-zа-яёәіңғқөұүһ0-9\- ]{2,}?)\s+(ауданы|аудан|ауд\.)\b",
    r"\bрайон\s+([a-zа-яёәіңғқөұүһ0-9\- ]{2,}?)\b",
    r"\bаудан\s+([a-zа-яёәіңғқөұүһ0-9\- ]{2,}?)\b",
]
CITY_PATTERNS = [
    r"\bг\.?\s*([a-zа-яёәіңғқөұүһ\- ]{2,})\b",
    r"\bгород\s+([a-zа-яёәіңғқөұүһ\- ]{2,})\b",
    r"\bқала\s+([a-zа-яёәіңғқөұүһ\- ]{2,})\b",
]

# ===================== УТИЛИТЫ =====================
def normalize(text: str) -> str:
    t = str(text).lower()
    t = re.sub(r"[^\w\s\-әіңғқөұүһ]", " ", t, flags=re.IGNORECASE)
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\bказахстанск(ая|ой|ому|ом|ие|их|ими|ую)\b", "казахстан", t)

    for src, dst in PLACE_ALIASES.items():
        t = re.sub(rf"\b{re.escape(src)}\b", dst, t)

    return re.sub(r"\s+", " ", t).strip()

def canon_region(region: str) -> str:
    r = str(region or "").strip()
    if not r:
        return r
    return REGION_CANON.get(normalize(r), r)

def split_parts(address: str, n: int = 4) -> list[str]:
    parts = [p.strip() for p in str(address).split(",")]
    return [p for p in parts[:n] if p]

def stem(word: str) -> str:
    w = normalize(word).replace("-", " ").strip()
    w = re.sub(r"\b(г|г\.)\b", "", w).strip()
    w = re.sub(r"(ская|ский|ского|скому|ским|ских|ское|ские|ое|ая|ий|ый)$", "", w)
    w = re.sub(r"(ов|ев|ин|ын)$", "", w)
    w = w.replace("ы", "и")
    return re.sub(r"\s+", " ", w).strip()

def tokens(text: str) -> list[str]:
    t = normalize(text).replace("-", " ")
    out = []
    for w in t.split():
        if not w or w.isdigit() or w in STOP_TOKENS:
            continue
        out.append(w)
    return out

def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()

def extract_by_patterns(text: str, patterns: list[str]) -> str | None:
    h = normalize(text)
    for pat in patterns:
        m = re.search(pat, h, flags=re.IGNORECASE)
        if m:
            return normalize(m.group(1))
    return None

def extract_address_entities(address: str) -> dict:
    parts = split_parts(address, 4)
    head = " ".join(parts)

    district = extract_by_patterns(head, DISTRICT_PATTERNS)
    city = extract_by_patterns(head, CITY_PATTERNS)

    guessed_city = None
    if not city and len(parts) >= 3:
        cand = normalize(parts[2])
        if not re.search(r"\b(район|р\-н|рн|аудан|ауданы)\b", cand) and not re.search(
            r"\b(ул|улица|просп|проспект|мкр|микрорайон|кв|квартира|дом|д\.)\b",
            cand
        ):
            guessed_city = cand

    first3 = " ".join(parts[:3]) if parts else ""
    cand_stems = {stem(w) for w in tokens(first3) if len(stem(w)) >= 3}

    return {
        "district": district,
        "city": city or guessed_city,
        "cand_stems": cand_stems
    }

# ===================== СУДЫ / ИНДЕКС =====================
@dataclass
class CourtRec:
    region: str
    name: str
    stems: set[str]
    norm: str

def load_courts(json_path: str):
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    by_region: dict[str, list[CourtRec]] = defaultdict(list)
    all_courts: list[CourtRec] = []

    for item in data:
        region_raw = str(item.get("Регион", "")).strip()
        court = str(item.get("СУД", "")).strip()
        if not region_raw or not court:
            continue

        region = canon_region(region_raw)
        cstems = {stem(w) for w in tokens(court) if len(stem(w)) >= 3}
        rec = CourtRec(region=region, name=court, stems=cstems, norm=normalize(court))

        by_region[region].append(rec)
        all_courts.append(rec)

    return by_region, all_courts

def build_place_to_region_index(all_courts: list[CourtRec]) -> dict[str, Counter]:
    idx: dict[str, Counter] = defaultdict(Counter)

    for rec in all_courts:
        d = extract_by_patterns(rec.norm, DISTRICT_PATTERNS)
        if d:
            for w in tokens(d):
                sw = stem(w)
                if len(sw) >= 3:
                    idx[sw][rec.region] += 3

        c = extract_by_patterns(rec.norm, CITY_PATTERNS)
        if c:
            for w in tokens(c):
                sw = stem(w)
                if len(sw) >= 3:
                    idx[sw][rec.region] += 2

        for sw in rec.stems:
            idx[sw][rec.region] += 1

    return idx

# Глобальные хранилища судов (инициализируются при старте)
COURTS_BY_REGION: dict[str, list[CourtRec]] = {}
ALL_COURTS: list[CourtRec] = []
PLACE_TO_REGION: dict[str, Counter] = {}

# ===================== РЕГИОН =====================
def detect_region(address: str) -> str:
    parts = split_parts(address, 3)
    head = normalize(" ".join(parts))

    for key, reg in REGION_ANCHORS.items():
        if key in head:
            return canon_region(reg)

    ent = extract_address_entities(address)
    candidates = []

    if ent["district"]:
        candidates += [stem(w) for w in tokens(ent["district"]) if stem(w)]
    if ent["city"]:
        candidates += [stem(w) for w in tokens(ent["city"]) if stem(w)]
    candidates += list(ent["cand_stems"])

    vote = Counter()
    for sw in candidates:
        vote.update(PLACE_TO_REGION.get(sw, Counter()))

    if vote:
        top_region, top_score = vote.most_common(1)[0]
        total = sum(vote.values())
        if total > 0 and (top_score / total) >= TH_REGION_FROM_PLACE:
            return canon_region(top_region)

    return "Не определён"

# ===================== СКОРИНГ / ВЫБОР СУДА =====================
def score_court(rec: CourtRec, district: str | None, city: str | None, addr_stems: set[str]) -> float:
    s = 0.0
    cstems = rec.stems
    nc = rec.norm

    if district:
        dstems = {stem(w) for w in tokens(district) if len(stem(w)) >= 3}
        if dstems:
            overlap = len(dstems & cstems)
            if overlap:
                s += overlap * 3.5
            else:
                best_sim = max((similarity(ds, cs) for ds in dstems for cs in cstems), default=0.0)
                s += best_sim * 2.4
            if ("район" in nc) or ("аудан" in nc) or ("р-н" in nc) or ("рн" in nc):
                s += 0.4

    if city:
        cst_city = {stem(w) for w in tokens(city) if len(stem(w)) >= 3}
        if cst_city:
            overlap = len(cst_city & cstems)
            if overlap:
                s += overlap * 2.5
            else:
                best_sim = max((similarity(a, b) for a in cst_city for b in cstems), default=0.0)
                s += best_sim * 1.8
            if ("город" in nc) or ("города" in nc) or ("г." in nc):
                s += 0.2

    s += len(addr_stems & cstems) * 1.0
    if "област" in nc or "вко" in nc or "ско" in nc or "зко" in nc:
        s += 0.1

    return s

def best_match(courts: list[CourtRec], district: str | None, city: str | None, addr_stems: set[str]):
    best = None
    best_score = -1.0
    for rec in courts:
        sc = score_court(rec, district, city, addr_stems)
        if sc > best_score:
            best_score = sc
            best = rec
    return best_score, best

def detect_court(address: str, region: str) -> tuple[str, float, str | None]:
    ent = extract_address_entities(address)
    district, city, addr_stems = ent["district"], ent["city"], ent["cand_stems"]

    head = normalize(address)

    # 0) Принудительный суд по якорю (Ордабасы и т.п.)
    for key, court_name in FORCED_COURT_BY_ANCHOR.items():
        if key in head:
            # Возвращаем высокий score, чтобы всегда победило
            return court_name, 10_000.0, canon_region("Туркестанская область")

    region = canon_region(region)
    local_list = COURTS_BY_REGION.get(region, [])

    if local_list and len(local_list) == 1:
        return local_list[0].name, 999.0, region

    if local_list:
        sc, rec = best_match(local_list, district, city, addr_stems)
        if rec and sc >= TH_TAKE_LOCAL:
            return rec.name, sc, rec.region

    scg, recg = best_match(ALL_COURTS, district, city, addr_stems)
    if recg and scg >= TH_TAKE_GLOBAL:
        return recg.name, scg, recg.region

    if local_list:
        sc, rec = best_match(local_list, district, city, addr_stems)
        if rec:
            return rec.name, sc, rec.region

    return "Не определён", 0.0, None

# ===================== EXCEL HELPERS =====================
def guess_address_column(df: pd.DataFrame) -> str | None:
    cols = list(df.columns)

    if ADDRESS_COL_PRIMARY in cols:
        return ADDRESS_COL_PRIMARY

    lowmap = {str(c).strip().lower(): c for c in cols}
    if "адрес" in lowmap:
        return lowmap["адрес"]

    for alias in ADDRESS_COL_ALIASES:
        if alias in lowmap:
            return lowmap[alias]

    best_col = None
    best_score = -1
    sample = df.head(150) if len(df) > 150 else df

    for c in cols:
        ser = sample[c].astype(str).fillna("")
        if ser.empty:
            continue

        hits = 0
        for v in ser.tolist():
            t = normalize(v)
            if len(t) < 8:
                continue
            if ("," in v) or (" г." in v.lower()) or ("город" in t) or ("район" in t) or ("аудан" in t) or ("область" in t):
                hits += 1

        score = hits / max(1, len(ser))
        if score > best_score:
            best_score = score
            best_col = c

    if best_score >= 0.25:
        return best_col

    return None

def safe_read_excel(file_path: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось прочитать Excel:\n{e}")
        return None

def make_output_path(file_path: str) -> str:
    base = re.sub(r"\.xlsx$", "", file_path, flags=re.IGNORECASE)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_courts_{ts}.xlsx"

# ===================== PROCESSING =====================
def process_excel(file_path: str, address_col: str):
    df = safe_read_excel(file_path)
    if df is None:
        return
    if df.empty:
        messagebox.showwarning("Пусто", "Файл пустой.")
        return
    if address_col not in df.columns:
        messagebox.showerror(
            "Ошибка",
            f"Столбец '{address_col}' не найден.\nНайденные столбцы:\n{', '.join(map(str, df.columns))}"
        )
        return

    set_status(f"Читаю столбец: {address_col}")
    root.update_idletasks()

    addr_series = df[address_col].fillna("").astype(str)

    set_status("Определяю регионы...")
    root.update_idletasks()
    df["Регион"] = addr_series.apply(detect_region)

    set_status("Определяю суды...")
    root.update_idletasks()
    courts = df.apply(lambda r: detect_court(r[address_col], r["Регион"]), axis=1)

    df["Суд"] = courts.map(lambda x: x[0])

    out_path = make_output_path(file_path)
    try:
        df.to_excel(out_path, index=False)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
        set_status("Ошибка сохранения.")
        return

    set_status("Готово ✅")
    messagebox.showinfo("Готово", f"Файл сохранён:\n{out_path}")

# ===================== GUI =====================
def set_status(text: str):
    status_var.set(text)

def check_json_exists() -> bool:
    try:
        with open(COURTS_JSON_PATH, "r", encoding="utf-8") as _:
            return True
    except Exception:
        return False

def init_courts_or_die():
    global COURTS_BY_REGION, ALL_COURTS, PLACE_TO_REGION

    if not check_json_exists():
        messagebox.showerror("Ошибка", f"Не найден файл: {COURTS_JSON_PATH}")
        raise SystemExit(1)

    try:
        COURTS_BY_REGION, ALL_COURTS = load_courts(COURTS_JSON_PATH)
        PLACE_TO_REGION = build_place_to_region_index(ALL_COURTS)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить справочник судов:\n{e}")
        raise SystemExit(1)

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    df = safe_read_excel(file_path)
    if df is None or df.empty:
        return

    cols = list(df.columns)
    guessed = guess_address_column(df)

    pick_column_window(file_path, cols, guessed)

def pick_column_window(file_path: str, columns: list, guessed: str | None):
    win = tk.Toplevel(root)
    win.title("Выбор столбца адреса")
    win.geometry("560x240")
    win.resizable(False, False)

    lbl = tk.Label(
        win,
        text="Выбери столбец, где находятся адреса (будет обработан именно он):",
        justify="left",
        font=("Arial", 10)
    )
    lbl.pack(pady=12, padx=12, anchor="w")

    frame = tk.Frame(win)
    frame.pack(pady=6, padx=12, fill="x")

    tk.Label(frame, text="Столбец:", font=("Arial", 10)).pack(side="left")

    default_val = guessed if guessed else (ADDRESS_COL_PRIMARY if ADDRESS_COL_PRIMARY in columns else str(columns[0]))
    col_var = tk.StringVar(value=default_val)

    combo = ttk.Combobox(frame, textvariable=col_var, values=[str(c) for c in columns], state="readonly", width=50)
    combo.pack(side="left", padx=10)

    hint = "Авто-определение: " + (f"похоже на «{guessed}»" if guessed else "не уверено, выбери вручную")
    tk.Label(win, text=hint, fg="#555", font=("Arial", 9)).pack(pady=4, padx=12, anchor="w")

    def run():
        chosen = col_var.get()
        win.destroy()
        set_status("Запуск обработки...")
        root.update_idletasks()
        process_excel(file_path, chosen)

    btn_frame = tk.Frame(win)
    btn_frame.pack(pady=14)

    tk.Button(btn_frame, text="Обработать", command=run, width=16, height=2).pack(side="left", padx=8)
    tk.Button(btn_frame, text="Отмена", command=win.destroy, width=16, height=2).pack(side="left", padx=8)

# ===================== MAIN

def main():
    root = tk.Tk()
    root.title("Определение региона и суда по адресу")
    root.geometry("580x260")
    root.resizable(False, False)

    global status_var
    status_var = tk.StringVar(value="Готов к работе.")
    init_courts_or_die()

    info = tk.Label(
        root,
        font=("Arial", 10),
        justify="left"
    )
    info.pack(pady=12, padx=14, anchor="w")

    btn = tk.Button(
        root,
        text="Загрузить Excel с адресами",
        font=("Arial", 10),
        command=open_file,
        width=32,
        height=2
    )
    btn.pack(pady=10)

    status = tk.Label(root, textvariable=status_var, font=("Arial", 9), fg="#333")
    status.pack(pady=6)

    root.mainloop()


if __name__ == "__main__":
    main()
