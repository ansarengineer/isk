from __future__ import annotations

import os
import re
import sys
import json
import time
import hashlib
import logging
import traceback
from dataclasses import dataclass
from typing import Optional, Tuple, List, Dict, Any

import pandas as pd

# PDF extraction
import fitz  # PyMuPDF

# OCR
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import pytesseract

# UI
import tkinter as tk
from tkinter import filedialog, messagebox

# Parallelism
from concurrent.futures import ProcessPoolExecutor, as_completed


# ----------------------------- НАСТРОЙКИ -----------------------------

# Если на Windows Tesseract не находится — укажите путь:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# OCR языки: сначала пробуем rus+kaz+eng (если kaz не установлен, tesseract может ругаться — тогда fallback).
OCR_LANG_PRIMARY = "rus+kaz+eng"
OCR_LANG_FALLBACK = "rus+eng"

MAX_OCR_PAGES = 60
MIN_TEXT_CHARS_TO_SKIP_OCR = 350
OCR_DPI = 300

# Параллелизм
DEFAULT_WORKERS = max(1, (os.cpu_count() or 2) - 1)

# Кэш
ENABLE_DISK_CACHE = True
CACHE_FILENAME = "pdf_text_cache.json"  # будет рядом с исходным Excel
MAX_CACHE_ITEMS = 20000  # чтобы не разрастался бесконечно

# Ограничение размера сохраняемого текста в кэше (если договоры гигантские)
CACHE_TEXT_MAX_CHARS = 350_000

# Колонки Excel
REQUIRED_COLS = ["ВНД", "Кредитор", "Путь к договору", "Тип договора"]
TYPE_TARGET = "Договорная"

# Чтобы колонка "Фрагмент_где_найдено" была читабельной
FRAGMENT_MAX_CHARS = 600


# ----------------------------- СЛОВАРИ (РК) -----------------------------

# Города/регионы — не идеально, но сильно помогает скорингу и распознаванию
KZ_CITIES = {
    "алматы", "астана", "шымкент", "актобе", "атырау", "актау", "конаев", "талдыкорган",
    "павлодар", "усть-каменогорск", "өскемен", "семей", "караганда", "қазақстан",
    "костанай", "қостанай", "петропавловск", "петропавл", "кокшетау", "қокшетау",
    "тараз", "туркестан", "орал", "уральск", "кызылорда", "қызылорда",
    "жезказган", "екібастұз", "экибастуз", "рудный", "темиртау",
}

KZ_REGIONS = {
    "алматинской области", "алматинская область", "акмолинской области", "акмолинская область",
    "актюбинской области", "актюбинская область", "атырауской области", "атырауская область",
    "восточно-казахстанской области", "вко", "жамбылской области", "жамбылская область",
    "западно-казахстанской области", "зко", "караганды", "караганды области", "карагандинской области",
    "костанайской области", "кызылординской области", "мангистаускои области", "мангистауская область",
    "павлодарской области", "северо-казахстанской области", "ско",
    "туркестанской области", "ұлытауской области", "улытауской области",
    "абайской области", "жетысуской области", "жетісуской области",
}

KZ_DISTRICTS_ALMATY = {
    "медеуский", "бостандыкский", "аэзовский", "алмалинский", "турксибский", "наурызбайский", "жетауский"
}

KZ_DISTRICTS_ASTANA = {
    "байконур", "сарыарка", "есиль", "алматы", "нура"
}


# ----------------------------- МОДЕЛЬ -----------------------------

@dataclass
class ExtractResult:
    full_text: str
    court_or_venue: str
    confidence: str  # high | medium | low
    notes: str
    fragment: str  # кусок текста, где нашли


# ----------------------------- УТИЛИТЫ -----------------------------

def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.replace("\u00A0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\r\n|\r", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def safe_lower(s: str) -> str:
    return (s or "").lower()


def short_fragment(s: str, max_chars: int = FRAGMENT_MAX_CHARS) -> str:
    s = normalize_text(s)
    if len(s) <= max_chars:
        return s
    return s[:max_chars] + "..."


def file_sig(path: str) -> Tuple[str, int, int]:
    """
    Сигнатура файла для кэша: abs_path, mtime, size
    """
    ap = os.path.abspath(path)
    st = os.stat(ap)
    return ap, int(st.st_mtime), int(st.st_size)


def cache_key(sig: Tuple[str, int, int]) -> str:
    ap, mtime, size = sig
    raw = f"{ap}|{mtime}|{size}".encode("utf-8", errors="ignore")
    return hashlib.sha256(raw).hexdigest()


# ----------------------------- ЛОГГИНГ -----------------------------

def setup_logger(log_path: str) -> logging.Logger:
    logger = logging.getLogger("pdf_court_bot")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    return logger


# ----------------------------- КЭШ -----------------------------

def load_disk_cache(cache_path: str) -> Dict[str, Any]:
    if not os.path.exists(cache_path):
        return {}
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        return {}
    return {}


def save_disk_cache(cache_path: str, cache: Dict[str, Any]) -> None:
    # подрежем, если слишком большой
    if len(cache) > MAX_CACHE_ITEMS:
        # примитивная чистка: оставим последние N по "ts"
        items = list(cache.items())
        items.sort(key=lambda kv: kv[1].get("ts", 0), reverse=True)
        cache = dict(items[:MAX_CACHE_ITEMS])

    tmp = cache_path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)
    os.replace(tmp, cache_path)


# ----------------------------- PDF -> TEXT -----------------------------

def extract_text_from_pdf_native(pdf_path: str) -> str:
    doc = fitz.open(pdf_path)
    parts: List[str] = []
    try:
        for page in doc:
            parts.append(page.get_text("text"))
    finally:
        doc.close()
    return normalize_text("\n".join(parts))


def preprocess_for_ocr(img: Image.Image) -> Image.Image:
    """
    Предобработка для OCR:
    - grayscale
    - autocontrast
    - усиление контраста
    - резкость
    - binarize (порог)
    """
    # grayscale
    g = img.convert("L")

    # autocontrast
    g = ImageOps.autocontrast(g)

    # contrast boost
    g = ImageEnhance.Contrast(g).enhance(1.8)

    # slight sharpen
    g = g.filter(ImageFilter.UnsharpMask(radius=2, percent=180, threshold=3))

    # binarize
    # adaptive порога без numpy: простой глобальный порог
    # (на практике это уже ощутимо помогает)
    threshold = 185
    g = g.point(lambda p: 255 if p > threshold else 0)

    return g


def ocr_pdf(pdf_path: str, max_pages: int = MAX_OCR_PAGES) -> str:
    doc = fitz.open(pdf_path)
    parts: List[str] = []
    try:
        total = doc.page_count
        pages_to_do = min(total, max_pages) if max_pages else total

        zoom = OCR_DPI / 72.0
        mat = fitz.Matrix(zoom, zoom)

        # Попробуем primary язык, если упадёт — fallback
        lang_to_try = [OCR_LANG_PRIMARY, OCR_LANG_FALLBACK]

        for i in range(pages_to_do):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = preprocess_for_ocr(img)

            text_i = None
            last_err = None
            for lang in lang_to_try:
                try:
                    text_i = pytesseract.image_to_string(img, lang=lang)
                    break
                except Exception as e:
                    last_err = e
                    continue

            if text_i is None:
                raise RuntimeError(f"OCR failed: {last_err}")

            parts.append(text_i)

    finally:
        doc.close()

    return normalize_text("\n".join(parts))


def extract_pdf_text_smart(pdf_path: str) -> Tuple[str, str]:
    if not pdf_path or not os.path.exists(pdf_path):
        return "", "PDF path missing or not found"

    notes = []
    native = ""
    try:
        native = extract_text_from_pdf_native(pdf_path)
        notes.append(f"native_chars={len(native)}")
    except Exception as e:
        notes.append(f"native_failed={type(e).__name__}:{e}")

    if len(native) >= MIN_TEXT_CHARS_TO_SKIP_OCR:
        return native, "; ".join(notes)

    try:
        ocr = ocr_pdf(pdf_path)
        notes.append(f"ocr_chars={len(ocr)}")
        combined = normalize_text((native + "\n\n" + ocr).strip())
        return combined, "; ".join(notes)
    except Exception as e:
        notes.append(f"ocr_failed={type(e).__name__}:{e}")
        return native, "; ".join(notes)


# ----------------------------- СЕКЦИИ / ЗАГОЛОВКИ -----------------------------

HEADING_MARKERS = [
    "РАЗРЕШЕНИЕ СПОРОВ",
    "ПОРЯДОК РАЗРЕШЕНИЯ СПОРОВ",
    "ПОДСУДНОСТЬ",
    "АРБИТРАЖ",
    "ТРЕТЕЙСК",
    "ЮРИСДИКЦ",
    "СПОРЫ",
]

def is_heading_line(line: str) -> bool:
    """
    Определяем "заголовок" по:
    - много букв в верхнем регистре
    - длина разумная
    - содержит маркер (или полностью uppercase и коротко)
    """
    l = line.strip()
    if not l:
        return False

    up = l.upper()
    if any(m in up for m in HEADING_MARKERS):
        return True

    # эвристика: "всё капсом" и мало знаков
    letters = sum(ch.isalpha() for ch in l)
    upper_letters = sum(ch.isalpha() and ch == ch.upper() for ch in l)
    if letters >= 6 and upper_letters / max(letters, 1) > 0.85 and len(l) <= 60:
        return True

    return False


def extract_sections(text: str) -> List[Tuple[str, str]]:
    """
    Режем текст на секции по заголовкам.
    Возвращает список (heading, body)
    """
    t = normalize_text(text)
    if not t:
        return []

    lines = t.split("\n")
    sections: List[Tuple[str, List[str]]] = []
    current_heading = "FULL"
    current_body: List[str] = []

    for line in lines:
        if is_heading_line(line):
            # сохранить прошлую секцию
            if current_body:
                sections.append((current_heading, current_body))
            current_heading = normalize_text(line)
            current_body = []
        else:
            current_body.append(line)

    if current_body:
        sections.append((current_heading, current_body))

    # склеиваем
    out: List[Tuple[str, str]] = []
    for h, body_lines in sections:
        body = normalize_text("\n".join(body_lines))
        out.append((h, body))
    return out


def build_relevant_text(text: str) -> str:
    """
    Берём:
    - все секции, где заголовок содержит маркеры
    - плюс немного текста вокруг (fallback)
    """
    sections = extract_sections(text)
    if not sections:
        return text

    picked: List[str] = []
    for h, body in sections:
        hup = h.upper()
        if any(m in hup for m in HEADING_MARKERS):
            if body:
                picked.append(f"{h}\n{body}")

    if picked:
        return normalize_text("\n\n----\n\n".join(picked))

    # если не нашли заголовки-маркеры — возвращаем полный текст
    return text


# ----------------------------- ПОИСК СУДА / ПОДСУДНОСТИ -----------------------------

def score_court_candidate(s: str) -> int:
    """
    Скоринг кандидата: чем больше "конкретики" (город, район, область) — тем лучше.
    """
    low = safe_lower(s)
    sc = len(s)

    # суд/арбитраж — must-have
    if "суд" in low:
        sc += 80
    if "третей" in low or "арбитраж" in low:
        sc += 60

    # география
    for city in KZ_CITIES:
        if city in low:
            sc += 40
    for reg in KZ_REGIONS:
        if reg in low:
            sc += 35

    # районы Алматы/Астаны
    for d in KZ_DISTRICTS_ALMATY:
        if d in low:
            sc += 35
    for d in KZ_DISTRICTS_ASTANA:
        if d in low:
            sc += 25

    # наказание за мусор
    if "настоящ" in low and "договор" in low and "суд" not in low:
        sc -= 60

    # бонус если выглядит как “... районный суд ... города ...”
    if "район" in low and ("город" in low or "г." in low):
        sc += 40

    return sc


def find_court_or_venue(full_text: str) -> ExtractResult:
    text = normalize_text(full_text)
    if not text:
        return ExtractResult(full_text="", court_or_venue="", confidence="low", notes="empty_text", fragment="")

    relevant = normalize_text(build_relevant_text(text))
    w = relevant

    # --- 1) Явное название суда (сразу high) ---
    court_patterns = [
        # "в ... суде ..."
        r"(?:в|во)\s+((?:[А-ЯЁ][а-яё]+\s+){0,7}(?:районн(?:ом|ый)|городск(?:ом|ой)|областн(?:ом|ой)|межрайонн(?:ом|ой)|специализированн(?:ом|ый)|экономическ(?:ом|ий)|административн(?:ом|ый)|гражданск(?:ом|ий)|уголовн(?:ом|ый)|апелляционн(?:ом|ый)|верховн(?:ом|ый))?\s*(?:суд|суде)\s*(?:[^,.;\n]{0,180}))",
        # "в районном суде по месту..."
        r"(?:в|во)\s+((?:районн(?:ом|ый)\s+)?(?:суд|суде)\s+по\s+месту\s+[^,.;\n]{0,140})",
        # Третейский/Арбитражный суд
        r"((?:Постоянно\s+действующ(?:ий|его)\s+)?(?:Третейск(?:ий|ого)|Арбитражн(?:ый|ого))\s+(?:суд|суда)\s*(?:[^,.;\n]{0,220}))",
        # СМЭС / специализированный межрайонный
        r"((?:Специализированн(?:ый|ого)\s+)?(?:межрайонн(?:ый|ого)\s+)?(?:суд|суда)\s*(?:[^,.;\n]{0,260}))",
    ]

    candidates: List[Tuple[str, str]] = []  # (candidate, fragment)
    for pat in court_patterns:
        for m in re.finditer(pat, w, flags=re.IGNORECASE | re.UNICODE):
            cand = normalize_text(m.group(1))
            if len(cand) < 10:
                continue
            # фрагмент — возьмём окружение
            span_start, span_end = m.span(1)
            frag = w[max(0, span_start - 200): min(len(w), span_end + 200)]
            candidates.append((cand, frag))

    # dedup (по lower нормализованному)
    deduped: Dict[str, Tuple[str, str]] = {}
    for cand, frag in candidates:
        key = re.sub(r"\s+", " ", cand.lower()).strip()
        if key not in deduped:
            deduped[key] = (cand, frag)

    if deduped:
        best_key = max(deduped.keys(), key=lambda k: score_court_candidate(deduped[k][0]))
        best_cand, best_frag = deduped[best_key]
        return ExtractResult(
            full_text=text,
            court_or_venue=best_cand,
            confidence="high",
            notes="explicit_court_found",
            fragment=short_fragment(best_frag),
        )

    # --- 2) Подсудность “по месту ...” ---
    venue_patterns = [
        r"(по месту нахождения\s+[^,.;\n]{0,120})",
        r"(по месту жительства\s+[^,.;\n]{0,120})",
        r"(по месту проживания\s+[^,.;\n]{0,120})",
        r"(по месту регистрации\s+[^,.;\n]{0,120})",
        r"(по месту нахождения\s+ответчика)",
        r"(по месту жительства\s+ответчика)",
        r"(по месту проживания\s+ответчика)",
        r"(по месту нахождения\s+(?:МФО|Займодателя|Кредитора|Заимодавца|Займодавца))",
        r"(по месту нахождения\s+(?:заемщика|заёмщика))",
        r"(по месту жительства\s+(?:заемщика|заёмщика))",
        r"(по месту нахождения\s+истца)",
        r"(по месту нахождения\s+ответчика)",
    ]

    for pat in venue_patterns:
        m = re.search(pat, w, flags=re.IGNORECASE | re.UNICODE)
        if m:
            hit = normalize_text(m.group(1) if m.lastindex else m.group(0))
            frag = w[max(0, m.start() - 250): min(len(w), m.end() + 250)]
            return ExtractResult(
                full_text=text,
                court_or_venue=hit,
                confidence="medium",
                notes="venue_phrase_found",
                fragment=short_fragment(frag),
            )

    # --- 3) Fallback: “предложение со словом суд” ---
    sent_like = re.split(r"(?<=[.!?])\s+|\n+", w)
    sent_like = [normalize_text(s) for s in sent_like if "суд" in safe_lower(s)]
    sent_like = [s for s in sent_like if len(s) > 15]

    if sent_like:
        best = max(sent_like, key=len)
        best_short = short_fragment(best, 300)
        return ExtractResult(
            full_text=text,
            court_or_venue=best_short,
            confidence="low",
            notes="fallback_sentence_with_sud",
            fragment=best_short,
        )

    return ExtractResult(
        full_text=text,
        court_or_venue="",
        confidence="low",
        notes="not_found",
        fragment="",
    )


# ----------------------------- WORKER ДЛЯ ПАРАЛЛЕЛИЗМА -----------------------------

def worker_process_pdf(pdf_path: str) -> Dict[str, str]:
    """
    Worker обязан возвращать только сериализуемое.
    """
    text, notes_extract = extract_pdf_text_smart(pdf_path)
    res = find_court_or_venue(text)
    return {
        "pdf_path": pdf_path,
        "text": res.full_text,
        "court_or_venue": res.court_or_venue,
        "confidence": res.confidence,
        "notes": f"{res.notes}; {notes_extract}",
        "fragment": res.fragment,
    }


# ----------------------------- EXCEL PIPELINE -----------------------------

def ensure_required_columns(df: pd.DataFrame) -> None:
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"В Excel не найдены обязательные колонки: {missing}")


def insert_column_after(df: pd.DataFrame, after_col: str, new_col: str, default_value=None) -> pd.DataFrame:
    if new_col in df.columns:
        return df
    cols = list(df.columns)
    if after_col not in cols:
        df[new_col] = default_value
        return df
    idx = cols.index(after_col) + 1
    cols.insert(idx, new_col)
    df[new_col] = default_value
    return df[cols]


def process_excel(excel_path: str) -> str:
    # Папка для результата/кэша/лога
    base_dir = os.path.dirname(os.path.abspath(excel_path))
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    ts = time.strftime("%Y%m%d_%H%M%S")

    out_path = os.path.join(base_dir, f"{base_name}_with_courts_{ts}.xlsx")
    log_path = os.path.join(base_dir, f"{base_name}_with_courts_{ts}.log")
    cache_path = os.path.join(base_dir, CACHE_FILENAME)

    logger = setup_logger(log_path)
    logger.info("Start processing: %s", excel_path)
    logger.info("Workers: %s", DEFAULT_WORKERS)

    disk_cache: Dict[str, Any] = load_disk_cache(cache_path) if (ENABLE_DISK_CACHE) else {}
    mem_cache: Dict[str, Any] = {}

    df = pd.read_excel(excel_path, dtype=str).fillna("")
    ensure_required_columns(df)

    # Колонки результата
    if "Текст_из_PDF" not in df.columns:
        df["Текст_из_PDF"] = ""

    df = insert_column_after(df, "Тип договора", "Суд_или_подсудность", default_value="")
    df = insert_column_after(df, "Суд_или_подсудность", "Качество_поиска", default_value="")
    df = insert_column_after(df, "Качество_поиска", "Фрагмент_где_найдено", default_value="")
    df = insert_column_after(df, "Фрагмент_где_найдено", "Тех_заметки", default_value="")

    # Только "Договорная"
    mask = df["Тип договора"].astype(str).str.strip().str.lower() == TYPE_TARGET.lower()
    indices = list(df.index[mask])
    total = len(indices)

    if total == 0:
        raise ValueError(f'Нет строк, где "Тип договора" == "{TYPE_TARGET}".')

    # Список заданий: будем параллелить только уникальные pdf_path
    idx_to_pdf: Dict[int, str] = {}
    unique_pdf: Dict[str, List[int]] = {}  # pdf_path -> список индексов, где он встречается

    for idx in indices:
        pdf_path = str(df.at[idx, "Путь к договору"]).strip()
        idx_to_pdf[idx] = pdf_path
        unique_pdf.setdefault(pdf_path, []).append(idx)

    logger.info("Rows to process: %d", total)
    logger.info("Unique PDF paths: %d", len(unique_pdf))

    # Сначала попробуем закрыть часть через кэш
    results_by_pdf: Dict[str, Dict[str, str]] = {}
    to_compute: List[str] = []

    for pdf_path in unique_pdf.keys():
        if not pdf_path or not os.path.exists(pdf_path):
            results_by_pdf[pdf_path] = {
                "pdf_path": pdf_path,
                "text": "",
                "court_or_venue": "",
                "confidence": "low",
                "notes": "PDF path missing or not found",
                "fragment": "",
            }
            continue

        sig = file_sig(pdf_path)
        key = cache_key(sig)

        # mem_cache
        if key in mem_cache:
            results_by_pdf[pdf_path] = mem_cache[key]
            continue

        # disk_cache
        if ENABLE_DISK_CACHE and key in disk_cache:
            results_by_pdf[pdf_path] = disk_cache[key]
            mem_cache[key] = disk_cache[key]
            continue

        to_compute.append(pdf_path)

    logger.info("From cache: %d PDFs", len(unique_pdf) - len(to_compute))
    logger.info("To compute: %d PDFs", len(to_compute))

    # Параллельная обработка
    done = 0
    if to_compute:
        with ProcessPoolExecutor(max_workers=DEFAULT_WORKERS) as ex:
            fut_map = {ex.submit(worker_process_pdf, p): p for p in to_compute}

            for fut in as_completed(fut_map):
                pdf_path = fut_map[fut]
                done += 1
                try:
                    r = fut.result()
                except Exception as e:
                    logger.error("Failed PDF: %s | %s", pdf_path, e)
                    r = {
                        "pdf_path": pdf_path,
                        "text": "",
                        "court_or_venue": "",
                        "confidence": "low",
                        "notes": f"worker_failed: {type(e).__name__}:{e}",
                        "fragment": "",
                    }

                results_by_pdf[pdf_path] = r

                # обновим кэш
                if pdf_path and os.path.exists(pdf_path):
                    sig = file_sig(pdf_path)
                    key = cache_key(sig)

                    # подрежем текст для кэша, если слишком большой
                    text_to_store = r.get("text", "")
                    if len(text_to_store) > CACHE_TEXT_MAX_CHARS:
                        text_to_store = text_to_store[:CACHE_TEXT_MAX_CHARS] + "\n...<TRUNCATED_FOR_CACHE>..."
                    stored = dict(r)
                    stored["text"] = text_to_store
                    stored["ts"] = int(time.time())

                    mem_cache[key] = stored
                    if ENABLE_DISK_CACHE:
                        disk_cache[key] = stored

                if done % 5 == 0 or done == len(to_compute):
                    logger.info("Progress PDFs computed: %d/%d", done, len(to_compute))

    # Запись результатов обратно в строки
    for pdf_path, idx_list in unique_pdf.items():
        r = results_by_pdf.get(pdf_path, None)
        if r is None:
            r = {
                "text": "",
                "court_or_venue": "",
                "confidence": "low",
                "notes": "missing_result",
                "fragment": "",
            }
        for idx in idx_list:
            df.at[idx, "Текст_из_PDF"] = r.get("text", "")
            df.at[idx, "Суд_или_подсудность"] = r.get("court_or_venue", "")
            df.at[idx, "Качество_поиска"] = r.get("confidence", "")
            df.at[idx, "Фрагмент_где_найдено"] = r.get("fragment", "")
            df.at[idx, "Тех_заметки"] = r.get("notes", "")

    # Сохранение Excel
    df.to_excel(out_path, index=False)
    logger.info("Saved: %s", out_path)
    logger.info("Log: %s", log_path)

    # Сохранение кэша на диск
    if ENABLE_DISK_CACHE:
        try:
            save_disk_cache(cache_path, disk_cache)
            logger.info("Cache saved: %s (items=%d)", cache_path, len(disk_cache))
        except Exception as e:
            logger.error("Cache save failed: %s", e)

    return out_path


# ----------------------------- UI -----------------------------

def choose_excel_file() -> Optional[str]:
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Выбор Excel",
        "Выберите Excel-файл со столбцами:\n"
        "1) ВНД\n2) Кредитор\n3) Путь к договору\n4) Тип договора\n\n"
        f"Робот обработает строки, где Тип договора = '{TYPE_TARGET}'."
    )

    file_path = filedialog.askopenfilename(
        title="Выберите Excel файл",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    root.destroy()
    return file_path or None


def main():
    try:
        excel_path = choose_excel_file()
        if not excel_path:
            print("Файл не выбран. Выход.")
            return

        out_path = process_excel(excel_path)

        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Готово", f"Готово!\nФайл сохранён:\n{out_path}")
        root.destroy()

    except Exception as e:
        err = f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}"
        print(err)

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Ошибка", f"Ошибка:\n{type(e).__name__}: {e}")
        root.destroy()


if __name__ == "__main__":
    # ВАЖНО для Windows multiprocessing
    main()
