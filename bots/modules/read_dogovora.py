
import os
import re
import json
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import pdfplumber

# OCR (для сканов)
# pip install pytesseract pymupdf pillow
# + установить Tesseract OCR в ОС и языки rus + kaz (по желанию eng)
try:
    import pytesseract
    import fitz  # PyMuPDF
    from PIL import Image
    import io
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False


# ---------------- НАСТРОЙКИ ----------------
RESULT_COL = "Режим"

# Рекомендуется задавать путь как относительный без "./"
COURTS_JSON_REL = os.path.join("modules", "courts_merged.json")  # файл со списком судов (ключ "СУД")
REGEX_CHUNK_SIZE = 200
OCR_LANG = "rus+kaz+eng"  # если eng не нужен: "rus+kaz"


# ---------------- TESSERACT PATH / ПРОВЕРКА ЯЗЫКОВ ----------------
def configure_tesseract():
    """
    Настраивает путь к tesseract.exe (особенно важно на Windows).
    Приоритет:
      1) переменная окружения TESSERACT_CMD
      2) типовые пути установки на Windows
      3) PATH (если tesseract добавлен)
    """
    if not OCR_AVAILABLE:
        return

    env_cmd = os.environ.get("TESSERACT_CMD", "").strip()
    if env_cmd:
        pytesseract.pytesseract.tesseract_cmd = env_cmd
        return

    # Типовые пути Windows
    possible = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        # иногда ставят в LocalAppData
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
    ]
    for p in possible:
        if p and os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            return
    # Иначе: надеемся на PATH


def check_tesseract_langs(required_langs: str) -> tuple[bool, str]:
    """
    Проверяет, что tesseract видит нужные языки (eng/rus/kaz).
    Возвращает (ok, message).
    """
    if not OCR_AVAILABLE:
        return False, "OCR модули не установлены (pytesseract/pymupdf/pillow)."

    try:
        langs = pytesseract.get_languages(config="")
        need = [x.strip() for x in required_langs.split("+") if x.strip()]
        missing = [x for x in need if x not in langs]
        if missing:
            return False, (
                f"Не найдены языки в Tesseract: {', '.join(missing)}. "
                f"Доступно (первые 30): {', '.join(langs[:30])}"
            )
        return True, f"Языки Tesseract OK: {', '.join(need)}"
    except Exception as e:
        return False, f"Не удалось проверить языки Tesseract: {e}"


if OCR_AVAILABLE:
    configure_tesseract()


# ---------------- РЕГУЛЯРКИ ----------------
ARB_RE = re.compile(r"\b(?:арбитраж\w*|төрелік\w*)\b", re.IGNORECASE)
COURT_WORD_RE = re.compile(r"\b(?:суд|сот)\w*\b", re.IGNORECASE)

VENUE_RE = re.compile(
    r"\b(?:договорн\w*\s+подсудн\w*|шартт\w*\s+сотт\w*)\b",
    re.IGNORECASE
)

COURT_HINT_FAST_RE = re.compile(r"(суд|сот|подсудн|сотт)", re.IGNORECASE)


# ---------------- ВСПОМОГАТЕЛЬНОЕ ----------------
def normalize_text(s: str) -> str:
    """
    Нормализация текста для устойчивого поиска:
    - lower
    - ё -> е
    - 'cуд' (латинская c) -> 'суд'
    - 'coт' (латинская c/o) -> 'сот'
    - сжать пробелы
    """
    if s is None:
        return ""
    s = str(s).lower().replace("ё", "е")
    # латинские двойники
    s = s.replace("cуд", "суд")
    s = s.replace("coт", "сот")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def resolve_courts_json_path() -> str:
    """
    Абсолютный путь к справочнику судов относительно папки скрипта.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(script_dir, COURTS_JSON_REL)


def load_court_names(json_path: str) -> list[str]:
    """
    Загружает список судов из JSON.
    Ожидается: массив объектов, у каждого ключ "СУД".
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("courts_merged.json должен быть массивом объектов (list).")

    names: list[str] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        name = item.get("СУД")
        if name:
            names.append(normalize_text(name))

    # убираем дубли, сортируем длинные сначала (уменьшает ложные совпадения)
    names = sorted({n for n in names if n}, key=len, reverse=True)
    return names


def build_court_regex_chunks(court_names: list[str], chunk_size: int = 200) -> list[re.Pattern]:
    """
    Компилируем регулярки чанками, чтобы не делать одну гигантскую.
    Ищем как подстроку, т.к. в PDF бывают переносы/знаки.
    """
    patterns: list[re.Pattern] = []
    if not court_names:
        return patterns

    for i in range(0, len(court_names), chunk_size):
        chunk = court_names[i:i + chunk_size]
        if not chunk:
            continue
        joined = "|".join(re.escape(x) for x in chunk)
        patterns.append(re.compile(joined, re.IGNORECASE))
    return patterns


# ---------------- ЗАГРУЗКА СУДОВ (ОДИН РАЗ) ----------------
COURT_NAMES: list[str] = []
COURT_PATTERNS: list[re.Pattern] = []
COURTS_READY_ERROR: str | None = None

try:
    courts_path = resolve_courts_json_path()
    COURT_NAMES = load_court_names(courts_path)
    COURT_PATTERNS = build_court_regex_chunks(COURT_NAMES, chunk_size=REGEX_CHUNK_SIZE)
except Exception as e:
    COURTS_READY_ERROR = f"Не удалось загрузить {COURTS_JSON_REL}: {e}"


def text_contains_court_by_list(norm_text: str) -> bool:
    """
    Проверка по точному списку судов из JSON.
    Быстрая отсечка: если нет намёка на суд/подсудность — False.
    """
    if not COURT_HINT_FAST_RE.search(norm_text):
        return False
    for pat in COURT_PATTERNS:
        if pat.search(norm_text):
            return True
    return False


# ---------------- OCR ----------------
def ocr_page_text_from_doc(doc, page_index: int) -> str:
    """
    OCR одной страницы через PyMuPDF -> PIL image -> pytesseract.
    doc = fitz.Document уже открыт.
    """
    if not OCR_AVAILABLE:
        return ""

    page = doc.load_page(page_index)

    # масштаб: 2.0-3.0 обычно ок (чем больше, тем медленнее, но точнее)
    mat = fitz.Matrix(2.5, 2.5)
    pix = page.get_pixmap(matrix=mat, alpha=False)

    img_bytes = pix.tobytes("png")
    img = Image.open(io.BytesIO(img_bytes))

    # лёгкая предобработка
    img = img.convert("L")  # grayscale

    config = "--oem 3 --psm 6"
    txt = pytesseract.image_to_string(img, lang=OCR_LANG, config=config)
    return txt or ""


def looks_like_garbage(norm_txt: str) -> bool:
    """
    Определяет, что "извлечённый текст" похож на мусор:
    - пусто
    - слишком мало букв
    """
    if not norm_txt:
        return True
    letters = sum(ch.isalpha() for ch in norm_txt)
    return letters < 10


def pdf_iter_page_text(pdf_path: str, use_ocr_for_scans: bool) -> list[str]:
    """
    Возвращает список текста по страницам.
    Сначала пытается вытащить текст обычным способом.
    Если текст пустой/мусорный и включен OCR — делаем OCR этой страницы.
    """
    texts: list[str] = []

    doc = None
    if use_ocr_for_scans and OCR_AVAILABLE:
        try:
            doc = fitz.open(pdf_path)
        except Exception:
            doc = None  # OCR просто будет недоступен для этого файла

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                txt = page.extract_text() or ""
                norm_txt = normalize_text(txt)

                if looks_like_garbage(norm_txt):
                    if use_ocr_for_scans and OCR_AVAILABLE and doc is not None:
                        ocr_txt = ocr_page_text_from_doc(doc, i)
                        texts.append(ocr_txt)
                    else:
                        texts.append("")
                else:
                    texts.append(txt)
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass

    return texts


# ---------------- ОСНОВНАЯ ЛОГИКА PDF ----------------
def pdf_mode(pdf_path: str, use_ocr_for_scans: bool) -> str:
    """
    Для НЕ-банка:
    - если где-то в PDF есть Арбитраж/Төрелік -> "Арбитраж" (приоритет)
    - иначе если есть:
        a) договорная подсудность / шарттық соттылық, ИЛИ
        b) суд/сот, ИЛИ
        c) название суда из courts_merged.json (ключ "СУД")
      -> "Договорная"
    - иначе -> ""
    """
    has_arbitrage = False
    has_contract_venue = False

    page_texts = pdf_iter_page_text(pdf_path, use_ocr_for_scans=use_ocr_for_scans)

    for txt in page_texts:
        if not txt.strip():
            continue

        norm = normalize_text(txt)

        if ARB_RE.search(norm):
            has_arbitrage = True

        if not has_contract_venue:
            if VENUE_RE.search(norm) or COURT_WORD_RE.search(norm):
                has_contract_venue = True
            else:
                if (COURTS_READY_ERROR is None) and text_contains_court_by_list(norm):
                    has_contract_venue = True

    if has_arbitrage:
        return "Арбитраж"
    if has_contract_venue:
        return "Договорная"
    return ""


# ---------------- EXCEL ОБРАБОТКА ----------------
def process_excel(input_xlsx: str, output_xlsx: str, log_fn, progress_fn, use_ocr_for_scans: bool):
    """
    progress_fn(percent:int, done:int, left:int, total:int)
    """
    df = pd.read_excel(input_xlsx)

    if df.shape[1] < 3:
        raise ValueError("В Excel должно быть минимум 3 столбца по порядку: ВНД | Кредитор | Путь к договору.")

    col_creditor = df.columns[1]
    col_path = df.columns[2]

    if RESULT_COL not in df.columns:
        df[RESULT_COL] = ""

    if COURTS_READY_ERROR:
        log_fn(
            f"ВНИМАНИЕ: список судов не загружен ({COURTS_READY_ERROR}). "
            f"Будем искать по 'суд/сот' и маркерам подсудности."
        )
    else:
        log_fn(f"Список судов загружен: {len(COURT_NAMES)} шт. ({COURTS_JSON_REL})")

    if use_ocr_for_scans:
        if not OCR_AVAILABLE:
            log_fn("ВНИМАНИЕ: OCR включен, но зависимости не установлены (pytesseract/pymupdf/pillow). OCR отключен.")
            use_ocr_for_scans = False
        else:
            ok, msg = check_tesseract_langs(OCR_LANG)
            log_fn(msg)
            if not ok:
                log_fn("ВНИМАНИЕ: не хватает языков Tesseract. Поставь языки (rus/kaz/eng) или поменяй OCR_LANG.")

    total = len(df)
    progress_fn(0, 0, total, total)

    for idx, row in df.iterrows():
        creditor_raw = normalize_text(row[col_creditor])
        pdf_path = str(row[col_path]).strip()

        done = idx + 1
        left = total - done
        percent = int(done / max(total, 1) * 100)
        progress_fn(percent, done, left, total)

        if "банк" in creditor_raw:
            df.at[idx, RESULT_COL] = "Общеустановленная"
        else:
            try:
                if not pdf_path:
                    df.at[idx, RESULT_COL] = "[ОШИБКА: пустой путь к PDF]"
                elif not os.path.exists(pdf_path):
                    df.at[idx, RESULT_COL] = "[ОШИБКА: PDF не найден]"
                else:
                    df.at[idx, RESULT_COL] = pdf_mode(pdf_path, use_ocr_for_scans=use_ocr_for_scans)
            except Exception as e:
                df.at[idx, RESULT_COL] = f"[ОШИБКА PDF: {e}]"

        if (idx + 1) % 10 == 0 or (idx + 1) == total:
            log_fn(f"Обработано: {idx + 1}/{total}")

    df.to_excel(output_xlsx, index=False)
    log_fn(f"Готово. Сохранено: {output_xlsx}")


# ---------------- UI (ОКНО) ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Робот договоров (Excel → PDF)")
        self.geometry("860x470")
        self.resizable(False, False)

        self.file_path = tk.StringVar(value="")
        self.use_ocr = tk.BooleanVar(value=True)
        self.lang_hint = tk.StringVar(value=OCR_LANG)

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(
            frm,
            text="Выбери Excel файл (3 столбца по порядку: ВНД | Кредитор | Путь к договору):"
        ).pack(anchor="w")

        row = ttk.Frame(frm)
        row.pack(fill="x", pady=8)

        ttk.Entry(row, textvariable=self.file_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="Выбрать…", command=self.pick_file).pack(side="left", padx=8)

        opts = ttk.Frame(frm)
        opts.pack(fill="x", pady=6)

        ttk.Checkbutton(
            opts,
            text="OCR для сканов (если страница без текста) — RU/KZ/ENG",
            variable=self.use_ocr
        ).pack(side="left")

        ttk.Label(opts, textvariable=self.lang_hint).pack(side="left", padx=10)

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill="x", pady=8)

        self.run_btn = ttk.Button(btn_row, text="Запустить", command=self.run)
        self.run_btn.pack(side="left")

        self.progress = ttk.Progressbar(btn_row, orient="horizontal", length=520, mode="determinate")
        self.progress.pack(side="left", padx=12)

        self.counter_var = tk.StringVar(value="Обработано: 0/0 | Осталось: 0")
        self.counter_lbl = ttk.Label(btn_row, textvariable=self.counter_var)
        self.counter_lbl.pack(side="left")

        ttk.Label(frm, text="Лог:").pack(anchor="w", pady=(10, 0))
        self.log = tk.Text(frm, height=16, wrap="word", state="disabled")
        self.log.pack(fill="both", expand=True)

        self.status = ttk.Label(frm, text="Готов к работе.")
        self.status.pack(anchor="w", pady=(8, 0))

        # Стартовые сообщения
        if COURTS_READY_ERROR:
            self.append_log("ВНИМАНИЕ: список судов не загружен.")
            self.append_log(COURTS_READY_ERROR)
        else:
            self.append_log(f"Суды загружены: {len(COURT_NAMES)} шт. ({COURTS_JSON_REL})")

        if self.use_ocr.get():
            if OCR_AVAILABLE:
                ok, msg = check_tesseract_langs(OCR_LANG)
                self.append_log(f"OCR доступен. Языки: {OCR_LANG}")
                self.append_log(msg)
                if not ok:
                    self.append_log(
                        "Подсказка: установи Tesseract и языки rus/kaz/eng "
                        "или задай путь через переменную окружения TESSERACT_CMD."
                    )
                self.append_log(f"Tesseract cmd: {getattr(pytesseract.pytesseract, 'tesseract_cmd', 'PATH')}")
            else:
                self.append_log("OCR НЕ доступен: установи pytesseract + pymupdf + pillow и Tesseract OCR (rus/kaz/eng).")

    def pick_file(self):
        path = filedialog.askopenfilename(
            title="Выбери Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.file_path.set(path)

    def append_log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def set_progress(self, percent: int, done: int = 0, left: int = 0, total: int = 0):
        self.progress["value"] = max(0, min(100, percent))
        self.counter_var.set(f"Обработано: {done}/{total} | Осталось: {left}")
        self.update_idletasks()

    def run(self):
        input_xlsx = self.file_path.get().strip()
        if not input_xlsx or not os.path.exists(input_xlsx):
            messagebox.showerror("Ошибка", "Выбери существующий Excel файл.")
            return

        base, _ = os.path.splitext(input_xlsx)
        output_xlsx = base + "_output.xlsx"

        use_ocr_for_scans = bool(self.use_ocr.get())

        self.run_btn.configure(state="disabled")
        self.set_progress(0, 0, 0, 0)
        self.status.configure(text="В работе…")
        self.append_log(f"Старт: {input_xlsx}")
        self.append_log(f"OCR для сканов: {'ВКЛ' if use_ocr_for_scans else 'ВЫКЛ'}")

        def worker():
            try:
                process_excel(
                    input_xlsx=input_xlsx,
                    output_xlsx=output_xlsx,
                    log_fn=lambda m: self.after(0, self.append_log, m),
                    progress_fn=lambda percent, done, left, total: self.after(
                        0, self.set_progress, percent, done, left, total
                    ),
                    use_ocr_for_scans=use_ocr_for_scans
                )
                # ВАЖНО: configure вызываем с kwargs, не с dict позиционно
                self.after(0, self.status.configure, text="Готово ✅")
                self.after(0, messagebox.showinfo, "Готово", f"Сохранено:\n{output_xlsx}")
            except Exception as e:
                self.after(0, self.status.configure, text="Ошибка ❌")
                self.after(0, messagebox.showerror, "Ошибка", str(e))
            finally:
                self.after(0, self.run_btn.configure, state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
