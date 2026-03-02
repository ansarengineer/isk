# Запуск - streamlit run app.py
from __future__ import annotations

import io
import sys
import zipfile
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

import streamlit as st

APP_TITLE = "ISK builder • local website • by Ansar"
BASE_DIR = Path(__file__).resolve().parent
BOTS_DIR = BASE_DIR / "bots" / "modules"
TEMPLATES_DIR = BASE_DIR / "templates"
EXAMPLE_XLSX = BASE_DIR / "example" / "example.xlsx"
ASSETS_DIR = BASE_DIR / "assets"


# -------------------- UI --------------------
def inject_css() -> None:
    """
    Minimal, non-invasive CSS to add hierarchy and reduce 'default Streamlit' look.
    Keep it light to avoid breaking theming.
    """
    st.markdown(
        """
        <style>
          /* Layout */
          .block-container { padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1250px; }
          h1, h2, h3 { letter-spacing: -0.02em; }

          /* Header */
          .app-title {
            text-align: center;
            font-size: 2.05rem;
            font-weight: 850;
            margin: 0.2rem 0 0.2rem 0;
          }
          .subtitle {
            text-align: center;
            color: rgba(250,250,250,0.75);
            font-size: 1.02rem;
            margin: 0 0 1.0rem 0;
          }

          /* Muted text */
          .muted { color: rgba(250,250,250,0.72); font-size: 0.98rem; }

          /* Sidebar typography + spacing */
          [data-testid="stSidebar"] .stRadio label {
            font-size: 1.05rem !important;
            padding: 8px 8px !important;
            margin: 0 !important;
          }
          [data-testid="stSidebar"] .stRadio div[role="radiogroup"]{
            gap: 6px;
          }
          [data-testid="stSidebar"] .stMarkdown, 
          [data-testid="stSidebar"] p, 
          [data-testid="stSidebar"] span,
          [data-testid="stSidebar"] label {
            font-size: 1.0rem !important;
          }

          /* Soft "cards" */
          .card {
            border: 1px solid rgba(255,255,255,0.14);
            border-radius: 18px;
            padding: 16px 16px;
            background: rgba(255,255,255,0.03);
            height: 100%;
          }
          .badge {
            display:inline-flex;
            align-items:center;
            justify-content:center;
            width: 34px;
            height: 34px;
            border-radius: 999px;
            border: 1px solid rgba(255,255,255,0.20);
            margin-right: 10px;
            font-weight: 850;
            font-size: 1.02rem;
          }
          .bot-title {
            font-size: 1.12rem;
            font-weight: 780;
          }
          .hr { height: 1px; background: rgba(255,255,255,0.08); margin: 14px 0; }

          /* Back button container */
          .top-right {
            display:flex;
            justify-content:flex-end;
            align-items:center;
            margin-top: -6px;
          }

          /* Small callout */
          .callout {
            border: 1px solid rgba(255,255,255,0.14);
            border-radius: 14px;
            padding: 12px 14px;
            background: rgba(255,255,255,0.02);
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


def hline() -> None:
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)


def get_query_param(name: str) -> Optional[str]:
    """
    Compatible with older Streamlit versions.
    Returns None if missing.
    """
    try:
        qp = st.query_params
        v = qp.get(name)
        if isinstance(v, list):
            return v[0] if v else None
        return v
    except Exception:
        qp = st.experimental_get_query_params()
        v = qp.get(name)
        return v[0] if v else None


def _normalize_query_value(v) -> Optional[str]:
    if v is None:
        return None
    if isinstance(v, list):
        return v[0] if v else None
    return str(v)


def set_query_param(**params) -> None:
    """
    Compatible with older Streamlit versions.
    Updates given params (does not necessarily clear others on older versions).
    If you need to CLEAR/REPLACE params, use set_query_params_exact().
    """
    try:
        # New Streamlit API behaves like a mutable mapping
        st.query_params.update(params)
    except Exception:
        st.experimental_set_query_params(**params)


def set_query_params_exact(**params) -> None:
    """
    Replace query params with exactly the given params.
    This solves cases where we must clear old params (e.g., bot=...).
    Works on new + old Streamlit versions.
    """
    # Clean None values: treat them as "remove"
    clean_params = {k: v for k, v in params.items() if v is not None}
    try:
        st.query_params.clear()
        st.query_params.update(clean_params)
    except Exception:
        st.experimental_set_query_params(**clean_params)


# -------------------- Model --------------------
@dataclass(frozen=True)
class BotSpec:
    order: int
    bot_id: str
    title: str
    filename: str
    category: str
    required_columns: List[str]
    notes: List[str]


BOTS: List[BotSpec] = [
    BotSpec(
        order=1,
        bot_id="find_sud",
        title="Определение подсудности",
        filename="find_sud.py",
        category="1) Определение суда",
        required_columns=["ВНД", "Адрес"],
        notes=[
            'Чтобы использовать бота, нужен Excel-файл с колонками "ВНД" и "Адрес". '
            '"ВНД" берётся из списка, полученного от Енлик. Для получения актуальных адресов ответчиков '
            'Excel с ВНД нужно отправить Анеле для выгрузки.',
            "После завершения работы бота обязательно вручную проверить, что все суды соответствуют адресу проживания ответчика.",
        ],
    ),
    BotSpec(
        order=2,
        bot_id="read_dogovora",
        title="Чтение договора для определения типа договора",
        filename="read_dogovora.py",
        category="2) Проверка договоров",
        required_columns=["ВНД", "Кредитор", "Путь к договору"],
        notes=[
            'Excel должен содержать колонки "ВНД", "Кредитор", "Путь к договору".',
            '"ВНД" и "Кредитор" берутся из Енлик или ПИР.',
            '"Путь к договору" нужно получить у Райымбека (Сбор).',
            'После завершения работы проверить, что новый столбец "Тип договора" заполнен. '
            "Если есть пустые значения или PDF не найден — вернуть Райымбеку для повторной проверки.",
            'Отфильтровать тип договора "Договорная" и сохранить в новый Excel для следующего бота.',
        ],
    ),
    BotSpec(
        order=3,
        bot_id="dogovornaya_where",
        title='Проверка подсудности договоров по типу "Договорная"',
        filename="dogovornaya_where.py",
        category="2) Проверка договоров",
        required_columns=["ВНД", "Кредитор", "Путь к договору", "Тип договора"],
        notes=[
            'Excel должен содержать "ВНД", "Кредитор", "Путь к договору", "Тип договора".',
            "Данные берутся с предыдущего бота.",
            "После получения результата проверить по тексту договора и определить суд.",
            'Нур-Султан / Астана / районы города → "Межрайонный суд по гражданским делам г.Астаны".',
            'Если указано "По месту нахождения МФО, Займодателя, Гаранта, Заемщика" — суд определять по адресу.',
        ],
    ),
    BotSpec(
        order=4,
        bot_id="split_astana_medeu",
        title="Разделение Астана и Медеу",
        filename="split_astana_medeu.py",
        category="3) Сформирование исков",
        required_columns=[],
        notes=[
            'Нужно закончить основную таблицу и использовать бота "isk_sum2".',
            'Результат от "isk_sum2" открыть и удалить два столбца: '
            '"Сумма расходов с пробелом" и "Сумма расходов с прописью".',
            'Далее выбрать только строки где "Подача да/нет" = "да", выбрать только объединённый шаблон и сохранить.',
            "В конце получите результаты в Excel файлах по физическим и юридическим лицам.",
        ],
    ),
    BotSpec(
        order=5,
        bot_id="isk_generator_all_and_zaimscoring",
        title="Сформирование исков по общему шаблону",
        filename="isk_generator_all_and_zaimscoring.py",
        category="3) Сформирование исков",
        required_columns=[],
        notes=[
            'Для запуска нужно закончить основную таблицу и использовать бота "isk_sum2". '
            'В результате удалить два столбца: "Сумма расходов с пробелом" и "Сумма расходов с прописью".',
            'Выбрать только строки где "Подача да/нет" = "да" и выбрать только общий шаблон, займ+скоринг, сохранить.',
            'Загрузить шаблон "Общий шаблон исков" и выбрать папку для сохранения результата.',
        ],
    ),
    BotSpec(
        order=6,
        bot_id="isk_generator_ND",
        title="Сформирование исков НД",
        filename="isk_generator_ND.py",
        category="3) Сформирование исков",
        required_columns=[],
        notes=[
            'Для запуска нужно закончить основную таблицу и использовать бота "isk_sum2". '
            'В результате удалить два столбца: "Сумма расходов с пробелом" и "Сумма расходов с прописью".',
            'Выбрать только строки где "Подача да/нет" = "да" и выбрать только НД, сохранить.',
            'Загрузить шаблон "Шаблон исков НД" и выбрать папку для сохранения результата.',
        ],
    ),
    BotSpec(
        order=7,
        bot_id="isk_generator_combined",
        title="Сформирование объединённых исков",
        filename="isk_generator_combined.py",
        category="3) Сформирование исков",
        required_columns=[],
        notes=[
            "Вложить Excel файлы, полученные после разделения Астана и Медеу.",
            "Выбрать шаблон для объединённых исков.",
            'После завершения будут получены иски и реестр. В реестре указывается, с каким "ВНД" объединены иски.',
        ],
    ),
    BotSpec(
        order=8,
        bot_id="spravka",
        title="Сформирование справок",
        filename="spravka.py",
        category="4) Справки и ходатайства",
        required_columns=[],
        notes=["Нужно вложить Excel файлы и шаблон справок ЭП и физических лиц."],
    ),
    BotSpec(
        order=9,
        bot_id="hodataistva",
        title="Сформирование ходатайств",
        filename="hodataistva.py",
        category="4) Справки и ходатайства",
        required_columns=[],
        notes=["Нужно вложить Excel файлы и шаблон ходатайство."],
    ),
]

CATEGORIES = ["Все"] + sorted({b.category for b in BOTS})


# -------------------- Download helpers (cached) --------------------
EXCLUDE_DIRS = {
    ".streamlit",
    ".venv",
    "venv",
    ".git",
    "__pycache__",
    ".idea",
    ".vscode",
}
EXCLUDE_SUFFIXES = {".pyc", ".log"}


def _should_exclude(path: Path) -> bool:
    if any(part in EXCLUDE_DIRS for part in path.parts):
        return True
    if path.suffix.lower() in EXCLUDE_SUFFIXES:
        return True
    return False


@st.cache_data(show_spinner=False)
def zip_project_bytes_cached() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for path in BASE_DIR.rglob("*"):
            if path.is_dir():
                continue
            if _should_exclude(path):
                continue
            arcname = path.relative_to(BASE_DIR).as_posix()
            z.write(path, arcname)
    buf.seek(0)
    return buf.read()


@st.cache_data(show_spinner=False)
def zip_templates_bytes_cached() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        if TEMPLATES_DIR.exists():
            for p in TEMPLATES_DIR.rglob("*"):
                if p.is_file():
                    z.write(p, p.relative_to(TEMPLATES_DIR).as_posix())
    buf.seek(0)
    return buf.read()


@st.cache_data(show_spinner=False)
def file_bytes_cached(path_str: str) -> bytes:
    return Path(path_str).read_bytes()


# -------------------- Bot launching --------------------
def launch_bot(script_name: str) -> Tuple[bool, str]:
    script_path = BOTS_DIR / script_name
    if not script_path.exists():
        return False, f"Файл не найден: {script_path}"

    try:
        p = subprocess.Popen(
            [sys.executable, str(script_path)],
            cwd=str(BOTS_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        # Bots can open their own file dialogs. We only try to capture early output.
        try:
            out, err = p.communicate(timeout=1.0)
        except Exception:
            out, err = "", ""
        return True, (out + ("\n" + err if err else "")).strip()
    except Exception as e:
        return False, str(e)


# -------------------- Reusable UI blocks --------------------
def top_header() -> None:
    st.markdown(f"<div class='app-title'>⚖️ {APP_TITLE}</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='subtitle'>Запуск локальных юридических ботов, шаблоны и примеры — в одном месте</div>",
        unsafe_allow_html=True,
    )


def sidebar() -> None:
    """
    Sidebar = navigation + global controls.
    Important: Avoid writing to query params on every rerun (can cause extra reruns).
    """
    with st.sidebar:
        st.markdown("### Навигация")

        MENU = [
            ("bots", "🤖 Боты"),
            ("templates", "📄 Шаблоны"),
            ("excel", "📊 Excel"),
            ("deps", "📦 Зависимости"),
            ("download", "⬇️ Скачать проект"),
        ]

        current_page = get_query_param("page") or "bots"
        labels = [m[1] for m in MENU]
        key_by_label = {m[1]: m[0] for m in MENU}
        pages = [m[0] for m in MENU]
        default_idx = pages.index(current_page) if current_page in pages else 0

        choice = st.radio("Раздел", labels, index=default_idx)
        new_page = key_by_label[choice]

        # Only update URL if page actually changed
        if new_page != current_page:
            # When switching section, clear bot selection so we don't "stick" to details.
            set_query_params_exact(page=new_page)
            st.rerun()

        st.markdown("---")

        # Global Quick Actions (download buttons) – reduces hunting in other pages
        st.markdown("### Быстрые действия")
        if TEMPLATES_DIR.exists():
            st.download_button(
                "⬇️ Шаблоны (ZIP)",
                data=zip_templates_bytes_cached(),
                file_name="templates.zip",
                mime="application/zip",
                use_container_width=True,
            )
        else:
            st.caption("Шаблоны: папка templates не найдена")

        if EXAMPLE_XLSX.exists():
            st.download_button(
                "⬇️ example.xlsx",
                data=file_bytes_cached(str(EXAMPLE_XLSX)),
                file_name="example.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.caption("Excel пример: файл example.xlsx не найден")

        st.download_button(
            "⬇️ Весь проект (ZIP)",
            data=zip_project_bytes_cached(),
            file_name="legal-bots-streamlit.zip",
            mime="application/zip",
            use_container_width=True,
        )

        st.markdown("---")
        with st.expander("ℹ️ Подсказка по интерфейсу", expanded=False):
            st.write(
                "• В центре — только основные данные/действия.\n"
                "• В сайдбаре — навигация и быстрые загрузки.\n"
                "• Внутри бота используй вкладки: требования → примечания → запуск."
            )


def bot_card(bot: BotSpec) -> None:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown(
        f'<div><span class="badge">{bot.order}</span>'
        f'<span class="bot-title">{bot.title}</span></div>',
        unsafe_allow_html=True,
    )
    st.markdown(f"<div class='muted'>{bot.category} • <code>{bot.filename}</code></div>", unsafe_allow_html=True)

    if bot.required_columns:
        st.markdown("**Колонки Excel:** " + ", ".join([f"`{c}`" for c in bot.required_columns]))
    else:
        st.markdown("**Колонки Excel:** зависит от входного файла (см. примечания).")

    if st.button("Открыть →", key=f"open_{bot.bot_id}", use_container_width=True):
        # Use exact set to guarantee we don't keep stale params
        set_query_params_exact(page="bots", bot=bot.bot_id)
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


# -------------------- Pages --------------------
def page_dependencies() -> None:
    st.header("📦 Установка зависимостей")

    left, right = st.columns([1.2, 1.0], gap="large")

    with left:
        st.subheader("Python-зависимости")
        st.caption("Скопируй команды и выполни в терминале из папки проекта.")
        st.code(
            "python -m venv .venv\n"
            ".venv\\Scripts\\activate\n"
            "pip install -r requirements.txt\n"
            "streamlit run app.py",
            language="bash",
        )
        st.info("Если используешь PowerShell — активация может отличаться: .venv\\Scripts\\Activate.ps1")

    with right:
        st.subheader("OCR (опционально)")
        st.caption("Если PDF — сканы, понадобится Tesseract.")
        st.markdown(
            """
<div class="callout">
<b>Windows:</b> скачай установщик Tesseract-OCR и добавь в PATH<br/>
Путь по умолчанию: <code>C:\\Program Files\\Tesseract-OCR</code>
</div>
            """.strip(),
            unsafe_allow_html=True,
        )
        with st.expander("Пошаговая инструкция", expanded=False):
            st.markdown(
                """
1. Скачай и установи **Tesseract-OCR** (обычно `.exe`) с SourceForge (официальное зеркало).
2. Установи в: `C:\\Program Files\\Tesseract-OCR`
3. Добавь в PATH:
   - **Пуск → Environment Variables → Path → Edit → New**
   - Добавь: `C:\\Program Files\\Tesseract-OCR`
4. Проверь установку в новом терминале:
```bash
tesseract --version
```
5. Если `pytesseract` не видит tesseract — перезапусти терминал/ПК.
                """.strip()
            )
        st.warning("Если OCR не нужен — Tesseract можно не ставить.")


def page_templates() -> None:
    st.header("📄 Шаблоны исков / справок / ходатайств")
    if not TEMPLATES_DIR.exists():
        st.error("Папка templates не найдена.")
        return

    files = sorted([p for p in TEMPLATES_DIR.rglob("*") if p.is_file()])
    c1, c2, c3 = st.columns([1, 1, 1], gap="large")
    with c1:
        st.metric("Файлов", value=len(files))
    with c2:
        st.metric("Папка", value="templates")
    with c3:
        st.metric("Форматы", value=len(sorted({p.suffix.lower() for p in files})))

    hline()

    st.download_button(
        "Скачать шаблоны (ZIP)",
        data=zip_templates_bytes_cached(),
        file_name="templates.zip",
        mime="application/zip",
        use_container_width=True,
    )

    with st.expander("Список файлов", expanded=False):
        for p in files:
            st.write(f"• {p.relative_to(TEMPLATES_DIR).as_posix()}")


def page_excel() -> None:
    st.header("📊 Шаблоны Excel")
    if not EXAMPLE_XLSX.exists():
        st.error("Файл example.xlsx не найден.")
        return

    left, right = st.columns([1.0, 1.2], gap="large")

    with left:
        st.metric("Файл", "example.xlsx")
        st.metric("Размер", f"{EXAMPLE_XLSX.stat().st_size / 1024:.1f} KB")
        st.download_button(
            "Скачать example.xlsx",
            data=file_bytes_cached(str(EXAMPLE_XLSX)),
            file_name="example.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with right:
        st.info("Используй этот файл как пример структуры колонок для ботов (см. требования конкретного бота).")
        st.caption("Подсказка: если бот пишет новые колонки — сначала прогоняй на малом наборе строк, чтобы проверить формат.")


def page_download_project() -> None:
    st.header("⬇️ Скачать проект")
    left, right = st.columns([1.2, 0.8], gap="large")

    with left:
        st.markdown(
            """
<div class="callout">
Скачай полный проект одним архивом: <code>app.py</code>, <code>requirements.txt</code>, боты, шаблоны и пример Excel.
</div>
            """.strip(),
            unsafe_allow_html=True,
        )

    with right:
        st.metric("Архив", "legal-bots-streamlit.zip")

    st.download_button(
        "Скачать весь проект (ZIP)",
        data=zip_project_bytes_cached(),
        file_name="legal-bots-streamlit.zip",
        mime="application/zip",
        use_container_width=True,
    )


def show_bot_detail(bot: BotSpec) -> None:
    # Top row with back button on the right
    top_l, top_r = st.columns([6, 1], gap="large")
    with top_l:
        st.subheader(f"{bot.order}. {bot.title}")
        st.markdown(
            f"<div class='muted'>Категория: <b>{bot.category}</b> • Файл: <code>{bot.filename}</code></div>",
            unsafe_allow_html=True,
        )
    with top_r:
        st.markdown("<div class='top-right'>", unsafe_allow_html=True)
        if st.button("← Назад", use_container_width=True):
            # CRITICAL FIX: clear bot param, otherwise we stay in detail page
            set_query_params_exact(page="bots")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    hline()

    # Key metrics row
    m1, m2, m3 = st.columns([1, 1, 1], gap="large")
    with m1:
        st.metric("№ в списке", bot.order)
    with m2:
        st.metric("Колонок требуется", len(bot.required_columns))
    with m3:
        st.metric("Примечаний", len(bot.notes))

    tabs = st.tabs(["✅ Требования", "📝 Примечания", "▶ Запуск", "ℹ️ О боте"])

    with tabs[0]:
        st.markdown("### Требования к входным данным")
        if bot.required_columns:
            st.markdown("- " + "\n- ".join([f"**{c}**" for c in bot.required_columns]))
            st.caption("Проверь, что названия колонок совпадают **точно** (включая регистр/пробелы).")
        else:
            st.info("См. примечания — у разных сценариев разные входные файлы/шаблоны.")

    with tabs[1]:
        st.markdown("### Примечания и инструкции")
        with st.expander("Показать/скрыть примечания", expanded=True):
            for i, note in enumerate(bot.notes, start=1):
                st.markdown(f"**{i})** {note}")

    with tabs[2]:
        st.markdown("### Запуск бота")
        st.warning(
            "Бот откроет отдельное окно для выбора Excel/шаблонов/папки (это нормально). "
            "После запуска следуй подсказкам в появившемся окне."
        )

        c1, c2 = st.columns([1, 1], gap="large")
        with c1:
            run = st.button("▶ Запустить бота", type="primary", use_container_width=True)
        with c2:
            st.caption("Если ничего не открылось — проверь, что файл бота существует и нет блокировок антивируса.")

        if run:
            ok, msg = launch_bot(bot.filename)
            if ok:
                st.success("Бот запущен.")
                if msg:
                    with st.expander("Лог запуска (если есть)", expanded=False):
                        st.code(msg)
            else:
                st.error("Не удалось запустить бот.")
                st.code(msg)

    with tabs[3]:
        st.markdown("### Контекст")
        st.markdown(
            """
- **Назначение:** автоматизация рутины по Excel/шаблонам (подготовка судов, проверка договоров, генерация исков/справок).
- **Важно:** результат некоторых ботов требует **ручной валидации** (см. примечания).
            """.strip()
        )


def page_bots() -> None:
    st.header("🤖 Боты")

    # If bot is selected via query param — show details
    bot_id = get_query_param("bot")
    if bot_id:
        bot = next((b for b in BOTS if b.bot_id == bot_id), None)
        if not bot:
            st.error("Бот не найден.")
            return
        show_bot_detail(bot)
        return

    # Compact filter row
    f1, f2, f3 = st.columns([1.1, 1.2, 0.7], gap="large")

    with f1:
        cat = st.selectbox("Категория", CATEGORIES, index=0)
    with f2:
        q = (st.text_input("Поиск", placeholder="например: договор, подсудность, справка…") or "").strip().lower()
    with f3:
        sort_mode = st.selectbox("Сортировка", ["По порядку", "По названию"], index=0)

    filtered = BOTS if cat == "Все" else [b for b in BOTS if b.category == cat]
    if q:
        filtered = [
            b
            for b in filtered
            if q in b.title.lower()
            or q in b.category.lower()
            or q in b.filename.lower()
            or q in b.bot_id.lower()
        ]

    if sort_mode == "По названию":
        filtered = sorted(filtered, key=lambda b: b.title.lower())
    else:
        filtered = sorted(filtered, key=lambda b: b.order)

    # KPI row
    k1, k2, k3, k4 = st.columns([1, 1, 1, 1], gap="large")
    with k1:
        st.metric("Всего ботов", len(BOTS))
    with k2:
        st.metric("Категорий", len(set(b.category for b in BOTS)))
    with k3:
        st.metric("Показано", len(filtered))
    with k4:
        st.metric("Фильтр", "Все" if cat == "Все" else cat)

    hline()

    if not filtered:
        st.info("Ничего не найдено по фильтрам. Попробуй изменить категорию или запрос.")
        return

    cols = st.columns(3, gap="large")
    for i, bot in enumerate(filtered):
        with cols[i % 3]:
            bot_card(bot)


# -------------------- App entry --------------------
def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="⚖️", layout="wide")
    inject_css()

    if not BOTS_DIR.exists():
        st.error(f"Папка с ботами не найдена: {BOTS_DIR}")
        st.stop()

    top_header()
    sidebar()

    page = get_query_param("page") or "bots"
    if page == "deps":
        page_dependencies()
    elif page == "templates":
        page_templates()
    elif page == "excel":
        page_excel()
    elif page == "download":
        page_download_project()
    else:
        page_bots()


if __name__ == "__main__":
    main()
