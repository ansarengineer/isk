\
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Any, Iterable

# ----------------------
# Data cleaning helpers
# ----------------------

# Common Latin->Cyrillic look-alikes that often appear in KZ/RU datasets
LATIN_TO_CYR = str.maketrans({
    "A": "А", "a": "а",
    "B": "В",
    "C": "С", "c": "с",
    "E": "Е", "e": "е",
    "H": "Н",
    "K": "К",
    "M": "М",
    "O": "О", "o": "о",
    "P": "Р", "p": "р",
    "T": "Т",
    "X": "Х", "x": "х",
    "Y": "У",
    # Note: "I" is ambiguous (І/И), "V/W" not safe.
})

# Normalization for court/region comparisons
def _norm_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.translate(LATIN_TO_CYR)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _norm_key(s: str) -> str:
    s = _norm_text(s).lower()
    # keep letters/digits/№ and spaces, normalize punctuation
    s = re.sub(r"[^\w\s№\-әіңғқөұүһё]", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    # normalize № spacing: "№ 2" -> "№2"
    s = re.sub(r"№\s+(\d+)", r"№\1", s)
    return s

def canon_region(region: str) -> str:
    """Canonicalize region field to reduce duplicates / mismatches."""
    r = _norm_key(region)
    if not r:
        return ""

    canon = {
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
        "область ұлытау": "Область Ұлытау",
        "область улытау": "Область Ұлытау",
        "область абай": "Область Абай",
        "алматы": "Алматы",
        "астана": "Астана",
        "шымкент": "Шымкент",
    }
    # restore original casing for known keys
    return canon.get(r, _norm_text(region))

def fix_common_typos_court_name(name: str) -> str:
    """Fix a few common spelling variants."""
    s = _norm_text(name)
    # normalize multiple spaces and '№ 2' formatting
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"№\s+(\d+)", r"№\1", s)

    # very common: 'Тюлькубаский' -> 'Тюлькубасский'
    s = re.sub(r"\bТюлькубаск(?:ий|ый)\b", "Тюлькубасский", s, flags=re.IGNORECASE)
    # common: 'Cуд' (latin C) already fixed by translate, but ensure casing
    s = re.sub(r"^\s*суд\b", "Суд", s, flags=re.IGNORECASE)
    return s

@dataclass
class CleanReport:
    total_in: int
    total_out: int
    removed_duplicates: int
    fixed_region: int
    fixed_court_name: int
    fixed_latin_lookalikes: int
    invalid_rows: int
    duplicates_examples: list[dict[str, Any]]

def validate_and_clean_courts(
    input_path: str,
    output_path: str | None = None,
    *,
    drop_rows_without_required: bool = True,
    required_fields: tuple[str, str] = ("СУД", "Регион"),
    dedupe_key: str = "court+region",  # options: "court+region" or "court_only"
    keep: str = "first",               # options: "first" or "best_desc"
    save_pretty: bool = True
) -> CleanReport:
    """
    Validate and clean courts_merged.json.

    - Fixes Latin look-alike characters in 'СУД' and 'Регион'
    - Canonicalizes 'Регион'
    - Normalizes spaces and '№' formatting
    - Fixes a few known typos in court names
    - Optionally drops rows missing required fields
    - De-duplicates by normalized key

    Returns CleanReport and writes cleaned JSON to output_path (defaults to '<input>_cleaned.json').
    """
    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        raise ValueError("Expected JSON array (list of objects).")

    if output_path is None:
        output_path = re.sub(r"\.json$", "", input_path, flags=re.IGNORECASE) + "_cleaned.json"

    total_in = len(data)
    invalid_rows = 0
    removed_duplicates = 0
    fixed_region = 0
    fixed_court_name = 0
    fixed_latin = 0

    seen: dict[str, dict[str, Any]] = {}
    dup_examples: list[dict[str, Any]] = []

    def make_key(item: dict[str, Any]) -> str:
        court = _norm_key(item.get("СУД"))
        region = _norm_key(item.get("Регион"))
        if dedupe_key == "court_only":
            return court
        return f"{court}||{region}"

    def choose(a: dict[str, Any], b: dict[str, Any]) -> dict[str, Any]:
        """Choose which duplicate record to keep."""
        if keep == "first":
            return a
        # keep == "best_desc": prefer record with non-empty description (longer)
        da = _norm_text(a.get("Описание") or "")
        db = _norm_text(b.get("Описание") or "")
        if len(db) > len(da):
            return b
        return a

    cleaned_list: list[dict[str, Any]] = []

    for raw in data:
        if not isinstance(raw, dict):
            invalid_rows += 1
            continue

        item = dict(raw)  # shallow copy

        # detect latin look-alikes by checking if translation changes the string
        orig_court = "" if item.get("СУД") is None else str(item.get("СУД"))
        orig_region = "" if item.get("Регион") is None else str(item.get("Регион"))

        trans_court = orig_court.translate(LATIN_TO_CYR)
        trans_region = orig_region.translate(LATIN_TO_CYR)
        if trans_court != orig_court or trans_region != orig_region:
            fixed_latin += 1

        # apply fixes
        item["СУД"] = fix_common_typos_court_name(trans_court)
        item["Регион"] = canon_region(trans_region)

        if item["СУД"] != orig_court:
            fixed_court_name += 1
        if item["Регион"] != orig_region:
            fixed_region += 1

        # required fields validation
        missing = False
        for field in required_fields:
            v = item.get(field)
            if v is None or _norm_text(v) == "":
                missing = True
                break

        if missing:
            invalid_rows += 1
            if drop_rows_without_required:
                continue

        k = make_key(item)
        if not k:
            invalid_rows += 1
            if drop_rows_without_required:
                continue

        if k in seen:
            removed_duplicates += 1
            if len(dup_examples) < 10:
                dup_examples.append({"key": k, "kept": seen[k], "dropped": item})
            # choose which to keep
            kept = choose(seen[k], item)
            seen[k] = kept
        else:
            seen[k] = item

    # preserve insertion order based on first occurrence (approx.)
    # If keep=="best_desc", the kept version may come from later, but position stays first.
    cleaned_list = list(seen.values())

    # Write output
    with open(output_path, "w", encoding="utf-8") as f:
        if save_pretty:
            json.dump(cleaned_list, f, ensure_ascii=False, indent=2)
        else:
            json.dump(cleaned_list, f, ensure_ascii=False)

    return CleanReport(
        total_in=total_in,
        total_out=len(cleaned_list),
        removed_duplicates=removed_duplicates,
        fixed_region=fixed_region,
        fixed_court_name=fixed_court_name,
        fixed_latin_lookalikes=fixed_latin,
        invalid_rows=invalid_rows,
        duplicates_examples=dup_examples
    )

# ----------------------
# CLI usage (optional)
# ----------------------
if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Validate & clean courts_merged.json")
    ap.add_argument("input", help="Path to courts_merged.json")
    ap.add_argument("--out", default=None, help="Output path (default: <input>_cleaned.json)")
    ap.add_argument("--keep", choices=["first", "best_desc"], default="best_desc")
    ap.add_argument("--dedupe", choices=["court+region", "court_only"], default="court+region")
    ap.add_argument("--no-pretty", action="store_true", help="Write minified JSON")
    args = ap.parse_args()

    rep = validate_and_clean_courts(
        args.input,
        args.out,
        keep=args.keep,
        dedupe_key=args.dedupe,
        save_pretty=not args.no_pretty,
    )

    print("Done.")
    print(rep)
