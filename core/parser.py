"""
core/parser.py
Business logic for parsing biometric time records.
"""
import re
from typing import List, Optional


# ─────────────────────────────────────────────────────────────────────────────
# 1. Time parsing
# ─────────────────────────────────────────────────────────────────────────────

def parse_times(raw: Optional[str], person: str = "", day: str = "") -> List[str]:
    """
    Extract and clean time strings from a raw biometric cell.

    Handles:
    - Space-separated:  "09:57 12:33 19:04"
    - Glued together:   "09:5712:3319:04"
    - Mixed garbage:    "09:57abc12:33"

    Returns a list of "HH:MM" strings, consecutive exact-duplicates removed.
    """
    if not raw:
        return []
    raw = str(raw).strip()
    if not raw or raw.lower() in ("nan", "-", "—", "–"):
        return []

    # Split on whitespace if spaces present, otherwise regex-extract
    if " " in raw:
        parts = raw.split()
    else:
        parts = re.findall(r"\d{1,2}:\d{2}", raw)

    times: List[str] = []
    for p in parts:
        p = p.strip()
        if ":" not in p:
            continue
        try:
            h_str, m_str = p.split(":", 1)
            h, m = int(h_str), int(m_str)
            if 0 <= h <= 23 and 0 <= m <= 59:
                times.append(f"{h:02d}:{m:02d}")
        except (ValueError, TypeError):
            pass

    # Remove consecutive exact duplicates only
    deduped: List[str] = []
    for t in times:
        if not deduped or t != deduped[-1]:
            deduped.append(t)

    return deduped


# ─────────────────────────────────────────────────────────────────────────────
# 2. Column assignment
# ─────────────────────────────────────────────────────────────────────────────

def get_column_schema(max_marks: int) -> List[str]:
    """
    Return the ordered list of column headers based on the maximum number
    of marks detected for a person.

    Skill rules:
        1 mark  → INGRESO
        2 marks → INGRESO, SALIDA FINAL
        3 marks → INGRESO, SALIDA, SALIDA FINAL
        4 marks → INGRESO, SALIDA, RETORNO, SALIDA FINAL
        5 marks → INGRESO, SALIDA, RETORNO, SALIDA FINAL  (+warn)
        6 marks → INGRESO, SALIDA, RETORNO, INGRESO 2, SALIDA 2, SALIDA FINAL
    """
    # Round up to next even number
    pairs = max(1, max_marks)
    if pairs <= 1:
        return ["INGRESO"]
    if pairs == 2:
        return ["INGRESO", "SALIDA FINAL"]
    if pairs == 3:
        return ["INGRESO", "SALIDA", "SALIDA FINAL"]
    if pairs == 4:
        return ["INGRESO", "SALIDA", "RETORNO", "SALIDA FINAL"]
    if pairs == 5:
        return ["INGRESO", "SALIDA", "RETORNO", "SALIDA FINAL"]  # 5th mark is extra
    # 6+
    return ["INGRESO", "SALIDA", "RETORNO", "INGRESO 2", "SALIDA 2", "SALIDA FINAL"]


def assign_marks_to_columns(times: List[str], schema: List[str]) -> dict:
    """
    Map a list of parsed times to the correct column names.

    Special rule for 2 columns: first→INGRESO, last→SALIDA FINAL.
    For all others: fill left-to-right, skipping extras past schema length.
    """
    result = {col: "" for col in schema}
    n = len(schema)

    if n == 1:
        if times:
            result["INGRESO"] = times[0]
        return result

    if n == 2:
        if len(times) >= 1:
            result["INGRESO"] = times[0]
        if len(times) >= 2:
            result["SALIDA FINAL"] = times[-1]
        return result

    for i, col in enumerate(schema):
        if i < len(times):
            result[col] = times[i]

    return result
