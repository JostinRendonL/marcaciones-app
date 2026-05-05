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

    Rules:
        1 mark  → INGRESO
        2 marks → INGRESO, SALIDA FINAL
        3 marks → INGRESO, SALIDA, SALIDA FINAL
        4 marks → INGRESO, SALIDA, INGRESO, SALIDA FINAL
        5 marks → INGRESO, SALIDA, INGRESO, SALIDA FINAL  (+warn)
        6 marks → INGRESO, SALIDA, INGRESO, SALIDA, INGRESO, SALIDA FINAL
    """
    n = max(1, max_marks)
    if n <= 1:
        return ["INGRESO"]
    if n == 2:
        return ["INGRESO", "SALIDA FINAL"]
    if n == 3:
        return ["INGRESO", "SALIDA", "SALIDA FINAL"]
    if n == 4:
        return ["INGRESO", "SALIDA", "INGRESO", "SALIDA FINAL"]
    if n == 5:
        return ["INGRESO", "SALIDA", "INGRESO", "SALIDA FINAL"]  # 5th mark is extra
    # 6+
    return ["INGRESO", "SALIDA", "INGRESO", "SALIDA", "INGRESO", "SALIDA FINAL"]


def assign_marks_to_columns(times: List[str], schema: List[str]) -> dict:
    """
    Map a list of parsed times to column names (positional, left-to-right).
    Schema may contain repeated column names; we use a list of (name, index)
    pairs so every slot is filled independently.
    """
    # Build indexed slots to handle repeated column names
    slots = list(schema)   # positional list
    result = {col: "" for col in schema}  # last-write wins for display
    _indexed: list = []
    for pos, col in enumerate(slots):
        _indexed.append((pos, col))

    n = len(slots)

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

    # For schemas with repeated names we need per-position storage.
    # Return a special ordered dict keyed by position.
    pos_result: dict = {}
    for i, (pos, col) in enumerate(_indexed):
        pos_result[pos] = times[i] if i < len(times) else ""

    # Also populate the plain col→value dict (last position wins)
    for pos, col in _indexed:
        if pos_result.get(pos):
            result[col] = pos_result[pos]

    # Attach positional data as a private attribute so exporter can use it
    result["_pos"] = pos_result
    return result


def get_entry_exit_pairs(schema: List[str]) -> List[tuple]:
    """
    Return (entry_col_index, exit_col_index) zero-based position pairs
    for every INGRESO/SALIDA pair in the schema, used for pay calculation.

    Examples:
      [INGRESO, SALIDA FINAL]                       → [(0, 1)]
      [INGRESO, SALIDA, SALIDA FINAL]               → [(0, 2)]
      [INGRESO, SALIDA, INGRESO, SALIDA FINAL]      → [(0, 1), (2, 3)]
      [INGRESO, SALIDA, INGRESO, SALIDA, INGRESO, SALIDA FINAL]
                                                    → [(0, 1), (2, 3), (4, 5)]
    """
    n = len(schema)
    if n == 1:
        return []  # only entry, no exit → cannot compute
    if n == 2:   # INGRESO, SALIDA FINAL
        return [(0, 1)]
    if n == 3:   # INGRESO, SALIDA, SALIDA FINAL
        return [(0, 2)]
    if n == 4:   # INGRESO, SALIDA, INGRESO, SALIDA FINAL
        return [(0, 1), (2, 3)]
    # 6:  INGRESO, SALIDA, INGRESO, SALIDA, INGRESO, SALIDA FINAL
    return [(0, 1), (2, 3), (4, 5)]
