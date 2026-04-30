"""
core/ocr.py
OCR integration for extracting attendance data from images.
Uses pytesseract when available; falls back to a structured manual-entry hint.
"""
import re
import io
from typing import Optional

import pandas as pd

try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────────
# OCR helpers
# ─────────────────────────────────────────────────────────────────────────────

def extract_text_from_image(image_bytes: bytes) -> str:
    """
    Run OCR on image bytes. Returns extracted text or empty string.
    """
    if not OCR_AVAILABLE:
        return ""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        # Preprocess: grayscale
        img = img.convert("L")
        text = pytesseract.image_to_string(img, config="--psm 6")
        return text
    except Exception as exc:
        return f"__OCR_ERROR__: {exc}"


def parse_ocr_text_to_df(raw_text: str) -> pd.DataFrame:
    """
    Parse OCR output that resembles:
        DIA  FECHA       INGRESO  SALIDA
        Lun  01/07/2025  08:00    17:00
        ...

    Returns a DataFrame with columns: DIA, FECHA, INGRESO, SALIDA, RETORNO, SALIDA FINAL.
    Rows that cannot be parsed are returned with empty time columns for manual correction.
    """
    columns = ["DIA", "FECHA", "INGRESO", "SALIDA", "RETORNO", "SALIDA FINAL"]
    rows = []

    time_re = re.compile(r"\d{1,2}:\d{2}")
    date_re = re.compile(r"\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}")

    for line in raw_text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Skip header-like lines
        if re.search(r"(?i)(dia|fecha|ingreso|salida|hora)", line) and len(line) < 60:
            continue

        times = time_re.findall(line)
        date_m = date_re.search(line)

        if not date_m and not times:
            continue

        fecha = ""
        if date_m:
            raw_date = date_m.group()
            parts = re.split(r"[/\-]", raw_date)
            if len(parts) == 3:
                d, m, y = parts
                if len(y) == 2:
                    y = "20" + y
                fecha = f"{int(d):02d}/{int(m):02d}/{y}"

        # Day abbreviation: first word if non-numeric
        tokens = line.split()
        dia = ""
        if tokens and not tokens[0][0].isdigit():
            dia = tokens[0][:3].capitalize()

        # Assign up to 4 times
        ingreso = times[0] if len(times) > 0 else ""
        salida = times[1] if len(times) > 1 else ""
        retorno = times[2] if len(times) > 2 else ""
        salida_final = times[3] if len(times) > 3 else ""

        rows.append([dia, fecha, ingreso, salida, retorno, salida_final])

    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns)


def ocr_df_to_records(df: pd.DataFrame, person_name: str) -> list[dict]:
    """
    Convert an OCR-parsed DataFrame into the same record format used by reader.py.
    """
    records = []
    month, year = None, None
    days_dict = {}

    date_re = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")

    for _, row in df.iterrows():
        fecha = str(row.get("FECHA", "")).strip()
        m = date_re.match(fecha)
        if not m:
            continue
        day_num, mon, yr = int(m.group(1)), int(m.group(2)), int(m.group(3))
        month, year = mon, yr

        times = []
        for col in ["INGRESO", "SALIDA", "RETORNO", "SALIDA FINAL"]:
            val = str(row.get(col, "")).strip()
            if val:
                times.append(val)

        if times:
            days_dict[day_num] = " ".join(times)

    if days_dict:
        records.append({
            "name": person_name.upper(),
            "month": month,
            "year": year,
            "days": days_dict,
        })
    return records
