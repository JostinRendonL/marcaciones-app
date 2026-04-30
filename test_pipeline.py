"""Test the reader with all 3 real file types."""
from core.reader import read_biometric_xlsx
from core.exporter import export_to_xlsx, build_sheet_name

FILES = [
    (r"C:\Users\Home\Downloads\Marcaciones Nube.xlsx",        "CRUDO con Periodo"),
    (r"C:\Users\Home\Downloads\marcacioness.xlsx",            "CRUDO sin Periodo"),
    (r"C:\Users\Home\Downloads\MARCACIONES VARIOS MESES.xlsx", "PRE-PROCESADO"),
]

for path, label in FILES:
    print(f"\n{'='*60}")
    print(f"FILE: {label}")
    print(f"PATH: {path}")
    print(f"{'='*60}")

    with open(path, "rb") as f:
        data = f.read()

    records = read_biometric_xlsx(data)
    print(f"  Records: {len(records)}")
    for r in records:
        sheet = build_sheet_name(r["name"], r["month"] or 1, r["year"] or 2025)
        n = len(r["days"])
        # Show first 3 days as sample
        sample = dict(list(sorted(r["days"].items()))[:3])
        print(f"    {sheet:30s}  {n:2d} days  sample={sample}")

    # Try generating report
    try:
        xlsx = export_to_xlsx(records)
        print(f"  EXPORT OK: {len(xlsx):,} bytes")
    except Exception as e:
        print(f"  EXPORT FAILED: {e}")

print("\n\nAll tests done!")
