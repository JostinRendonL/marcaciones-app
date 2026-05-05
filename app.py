"""
app.py  –  Marcaciones Biométricas · Streamlit App
Senior Full-stack build | Ecuador locale (DD/MM/YYYY)
"""

import io
import calendar
from datetime import date

import pandas as pd
import streamlit as st

from core.reader   import read_biometric_xlsx, records_to_dataframe, month_short
from core.exporter import export_to_xlsx, build_sheet_name
from core.ocr      import (
    OCR_AVAILABLE, extract_text_from_image,
    parse_ocr_text_to_df, ocr_df_to_records,
)
from core.parser   import parse_times, get_column_schema, assign_marks_to_columns, get_entry_exit_pairs


# ─────────────────────────────────────────────────────────────────────────────
# Page config – must be the very first Streamlit call
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Marcaciones Biométricas",
    page_icon="🕐",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ─────────────────────────────────────────────────────────────────────────────
# Custom CSS – premium dark theme
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* ── Main background ────────────────────────────────── */
.stApp {
    background: linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 50%, #16213e 100%);
    color: #e8eaf6;
}

/* ── Sidebar ────────────────────────────────────────── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #12122a 0%, #1c1c3a 100%);
    border-right: 1px solid #2a2a5a;
}
section[data-testid="stSidebar"] * { color: #c5cae9 !important; }

/* ── Cards / metric boxes ────────────────────────────── */
[data-testid="metric-container"] {
    background: linear-gradient(135deg, #1e1e3f, #252550);
    border: 1px solid #3f3f8f;
    border-radius: 12px;
    padding: 16px !important;
    box-shadow: 0 4px 20px rgba(0,0,0,0.4);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
[data-testid="metric-container"]:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 30px rgba(100,100,255,0.2);
}
[data-testid="metric-container"] label { color: #7986cb !important; font-weight: 500; }
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #e8eaf6 !important; font-size: 2rem; font-weight: 700;
}

/* ── Buttons ─────────────────────────────────────────── */
.stButton > button {
    background: linear-gradient(135deg, #3949ab, #5c6bc0);
    color: white !important;
    border: none;
    border-radius: 8px;
    padding: 0.5rem 1.5rem;
    font-weight: 600;
    letter-spacing: 0.5px;
    transition: all 0.2s ease;
    box-shadow: 0 2px 10px rgba(63,81,181,0.4);
}
.stButton > button:hover {
    background: linear-gradient(135deg, #5c6bc0, #7986cb);
    transform: translateY(-1px);
    box-shadow: 0 4px 20px rgba(92,107,192,0.5);
}

/* ── Download button ─────────────────────────────────── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #00897b, #00bfa5) !important;
    color: white !important;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 2rem;
    font-weight: 700;
    font-size: 1rem;
    box-shadow: 0 2px 12px rgba(0,191,165,0.4);
    transition: all 0.2s ease;
}
.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 24px rgba(0,191,165,0.5);
}

/* ── DataFrames ──────────────────────────────────────── */
.stDataFrame { border-radius: 8px; overflow: hidden; }

/* ── File uploader ───────────────────────────────────── */
[data-testid="stFileUploader"] {
    background: rgba(30,30,63,0.6);
    border: 2px dashed #3949ab;
    border-radius: 12px;
    padding: 1rem;
    transition: border-color 0.2s ease;
}
[data-testid="stFileUploader"]:hover { border-color: #7986cb; }

/* ── Expanders ───────────────────────────────────────── */
[data-testid="stExpander"] {
    background: rgba(25,25,50,0.7);
    border: 1px solid #2a2a5a;
    border-radius: 10px;
}

/* ── Divider ─────────────────────────────────────────── */
hr { border-color: #2a2a5a !important; }

/* ── Headers ─────────────────────────────────────────── */
h1, h2, h3 { color: #c5cae9 !important; }
h1 { font-size: 1.9rem !important; font-weight: 700 !important; }

/* ── Alerts ──────────────────────────────────────────── */
.stAlert { border-radius: 8px; }

/* ── Tabs ────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(20,20,45,0.8);
    border-radius: 10px;
    gap: 4px;
    padding: 4px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    color: #9fa8da !important;
    font-weight: 500;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #3949ab, #5c6bc0) !important;
    color: white !important;
}

/* ── Section header cards ────────────────────────────── */
.section-card {
    background: linear-gradient(135deg, rgba(30,30,70,0.8), rgba(40,40,90,0.8));
    border: 1px solid #3a3a7a;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 15px rgba(0,0,0,0.3);
}

/* ── Person badge ────────────────────────────────────── */
.person-badge {
    display: inline-block;
    background: linear-gradient(135deg, #3949ab, #5c6bc0);
    color: white;
    border-radius: 20px;
    padding: 2px 12px;
    font-size: 0.85rem;
    font-weight: 600;
    margin: 2px;
}

/* ── Status chip ─────────────────────────────────────── */
.chip-ok   { color: #a5d6a7; font-weight: 600; }
.chip-warn { color: #fff176; font-weight: 600; }
.chip-err  { color: #ef9a9a; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Session state initialisation
# ─────────────────────────────────────────────────────────────────────────────
def _init_state():
    defaults = {
        "records":          [],        # raw parsed records
        "flat_df":          None,      # tidy DataFrame for editing
        "edited_df":        None,      # user-corrected DataFrame
        "processed":        False,
        "ocr_df":           None,      # OCR preview for images
        "ocr_person_name":  "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🕐 Marcaciones")
    st.markdown("**Procesador biométrico**")
    st.markdown("---")
    st.markdown(
        "Soporta archivos `.xlsx` del reloj biométrico y fotos de registros manuales."
    )
    st.markdown("---")

    st.markdown("### ⚙️ Opciones")
    show_raw_times = st.checkbox("Mostrar horas crudas en vista previa", value=True)
    warn_ambiguous = st.checkbox("Alertar marcaciones ≥ 5 por día", value=True)

    st.markdown("---")
    st.markdown(
        "<small style='color:#7986cb'>Optimizado para equipos de hasta 6 usuarios.<br>"
        "Formato Ecuador: DD/MM/YYYY</small>",
        unsafe_allow_html=True,
    )
    st.markdown("---")
    if st.button("🔄 Reiniciar sesión", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        _init_state()
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(
    '<h1>🕐 Procesador de Marcaciones Biométricas</h1>',
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='color:#9fa8da;margin-top:-0.5rem;'>Automatización de registros de asistencia · "
    "Formato Ecuador DD/MM/YYYY</p>",
    unsafe_allow_html=True,
)
st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab_upload, tab_preview, tab_export = st.tabs([
    "📂 1. Cargar Archivos",
    "🔍 2. Vista Previa y Correcciones",
    "📊 3. Exportar Reporte",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.markdown("### Selecciona el tipo de archivo a procesar")

    input_mode = st.radio(
        "Modo de ingesta",
        ["📊 Archivo Excel biométrico (.xlsx)", "🖼️ Imagen de registro manual (.png/.jpg)"],
        horizontal=True,
        label_visibility="collapsed",
    )

    st.markdown("---")

    # ── Excel upload ─────────────────────────────────────────────────────────
    if input_mode.startswith("📊"):
        uploaded_xlsx = st.file_uploader(
            "Arrastra o selecciona el archivo .xlsx del biométrico",
            type=["xlsx"],
            key="xlsx_uploader",
            help="El archivo debe tener la estructura estándar: fila de días, pares de nombre+datos",
        )

        if uploaded_xlsx:
            with st.spinner("Analizando archivo Excel…"):
                try:
                    file_bytes = uploaded_xlsx.read()
                    records    = read_biometric_xlsx(file_bytes)
                    flat_df    = records_to_dataframe(records)

                    st.session_state["records"]   = records
                    st.session_state["flat_df"]   = flat_df
                    st.session_state["edited_df"] = flat_df.copy()
                    st.session_state["processed"] = True

                    st.success(f"✅ Archivo procesado: **{len(records)} bloque(s) persona/mes** detectados.")

                except Exception as exc:
                    st.error(f"❌ Error al leer el archivo: `{exc}`")
                    st.info("Verifica que el archivo tenga el formato biométrico estándar y no esté corrupto.")

    # ── Image / OCR upload ───────────────────────────────────────────────────
    else:
        if not OCR_AVAILABLE:
            st.warning(
                "⚠️ **pytesseract** no está instalado. "
                "Instálalo con `pip install pytesseract Pillow` y asegúrate de "
                "[instalar Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)."
                "\n\nPor ahora puedes **cargar la imagen** y editar la tabla manualmente."
            )

        uploaded_img = st.file_uploader(
            "Arrastra o selecciona la imagen del registro manual",
            type=["png", "jpg", "jpeg"],
            key="img_uploader",
        )

        ocr_person_name = st.text_input(
            "Nombre de la persona en esta imagen",
            value=st.session_state["ocr_person_name"],
            placeholder="Ej: VICKY SUÁREZ",
            key="ocr_name_input",
        )
        st.session_state["ocr_person_name"] = ocr_person_name

        if uploaded_img and ocr_person_name:
            img_bytes = uploaded_img.read()

            col_img, col_table = st.columns([1, 1])
            with col_img:
                st.image(img_bytes, caption="Imagen cargada", use_container_width=True)

            with col_table:
                if OCR_AVAILABLE:
                    with st.spinner("Ejecutando OCR…"):
                        raw_text = extract_text_from_image(img_bytes)
                    if raw_text.startswith("__OCR_ERROR__"):
                        st.error(raw_text)
                        ocr_df = pd.DataFrame(
                            columns=["DIA", "FECHA", "INGRESO", "SALIDA", "RETORNO", "SALIDA FINAL"]
                        )
                    else:
                        ocr_df = parse_ocr_text_to_df(raw_text)
                        st.success("OCR completado. Revisa y corrige la tabla:")
                else:
                    ocr_df = pd.DataFrame(
                        columns=["DIA", "FECHA", "INGRESO", "SALIDA", "RETORNO", "SALIDA FINAL"]
                    )
                    st.info("Completa la tabla manualmente:")

                edited_ocr = st.data_editor(
                    ocr_df,
                    use_container_width=True,
                    num_rows="dynamic",
                    key="ocr_editor",
                )
                st.session_state["ocr_df"] = edited_ocr

            if st.button("✅ Confirmar y agregar a registros", key="btn_confirm_ocr"):
                try:
                    new_records = ocr_df_to_records(edited_ocr, ocr_person_name)
                    st.session_state["records"] += new_records

                    flat_df = records_to_dataframe(st.session_state["records"])
                    st.session_state["flat_df"]   = flat_df
                    st.session_state["edited_df"] = flat_df.copy()
                    st.session_state["processed"] = True
                    st.success(
                        f"✅ Registros de {ocr_person_name} agregados. "
                        f"Total registros: {len(st.session_state['records'])}"
                    )
                except Exception as exc:
                    st.error(f"Error procesando la imagen: {exc}")

        elif uploaded_img and not ocr_person_name:
            st.warning("👆 Ingresa el nombre de la persona antes de continuar.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PREVIEW & CORRECTIONS
# ══════════════════════════════════════════════════════════════════════════════
with tab_preview:
    if not st.session_state["processed"]:
        st.info("⬅️ Carga un archivo en la pestaña **1. Cargar Archivos** para comenzar.")
    else:
        records  = st.session_state["records"]
        flat_df  = st.session_state["flat_df"]

        # ── Executive summary ─────────────────────────────────────────────
        st.markdown("### 📋 Resumen Ejecutivo")

        people     = [r["name"] for r in records]
        unique_ppl = list(dict.fromkeys(people))  # preserve order, deduplicate

        months_years = [(r["month"], r["year"]) for r in records if r["month"] and r["year"]]
        all_days = []
        for r in records:
            m, y = r.get("month") or 1, r.get("year") or 2025
            for d in r["days"].keys():
                try:
                    all_days.append(date(y, m, d))
                except ValueError:
                    pass

        fecha_min = min(all_days).strftime("%d/%m/%Y") if all_days else "—"
        fecha_max = max(all_days).strftime("%d/%m/%Y") if all_days else "—"

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("👥 Personas", len(unique_ppl))
        m2.metric("📅 Registros (persona/mes)", len(records))
        m3.metric("📌 Fecha inicio detectada", fecha_min)
        m4.metric("📌 Fecha fin detectada", fecha_max)

        st.markdown("**Personas detectadas:**")
        badges_html = " ".join(f'<span class="person-badge">{p}</span>' for p in unique_ppl)
        st.markdown(badges_html, unsafe_allow_html=True)

        st.markdown("---")

        # ── Ambiguous marks warning ───────────────────────────────────────
        if warn_ambiguous and flat_df is not None and not flat_df.empty:
            warnings = []
            for _, row in flat_df.iterrows():
                parsed = parse_times(str(row.get("Marcaciones Raw", "")))
                if len(parsed) >= 5:
                    warnings.append(
                        f"**{row['Persona']}** · día {row['Día']}/{row.get('Mes','?')}: "
                        f"{len(parsed)} marcaciones → {', '.join(parsed)}"
                    )
            if warnings:
                with st.expander(f"⚠️ {len(warnings)} día(s) con 5+ marcaciones — revisar", expanded=True):
                    for w in warnings:
                        st.markdown(f"- {w}")
                    st.caption("Corrige las marcaciones en la tabla inferior si es necesario.")

        # ── Editable dataframe ────────────────────────────────────────────
        st.markdown("### ✏️ Vista Previa y Corrección de Datos")
        st.caption(
            "Puedes editar directamente la columna **Marcaciones Raw** antes de exportar. "
            "Formato: `HH:MM HH:MM` separado por espacios."
        )

        columns_to_show = ["Persona", "Mes", "Año", "Día", "Marcaciones Raw"]
        if not show_raw_times:
            columns_to_show = [c for c in columns_to_show if c != "Marcaciones Raw"]

        edited_df = st.data_editor(
            st.session_state["edited_df"][columns_to_show],
            use_container_width=True,
            num_rows="fixed",
            key="main_editor",
            column_config={
                "Persona":          st.column_config.TextColumn("Persona", disabled=True),
                "Mes":              st.column_config.NumberColumn("Mes", disabled=True, format="%d"),
                "Año":              st.column_config.NumberColumn("Año", disabled=True, format="%d"),
                "Día":              st.column_config.NumberColumn("Día", disabled=True),
                "Marcaciones Raw":  st.column_config.TextColumn(
                    "Marcaciones Raw",
                    help="Edita las horas separadas por espacios, ej: 08:00 12:30 13:30 17:45",
                ),
            },
            height=420,
        )

        # Merge edits back with all columns
        merged = st.session_state["edited_df"].copy()
        for col in columns_to_show:
            if col in edited_df.columns:
                merged[col] = edited_df[col].values
        st.session_state["edited_df"] = merged

        st.markdown("---")

        # ── Per-person detail ─────────────────────────────────────────────
        with st.expander("🔎 Ver detalle por persona"):
            sel_person = st.selectbox(
                "Selecciona persona",
                options=unique_ppl,
                key="detail_person",
            )
            if sel_person:
                person_df = merged[merged["Persona"] == sel_person].copy()
                person_df["Horas Parseadas"] = person_df["Marcaciones Raw"].apply(
                    lambda x: "  |  ".join(parse_times(str(x))) if x else "—"
                )
                person_df["# Marcaciones"] = person_df["Marcaciones Raw"].apply(
                    lambda x: len(parse_times(str(x)))
                )
                max_m = int(person_df["# Marcaciones"].max()) if not person_df.empty else 0
                schema = get_column_schema(max_m)
                st.info(f"Máximo marcaciones: **{max_m}** → Columnas: `{'  |  '.join(schema)}`")
                st.dataframe(
                    person_df[["Día", "Marcaciones Raw", "Horas Parseadas", "# Marcaciones"]],
                    use_container_width=True,
                    hide_index=True,
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with tab_export:
    if not st.session_state["processed"]:
        st.info("⬅️ Carga y revisa los datos primero.")
    else:
        records  = st.session_state["records"]
        edited   = st.session_state["edited_df"]

        st.markdown("### 📊 Configuración del Reporte")

        col_left, col_right = st.columns([2, 1])

        with col_left:
            st.markdown("**Hojas que se generarán:**")
            for rec in records:
                name  = rec["name"]
                month = rec["month"] or 1
                year  = rec["year"] or 2025
                sheet = build_sheet_name(name, month, year)
                days_count = len(rec["days"])
                st.markdown(
                    f'<span class="person-badge">{sheet}</span> '
                    f'<small style="color:#7986cb">{days_count} día(s) con datos</small>',
                    unsafe_allow_html=True,
                )

        with col_right:
            st.markdown("**Formato de salida:**")
            st.markdown("- ✅ Una hoja por persona/mes")
            st.markdown("- ✅ Encabezados azul claro")
            st.markdown("- ✅ Fines de semana gris (#F2F2F2)")
            st.markdown("- ✅ Celdas formato texto")
            st.markdown("- ✅ Fechas DD/MM/YYYY (Ecuador)")
            st.markdown("- ✅ Columna **PAGAR** automatizada")

        st.markdown("---")

        # ── Build edited overrides dict ───────────────────────────────────
        edited_overrides = {}
        if edited is not None and not edited.empty:
            for _, row in edited.iterrows():
                try:
                    key = (
                        str(row["Persona"]),
                        int(row["Mes"]),
                        int(row["Año"]),
                        int(row["Día"]),
                    )
                    edited_overrides[key] = str(row.get("Marcaciones Raw", "") or "")
                except (ValueError, TypeError, KeyError):
                    pass

        # ── Generate button ───────────────────────────────────────────────
        st.markdown("### 🚀 Generar y Descargar")

        if st.button("⚙️ Generar Reporte Excel", use_container_width=True, key="btn_generate"):
            with st.spinner("Generando reporte… por favor espera."):
                try:
                    xlsx_bytes = export_to_xlsx(records, edited_overrides)
                    from datetime import datetime
                    gen_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.session_state["xlsx_bytes"]    = xlsx_bytes
                    st.session_state["xlsx_filename"] = f"Asistencia_{gen_time}.xlsx"
                    st.session_state["xlsx_gen_id"]   = gen_time  # unique per generation
                    st.success("✅ Reporte generado correctamente.")
                except Exception as exc:
                    st.error(f"❌ Error generando el reporte: `{exc}`")
                    st.session_state["xlsx_bytes"] = None

        if st.session_state.get("xlsx_bytes"):
            import base64
            xlsx_data = st.session_state["xlsx_bytes"]
            fname = st.session_state.get("xlsx_filename", "Asistencia.xlsx")

            # ── Provide a clickable HTML download link ───────────────
            b64 = base64.b64encode(xlsx_data).decode()
            href = (
                f'<a href="data:application/vnd.openxmlformats-officedocument'
                f'.spreadsheetml.sheet;base64,{b64}" '
                f'download="{fname}" '
                f'style="display:inline-block;background:linear-gradient(135deg,#00897b,#00bfa5);'
                f'color:white;padding:12px 32px;border-radius:8px;font-weight:700;'
                f'font-size:1rem;text-decoration:none;text-align:center;margin:8px 0;'
                f'box-shadow:0 2px 12px rgba(0,191,165,0.4);transition:all 0.2s ease;"'
                f'>⬇️ Descargar {fname}</a>'
            )
            st.markdown(href, unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### 📋 Vista Previa del Reporte")
            st.caption("Primeras filas de cada persona procesada")

            for rec in records:
                name  = rec["name"]
                month = rec["month"] or 1
                year  = rec["year"] or 2025


                days = dict(rec["days"])
                # Apply edits
                for (n, m, y, d), raw in edited_overrides.items():
                    if n == name and m == month and y == year:
                        days[d] = raw

                all_parsed = {d: parse_times(raw) for d, raw in days.items()}
                max_m = max((len(v) for v in all_parsed.values()), default=0)
                schema = get_column_schema(max_m)

                preview_rows = []
                import calendar as cal_mod
                days_in_month = cal_mod.monthrange(year, month)[1]
                _DAYS_ES_PREV = {
                    0: "Lunes", 1: "Martes", 2: "Miércoles", 3: "Jueves",
                    4: "Viernes", 5: "Sábado", 6: "Domingo",
                }

                # Build unique column display headers for repeated names
                # e.g. [INGRESO, SALIDA, INGRESO, SALIDA, INGRESO, SALIDA FINAL]
                # → [INGRESO, SALIDA, INGRESO_2, SALIDA_2, INGRESO_3, SALIDA FINAL]
                def _unique_headers(s):
                    seen = {}
                    result = []
                    for col in s:
                        if col not in seen:
                            seen[col] = 1
                            result.append(col)
                        else:
                            seen[col] += 1
                            result.append(f"{col}_{seen[col]}")
                    return result

                uniq_schema = _unique_headers(schema)

                for day_num in range(1, min(days_in_month + 1, 8)):
                    try:
                        d_obj   = date(year, month, day_num)
                        raw_t   = days.get(day_num, "")
                        parsed  = parse_times(raw_t)
                        mapped  = assign_marks_to_columns(parsed, schema)
                        pos     = mapped.get("_pos", {})
                        row_data = {
                            "FECHA": d_obj.strftime("%d/%m/%Y"),
                            "DIA":   _DAYS_ES_PREV[d_obj.weekday()],
                        }
                        for pos_i, uniq_col in enumerate(uniq_schema):
                            row_data[uniq_col] = pos.get(pos_i, "") if pos else mapped.get(schema[pos_i], "")
                        preview_rows.append(row_data)
                    except ValueError:
                        pass

                with st.expander(f"👤 {build_sheet_name(name, month, year)}", expanded=False):
                    if preview_rows:
                        preview_df = pd.DataFrame(preview_rows)

                        # ── Add PAGAR column to preview ───────────────────
                        _RATE_WD = 3.26
                        _RATE_WE = 6.27
                        pairs_schema = get_entry_exit_pairs(schema)

                        def _compute_pagar(row):
                            """Return pay string or warning for one preview row."""
                            if not pairs_schema:
                                return ""
                            # Detect weekend from DIA column
                            is_we = row.get("DIA", "") in ("Sábado", "Domingo")
                            rate = _RATE_WE if is_we else _RATE_WD

                            total_h = 0.0
                            warn = False
                            for (ei, xi) in pairs_schema:
                                # Use unique column names to look up correct slot
                                entry_col = uniq_schema[ei] if ei < len(uniq_schema) else ""
                                exit_col  = uniq_schema[xi] if xi < len(uniq_schema) else ""
                                entry_v = str(row.get(entry_col, "") or "").strip()
                                exit_v  = str(row.get(exit_col,  "") or "").strip()
                                if bool(entry_v) != bool(exit_v):
                                    warn = True
                                    break
                                if entry_v and exit_v:
                                    try:
                                        eh, em = map(int, entry_v.split(":"))
                                        xh, xm = map(int, exit_v.split(":"))
                                        dh = (xh * 60 + xm - eh * 60 - em) / 60.0
                                        if dh < 0:
                                            dh += 24
                                        total_h += dh
                                    except (ValueError, TypeError):
                                        warn = True
                                        break
                            if warn:
                                return "⚠️ Revisar"
                            if total_h == 0:
                                return ""
                            
                            return f"${total_h * rate:.2f}"

                        preview_df["PAGAR"] = preview_df.apply(_compute_pagar, axis=1)

                        # Color-code PAGAR column
                        def _style_pagar(val):
                            if str(val).startswith("⚠️"):
                                return "background-color:#FF6600;color:white;font-weight:bold"
                            if str(val).startswith("$"):
                                return "background-color:#E2EFDA;color:#375623;font-weight:600"
                            return ""

                        st.dataframe(
                            preview_df.style.applymap(_style_pagar, subset=["PAGAR"]),
                            use_container_width=True,
                            hide_index=True,
                        )
                    else:
                        st.caption("Sin datos para mostrar.")

# End of file (Force Streamlit reload)
