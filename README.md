# 🕐 Procesador de Marcaciones Biométricas

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)
![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![License](https://img.shields.io/badge/License-MIT-green)

**App web profesional para automatizar el procesamiento de marcaciones biométricas**, convirtiendo archivos `.xlsx` del reloj biométrico (y registros manuales en imagen) en reportes de asistencia formateados por persona y mes, listos para RR.HH.

---

## 🚀 Demo rápida

> Visita: **`https://marcaciones-app-vy6eabdnnha95uhrzu7pwh.streamlit.app/`**

---

## ✨ Características principales

| Funcionalidad               | Detalle                                                              |
| --------------------------- | -------------------------------------------------------------------- |
| 📊 **Ingesta Excel**        | Lee archivos `.xlsx` del biométrico con horas pegadas (`09:5712:33`) |
| 🖼️ **OCR de imágenes**      | Extrae registros manuales desde fotos `.png/.jpg` vía pytesseract    |
| 🔁 **Parseo regex robusto** | Separa horas pegadas, elimina duplicados exactos consecutivos        |
| 📋 **Asignación dinámica**  | INGRESO / SALIDA / RETORNO / SALIDA FINAL según máximo detectado     |
| ✏️ **Corrección en línea**  | Edición celda a celda antes de exportar                              |
| 📈 **Dashboard ejecutivo**  | Resumen de personas, fechas y alertas de ambigüedad                  |
| 🎨 **Formato profesional**  | Encabezados azules, fines de semana grises, celdas como texto        |
| 🇪🇨 **Locale Ecuador**       | Fechas `DD/MM/YYYY` en todas las salidas                             |

---

## 🗂️ Estructura del proyecto

```
marcaciones-app/
├── app.py                  # Punto de entrada Streamlit (UI)
├── requirements.txt        # Dependencias Python
├── .streamlit/
│   └── config.toml         # Tema dark premium
├── core/
│   ├── __init__.py
│   ├── parser.py           # Lógica de parseo de horas y asignación de columnas
│   ├── reader.py           # Lectura y normalización del .xlsx biométrico
│   ├── ocr.py              # Extracción de texto desde imágenes
│   └── exporter.py         # Generación del reporte .xlsx formateado
└── README.md
```

---

## 📋 Formato de entrada esperado

### Excel biométrico (`.xlsx`)

El procesador autodetecta **3 formatos** distintos por cada hoja:

**Formato A: Pre-procesado (Tabla limpia)**

```
Fila 1: "VICKY – Julio 2025" (Nombre y mes)
Fila 2: "Período: 01/07/2025 ~ 31/07/2025"
Fila 4: FECHA | DIA | INGRESO | SALIDA ...
Fila 5+: Datos diarios
```

**Formato B: Matriz Cruda (Con fila de Periodo)**

```
Fila 1: Periodo: | 2025-07-01 ~ 2025-07-31
Fila 2: 1 | 2 | 3 | 4 ... 31                  ← Números de día en cabecera
Fila 3: ID: | cédula | Nombre: | VICKY        ← Datos de la persona
Fila 4: 09:5712:33 | 08:0017:00 | ...         ← Horas pegadas
(Se repite bloque de ID y horas por persona)
```

**Formato C: Matriz Cruda (Sin Periodo)**
Idéntico al Formato B, pero sin la primera fila de `Periodo:`. Inicia directamente con la numeración de los días del 1 al 31.

### Imagen de registro manual

Tabla con columnas: `DIA | FECHA | HORA DE INGRESO | HORA DE SALIDA`

---

## 📊 Reglas de asignación de columnas

| Marcaciones | Columnas generadas                                               |
| ----------- | ---------------------------------------------------------------- |
| 1           | INGRESO                                                          |
| 2           | INGRESO · SALIDA FINAL                                           |
| 3           | INGRESO · SALIDA · SALIDA FINAL                                  |
| 4           | INGRESO · SALIDA · RETORNO · SALIDA FINAL                        |
| 5           | INGRESO · SALIDA · RETORNO · SALIDA FINAL + alerta               |
| 6           | INGRESO · SALIDA · RETORNO · INGRESO 2 · SALIDA 2 · SALIDA FINAL |

---

## 🎨 Formato del reporte de salida

- **Una hoja por persona/mes** → nombre: `VICKY Jul25`
- **Fila 1**: Nombre en rojo/negrita/fusionado
- **Fila 2**: `Período: 01/07/2025 ~ 31/07/2025`
- **Fila 4**: Encabezados en azul claro (`D9E1F2`)
- **Filas 5+**: Un día por fila, fines de semana en gris (`F2F2F2`)
- Todas las celdas con `number_format = '@'` (texto puro)

---

## 🛠️ Instalación local

### Prerequisitos

- Python 3.10+
- (Opcional para OCR) [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)

### Pasos

```bash
# 1. Clona o descarga el proyecto
cd marcaciones-app

# 2. Crea un entorno virtual (recomendado)
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # Linux/Mac

# 3. Instala dependencias
pip install -r requirements.txt

# 4. Ejecuta la app
streamlit run app.py
```

La app estará disponible en `http://localhost:8501`

---

## ☁️ Despliegue en Streamlit Cloud

### Pasos:

1. **Sube el proyecto a GitHub** (repositorio público o privado)

   ```bash
   git init
   git add .
   git commit -m "feat: initial marcaciones app"
   git remote add origin https://github.com/tu-usuario/marcaciones-app.git
   git push -u origin main
   ```

2. **Ve a [share.streamlit.io](https://share.streamlit.io)** e inicia sesión con GitHub

3. Clic en **"New app"** → selecciona tu repositorio

4. Configura:
   - **Main file path**: `app.py`
   - **Python version**: `3.11`

5. Clic en **"Deploy!"**

> ⚠️ **Nota sobre OCR en Streamlit Cloud**: `pytesseract` requiere que Tesseract esté instalado en el servidor. Para activarlo, crea el archivo `packages.txt` en la raíz:
>
> ```
> tesseract-ocr
> tesseract-ocr-spa
> ```

---

## 👥 Optimización multi-usuario

La app está diseñada para **hasta 6 usuarios concurrentes**:

- Estado de sesión completamente aislado por usuario via `st.session_state`
- Sin escritura en disco durante el procesamiento (todo en memoria con `io.BytesIO`)
- Excepciones capturadas por persona para no bloquear el reporte completo

---

## 🔧 Casos especiales manejados

| Caso                               | Comportamiento                                                    |
| ---------------------------------- | ----------------------------------------------------------------- |
| Auto-detección de formato          | Diferencia matrices crudas vs tablas preprocesadas                |
| Nombres de Excel inválidos         | Limpia automáticamente caracteres prohibidos (`: / \ ? * [ ]`)    |
| Horas pegadas `09:5712:33`         | Separadas automáticamente con regex                               |
| Duplicados exactos consecutivos    | Eliminados (ej: `08:00 08:00` → `08:00`)                          |
| Duplicados distintos `14:20 14:22` | Conservados + alerta visual                                       |
| Día sin marcaciones                | Celda vacía (no ceros ni guiones)                                 |
| Archivo corrupto                   | Error por persona, no rompe el reporte (genera "Hoja ERROR")      |
| Bug de descarga Streamlit          | Guardado directo a carpeta Descargas + enlace HTML base64 forzado |
| Múltiples períodos en un .xlsx     | Detectados como registros separados                               |
| Hoja con nombre >31 chars          | Truncada automáticamente                                          |

---

## 📦 Dependencias

```
streamlit>=1.32.0
pandas>=2.0.0
openpyxl>=3.1.2
Pillow>=10.0.0
pytesseract>=0.3.10  # Requiere Tesseract instalado en el sistema
```

---

## 📄 Licencia

MIT License — libre para uso comercial y personal.

---

_Desarrollado con ❤️ | Formato Ecuador 🇪🇨 DD/MM/YYYY_
