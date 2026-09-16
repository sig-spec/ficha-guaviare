# -*- coding: utf-8 -*-
# ============================================================================
# Ficha de Cooperacion Internacional - GUAVIARE
# Version simplificada, adaptada a partir de la Ficha nacional de APC-Colombia,
# para trabajar con el archivo de mapeo de proyectos de cooperacion
# internacional en el departamento del Guaviare (formulario de campo 2025-2026).
# ============================================================================
import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
import altair as alt

st.set_page_config(
    page_title="Ficha de Cooperacion Internacional | Guaviare",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --------------------------------------------------------------------------
# ARCHIVOS DE DATOS Y LOGOS (ajusta estos nombres si cambian)
# --------------------------------------------------------------------------
FILE = "Mapeo_CI_Proyectos_a_2025_-_2026.xlsx"
SHEET = "COOPERACION_INTERNACIONAL_M_0"
LOGO_1 = "logo_gobernacion.png"   # opcional, si no existe se omite (derecha)
LOGO_2 = "logo_planeacion.png"    # opcional, si no existe se omite (izquierda)

# ============================================================================
# ESTILOS (identidad visual institucional, sin franja/acentos en rojo)
# ============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&family=Source+Sans+3:wght@400;600&display=swap');

:root {
    --apc-blue: #0765AD;
    --apc-blue-dark: #054F82;
    --apc-green: #00A859;
    --apc-yellow: #FDBC2D;
    --apc-light: #EAF4FB;
    --apc-gray: #F7F8FA;
    --apc-border: #D6E6F0;
    --apc-text: #0B3B5C;
    --apc-muted: #5A7185;
}

html, body, [class*="css"] {
    font-family: 'Source Sans 3', sans-serif;
    color: var(--apc-text);
}
:root { color-scheme: light only; }
.stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], body {
    background-color: #FFFFFF !important;
}
div[data-testid="stToolbar"], div[data-testid="stDecoration"], header[data-testid="stHeader"] {
    display: none !important;
}

/* Fuerza texto legible aunque el navegador este en modo oscuro */
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stCaptionContainer"] p,
[data-testid="stRadio"] label span,
[data-testid="stWidgetLabel"] p,
[data-testid="stSelectbox"] label p,
label p {
    color: var(--apc-text) !important;
}
[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    color: var(--apc-text) !important;
    border-color: var(--apc-border) !important;
}
[data-baseweb="select"] * {
    color: var(--apc-text) !important;
}
[data-baseweb="popover"], [data-baseweb="menu"] {
    background-color: #FFFFFF !important;
}
li[role="option"] {
    background-color: #FFFFFF !important;
    color: var(--apc-text) !important;
}
li[role="option"]:hover, li[aria-selected="true"] {
    background-color: var(--apc-light) !important;
}
/* Texto de los botones de descarga: debe quedar blanco sobre el boton azul */
[data-testid="stDownloadButton"] p,
[data-testid="stDownloadButton"] div[data-testid="stMarkdownContainer"] p {
    color: #FFFFFF !important;
}
/* Excepciones: estos textos van sobre fondo azul, deben quedar blancos */
[data-testid="stMarkdownContainer"] p.apc-header-title {
    color: #FFFFFF !important;
}
[data-testid="stMarkdownContainer"] p.apc-header-subtitle {
    color: rgba(255,255,255,0.78) !important;
}

.apc-header {
    background: var(--apc-blue);
    padding: 1.2rem 2.2rem 1rem 2.2rem;
}
.apc-header-title {
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.4rem;
    font-weight: 700;
    margin: 0;
}
.apc-header-subtitle {
    color: rgba(255,255,255,0.78);
    font-size: 0.85rem;
    margin-top: 4px;
}
.apc-flag-bar {
    height: 5px;
    background: linear-gradient(90deg, var(--apc-green) 33.3%, var(--apc-blue) 33.3% 66.6%, var(--apc-yellow) 66.6%);
    margin-bottom: 1.4rem;
}

.dept-title-banner {
    background: var(--apc-blue);
    color: white;
    font-family: 'Montserrat', sans-serif;
    font-size: 1.15rem;
    font-weight: 700;
    padding: 0.75rem 1.5rem;
    border-radius: 6px;
    margin-bottom: 1.2rem;
    border-left: 6px solid var(--apc-yellow);
}

.section-header {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 0.92rem;
    color: var(--apc-blue);
    text-transform: uppercase;
    letter-spacing: 1px;
    border-bottom: 3px solid var(--apc-yellow);
    padding-bottom: 6px;
    margin: 1.8rem 0 1rem 0;
}

div[data-testid="stMetric"] {
    background: white;
    border: 1px solid var(--apc-border);
    border-left: 5px solid var(--apc-blue);
    border-radius: 6px;
    padding: 1rem 1.1rem !important;
    box-shadow: 0 1px 6px rgba(0,48,135,0.06);
}
div[data-testid="stMetricLabel"] p {
    font-family: 'Montserrat', sans-serif;
    font-weight: 600;
    font-size: 0.65rem !important;
    color: var(--apc-muted) !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
div[data-testid="stMetricValue"] {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    font-size: 1.5rem !important;
    color: var(--apc-blue) !important;
}

button[data-baseweb="tab"] {
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    color: var(--apc-muted) !important;
}
button[data-baseweb="tab"][aria-selected="true"] { color: var(--apc-blue) !important; }

div[data-testid="stDataFrame"] {
    border-radius: 6px;
    border: 1px solid var(--apc-border);
    overflow: hidden;
}
[data-testid="stArrowVegaLiteChart"], [data-testid="stVegaLiteChart"] {
    background-color: #FFFFFF !important;
}

div[data-testid="stDownloadButton"] button {
    background: var(--apc-blue) !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    text-transform: uppercase;
}
div[data-testid="stDownloadButton"] button:hover { background: var(--apc-blue-dark) !important; }

.guia-card {
    background: white;
    border: 1px solid var(--apc-border);
    border-radius: 8px;
    padding: 2rem 2.4rem;
    line-height: 1.8;
    font-size: 0.97rem;
}
.guia-card p { margin-bottom: 1rem; }
.guia-intro {
    font-family: 'Montserrat', sans-serif;
    font-weight: 700;
    color: var(--apc-blue);
    background: var(--apc-light);
    border-left: 5px solid var(--apc-yellow);
    border-radius: 0 6px 6px 0;
    padding: 0.8rem 1.2rem;
    margin-bottom: 1.4rem;
    text-transform: uppercase;
}

.apc-footer {
    text-align: center;
    color: var(--apc-muted);
    font-size: 0.76rem;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid var(--apc-border);
}
</style>
""", unsafe_allow_html=True)

# ============================================================================
# MAPEOS DE ETIQUETAS (los datos vienen codificados desde el formulario)
# ============================================================================
MUNICIPIOS_MAP = {
    "san_jose_del_guaviare": "San Jose del Guaviare",
    "calamar": "Calamar",
    "el_retorno": "El Retorno",
    "miraflores": "Miraflores",
}

POBLACION_MAP = {
    "ninos": "Niños", "ninas": "Niñas",
    "jovenes": "Jóvenes", "mujeres": "Mujeres", "hombres": "Hombres", "adultos": "Adultos",
    "etnicos": "Étnicos", "afro": "Afro", "desplazados": "Desplazados",
    "discapacitados": "Discapacitados", "reincorporados": "Reincorporados",
    "lgbtiq": "LGBTIQ+",
}

ODS_NOMBRES = {
    "1": "ODS 1 - Fin de la pobreza", "2": "ODS 2 - Hambre cero",
    "3": "ODS 3 - Salud y bienestar", "4": "ODS 4 - Educacion de calidad",
    "5": "ODS 5 - Igualdad de genero", "6": "ODS 6 - Agua limpia y saneamiento",
    "7": "ODS 7 - Energia asequible y no contaminante",
    "8": "ODS 8 - Trabajo decente y crecimiento economico",
    "9": "ODS 9 - Industria, innovacion e infraestructura",
    "10": "ODS 10 - Reduccion de las desigualdades",
    "11": "ODS 11 - Ciudades y comunidades sostenibles",
    "12": "ODS 12 - Produccion y consumo responsables",
    "13": "ODS 13 - Accion por el clima", "14": "ODS 14 - Vida submarina",
    "15": "ODS 15 - Vida de ecosistemas terrestres",
    "16": "ODS 16 - Paz, justicia e instituciones solidas",
    "17": "ODS 17 - Alianzas para lograr los objetivos",
}


def normalize_key(s):
    """Normaliza texto a snake_case sin tildes, para poder comparar/mapear."""
    if not isinstance(s, str):
        return ""
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[\s\-]+", "_", s)
    s = re.sub(r"[^a-z0-9_]", "", s)
    return s


def split_multi_normalized(value):
    """Campos de seleccion multiple (ej. municipios) separados por coma."""
    if not isinstance(value, str) or not value.strip():
        return []
    return [normalize_key(p) for p in value.split(",") if p.strip()]


def split_multi_plain(value):
    """Campos de seleccion multiple donde se conserva el texto tal cual."""
    if not isinstance(value, str) or not value.strip():
        return []
    return [p.strip() for p in value.split(",") if p.strip()]


def extract_ods_list(value):
    if not isinstance(value, str) or not value.strip():
        return []
    return re.findall(r"\d+", value)


def municipio_label(x):
    if not x:
        return "Sin dato"
    return MUNICIPIOS_MAP.get(x, x.replace("_", " ").title())


def sector_label(x):
    if not x:
        return "Sin dato"
    return x.replace("_", " ").strip().title()


def poblacion_label(x):
    if not x:
        return "Sin dato"
    return POBLACION_MAP.get(normalize_key(x), x.strip().title())


def ods_label(x):
    return ODS_NOMBRES.get(x, f"ODS {x}")


MESES_NOMBRES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def format_int(n):
    try:
        n = int(n)
    except Exception:
        return str(n)
    return f"{n:,}".replace(",", ".")


# ============================================================================
# CARGA Y LIMPIEZA DE DATOS
# ============================================================================
COLMAP = {
    "1. Fecha de registro de la intervención": "fecha_registro",
    "5.Pais de origen del donante": "donante_pais",
    "6.Nombre del donante/cooperante": "donante_nombre",
    "8.Nombre de la entidad /organización que ejecuta el proyecto.": "organizacion_ejecutora",
    "9.Nombre del proyecto/programa": "nombre_intervencion",
    "10.Fecha de inicio": "fecha_inicial",
    "11.Fecha de finalización": "fecha_final",
    "13.Seleccione el sector que mejor describe el proyecto": "sector_raw",
    "14.Objetivo general": "objetivo",
    "19.Municipios a intervenir": "municipio_raw",
    "23.Qué población atiende el proyecto": "poblacion_raw",
    "29.Mencione el/los objetivos de Desarrollo Sostenible-ODS (Agenda 2030) a los que apunta el proyecto": "ods_raw",
}


@st.cache_data
def load_data():
    raw = pd.read_excel(FILE, sheet_name=SHEET)
    df = raw.rename(columns=COLMAP)
    for needed in COLMAP.values():
        if needed not in df.columns:
            df[needed] = ""

    text_cols = [
        "donante_pais", "donante_nombre", "organizacion_ejecutora", "nombre_intervencion",
        "sector_raw", "objetivo", "municipio_raw", "poblacion_raw", "ods_raw",
    ]
    # Nota: se limpia por columna (no por dtype) porque pandas >= 2.x puede
    # asignar dtype "str" en vez de "object" a columnas de texto, y un chequeo
    # por dtype == "object" las deja pasar sin limpiar.
    for c in text_cols:
        df[c] = df[c].where(df[c].notna(), "")
        df[c] = df[c].astype(str).str.strip()
        df[c] = df[c].replace({"nan": "", "None": "", "NA": "", "na": "", "N/A": ""})

    df["fecha_registro"] = pd.to_datetime(df["fecha_registro"], errors="coerce")
    df["fecha_inicial"] = pd.to_datetime(df["fecha_inicial"], errors="coerce")
    df["fecha_final"] = pd.to_datetime(df["fecha_final"], errors="coerce")
    df["anio_inicio"] = df["fecha_inicial"].dt.year
    df["anio_fin"] = df["fecha_final"].dt.year
    df["mes_fin"] = df["fecha_final"].dt.month

    df["municipio_list"] = df["municipio_raw"].apply(split_multi_normalized)
    df["sector_list"] = df["sector_raw"].apply(split_multi_plain)
    df["poblacion_list"] = df["poblacion_raw"].apply(split_multi_plain)
    df["ods_list"] = df["ods_raw"].apply(extract_ods_list)

    hoy = pd.Timestamp.now().normalize()
    df["activo_hoy"] = (df["fecha_inicial"] <= hoy) & (df["fecha_final"] >= hoy)
    return df


def counts_table(series, label_fn, value_name="intervenciones"):
    """Para columnas de un solo valor por fila (ej. donante, organizacion)."""
    s = series.replace("", pd.NA).dropna()
    if s.empty:
        return pd.DataFrame(columns=["clave", "etiqueta", value_name])
    out = s.value_counts().reset_index()
    out.columns = ["clave", value_name]
    out["etiqueta"] = out["clave"].map(label_fn)
    return out


def counts_from_lists(series_of_lists, label_fn, value_name="intervenciones"):
    """Para columnas de seleccion multiple (ej. sector, municipio, ODS)."""
    exploded = series_of_lists.explode()
    exploded = exploded.dropna()
    exploded = exploded[exploded != ""]
    if exploded.empty:
        return pd.DataFrame(columns=["clave", "etiqueta", value_name])
    out = exploded.value_counts().reset_index()
    out.columns = ["clave", value_name]
    out["etiqueta"] = out["clave"].map(label_fn)
    return out


def bar_chart(df, y_field, x_field, color="#0765AD"):
    text_dark = "#0B3B5C"
    grid_color = "#E3ECF2"
    chart = (
        alt.Chart(df)
        .mark_bar(color=color, cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
        .encode(
            y=alt.Y(f"{y_field}:N", sort="-x", title="",
                    axis=alt.Axis(labelColor=text_dark, labelFontSize=11, domainColor=grid_color, tickColor=grid_color)),
            x=alt.X(f"{x_field}:Q", title="Intervenciones",
                    axis=alt.Axis(labelColor=text_dark, titleColor=text_dark, gridColor=grid_color, domainColor=grid_color, tickColor=grid_color)),
            tooltip=[f"{y_field}:N", f"{x_field}:Q"],
        )
        .properties(height=max(180, 32 * len(df)), background="#FFFFFF")
        .configure_view(strokeWidth=0)
    )
    return chart


# ============================================================================
# EXPORTACION: EXCEL Y PDF
# ============================================================================
def build_export_df(df):
    d = df.copy()
    d["Sector(es)"] = d["sector_list"].apply(lambda l: ", ".join(sector_label(x) for x in l) if l else "Sin dato")
    d["Municipio(s)"] = d["municipio_list"].apply(lambda l: ", ".join(municipio_label(x) for x in l) if l else "Sin dato")
    d["Poblacion atendida"] = d["poblacion_list"].apply(lambda l: ", ".join(poblacion_label(x) for x in l) if l else "Sin dato")
    d["ODS"] = d["ods_list"].apply(lambda l: ", ".join(ods_label(x) for x in l) if l else "Sin dato")
    d["Fecha inicio"] = d["fecha_inicial"].dt.strftime("%Y-%m-%d").fillna("")
    d["Fecha finalizacion"] = d["fecha_final"].dt.strftime("%Y-%m-%d").fillna("")

    d = d.rename(columns={
        "nombre_intervencion": "Nombre del proyecto",
        "organizacion_ejecutora": "Organizacion ejecutora",
        "donante_nombre": "Donante / cooperante",
        "donante_pais": "Pais de origen del donante",
        "objetivo": "Objetivo general",
    })
    cols = [
        "Nombre del proyecto", "Organizacion ejecutora", "Donante / cooperante",
        "Pais de origen del donante", "Sector(es)", "Municipio(s)",
        "Fecha inicio", "Fecha finalizacion", "Poblacion atendida", "ODS",
        "Objetivo general",
    ]
    cols = [c for c in cols if c in d.columns]
    return d[cols]


def to_excel(df, sector_counts, poblacion_counts, ods_counts, mun_counts):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        build_export_df(df).to_excel(writer, sheet_name="Intervenciones", index=False)
        sector_counts.rename(columns={"etiqueta": "Sector", "intervenciones": "Intervenciones"})[
            ["Sector", "Intervenciones"]
        ].to_excel(writer, sheet_name="Sectores", index=False)
        poblacion_counts.rename(columns={"etiqueta": "Poblacion", "intervenciones": "Intervenciones"})[
            ["Poblacion", "Intervenciones"]
        ].to_excel(writer, sheet_name="Poblacion atendida", index=False)
        ods_counts.rename(columns={"etiqueta": "ODS", "intervenciones": "Intervenciones"})[
            ["ODS", "Intervenciones"]
        ].to_excel(writer, sheet_name="ODS", index=False)
        mun_counts.rename(columns={"etiqueta": "Municipio", "intervenciones": "Intervenciones"})[
            ["Municipio", "Intervenciones"]
        ].to_excel(writer, sheet_name="Municipios", index=False)
    output.seek(0)
    return output.getvalue()


def to_pdf(df_export, kpis):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.5 * cm, rightMargin=1.5 * cm
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleAPC", parent=styles["Heading1"],
                                  textColor=colors.HexColor("#0765AD"), fontSize=15, spaceAfter=4)
    sub_style = ParagraphStyle("SubAPC", parent=styles["Normal"],
                                textColor=colors.HexColor("#5A6A85"), fontSize=9, spaceAfter=10)
    section_style = ParagraphStyle("SectionAPC", parent=styles["Heading2"],
                                    textColor=colors.HexColor("#0765AD"), fontSize=11,
                                    spaceBefore=14, spaceAfter=6)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=7.5, leading=9)

    elements = [
        Paragraph("Ficha de Cooperacion Internacional - Guaviare", title_style),
        Paragraph("Mapeo de proyectos de cooperacion internacional en Guaviare. Periodo: 2025-2026.", sub_style),
        HRFlowable(width="100%", color=colors.HexColor("#FDBC2D"), thickness=2, spaceAfter=10),
    ]

    kpi_data = [["Indicador", "Valor"]] + kpis
    kpi_table = Table(kpi_data, colWidths=[9 * cm, 6 * cm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0765AD")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D9E6")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F8FA")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elements += [kpi_table, Spacer(1, 0.6 * cm), Paragraph("Listado de proyectos", section_style)]

    table_data = [["Proyecto", "Sector(es)", "Municipio(s)", "Organizacion", "Inicio", "Fin"]]
    for _, row in df_export.iterrows():
        table_data.append([
            Paragraph(str(row.get("Nombre del proyecto", ""))[:60], cell_style),
            Paragraph(str(row.get("Sector(es)", ""))[:35], cell_style),
            Paragraph(str(row.get("Municipio(s)", ""))[:35], cell_style),
            Paragraph(str(row.get("Organizacion ejecutora", ""))[:35], cell_style),
            str(row.get("Fecha inicio", "")),
            str(row.get("Fecha finalizacion", "")),
        ])
    detail_table = Table(table_data, colWidths=[4.3 * cm, 2.6 * cm, 2.6 * cm, 3.2 * cm, 1.9 * cm, 1.9 * cm], repeatRows=1)
    detail_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0765AD")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D1D9E6")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F8FA")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(detail_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================================
# CARGA DE DATOS
# ============================================================================
df = load_data()

# ============================================================================
# HEADER
# ============================================================================
col_logo_izq, col_title, col_logo_der = st.columns([1, 3, 1], vertical_alignment="center")
with col_logo_izq:
    try:
        st.image(LOGO_2, width=200)
    except Exception:
        pass
with col_title:
    st.markdown(
        '<div style="padding: 0.6rem 0 0.2rem 0; text-align:center;">'
        '<div style="font-family:Montserrat,sans-serif;font-weight:800;font-size:1.9rem;color:#0765AD;line-height:1.3;">'
        'Ficha de Cooperacion Internacional'
        '</div>'
        '<div style="font-size:1rem;color:#5A6A85;margin-top:6px;">'
        'Departamento del Guaviare &mdash; Mapeo de proyectos de cooperacion internacional 2025-2026'
        '</div></div>',
        unsafe_allow_html=True
    )
with col_logo_der:
    try:
        st.image(LOGO_1, width=200)
    except Exception:
        pass

st.markdown('<div class="apc-flag-bar"></div>', unsafe_allow_html=True)

nav_options = ["\U0001F4CA Panorama Guaviare", "\U0001F4D6 Guia de usuario"]
nav = st.radio("", nav_options, horizontal=True, label_visibility="collapsed", key="main_nav")
st.markdown("---")

# ============================================================================
# PANORAMA GUAVIARE
# ============================================================================
if nav == nav_options[0]:
    st.markdown('<div class="dept-title-banner">GUAVIARE &mdash; Cooperacion Internacional</div>', unsafe_allow_html=True)

    # -- Filtros: municipio y año --
    mun_keys_presentes = sorted(df["municipio_list"].explode().replace("", pd.NA).dropna().unique().tolist())
    mun_label_to_key = {"Todos los municipios": None}
    mun_label_to_key.update({municipio_label(k): k for k in mun_keys_presentes})
    mun_opciones = list(mun_label_to_key.keys())

    anios_inicio_presentes = sorted(df["anio_inicio"].dropna().unique().tolist())
    anio_ini_label_to_key = {"Todos los años": None}
    anio_ini_label_to_key.update({str(int(a)): int(a) for a in anios_inicio_presentes})
    anio_ini_opciones = list(anio_ini_label_to_key.keys())

    anios_fin_presentes = sorted(df["anio_fin"].dropna().unique().tolist())
    anio_fin_label_to_key = {"Todos los años": None}
    anio_fin_label_to_key.update({str(int(a)): int(a) for a in anios_fin_presentes})
    anio_fin_opciones = list(anio_fin_label_to_key.keys())

    meses_fin_presentes = sorted(df["mes_fin"].dropna().unique().tolist())
    mes_fin_label_to_key = {"Todos los meses": None}
    mes_fin_label_to_key.update({MESES_NOMBRES[int(m)]: int(m) for m in meses_fin_presentes})
    mes_fin_opciones = list(mes_fin_label_to_key.keys())

    col_filtro1, col_filtro2, col_filtro3, col_filtro4 = st.columns(4)
    with col_filtro1:
        mun_sel_label = st.selectbox("Filtrar por municipio", mun_opciones, key="filtro_municipio")
    with col_filtro2:
        anio_sel_label = st.selectbox("Filtrar por año de inicio", anio_ini_opciones, key="filtro_anio_inicio")
    with col_filtro3:
        anio_fin_sel_label = st.selectbox("Filtrar por año de finalizacion", anio_fin_opciones, key="filtro_anio_fin")
    with col_filtro4:
        mes_fin_sel_label = st.selectbox("Filtrar por mes de finalizacion", mes_fin_opciones, key="filtro_mes_fin")
    mun_sel_key = mun_label_to_key[mun_sel_label]
    anio_sel_key = anio_ini_label_to_key[anio_sel_label]
    anio_fin_sel_key = anio_fin_label_to_key[anio_fin_sel_label]
    mes_fin_sel_key = mes_fin_label_to_key[mes_fin_sel_label]

    df_view = df
    if mun_sel_key is not None:
        df_view = df_view[df_view["municipio_list"].apply(lambda lst: mun_sel_key in lst)]
    if anio_sel_key is not None:
        df_view = df_view[df_view["anio_inicio"] == anio_sel_key]
    if anio_fin_sel_key is not None:
        df_view = df_view[df_view["anio_fin"] == anio_fin_sel_key]
    if mes_fin_sel_key is not None:
        df_view = df_view[df_view["mes_fin"] == mes_fin_sel_key]
    df_view = df_view.reset_index(drop=True)

    total_intervenciones = len(df_view)
    total_organizaciones = df_view["organizacion_ejecutora"].replace("", pd.NA).dropna().nunique()
    total_donantes = df_view["donante_nombre"].replace("", pd.NA).dropna().nunique()
    total_municipios = df_view["municipio_list"].explode().replace("", pd.NA).dropna().nunique()
    total_activos = int(df_view["activo_hoy"].sum())

    sector_counts = counts_from_lists(df_view["sector_list"], sector_label)
    poblacion_counts = counts_from_lists(df_view["poblacion_list"], poblacion_label)
    mun_counts = counts_from_lists(df_view["municipio_list"], municipio_label)
    ods_counts = counts_from_lists(df_view["ods_list"], ods_label)
    donante_counts = counts_table(df_view["donante_nombre"], lambda x: x)
    org_counts = counts_table(df_view["organizacion_ejecutora"], lambda x: x)

    filtro_txt = []
    if mun_sel_key is not None:
        filtro_txt.append(f"en {mun_sel_label}")
    if anio_sel_key is not None:
        filtro_txt.append(f"con inicio en {anio_sel_label}")
    if anio_fin_sel_key is not None:
        filtro_txt.append(f"con finalizacion en {anio_fin_sel_label}")
    if mes_fin_sel_key is not None:
        filtro_txt.append(f"con finalizacion en {mes_fin_sel_label}")
    filtro_str = " ".join(filtro_txt) if filtro_txt else "en todo el departamento"

    st.caption(
        f"Fuente: mapeo de proyectos de cooperacion internacional en Guaviare. "
        f"Periodo: 2025-2026. {total_intervenciones} proyectos registrados {filtro_str}."
    )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Proyectos", format_int(total_intervenciones))
    c2.metric("Organizaciones ejecutoras", format_int(total_organizaciones))
    c3.metric("Donantes / cooperantes", format_int(total_donantes))
    c4.metric("Municipios", format_int(total_municipios))
    c5.metric("Proyectos vigentes hoy", format_int(total_activos))

    st.markdown('<div class="section-header">Proyectos por sector</div>', unsafe_allow_html=True)
    if not sector_counts.empty:
        st.altair_chart(bar_chart(sector_counts, "etiqueta", "intervenciones"), use_container_width=True, theme=None)
        st.dataframe(
            sector_counts[["etiqueta", "intervenciones"]].rename(
                columns={"etiqueta": "Sector", "intervenciones": "Proyectos"}),
            use_container_width=True, hide_index=True
        )
        st.caption(
            f"Un mismo proyecto puede abarcar varios sectores, asi que la suma de esta tabla "
            f"puede superar el total de {format_int(total_intervenciones)} proyectos."
        )
    else:
        st.info("No hay datos de sector registrados.")

    st.markdown('<div class="section-header">Poblacion atendida</div>', unsafe_allow_html=True)
    if not poblacion_counts.empty:
        st.altair_chart(bar_chart(poblacion_counts, "etiqueta", "intervenciones", color="#00A859"), use_container_width=True, theme=None)
        st.dataframe(
            poblacion_counts[["etiqueta", "intervenciones"]].rename(
                columns={"etiqueta": "Poblacion", "intervenciones": "Proyectos"}),
            use_container_width=True, hide_index=True
        )
        st.caption(
            f"Un mismo proyecto puede atender a varios tipos de poblacion, asi que la suma de esta "
            f"tabla puede superar el total de {format_int(total_intervenciones)} proyectos."
        )
    else:
        st.info("No hay datos de poblacion atendida registrados.")

    st.markdown('<div class="section-header">Objetivos de Desarrollo Sostenible (ODS)</div>', unsafe_allow_html=True)
    if not ods_counts.empty:
        st.altair_chart(bar_chart(ods_counts.sort_values("intervenciones", ascending=False).head(10), "etiqueta", "intervenciones"),
                         use_container_width=True, theme=None)
        st.dataframe(
            ods_counts[["etiqueta", "intervenciones"]].rename(
                columns={"etiqueta": "ODS", "intervenciones": "Proyectos"}),
            use_container_width=True, hide_index=True
        )
        st.caption(
            f"Un mismo proyecto puede apuntar a varios ODS, asi que la suma de esta tabla puede "
            f"superar el total de {format_int(total_intervenciones)} proyectos."
        )
    else:
        st.info("No hay datos de ODS registrados.")

    st.markdown('<div class="section-header">Donantes y organizaciones ejecutoras</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2)
    with col_c:
        st.markdown("**Proyectos por donante / cooperante**")
        if not donante_counts.empty:
            st.altair_chart(bar_chart(donante_counts, "etiqueta", "intervenciones"), use_container_width=True, theme=None)
        else:
            st.info("No hay datos de donante / cooperante registrados.")
    with col_d:
        st.markdown("**Proyectos por organizacion ejecutora**")
        if not org_counts.empty:
            st.altair_chart(bar_chart(org_counts, "etiqueta", "intervenciones", color="#00A859"), use_container_width=True, theme=None)
        else:
            st.info("No hay datos de organizacion ejecutora registrados.")

    if mun_sel_key is None:
        st.markdown('<div class="section-header">Municipios</div>', unsafe_allow_html=True)
        if not mun_counts.empty:
            st.altair_chart(bar_chart(mun_counts, "etiqueta", "intervenciones"), use_container_width=True, theme=None)
            st.caption(
                f"Un mismo proyecto puede intervenir en varios municipios, asi que la suma de estas "
                f"barras puede superar el total de {format_int(total_intervenciones)} proyectos. Usa el "
                f"filtro de municipio arriba para ver el total exacto de un municipio especifico."
            )
        else:
            st.info("No hay datos de municipio registrados.")

    st.markdown('<div class="section-header">Listado detallado de proyectos</div>', unsafe_allow_html=True)
    df_export = build_export_df(df_view)
    st.dataframe(df_export, use_container_width=True, hide_index=True)

    kpis_pdf = [
        ["Proyectos", format_int(total_intervenciones)],
        ["Organizaciones ejecutoras", format_int(total_organizaciones)],
        ["Donantes / cooperantes", format_int(total_donantes)],
        ["Municipios con proyectos", format_int(total_municipios)],
        ["Proyectos vigentes hoy", format_int(total_activos)],
    ]

    file_suffix = "guaviare"
    if mun_sel_key is not None:
        file_suffix += f"_{mun_sel_key}"
    if anio_sel_key is not None:
        file_suffix += f"_ini{anio_sel_key}"
    if anio_fin_sel_key is not None:
        file_suffix += f"_fin{anio_fin_sel_key}"
    if mes_fin_sel_key is not None:
        file_suffix += f"_mes{mes_fin_sel_key:02d}"

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            "\u2b07\ufe0f Descargar Excel",
            data=to_excel(df_view, sector_counts, poblacion_counts, ods_counts, mun_counts),
            file_name=f"ficha_cooperacion_{file_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_dl2:
        st.download_button(
            "\u2b07\ufe0f Descargar PDF",
            data=to_pdf(df_export, kpis_pdf),
            file_name=f"ficha_cooperacion_{file_suffix}.pdf",
            mime="application/pdf",
        )

    st.markdown(
        '<div class="apc-footer">Ficha de Cooperacion Internacional &mdash; Guaviare</div>',
        unsafe_allow_html=True
    )

# ============================================================================
# GUIA DE USUARIO
# ============================================================================
elif nav == nav_options[1]:
    st.markdown('<div class="dept-title-banner">Guia de usuario</div>', unsafe_allow_html=True)

    guia_html = (
        '<div class="guia-card">'
        '<div class="guia-intro">Que es esta herramienta?</div>'
        '<p>Esta ficha resume el mapeo de proyectos de cooperacion internacional que operan '
        'en el departamento del Guaviare, con datos del periodo 2025-2026. La informacion '
        'proviene de un formulario de campo diligenciado por las organizaciones ejecutoras.</p>'
        '<p><strong>De donde viene la informacion?</strong> Del archivo '
        '<code>Mapeo_CI_Proyectos_a_2025_-_2026.xlsx</code>, hoja <code>COOPERACION_INTERNACIONAL_M_0</code>. '
        'Cada fila es un proyecto reportado por una organizacion.</p>'
        '<p><strong>Como leer los indicadores?</strong> Las tarjetas superiores muestran totales '
        'unicos para el municipio y/o año seleccionado: numero de proyectos, organizaciones '
        'ejecutoras distintas, donantes o cooperantes distintos, municipios con al menos un '
        'proyecto, y los proyectos que se encuentran vigentes hoy segun sus fechas de inicio y '
        'finalizacion.</p>'
        '<p><strong>Filtros disponibles:</strong> se puede filtrar por municipio, por año de '
        'inicio, y por año y mes de finalizacion del proyecto. Un mismo proyecto puede abarcar varios municipios, varios sectores '
        'y varios ODS a la vez; por eso los conteos por municipio, sector, poblacion y ODS pueden '
        'sumar mas que el numero total de proyectos.</p>'
        '<p><strong>Limitaciones a tener en cuenta:</strong> esta version no incluye el valor '
        'estimado del aporte de cada proyecto ni el numero de participantes, porque el formulario '
        'actual no recoge esos datos de forma numerica comparable. El listado detallado al final '
        'del Panorama permite revisar cada proyecto tal como fue reportado.</p>'
        '</div>'
    )
    st.markdown(guia_html, unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.info("**\U0001F4CA Panorama Guaviare**\n\nIndicadores, sectores, ODS, donantes y listado completo. Filtra por municipio y año. Descarga en Excel y PDF.")
    with g2:
        st.info("**Fuente**\n\nMapeo de proyectos de cooperacion internacional en Guaviare. Periodo: 2025-2026.")

    st.markdown(
        '<div class="apc-footer">Ficha de Cooperacion Internacional &mdash; Guaviare</div>',
        unsafe_allow_html=True
    )
