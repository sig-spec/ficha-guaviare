# -*- coding: utf-8 -*-
# ============================================================================
# Ficha de Cooperacion Internacional - GUAVIARE
# Version simplificada, adaptada a partir de la Ficha nacional de APC-Colombia,
# para trabajar con el archivo de mapeo de actores de cooperacion internacional
# en el departamento del Guaviare (encuesta / formulario de campo).
# ============================================================================
import streamlit as st
import pandas as pd
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
FILE = "Primer_semestre_2026_mapeo.xlsx"
SHEET = "Mapeo_de_actoresV3_0"
LOGO_1 = "logo_gobernacion.png"   # opcional, si no existe se omite
LOGO_2 = "logo_sncic.png"         # opcional, si no existe se omite

# ============================================================================
# ESTILOS (identidad visual institucional, adaptada del manual original)
# ============================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&family=Source+Sans+3:wght@400;600&display=swap');

:root {
    --apc-blue: #003087;
    --apc-red: #C8102E;
    --apc-yellow: #F5A623;
    --apc-light: #EEF3FB;
    --apc-gray: #F7F8FA;
    --apc-border: #D1D9E6;
    --apc-text: #1C2B4A;
    --apc-muted: #5A6A85;
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

.apc-header {
    background: var(--apc-blue);
    padding: 1.2rem 2.2rem 1rem 2.2rem;
    border-bottom: 4px solid var(--apc-red);
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
    background: linear-gradient(90deg, var(--apc-yellow) 33.3%, var(--apc-blue) 33.3% 66.6%, var(--apc-red) 66.6%);
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
    border-left: 6px solid var(--apc-red);
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

div[data-testid="stDownloadButton"] button {
    background: var(--apc-blue) !important;
    color: white !important;
    border: none !important;
    border-radius: 4px !important;
    font-family: 'Montserrat', sans-serif !important;
    font-weight: 600 !important;
    text-transform: uppercase;
}
div[data-testid="stDownloadButton"] button:hover { background: var(--apc-red) !important; }

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
    border-left: 5px solid var(--apc-red);
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

SECTOR_MAP = {
    "otra": "Otro",
    "caracter_publico": "Caracter publico",
    "caracter_privado": "Caracter privado",
}

FASE_MAP = {
    "ejecucion": "En ejecucion",
    "finalizacion": "Finalizacion",
    "diseno": "En diseno",
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


def municipio_label(x):
    if not x:
        return "Sin dato"
    return MUNICIPIOS_MAP.get(x, x.replace("_", " ").title())


def sector_label(x):
    if not x:
        return "Sin dato"
    return SECTOR_MAP.get(x, x.replace("_", " ").title())


def fase_label(x):
    if not x:
        return "Sin dato"
    return FASE_MAP.get(x, x.replace("_", " ").capitalize())


def format_usd(n):
    try:
        n = float(n)
    except Exception:
        return "USD 0"
    return "USD " + f"{n:,.0f}".replace(",", ".")


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
    "1.FECHA REGISTRO DE LA INTERVENCIÓN": "fecha_registro",
    "3.Nombre de la entidad /organización": "organizacion_ejecutora",
    "6.Nombre Intervención": "nombre_intervencion",
    "7.Objetivo General": "objetivo",
    "8.Fecha Inicial": "fecha_inicial",
    "9.Fecha Final": "fecha_final",
    "10.Sector al que pertenece": "sector",
    "¿Cuál?": "sector_otro",
    "11. Estado Intervención": "estado",
    "13.Municipio": "municipio",
    "17. Valor Aporte (USD)": "valor_usd",
    "18.Mencione el número de participantes": "participantes",
    "19.A qué población atiende el proyecto": "poblacion",
    "22.¿En qué fase se encuentra el proyecto?": "fase",
    "23.¿Con que entidades/organizaciones ha articulado para la ejecucion del proyecto?": "entidades_articuladas",
    "30.Origen del actor": "origen_actor",
    "31.País actor": "actor_financiador",
    "32.ODS": "ods",
}


@st.cache_data
def load_data():
    raw = pd.read_excel(FILE, sheet_name=SHEET)
    df = raw.rename(columns=COLMAP)
    for needed in COLMAP.values():
        if needed not in df.columns:
            df[needed] = ""
    # Nota: se limpia por columna (no por dtype) porque pandas >= 2.x puede
    # asignar dtype "str" en vez de "object" a columnas de texto, y un chequeo
    # por dtype == "object" las deja pasar sin limpiar.
    text_cols = [c for c in COLMAP.values() if c not in ("valor_usd", "participantes")]
    for c in text_cols:
        df[c] = df[c].where(df[c].notna(), "")
        df[c] = df[c].astype(str).str.strip()
        df[c] = df[c].replace({"nan": "", "None": "", "NA": "", "na": "", "N/A": ""})
    df["valor_usd"] = pd.to_numeric(df["valor_usd"], errors="coerce").fillna(0)
    df["participantes"] = pd.to_numeric(df["participantes"], errors="coerce").fillna(0)
    return df


def counts_table(series, label_fn, value_name="intervenciones"):
    s = series.replace("", pd.NA).dropna()
    if s.empty:
        return pd.DataFrame(columns=["clave", "etiqueta", value_name])
    out = s.value_counts().reset_index()
    out.columns = ["clave", value_name]
    out["etiqueta"] = out["clave"].map(label_fn)
    return out


def exploded_ods_counts(series):
    s = series.replace("", pd.NA).dropna().astype(str)
    if s.empty:
        return pd.DataFrame(columns=["clave", "etiqueta", "intervenciones"])
    exp = s.str.split(",").explode().str.strip()
    exp = exp[exp != ""]
    if exp.empty:
        return pd.DataFrame(columns=["clave", "etiqueta", "intervenciones"])
    out = exp.value_counts().reset_index()
    out.columns = ["clave", "intervenciones"]
    out["etiqueta"] = out["clave"].map(lambda x: ODS_NOMBRES.get(x, f"ODS {x}"))
    return out


def bar_chart(df, y_field, x_field, color="#003087"):
    return (
        alt.Chart(df)
        .mark_bar(color=color, cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
        .encode(
            y=alt.Y(f"{y_field}:N", sort="-x", title=""),
            x=alt.X(f"{x_field}:Q", title="Intervenciones"),
            tooltip=[f"{y_field}:N", f"{x_field}:Q"],
        )
        .properties(height=max(180, 32 * len(df)))
    )


# ============================================================================
# EXPORTACION: EXCEL Y PDF
# ============================================================================
EXPORT_LABELS = {
    "nombre_intervencion": "Nombre intervencion",
    "organizacion_ejecutora": "Organizacion ejecutora",
    "sector_label": "Sector",
    "municipio_label": "Municipio",
    "fase_label": "Fase",
    "actor_financiador": "Actor / cooperante",
    "origen_actor": "Origen del actor",
    "valor_usd": "Valor aporte (USD)",
    "participantes": "Participantes",
    "ods": "ODS",
    "estado": "Estado intervencion",
    "objetivo": "Objetivo general",
}


def build_export_df(df):
    d = df.copy()
    d["sector_label"] = d["sector"].map(sector_label)
    d["municipio_label"] = d["municipio"].map(municipio_label)
    d["fase_label"] = d["fase"].map(fase_label)
    cols = [c for c in EXPORT_LABELS if c in d.columns]
    d = d[cols].rename(columns=EXPORT_LABELS)
    return d


def to_excel(df, sector_counts, fase_counts, ods_counts, mun_counts):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        build_export_df(df).to_excel(writer, sheet_name="Intervenciones", index=False)
        sector_counts.rename(columns={"etiqueta": "Sector", "intervenciones": "Intervenciones"})[
            ["Sector", "Intervenciones"]
        ].to_excel(writer, sheet_name="Sectores", index=False)
        fase_counts.rename(columns={"etiqueta": "Fase", "intervenciones": "Intervenciones"})[
            ["Fase", "Intervenciones"]
        ].to_excel(writer, sheet_name="Fase", index=False)
        ods_counts.rename(columns={"etiqueta": "ODS", "intervenciones": "Intervenciones"})[
            ["ODS", "Intervenciones"]
        ].to_excel(writer, sheet_name="ODS", index=False)
        mun_counts.rename(columns={"etiqueta": "Municipio", "intervenciones": "Intervenciones"})[
            ["Municipio", "Intervenciones"]
        ].to_excel(writer, sheet_name="Municipios", index=False)
    output.seek(0)
    return output.getvalue()


def to_pdf(df, kpis):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm, leftMargin=1.5 * cm, rightMargin=1.5 * cm
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleAPC", parent=styles["Heading1"],
                                  textColor=colors.HexColor("#003087"), fontSize=15, spaceAfter=4)
    sub_style = ParagraphStyle("SubAPC", parent=styles["Normal"],
                                textColor=colors.HexColor("#5A6A85"), fontSize=9, spaceAfter=10)
    section_style = ParagraphStyle("SectionAPC", parent=styles["Heading2"],
                                    textColor=colors.HexColor("#003087"), fontSize=11,
                                    spaceBefore=14, spaceAfter=6)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontSize=7.5, leading=9)

    elements = [
        Paragraph("Ficha de Cooperacion Internacional - Guaviare", title_style),
        Paragraph("Mapeo de actores de cooperacion internacional. Corte: primer semestre 2026.", sub_style),
        HRFlowable(width="100%", color=colors.HexColor("#F5A623"), thickness=2, spaceAfter=10),
    ]

    kpi_data = [["Indicador", "Valor"]] + kpis
    kpi_table = Table(kpi_data, colWidths=[9 * cm, 6 * cm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003087")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D9E6")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F8FA")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elements += [kpi_table, Spacer(1, 0.6 * cm), Paragraph("Listado de intervenciones", section_style)]

    table_data = [["Intervencion", "Sector", "Municipio", "Organizacion", "Aporte (USD)"]]
    for _, row in df.iterrows():
        table_data.append([
            Paragraph(str(row.get("nombre_intervencion", ""))[:70], cell_style),
            sector_label(row.get("sector", "")),
            municipio_label(row.get("municipio", "")),
            Paragraph(str(row.get("organizacion_ejecutora", ""))[:40], cell_style),
            format_usd(row.get("valor_usd", 0)),
        ])
    detail_table = Table(table_data, colWidths=[5.5 * cm, 2.6 * cm, 2.8 * cm, 3.7 * cm, 2.4 * cm], repeatRows=1)
    detail_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003087")),
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
st.markdown(
    '<div class="apc-header">'
    '<p class="apc-header-title">Ficha de Cooperacion Internacional</p>'
    '<p class="apc-header-subtitle">Departamento del Guaviare &mdash; Mapeo de actores, primer semestre 2026</p>'
    '</div>',
    unsafe_allow_html=True
)
st.markdown('<div class="apc-flag-bar"></div>', unsafe_allow_html=True)

nav_options = ["\U0001F4CA Panorama Guaviare", "\U0001F4D6 Guia de usuario"]
nav = st.radio("", nav_options, horizontal=True, label_visibility="collapsed", key="main_nav")
st.markdown("---")

# ============================================================================
# PANORAMA GUAVIARE
# ============================================================================
if nav == nav_options[0]:
    st.markdown('<div class="dept-title-banner">GUAVIARE &mdash; Cooperacion Internacional</div>', unsafe_allow_html=True)

    # -- Filtro por municipio --
    mun_keys_presentes = sorted(df["municipio"].replace("", pd.NA).dropna().unique().tolist())
    mun_label_to_key = {"Todos los municipios": None}
    mun_label_to_key.update({municipio_label(k): k for k in mun_keys_presentes})
    mun_opciones = list(mun_label_to_key.keys())
    col_filtro, _ = st.columns([1, 3])
    with col_filtro:
        mun_sel_label = st.selectbox("Filtrar por municipio", mun_opciones, key="filtro_municipio")
    mun_sel_key = mun_label_to_key[mun_sel_label]

    if mun_sel_key is not None:
        df_view = df[df["municipio"] == mun_sel_key].reset_index(drop=True)
    else:
        df_view = df

    total_intervenciones = len(df_view)
    total_organizaciones = df_view["organizacion_ejecutora"].replace("", pd.NA).dropna().nunique()
    total_municipios = df_view["municipio"].replace("", pd.NA).dropna().nunique()
    total_usd = df_view["valor_usd"].sum()
    total_participantes = df_view["participantes"].sum()
    total_cooperantes = df_view["actor_financiador"].replace("", pd.NA).dropna().nunique()

    sector_counts = counts_table(df_view["sector"], sector_label)
    fase_counts = counts_table(df_view["fase"], fase_label)
    mun_counts = counts_table(df_view["municipio"], municipio_label)
    ods_counts = exploded_ods_counts(df_view["ods"])
    actor_counts = counts_table(df_view["actor_financiador"], lambda x: x)
    org_counts = counts_table(df_view["organizacion_ejecutora"], lambda x: x)

    st.caption(
        f"Fuente: mapeo de actores de cooperacion internacional en Guaviare. "
        f"Corte: primer semestre 2026. {total_intervenciones} intervenciones registradas"
        + (f" en {mun_sel_label}." if mun_sel_key is not None else " en todo el departamento.")
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Intervenciones", format_int(total_intervenciones))
    c2.metric("Organizaciones ejecutoras", format_int(total_organizaciones))
    c3.metric("Actores / cooperantes", format_int(total_cooperantes))
    c4.metric("Municipios", format_int(total_municipios))
    c5.metric("Total aporte estimado", format_usd(total_usd))
    c6.metric("Participantes reportados", format_int(total_participantes))

    st.markdown('<div class="section-header">Sectores y fase de los proyectos</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Intervenciones por tipo de sector**")
        if not sector_counts.empty:
            st.altair_chart(bar_chart(sector_counts, "etiqueta", "intervenciones"), use_container_width=True)
            st.dataframe(
                sector_counts[["etiqueta", "intervenciones"]].rename(
                    columns={"etiqueta": "Sector", "intervenciones": "Intervenciones"}),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("No hay datos de sector registrados.")
    with col_b:
        st.markdown("**Fase de los proyectos**")
        if not fase_counts.empty:
            st.altair_chart(bar_chart(fase_counts, "etiqueta", "intervenciones", color="#1565C0"), use_container_width=True)
            st.dataframe(
                fase_counts[["etiqueta", "intervenciones"]].rename(
                    columns={"etiqueta": "Fase", "intervenciones": "Intervenciones"}),
                use_container_width=True, hide_index=True
            )
        else:
            st.info("No hay datos de fase registrados.")

    st.markdown('<div class="section-header">Objetivos de Desarrollo Sostenible (ODS)</div>', unsafe_allow_html=True)
    if not ods_counts.empty:
        st.altair_chart(bar_chart(ods_counts.sort_values("intervenciones", ascending=False).head(10), "etiqueta", "intervenciones"),
                         use_container_width=True)
        st.dataframe(
            ods_counts[["etiqueta", "intervenciones"]].rename(
                columns={"etiqueta": "ODS", "intervenciones": "Intervenciones"}),
            use_container_width=True, hide_index=True
        )
    else:
        st.info("No hay datos de ODS registrados.")

    st.markdown('<div class="section-header">Cooperantes y organizaciones ejecutoras</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2)
    with col_c:
        st.markdown("**Intervenciones por actor / cooperante**")
        if not actor_counts.empty:
            st.altair_chart(bar_chart(actor_counts, "etiqueta", "intervenciones"), use_container_width=True)
        else:
            st.info("No hay datos de actor / cooperante registrados.")
    with col_d:
        st.markdown("**Intervenciones por organizacion ejecutora**")
        if not org_counts.empty:
            st.altair_chart(bar_chart(org_counts, "etiqueta", "intervenciones", color="#1565C0"), use_container_width=True)
        else:
            st.info("No hay datos de organizacion ejecutora registrados.")

    if mun_sel_key is None:
        st.markdown('<div class="section-header">Municipios</div>', unsafe_allow_html=True)
        if not mun_counts.empty:
            st.altair_chart(bar_chart(mun_counts, "etiqueta", "intervenciones"), use_container_width=True)
        else:
            st.info("No hay datos de municipio registrados.")

    st.markdown('<div class="section-header">Listado detallado de intervenciones</div>', unsafe_allow_html=True)
    st.dataframe(build_export_df(df_view), use_container_width=True, hide_index=True)

    kpis_pdf = [
        ["Intervenciones", format_int(total_intervenciones)],
        ["Organizaciones ejecutoras", format_int(total_organizaciones)],
        ["Actores / cooperantes", format_int(total_cooperantes)],
        ["Municipios con intervencion", format_int(total_municipios)],
        ["Total aporte estimado", format_usd(total_usd)],
        ["Participantes reportados", format_int(total_participantes)],
    ]

    file_suffix = "guaviare" if mun_sel_key is None else mun_sel_key

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            "\u2b07\ufe0f Descargar Excel",
            data=to_excel(df_view, sector_counts, fase_counts, ods_counts, mun_counts),
            file_name=f"ficha_cooperacion_{file_suffix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_dl2:
        st.download_button(
            "\u2b07\ufe0f Descargar PDF",
            data=to_pdf(df_view, kpis_pdf),
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
        '<p>Esta ficha resume el mapeo de actores de cooperacion internacional que operan '
        'en el departamento del Guaviare, con corte al primer semestre de 2026. La informacion '
        'proviene de un formulario de campo diligenciado por las organizaciones ejecutoras '
        '(PNUD, WWF Colombia, Swisscontact, Rainforest Alliance, Caritas Guaviare, entre otras).</p>'
        '<p><strong>De donde viene la informacion?</strong> Del archivo '
        '<code>Primer_semestre_2026_mapeo.xlsx</code>, hoja <code>Mapeo_de_actoresV3_0</code>. '
        'Cada fila es una intervencion reportada por una organizacion.</p>'
        '<p><strong>Como leer los indicadores?</strong> Las tarjetas superiores muestran totales '
        'unicos: numero de intervenciones, organizaciones ejecutoras distintas, actores o '
        'cooperantes distintos, municipios con al menos una intervencion, el aporte total '
        'estimado en USD y los participantes reportados sumados.</p>'
        '<p><strong>Limitaciones a tener en cuenta:</strong> esta version no incluye comparativo '
        'con un periodo anterior (solo hay un corte disponible), y algunos campos del formulario '
        '(sector, fase, actor financiador) tienen categorias con poca estandarizacion porque '
        'provienen de texto libre o listas desplegadas de forma inconsistente. El listado '
        'detallado al final del Panorama permite revisar cada intervencion tal como fue '
        'reportada.</p>'
        '</div>'
    )
    st.markdown(guia_html, unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        st.info("**\U0001F4CA Panorama Guaviare**\n\nIndicadores, sectores, ODS, cooperantes y listado completo. Descarga en Excel y PDF.")
    with g2:
        st.info("**Fuente**\n\nMapeo de actores de cooperacion internacional en Guaviare. Corte: primer semestre 2026.")

    st.markdown(
        '<div class="apc-footer">Ficha de Cooperacion Internacional &mdash; Guaviare</div>',
        unsafe_allow_html=True
    )
