"""
Tablero Predial Girardota 2026 — Streamlit
==========================================
Lee Predial_Girardota_2026_Ley44.xlsx (generado por procesar_predial_girardota.py)
con los NUEVOS ORÍGENES:

  2025 · INFORME CATASTRO_DELTA_13082026.xlsx        (base de datos del municipio)
  2026 · REPORTE_REGISTRO_BASICO GIRARDOTA_JUNIO     (gestor catastral)
         REPORTE_REGISTRO_COMPLEMENTARIO_GIRARDOTA_JUNIO
         REPORTE_FICHAS_NUEVAS

Segmentación de predios nuevos por matrícula inmobiliaria:
  · CON matrícula  → excluidos del tope por Ley 44 Art.6, tarifa plena legítima
  · SIN matrícula  → apartados, NO se cuentan como ilegalidades

Ejecutar: streamlit run tablero_girardota.py
"""

import io
import os

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Predial Girardota 2026",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_EXCEL = os.path.join(_DIR, "Predial_Girardota_2026_Ley44.xlsx")
RUTA_WORD  = os.path.join(_DIR, "Informe_Predial_Girardota_2026.docx")

AZUL_OSC  = "#1F4E79"
AZUL_MED  = "#2E75B6"
AZUL_CLAR = "#BDD7EE"
VERDE     = "#70AD47"
AMBAR     = "#FFD966"
ROJO      = "#FF7C80"
NARANJA   = "#F4B183"
GRIS      = "#D9D9D9"
MORADO    = "#B4A7D6"

CLS_EXISTENTE = "EXISTENTE"
CLS_NUEVO_MAT = "NUEVO CON MATRÍCULA"
CLS_NUEVO_SIN = "NUEVO SIN MATRÍCULA"
CLS_NO_REPORT = "NO REPORTADO 2026"

ORDEN_RNG = ["0–5 M", "5–15 M", "15–50 M", "50–150 M", "150–500 M", ">500 M"]

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .stApp { background-color: #F0F4F8; }
  [data-testid="stSidebar"] {
      background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
  }
  [data-testid="stSidebar"] * { color: white !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stSlider label { color: #BDD7EE !important; font-weight: 600; }

  .kpi-card {
      background: white; border-radius: 12px; padding: 18px 14px;
      text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.10);
      border-top: 4px solid #2E75B6; min-height: 120px;
      display: flex; flex-direction: column; justify-content: center;
  }
  .kpi-card.verde  { border-top-color: #70AD47; }
  .kpi-card.ambar  { border-top-color: #FFD966; }
  .kpi-card.rojo   { border-top-color: #FF7C80; }
  .kpi-card.oscuro { border-top-color: #1F4E79; }
  .kpi-card.naranja{ border-top-color: #F4B183; }
  .kpi-card.morado { border-top-color: #B4A7D6; }
  .kpi-valor { font-size: 1.6rem; font-weight: 800; color: #1F4E79; line-height:1.1; }
  .kpi-label { font-size: 0.72rem; color: #666; margin-top:5px; font-weight:500;
               text-transform:uppercase; letter-spacing:0.5px; }
  .kpi-sub   { font-size: 0.76rem; color: #2E75B6; font-weight:600; margin-top:3px; }

  .sec-tit {
      background: linear-gradient(90deg, #1F4E79, #2E75B6);
      color: white; padding: 7px 16px; border-radius: 8px;
      font-size: 0.92rem; font-weight: 700; letter-spacing: 0.4px;
      margin: 20px 0 10px 0;
  }
  .header-main {
      background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 55%, #3A9AD9 100%);
      border-radius: 14px; padding: 24px 30px; color: white;
      margin-bottom: 20px; box-shadow: 0 4px 15px rgba(31,78,121,0.3);
  }
  .header-main h1 { margin:0; font-size:1.7rem; font-weight:800; }
  .header-main p  { margin:4px 0 0; opacity:0.88; font-size:0.87rem; }
  .nota-legal {
      background:#FFF9E6; border-left:4px solid #FFD966;
      border-radius:6px; padding:10px 14px; font-size:0.80rem; color:#555; margin-top:8px;
  }
  .nota-info {
      background:#EAF2FB; border-left:4px solid #2E75B6;
      border-radius:6px; padding:10px 14px; font-size:0.82rem; color:#333; margin-top:8px;
  }
  .nota-alerta {
      background:#FFF0F0; border-left:4px solid #FF7C80;
      border-radius:6px; padding:10px 14px; font-size:0.82rem; color:#333; margin-top:8px;
  }
</style>
""", unsafe_allow_html=True)


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
NUM_COLS = [
    "N_PROPIETARIOS", "AREA_TERR_2025_M2", "AREA_TERR_2026_M2",
    "AREA_CONS_2025_M2", "AREA_CONS_2026_M2", "N_UNIDADES_2026",
    "AVALUO_2025", "AVALUO_2026", "VAR_AVALUO_%", "TARIFA_MIL",
    "IMPTO_2025", "IMPTO_2026_PLENO", "LIMITE_LEY44", "IMPTO_CORRECTO_2026",
    "EXCESO_LEY44", "VAR_IMPTO_%", "LIMITE_LOCAL", "IMPTO_CON_LOCAL", "EXCESO_LOCAL",
]

NUM_COLS_SEG = [
    "N_PROPIETARIOS", "AREA_TERR_M2", "AREA_CONS_M2", "N_UNIDADES",
    "AVALUO_2025 ($)", "AVALUO_2026 ($)", "VAR_AVALUO_%",
    "TARIFA_MIL", "IMPTO_2025 ($)", "IMPTO_2026_PLENO ($)",
]


@st.cache_data(show_spinner="Cargando análisis predial de Girardota 2026 …")
def cargar():
    df = pd.read_excel(RUTA_EXCEL, sheet_name="Todos_Predios", header=1)
    df = df.drop_duplicates(subset=["FICHA"], keep="first").reset_index(drop=True)

    for c in NUM_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    def yn(col):
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        return df[col].astype(str).str.strip().str.upper() == "SÍ"

    df["_aplica"]       = yn("APLICA_LIMITE_LEY44")
    df["_excede"]       = yn("EXCEDE_LEY44")
    df["_aplica_local"] = yn("APLICA_LIMITE_LOCAL")
    df["_tiene_hist"]   = df["IMPTO_2025"].fillna(0) > 0
    df["_nuevo"]        = df["CLASIFICACION"].astype(str).str.strip() == CLS_NUEVO_MAT

    df["DEST_NOM_2026"] = df["DEST_NOM_2026"].replace({"—": np.nan})
    df["DESTINO"] = df["DEST_NOM_2026"].fillna(df["DEST_NOM_2025"]).fillna("Sin destino")
    df["RANGO_AVALUO_2026"] = df["RANGO_AVALUO_2026"].fillna("Sin dato")
    df["MATRICULA"] = df["MATRICULA"].fillna("").astype(str)
    df["_con_matricula"] = df["MATRICULA"].str.strip() != ""

    return df


@st.cache_data(show_spinner=False)
def cargar_segmento(hoja):
    try:
        d = pd.read_excel(RUTA_EXCEL, sheet_name=hoja, header=1)
    except Exception:
        return pd.DataFrame()
    # La hoja lleva una nota al pie que cae en la columna FICHA: se descarta por longitud
    d = d[d["FICHA"].notna()].copy()
    d = d[d["FICHA"].astype(str).str.strip().str.len() <= 20]
    for c in NUM_COLS_SEG:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce")
    if "MATRICULA" in d.columns:
        d["MATRICULA"] = d["MATRICULA"].fillna("").astype(str)
    return d.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def cargar_simple(hoja, header=0):
    try:
        return pd.read_excel(RUTA_EXCEL, sheet_name=hoja, header=header)
    except Exception:
        return pd.DataFrame()


if not os.path.exists(RUTA_EXCEL):
    st.error(
        f"No se encontró **{RUTA_EXCEL}**. "
        "Ejecute primero `procesar_predial_girardota.py` para generar el archivo."
    )
    st.stop()

df_all       = cargar()
df_nuevos    = cargar_segmento("Predios_Nuevos")
df_sin_mat   = cargar_segmento("Nuevos_Sin_Matricula")
df_no_report = cargar_segmento("No_Reportados_2026")
df_incons    = cargar_simple("Inconsistencias")
df_origenes  = cargar_simple("Origenes_Datos")


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_cop(v):
    if pd.isna(v):
        return "$0"
    a = abs(v)
    if a >= 1e12: return f"${v/1e12:,.2f} B"
    if a >= 1e9:  return f"${v/1e9:,.1f} MM"
    if a >= 1e6:  return f"${v/1e6:,.0f} M"
    return f"${v:,.0f}"


def kpi(col, val, lab, sub="", color=""):
    with col:
        st.markdown(f"""
        <div class="kpi-card {color}">
          <div class="kpi-valor">{val}</div>
          <div class="kpi-label">{lab}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)


def suma(d, col):
    return float(d[col].fillna(0).sum()) if col in d.columns else 0.0


def descargar(d, nombre, etiqueta, key=None):
    if d.empty:
        return
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        d.to_excel(w, index=False, sheet_name="Datos")
    st.download_button(etiqueta, data=buf.getvalue(), file_name=nombre, key=key,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏛️ Predial Girardota 2026")
    st.markdown("**Base municipio 2025 vs Gestor catastral 2026**")
    st.divider()
    st.markdown("#### Filtros globales")

    sel_zona = st.selectbox("Zona:", ["Todos", "Solo Urbano", "Solo Rural"])

    sel_clas = st.selectbox(
        "Clasificación del predio:",
        ["Todos", "Solo existentes", "Solo nuevos con matrícula"])

    sel_lim = st.selectbox(
        "Límite Ley 44:",
        ["Todos", "Excede el límite", "Dentro del límite", "No aplica límite"])

    destinos_disp = sorted(df_all["DESTINO"].dropna().unique())
    sel_dest = st.multiselect("Destino económico:", destinos_disp, placeholder="Todos")

    sel_rng = st.multiselect("Rango de avalúo 2026:", ORDEN_RNG, placeholder="Todos")

    sel_tar = st.multiselect(
        "Origen de la tarifa:",
        sorted(df_all["ORIGEN_TARIFA"].dropna().astype(str).unique()),
        placeholder="Todos")

    st.divider()
    umbral_av = st.slider("Umbral 'cambio extremo' de avalúo (%):", 20, 500, 100, 10)

    st.divider()
    st.markdown("""
    <div style='font-size:0.72rem; opacity:0.85; line-height:1.5;'>
    📂 <b>Orígenes</b><br>
    · 2025 — INFORME CATASTRO_DELTA<br>
    · 2026 — Registro Básico y Complementario<br>
    · Predios nuevos — REPORTE_FICHAS_NUEVAS<br>
    · Tarifas — base de facturación<br><br>
    📋 <b>Marco legal</b><br>
    · Ley 44/1990 Art. 6 — doble del impuesto 2025<br>
    · Acuerdo 49 — límites 25 % y 50 %<br>
    · UVT 2026: $52.374
    </div>""", unsafe_allow_html=True)


# ── FILTRAR ───────────────────────────────────────────────────────────────────
df = df_all.copy()

if sel_zona == "Solo Urbano":
    df = df[df["ZONA"] == "URBANO"]
elif sel_zona == "Solo Rural":
    df = df[df["ZONA"] == "RURAL"]

if sel_clas == "Solo existentes":
    df = df[~df["_nuevo"]]
elif sel_clas == "Solo nuevos con matrícula":
    df = df[df["_nuevo"]]

if sel_lim == "Excede el límite":
    df = df[df["_excede"]]
elif sel_lim == "Dentro del límite":
    df = df[df["_aplica"] & ~df["_excede"] & df["_tiene_hist"]]
elif sel_lim == "No aplica límite":
    df = df[~df["_aplica"]]

if sel_dest:
    df = df[df["DESTINO"].isin(sel_dest)]
if sel_rng:
    df = df[df["RANGO_AVALUO_2026"].isin(sel_rng)]
if sel_tar:
    df = df[df["ORIGEN_TARIFA"].astype(str).isin(sel_tar)]

# Universo del análisis Ley 44 = predios existentes dentro del filtro
df_ex = df[~df["_nuevo"]]


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-main">
  <h1>🏛️ Análisis Predial — Municipio de Girardota · Vigencia 2026</h1>
  <p>Base de datos del municipio (INFORME CATASTRO_DELTA · 2025) contra el reporte del gestor
  catastral (Registro Básico y Complementario · 2026) · Ley 44/1990 Art. 6 — el impuesto no
  puede superar el doble del año anterior · Acuerdo 49 · UVT 2026: $52.374</p>
</div>""", unsafe_allow_html=True)

n_f = df["FICHA"].nunique()
n_t = df_all["FICHA"].nunique()
n_u = df[df["ZONA"] == "URBANO"]["FICHA"].nunique()
n_r = df[df["ZONA"] == "RURAL"]["FICHA"].nunique()
if n_f < n_t:
    st.info(f"Mostrando **{n_f:,}** de **{n_t:,}** predios según filtros — "
            f"**{n_u:,} urbanos · {n_r:,} rurales**")
else:
    st.info(f"**{n_t:,}** predios en el análisis principal — "
            f"**{n_u:,} urbanos · {n_r:,} rurales** · "
            f"además {len(df_sin_mat):,} nuevos sin matrícula y "
            f"{len(df_no_report):,} no reportados en hojas aparte")


# ── KPIs PRINCIPALES ──────────────────────────────────────────────────────────
st.markdown('<div class="sec-tit">📊 Indicadores clave — predios existentes</div>',
            unsafe_allow_html=True)

av25   = suma(df_ex, "AVALUO_2025")
av26   = suma(df_ex, "AVALUO_2026")
i25    = suma(df_ex, "IMPTO_2025")
i26    = suma(df_ex, "IMPTO_2026_PLENO")
ic26   = suma(df_ex, "IMPTO_CORRECTO_2026")
exceso = suma(df_ex, "EXCESO_LEY44")
n_exc  = int(df_ex["_excede"].sum())
var_av = (av26 - av25) / av25 * 100 if av25 > 0 else 0
var_i  = (i26 - i25) / i25 * 100 if i25 > 0 else 0

k = st.columns(7)
kpi(k[0], f"{len(df_ex):,}",   "Predios existentes",     "En ambas bases",            "oscuro")
kpi(k[1], fmt_cop(av25),       "Avalúo 2025",            "Base del municipio",        "")
kpi(k[2], fmt_cop(av26),       "Avalúo 2026",            f"Variación {var_av:+.1f} %", "ambar")
kpi(k[3], fmt_cop(i25),        "Impuesto 2025",          "Facturado",                 "")
kpi(k[4], fmt_cop(i26),        "Impuesto 2026 pleno",    f"Variación {var_i:+.1f} %",
    "rojo" if var_i > 100 else "ambar")
kpi(k[5], fmt_cop(ic26),       "Impuesto con Ley 44",    "Tope legal aplicado",       "verde")
kpi(k[6], fmt_cop(exceso),     "Exceso sobre el tope",   f"{n_exc:,} predios",        "naranja")
st.markdown("<br>", unsafe_allow_html=True)

# ── KPIs SEGMENTACIÓN ─────────────────────────────────────────────────────────
st.markdown('<div class="sec-tit">🧭 Segmentación del universo de predios</div>',
            unsafe_allow_html=True)

n_nue = int(df["_nuevo"].sum())
av_nue = suma(df[df["_nuevo"]], "AVALUO_2026")
av_sin = suma(df_sin_mat, "AVALUO_2026 ($)")
av_nor = suma(df_no_report, "AVALUO_2025 ($)")
n_incons = len(df_incons)

s = st.columns(6)
kpi(s[0], f"{len(df_ex):,}",        "Existentes",              "Universo Ley 44",           "oscuro")
kpi(s[1], f"{n_nue:,}",             "Nuevos con matrícula",    fmt_cop(av_nue),             "verde")
kpi(s[2], f"{len(df_sin_mat):,}",   "Nuevos sin matrícula",    "Apartados · no ilegalidad", "morado")
kpi(s[3], f"{len(df_no_report):,}", "No reportados 2026",      fmt_cop(av_nor),             "naranja")
kpi(s[4], f"{n_incons:,}",          "Inconsistencias",         "Matrícula / zona / tarifa", "rojo")
cob = (df_all["ORIGEN_TARIFA"].astype(str) == "FACTURACIÓN").mean() * 100
kpi(s[5], f"{cob:.1f} %",           "Tarifa desde facturación", "Cobertura del cruce",      "ambar")
st.markdown("<br>", unsafe_allow_html=True)

st.markdown(
    "<div class='nota-info'>🔎 <b>Cómo leer la segmentación.</b> Los <b>predios nuevos con "
    "matrícula</b> se incorporan por primera vez al catastro y tienen folio de registro: el "
    "artículo 6 de la Ley 44 de 1990 los excluye del tope, de modo que van a tarifa plena de "
    "forma legítima y no suman al exceso. Los <b>predios nuevos sin matrícula</b> se apartan "
    "del análisis porque sin folio no hay título registrado que permita calificar la "
    "incorporación como irregular: <b>no se cuentan como ilegalidades</b>. Los <b>no "
    "reportados</b> figuran en la base del municipio y el gestor no los incluyó en 2026.</div>",
    unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# PESTAÑAS
# ═════════════════════════════════════════════════════════════════════════════
tabs = st.tabs([
    "📈 Panorama fiscal",
    "⚖️ Límite Ley 44",
    "🏛️ Acuerdo 49",
    "🆕 Predios nuevos",
    "🚫 Sin matrícula / No reportados",
    "🚨 Cambios extremos",
    "⚠️ Inconsistencias",
    "📋 Detalle",
])


# ── TAB 1: PANORAMA FISCAL ────────────────────────────────────────────────────
with tabs[0]:
    c1, c2, c3 = st.columns([2.2, 1.8, 1.8])

    with c1:
        fig = go.Figure(go.Bar(
            x=["Impuesto 2025", "2026 a tarifa plena", "2026 con Ley 44"],
            y=[i25, i26, ic26],
            marker_color=[AZUL_MED, ROJO if i26 > ic26 else AMBAR, VERDE],
            text=[fmt_cop(v) for v in [i25, i26, ic26]],
            textposition="outside", textfont=dict(size=11, color="black"),
        ))
        fig.update_layout(
            title=dict(text="Comparativo del impuesto ($)", font=dict(size=13, color=AZUL_OSC)),
            xaxis=dict(tickfont=dict(color="black")),
            yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                       tickfont=dict(color="black")),
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=50, b=20, l=10, r=20), height=330, showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        n_ok    = int((df_ex["_aplica"] & ~df_ex["_excede"] & df_ex["_tiene_hist"]).sum())
        n_sinb  = int((df_ex["_aplica"] & ~df_ex["_tiene_hist"]).sum())
        n_noapl = int((~df_ex["_aplica"]).sum())
        fig = go.Figure(go.Pie(
            labels=["Excede el límite", "Dentro del límite",
                    "Sin base 2025", "No aplica límite", "Nuevos con matrícula"],
            values=[n_exc, n_ok, n_sinb, n_noapl, n_nue],
            hole=0.52,
            marker_colors=[ROJO, VERDE, GRIS, AMBAR, AZUL_CLAR],
            textinfo="percent", textfont=dict(size=10, color="black"),
            hovertemplate="%{label}: %{value:,}<extra></extra>",
        ))
        fig.update_layout(
            title=dict(text="Aplicación del límite Ley 44", font=dict(size=13, color=AZUL_OSC)),
            showlegend=True,
            legend=dict(font=dict(size=9, color="black"), orientation="h", y=-0.12),
            margin=dict(t=50, b=10, l=0, r=0), height=350, paper_bgcolor="white",
            annotations=[dict(text=f"<b>{n_f:,}</b><br>predios", x=0.5, y=0.5,
                              font_size=12, showarrow=False, font_color=AZUL_OSC)],
        )
        st.plotly_chart(fig, use_container_width=True)

    with c3:
        av_u = suma(df_ex[df_ex["ZONA"] == "URBANO"], "AVALUO_2026")
        av_r = suma(df_ex[df_ex["ZONA"] == "RURAL"], "AVALUO_2026")
        fig = go.Figure(go.Pie(
            labels=["Urbano", "Rural"], values=[av_u, av_r], hole=0.55,
            marker_colors=[AZUL_MED, VERDE],
            textinfo="label+percent", textfont=dict(size=10),
            hovertemplate="%{label}: %{value:$,.0f}<extra></extra>",
        ))
        fig.update_layout(
            title=dict(text="Avalúo 2026 por zona", font=dict(size=13, color=AZUL_OSC)),
            showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=330,
            paper_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)

    c4, c5 = st.columns(2)

    with c4:
        d = df_ex[df_ex["VAR_AVALUO_%"].notna() & (df_ex["VAR_AVALUO_%"].abs() < 2000)]
        if len(d):
            fig = px.histogram(
                d, x="VAR_AVALUO_%", nbins=60, color_discrete_sequence=[AZUL_MED],
                labels={"VAR_AVALUO_%": "Variación del avalúo (%)"},
                title="Variación del avalúo — municipio 2025 → gestor 2026",
            )
            fig.add_vline(x=0, line_dash="dash", line_color=AZUL_OSC, annotation_text="0 %")
            fig.add_vline(x=100, line_dash="dot", line_color=ROJO,
                          annotation_text="+100 % (duplica)")
            fig.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(t=50, b=20, l=10, r=10), height=340, bargap=0.05,
                xaxis=dict(tickfont=dict(color="black")),
                yaxis=dict(title="N° predios", tickfont=dict(color="black")),
            )
            st.plotly_chart(fig, use_container_width=True)

    with c5:
        d = (df_ex.groupby("DESTINO")
             .agg(avaluo=("AVALUO_2026", "sum"))
             .reset_index().sort_values("avaluo").tail(14))
        fig = go.Figure(go.Bar(
            x=d["avaluo"], y=d["DESTINO"], orientation="h", marker_color=AZUL_MED,
            text=[fmt_cop(v) for v in d["avaluo"]], textposition="outside",
        ))
        fig.update_layout(
            title=dict(text="Avalúo 2026 por destino económico",
                       font=dict(size=13, color=AZUL_OSC)),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                       tickfont=dict(color="black")),
            yaxis=dict(tickfont=dict(color="black")),
            margin=dict(t=50, b=20, l=10, r=90), height=340,
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sec-tit">📌 Variación del avalúo por tramos</div>',
                unsafe_allow_html=True)
    tramos = [
        ("Disminuye",            lambda v: v < 0),
        ("Sin variación",        lambda v: v == 0),
        ("Aumenta hasta 25 %",   lambda v: (v > 0) & (v <= 25)),
        ("Aumenta 25 – 50 %",    lambda v: (v > 25) & (v <= 50)),
        ("Aumenta 50 – 100 %",   lambda v: (v > 50) & (v <= 100)),
        ("Aumenta 100 – 200 %",  lambda v: (v > 100) & (v <= 200)),
        ("Aumenta 200 – 500 %",  lambda v: (v > 200) & (v <= 500)),
        ("Aumenta más de 500 %", lambda v: v > 500),
    ]
    dv = df_ex[df_ex["VAR_AVALUO_%"].notna()]
    filas = []
    for lbl, f_ in tramos:
        sel = dv[f_(dv["VAR_AVALUO_%"])]
        filas.append({
            "Tramo": lbl, "Predios": len(sel),
            "% del total": round(len(sel) / max(len(dv), 1) * 100, 2),
            "Avalúo 2025": suma(sel, "AVALUO_2025"),
            "Avalúo 2026": suma(sel, "AVALUO_2026"),
            "Impuesto 2026 pleno": suma(sel, "IMPTO_2026_PLENO"),
            "Exceso Ley 44": suma(sel, "EXCESO_LEY44"),
        })
    df_tramos = pd.DataFrame(filas)
    st.dataframe(
        df_tramos, use_container_width=True, height=330, hide_index=True,
        column_config={
            "Avalúo 2025":         st.column_config.NumberColumn(format="$ %,.0f"),
            "Avalúo 2026":         st.column_config.NumberColumn(format="$ %,.0f"),
            "Impuesto 2026 pleno": st.column_config.NumberColumn(format="$ %,.0f"),
            "Exceso Ley 44":       st.column_config.NumberColumn(format="$ %,.0f"),
            "% del total":         st.column_config.NumberColumn(format="%.2f %%"),
        })


# ── TAB 2: LEY 44 ─────────────────────────────────────────────────────────────
with tabs[1]:
    st.markdown('<div class="sec-tit">⚖️ Cumplimiento del límite — Ley 44/1990 Art. 6</div>',
                unsafe_allow_html=True)

    n_apl = int(df_ex["_aplica"].sum())
    pct_exc = n_exc / n_apl * 100 if n_apl else 0
    st.markdown(
        f"<div class='nota-alerta'>De <b>{n_apl:,}</b> predios con límite aplicable, "
        f"<b>{n_exc:,}</b> ({pct_exc:.1f} %) superarían el tope legal si se liquidaran a tarifa "
        f"plena sobre el nuevo avalúo del gestor. El exceso asciende a <b>{fmt_cop(exceso)}</b>, "
        f"que no puede cobrarse en la vigencia 2026.</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        d = (df_ex[df_ex["_excede"]].groupby("DESTINO")
             .agg(exceso=("EXCESO_LEY44", "sum"), n=("FICHA", "count"))
             .reset_index().sort_values("exceso"))
        if len(d):
            fig = go.Figure(go.Bar(
                x=d["exceso"], y=d["DESTINO"], orientation="h", marker_color=ROJO,
                text=[fmt_cop(v) for v in d["exceso"]], textposition="outside",
                customdata=d["n"],
                hovertemplate="%{y}<br>Exceso: %{x:$,.0f}<br>Predios: %{customdata:,}<extra></extra>",
            ))
            fig.update_layout(
                title=dict(text="Exceso sobre el límite por destino",
                           font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                yaxis=dict(tickfont=dict(color="black")),
                margin=dict(t=50, b=20, l=10, r=90), height=380,
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.success("Ningún predio excede el límite de la Ley 44 en la selección actual.")

    with c2:
        d = (df_ex.groupby("RANGO_AVALUO_2026")
             .agg(i25=("IMPTO_2025", "sum"),
                  i26=("IMPTO_2026_PLENO", "sum"),
                  ic=("IMPTO_CORRECTO_2026", "sum"))
             .reset_index())
        d["_o"] = d["RANGO_AVALUO_2026"].apply(
            lambda x: ORDEN_RNG.index(x) if x in ORDEN_RNG else 99)
        d = d.sort_values("_o")
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Impuesto 2025", x=d["RANGO_AVALUO_2026"], y=d["i25"],
                             marker_color=AZUL_CLAR))
        fig.add_trace(go.Bar(name="2026 tarifa plena", x=d["RANGO_AVALUO_2026"], y=d["i26"],
                             marker_color=ROJO))
        fig.add_trace(go.Bar(name="2026 con Ley 44", x=d["RANGO_AVALUO_2026"], y=d["ic"],
                             marker_color=VERDE))
        fig.update_layout(
            title=dict(text="Impuesto por rango de avalúo", font=dict(size=13, color=AZUL_OSC)),
            barmode="group", plot_bgcolor="white", paper_bgcolor="white",
            yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                       tickfont=dict(color="black")),
            xaxis=dict(tickfont=dict(color="black")),
            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
            margin=dict(t=60, b=20, l=10, r=10), height=380,
        )
        st.plotly_chart(fig, use_container_width=True)

    st.markdown('<div class="sec-tit">📊 Matriz destino × rango de avalúo</div>',
                unsafe_allow_html=True)
    t1, t2, t3 = st.tabs(["N° predios", "Impuesto 2026 pleno ($)", "Exceso Ley 44 ($)"])
    dmat = df_ex[df_ex["RANGO_AVALUO_2026"].isin(ORDEN_RNG)].copy()

    def pivote(val_col, aggfunc="sum"):
        if dmat.empty or val_col not in dmat.columns:
            return pd.DataFrame()
        piv = pd.pivot_table(dmat, values=val_col, index="DESTINO",
                             columns="RANGO_AVALUO_2026", aggfunc=aggfunc,
                             fill_value=0, observed=True)
        cols = [c for c in ORDEN_RNG if c in piv.columns]
        piv = piv[cols]
        piv["TOTAL"] = piv.sum(axis=1)
        piv = piv.sort_values("TOTAL", ascending=False)
        tot = piv.sum(); tot.name = "▶ TOTAL"
        return pd.concat([piv, tot.to_frame().T])

    with t1:
        p = pivote("FICHA", "count")
        if not p.empty:
            cols = [c for c in ORDEN_RNG if c in p.columns]
            st.dataframe(p.astype(int).style.background_gradient(
                cmap="Blues", axis=None, subset=cols),
                use_container_width=True, height=400)
    with t2:
        p = pivote("IMPTO_2026_PLENO")
        if not p.empty:
            cols = [c for c in ORDEN_RNG if c in p.columns]
            st.dataframe(p.style.format("${:,.0f}").background_gradient(
                cmap="Oranges", axis=None, subset=cols),
                use_container_width=True, height=400)
    with t3:
        p = pivote("EXCESO_LEY44")
        if not p.empty:
            cols = [c for c in ORDEN_RNG if c in p.columns]
            st.dataframe(p.style.format("${:,.0f}").background_gradient(
                cmap="Reds", axis=None, subset=cols),
                use_container_width=True, height=400)

    st.markdown('<div class="sec-tit">📋 Resumen por zona y destino</div>',
                unsafe_allow_html=True)
    res = (df_ex.groupby(["ZONA", "DESTINO"])
           .agg(Predios=("FICHA", "count"),
                Aplica_limite=("_aplica", "sum"),
                Exceden=("_excede", "sum"),
                Avaluo_2025=("AVALUO_2025", "sum"),
                Avaluo_2026=("AVALUO_2026", "sum"),
                Impto_2025=("IMPTO_2025", "sum"),
                Impto_2026_pleno=("IMPTO_2026_PLENO", "sum"),
                Impto_correcto=("IMPTO_CORRECTO_2026", "sum"),
                Exceso=("EXCESO_LEY44", "sum"))
           .reset_index().sort_values(["ZONA", "Exceso"], ascending=[True, False]))
    res["Var_avaluo_%"] = ((res["Avaluo_2026"] - res["Avaluo_2025"]) /
                           res["Avaluo_2025"].replace(0, np.nan) * 100).round(2)
    res["%_exceden"] = (res["Exceden"] / res["Aplica_limite"].replace(0, np.nan) * 100).round(1)

    cfg = {c: st.column_config.NumberColumn(format="$ %,.0f") for c in
           ["Avaluo_2025", "Avaluo_2026", "Impto_2025", "Impto_2026_pleno",
            "Impto_correcto", "Exceso"]}
    cfg["Var_avaluo_%"] = st.column_config.NumberColumn(format="%.2f %%")
    cfg["%_exceden"]    = st.column_config.NumberColumn(format="%.1f %%")
    st.dataframe(res, use_container_width=True, height=400, hide_index=True, column_config=cfg)
    descargar(res, "girardota_resumen_destino.xlsx",
              "⬇️ Descargar resumen por destino (.xlsx)", key="dl_res_dest")


# ── TAB 3: ACUERDO 49 ─────────────────────────────────────────────────────────
with tabs[2]:
    st.markdown('<div class="sec-tit">🏛️ Límite local — Acuerdo 49</div>', unsafe_allow_html=True)
    st.markdown(
        "<div class='nota-info'>El límite local sólo se verifica sobre los predios que "
        "<b>no presentan cambio de destino económico ni de área</b> entre las dos vigencias, "
        "que son aquellos en los que el incremento del impuesto obedece exclusivamente a la "
        "actualización del avalúo. Los destinos habitacional (01) y agrícola (24) tienen tope "
        "del 25 %; los demás, del 50 %.</div>", unsafe_allow_html=True)

    dl = df_ex[df_ex["_aplica_local"]]
    n_l25 = int((dl["TIPO_LIMITE_LOCAL"] == "25%").sum())
    n_l50 = int((dl["TIPO_LIMITE_LOCAL"] == "50%").sum())
    exc_l = suma(dl, "EXCESO_LOCAL")

    kl = st.columns(5)
    kpi(kl[0], f"{len(dl):,}",  "Predios con Acuerdo 49", "Sin cambio destino/área", "oscuro")
    kpi(kl[1], f"{n_l25:,}",    "Destinos 01 y 24",       "Tope 25 %",               "ambar")
    kpi(kl[2], f"{n_l50:,}",    "Demás destinos",         "Tope 50 %",               "ambar")
    kpi(kl[3], fmt_cop(exc_l),  "Exceso s/ Acuerdo 49",   "Pleno vs tope local",     "naranja")
    kpi(kl[4], fmt_cop(suma(dl, "IMPTO_CON_LOCAL")), "Impuesto con tope local",
        "Recaudo resultante", "verde")
    st.markdown("<br>", unsafe_allow_html=True)

    if len(dl):
        c1, c2 = st.columns(2)
        with c1:
            d = (dl.groupby("DESTINO")
                 .agg(exceso=("EXCESO_LOCAL", "sum"), n=("FICHA", "count"))
                 .reset_index().sort_values("exceso").tail(14))
            fig = go.Figure(go.Bar(
                x=d["exceso"], y=d["DESTINO"], orientation="h", marker_color=NARANJA,
                text=[fmt_cop(v) for v in d["exceso"]], textposition="outside",
            ))
            fig.update_layout(
                title=dict(text="Exceso sobre el Acuerdo 49 por destino",
                           font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                yaxis=dict(tickfont=dict(color="black")),
                margin=dict(t=50, b=20, l=10, r=90), height=360,
            )
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = go.Figure(go.Bar(
                x=["Impuesto 2025", "2026 pleno", "2026 con Ley 44", "2026 con Acuerdo 49"],
                y=[suma(dl, "IMPTO_2025"), suma(dl, "IMPTO_2026_PLENO"),
                   suma(dl, "IMPTO_CORRECTO_2026"), suma(dl, "IMPTO_CON_LOCAL")],
                marker_color=[AZUL_MED, ROJO, AMBAR, VERDE],
                text=[fmt_cop(v) for v in
                      [suma(dl, "IMPTO_2025"), suma(dl, "IMPTO_2026_PLENO"),
                       suma(dl, "IMPTO_CORRECTO_2026"), suma(dl, "IMPTO_CON_LOCAL")]],
                textposition="outside",
            ))
            fig.update_layout(
                title=dict(text="Efecto de los topes en los predios con Acuerdo 49",
                           font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                xaxis=dict(tickfont=dict(color="black")),
                margin=dict(t=50, b=20, l=10, r=20), height=360, showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True)

        cols_l = [c for c in ["FICHA", "MATRICULA", "ZONA", "DESTINO", "IMPTO_2025",
                              "IMPTO_2026_PLENO", "TIPO_LIMITE_LOCAL", "LIMITE_LOCAL",
                              "IMPTO_CON_LOCAL", "EXCESO_LOCAL"] if c in dl.columns]
        cfg = {c: st.column_config.NumberColumn(format="$ %,.0f") for c in
               ["IMPTO_2025", "IMPTO_2026_PLENO", "LIMITE_LOCAL",
                "IMPTO_CON_LOCAL", "EXCESO_LOCAL"]}
        st.dataframe(dl[cols_l].sort_values("EXCESO_LOCAL", ascending=False),
                     use_container_width=True, height=380, hide_index=True, column_config=cfg)
        descargar(dl[cols_l], "girardota_acuerdo49.xlsx",
                  "⬇️ Descargar predios Acuerdo 49 (.xlsx)", key="dl_ac49")
    else:
        st.warning("Ningún predio cumple las condiciones del Acuerdo 49 en la selección actual.")


# ── TAB 4: PREDIOS NUEVOS ─────────────────────────────────────────────────────
with tabs[3]:
    st.markdown('<div class="sec-tit">🆕 Predios nuevos con matrícula inmobiliaria</div>',
                unsafe_allow_html=True)
    st.markdown(
        "<div class='nota-info'>Aparecen en el reporte 2026 del gestor y <b>no figuran en la "
        "base de datos del municipio</b>. Tienen folio de matrícula, de modo que su "
        "incorporación al catastro es verificable contra el registro. El artículo 6 de la "
        "Ley 44 de 1990 los excluye del tope porque no existe impuesto del año anterior con el "
        "cual comparar: se liquidan a tarifa plena de manera legítima y <b>no suman al "
        "exceso</b> calculado en el análisis.</div>", unsafe_allow_html=True)

    if not df_nuevos.empty:
        av_n = suma(df_nuevos, "AVALUO_2026 ($)")
        i_n  = suma(df_nuevos, "IMPTO_2026_PLENO ($)")
        kn = st.columns(5)
        kpi(kn[0], f"{len(df_nuevos):,}", "Predios nuevos", "Con matrícula", "verde")
        kpi(kn[1], fmt_cop(av_n), "Avalúo 2026", "Total del segmento", "oscuro")
        kpi(kn[2], fmt_cop(i_n), "Impuesto 2026 pleno", "Tarifa plena legítima", "ambar")
        n_urb = int((df_nuevos["ZONA"] == "URBANO").sum())
        kpi(kn[3], f"{n_urb:,}", "Urbanos", f"{len(df_nuevos)-n_urb:,} rurales", "")
        kpi(kn[4], fmt_cop(av_n / max(len(df_nuevos), 1)), "Avalúo promedio", "Por predio", "")
        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            d = (df_nuevos.groupby("DEST_NOM")
                 .agg(n=("FICHA", "count"), av=("AVALUO_2026 ($)", "sum"))
                 .reset_index().sort_values("av").tail(12))
            fig = go.Figure(go.Bar(
                x=d["av"], y=d["DEST_NOM"], orientation="h", marker_color=VERDE,
                text=[fmt_cop(v) for v in d["av"]], textposition="outside",
                customdata=d["n"],
                hovertemplate="%{y}<br>Avalúo: %{x:$,.0f}<br>Predios: %{customdata:,}<extra></extra>",
            ))
            fig.update_layout(
                title=dict(text="Avalúo 2026 por destino — predios nuevos",
                           font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                yaxis=dict(tickfont=dict(color="black")),
                margin=dict(t=50, b=20, l=10, r=90), height=360,
            )
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            d = (df_nuevos.groupby("RANGO_AVALUO")
                 .agg(n=("FICHA", "count")).reset_index())
            d["_o"] = d["RANGO_AVALUO"].apply(
                lambda x: ORDEN_RNG.index(x) if x in ORDEN_RNG else 99)
            d = d.sort_values("_o")
            fig = go.Figure(go.Bar(
                x=d["RANGO_AVALUO"], y=d["n"], marker_color=AZUL_MED,
                text=d["n"], textposition="outside",
            ))
            fig.update_layout(
                title=dict(text="Predios nuevos por rango de avalúo",
                           font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis=dict(title="N° predios", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                xaxis=dict(tickfont=dict(color="black")),
                margin=dict(t=50, b=20, l=10, r=20), height=360, showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True)

        busq_n = st.text_input("🔎 Buscar en predios nuevos (ficha, matrícula, dirección):",
                               key="b_nue", placeholder="Escriba para filtrar…")
        dn = df_nuevos.copy()
        if busq_n:
            m = dn.apply(lambda c: c.astype(str).str.contains(busq_n, case=False, na=False)).any(axis=1)
            dn = dn[m]
        st.markdown(f"**{len(dn):,}** predios nuevos con matrícula")
        cfg = {c: st.column_config.NumberColumn(format="$ %,.0f") for c in
               ["AVALUO_2025 ($)", "AVALUO_2026 ($)", "IMPTO_2025 ($)", "IMPTO_2026_PLENO ($)"]}
        cfg["VAR_AVALUO_%"] = st.column_config.NumberColumn(format="%.2f %%")
        cfg["TARIFA_MIL"]   = st.column_config.NumberColumn(format="%.2f ‰")
        st.dataframe(dn, use_container_width=True, height=420, hide_index=True, column_config=cfg)
        descargar(dn, "girardota_predios_nuevos.xlsx",
                  "⬇️ Descargar predios nuevos (.xlsx)", key="dl_nue")
    else:
        st.warning("No se encontró la hoja «Predios_Nuevos» en el archivo.")


# ── TAB 5: SIN MATRÍCULA / NO REPORTADOS ──────────────────────────────────────
with tabs[4]:
    st.markdown('<div class="sec-tit">🚫 Predios nuevos SIN matrícula inmobiliaria</div>',
                unsafe_allow_html=True)
    st.markdown(
        "<div class='nota-alerta'>Aparecen en el reporte del gestor y no figuran en la base del "
        "municipio, pero <b>carecen de folio de matrícula</b>. Se apartan del análisis: sin "
        "matrícula no hay título registrado que permita calificar la incorporación como "
        "irregular, de modo que <b>NO se cuentan como ilegalidades</b> ni se suman al exceso "
        "ni al recaudo proyectado.</div>", unsafe_allow_html=True)

    if not df_sin_mat.empty:
        av_s = suma(df_sin_mat, "AVALUO_2026 ($)")
        i_s  = suma(df_sin_mat, "IMPTO_2026_PLENO ($)")
        ks = st.columns(4)
        kpi(ks[0], f"{len(df_sin_mat):,}", "Predios sin matrícula", "Apartados", "morado")
        kpi(ks[1], fmt_cop(av_s), "Avalúo 2026", "No se liquida aquí", "oscuro")
        kpi(ks[2], fmt_cop(i_s), "Impuesto teórico", "Referencia únicamente", "ambar")
        n_urb = int((df_sin_mat["ZONA"] == "URBANO").sum())
        kpi(ks[3], f"{n_urb:,}", "Urbanos", f"{len(df_sin_mat)-n_urb:,} rurales", "")
        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            d = (df_sin_mat.groupby("DEST_NOM")
                 .agg(n=("FICHA", "count"), av=("AVALUO_2026 ($)", "sum"))
                 .reset_index().sort_values("av").tail(12))
            fig = go.Figure(go.Bar(
                x=d["av"], y=d["DEST_NOM"], orientation="h", marker_color=MORADO,
                text=[fmt_cop(v) for v in d["av"]], textposition="outside",
            ))
            fig.update_layout(
                title=dict(text="Avalúo 2026 por destino — sin matrícula",
                           font=dict(size=13, color=AZUL_OSC)),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee",
                           tickfont=dict(color="black")),
                yaxis=dict(tickfont=dict(color="black")),
                margin=dict(t=50, b=20, l=10, r=90), height=340,
            )
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            d = (df_sin_mat.groupby("ZONA").agg(n=("FICHA", "count"),
                                                av=("AVALUO_2026 ($)", "sum")).reset_index())
            fig = go.Figure(go.Pie(
                labels=d["ZONA"], values=d["av"], hole=0.5,
                marker_colors=[AZUL_MED, VERDE],
                textinfo="label+percent", textfont=dict(size=11),
                hovertemplate="%{label}: %{value:$,.0f}<extra></extra>",
            ))
            fig.update_layout(
                title=dict(text="Avalúo por zona — sin matrícula",
                           font=dict(size=13, color=AZUL_OSC)),
                showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=340,
                paper_bgcolor="white",
            )
            st.plotly_chart(fig, use_container_width=True)

        cfg = {c: st.column_config.NumberColumn(format="$ %,.0f") for c in
               ["AVALUO_2025 ($)", "AVALUO_2026 ($)", "IMPTO_2025 ($)", "IMPTO_2026_PLENO ($)"]}
        st.dataframe(df_sin_mat, use_container_width=True, height=380,
                     hide_index=True, column_config=cfg)
        descargar(df_sin_mat, "girardota_nuevos_sin_matricula.xlsx",
                  "⬇️ Descargar predios sin matrícula (.xlsx)", key="dl_sinmat")

    st.markdown('<div class="sec-tit">📭 Predios no reportados por el gestor en 2026</div>',
                unsafe_allow_html=True)
    st.markdown(
        "<div class='nota-info'>Figuran en la base de datos del municipio y el gestor no los "
        "incluyó en el registro básico de junio de 2026. Pueden corresponder a englobes, "
        "desenglobes o cancelaciones, pero también a omisiones del reporte: deben verificarse "
        "antes de retirarlos de la facturación.</div>", unsafe_allow_html=True)

    if not df_no_report.empty:
        kr = st.columns(4)
        kpi(kr[0], f"{len(df_no_report):,}", "Predios no reportados", "Requieren verificación",
            "naranja")
        kpi(kr[1], fmt_cop(suma(df_no_report, "AVALUO_2025 ($)")), "Avalúo 2025",
            "Base del municipio", "oscuro")
        kpi(kr[2], fmt_cop(suma(df_no_report, "IMPTO_2025 ($)")), "Impuesto 2025 asociado",
            "En riesgo de perderse", "rojo")
        n_urb = int((df_no_report["ZONA"] == "URBANO").sum())
        kpi(kr[3], f"{n_urb:,}", "Urbanos", f"{len(df_no_report)-n_urb:,} rurales", "")
        st.markdown("<br>", unsafe_allow_html=True)

        cfg = {c: st.column_config.NumberColumn(format="$ %,.0f") for c in
               ["AVALUO_2025 ($)", "AVALUO_2026 ($)", "IMPTO_2025 ($)", "IMPTO_2026_PLENO ($)"]}
        st.dataframe(df_no_report, use_container_width=True, height=380,
                     hide_index=True, column_config=cfg)
        descargar(df_no_report, "girardota_no_reportados.xlsx",
                  "⬇️ Descargar no reportados (.xlsx)", key="dl_norep")


# ── TAB 6: CAMBIOS EXTREMOS ───────────────────────────────────────────────────
with tabs[5]:
    st.markdown(f'<div class="sec-tit">🚨 Predios con variación de avalúo ≥ {umbral_av} %</div>',
                unsafe_allow_html=True)

    d_ext = df[(df["VAR_AVALUO_%"].abs().fillna(0) >= umbral_av) | df["_excede"]].copy()
    d_ext = d_ext.sort_values("VAR_AVALUO_%", ascending=False)
    n_ext_exc = int(d_ext["_excede"].sum())

    st.markdown(
        f"<p style='color:#222; font-size:0.95rem;'><b>{len(d_ext):,}</b> predios con cambio "
        f"extremo · de estos <b>{n_ext_exc:,}</b> exceden el límite de la Ley 44.</p>",
        unsafe_allow_html=True)

    dv = df[df["VAR_AVALUO_%"].notna()].copy()
    if len(dv):
        def cat_var(v):
            if v < umbral_av: return f"Sin cambio (< {umbral_av} %)"
            if v < 200:       return f"{umbral_av} – 200 %"
            if v < 500:       return "200 – 500 %"
            if v < 1000:      return "500 – 1000 %"
            return "> 1000 %"

        dv["_cat"] = dv["VAR_AVALUO_%"].apply(cat_var)
        orden_cat = [f"Sin cambio (< {umbral_av} %)", f"{umbral_av} – 200 %",
                     "200 – 500 %", "500 – 1000 %", "> 1000 %"]
        cnt = dv["_cat"].value_counts()
        labels = [c for c in orden_cat if c in cnt.index]
        values = [int(cnt[c]) for c in labels]
        colores = [GRIS, AMBAR, NARANJA, ROJO, "#8B0000"][:len(labels)]
        n_extremo = sum(v for c, v in zip(labels, values) if not c.startswith("Sin"))

        c1, c2 = st.columns([1.2, 1.8])
        with c1:
            fig = go.Figure(go.Pie(
                labels=labels, values=values, hole=0.52, marker_colors=colores,
                textinfo="percent", textfont=dict(size=10, color="black"),
                hovertemplate="%{label}: %{value:,} predios (%{percent})<extra></extra>",
            ))
            fig.update_layout(
                title=dict(text="Distribución de la variación del avalúo",
                           font=dict(size=13, color=AZUL_OSC)),
                showlegend=True,
                legend=dict(font=dict(color="black", size=9), orientation="h", y=-0.12),
                margin=dict(t=50, b=10, l=0, r=0), height=400, paper_bgcolor="white",
                annotations=[dict(text=f"<b>{n_extremo:,}</b><br>≥ {umbral_av} %",
                                  x=0.5, y=0.5, font_size=11, showarrow=False,
                                  font_color=AZUL_OSC)],
            )
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            d_sc = df[df["AVALUO_2026"].notna() & df["IMPTO_2026_PLENO"].notna()].copy()
            if len(d_sc):
                d_sc["_cat"] = np.where(d_sc["_excede"], "Excede el límite", "Dentro del límite")
                muestra = d_sc.sample(min(len(d_sc), 3000), random_state=42)
                fig = px.scatter(
                    muestra, x="AVALUO_2025", y="AVALUO_2026", color="_cat",
                    color_discrete_map={"Excede el límite": ROJO, "Dentro del límite": VERDE},
                    hover_data={"FICHA": True, "MATRICULA": True, "DESTINO": True,
                                "VAR_AVALUO_%": ":.1f", "_cat": False},
                    labels={"AVALUO_2025": "Avalúo 2025 — municipio ($)",
                            "AVALUO_2026": "Avalúo 2026 — gestor ($)", "_cat": "Estado Ley 44"},
                    title="Avalúo 2025 vs 2026 — muestra hasta 3.000 predios",
                    opacity=0.55, log_x=True, log_y=True,
                )
                fig.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white", height=400,
                    margin=dict(t=50, b=20, l=10, r=10),
                    legend=dict(orientation="h", yanchor="bottom", y=1.01),
                    xaxis=dict(tickfont=dict(color="black")),
                    yaxis=dict(tickfont=dict(color="black")),
                )
                fig.update_traces(marker=dict(size=5))
                st.plotly_chart(fig, use_container_width=True)

    cols_e = [c for c in ["FICHA", "MATRICULA", "CLASIFICACION", "ZONA", "DESTINO", "DIRECCION",
                          "AVALUO_2025", "AVALUO_2026", "VAR_AVALUO_%",
                          "IMPTO_2025", "IMPTO_2026_PLENO", "LIMITE_LEY44",
                          "EXCEDE_LEY44", "EXCESO_LEY44"] if c in d_ext.columns]
    cfg = {c: st.column_config.NumberColumn(format="$ %,.0f") for c in
           ["AVALUO_2025", "AVALUO_2026", "IMPTO_2025", "IMPTO_2026_PLENO",
            "LIMITE_LEY44", "EXCESO_LEY44"]}
    cfg["VAR_AVALUO_%"] = st.column_config.NumberColumn(format="%.2f %%")
    st.dataframe(d_ext[cols_e], use_container_width=True, height=400,
                 hide_index=True, column_config=cfg)
    descargar(d_ext[cols_e], "girardota_cambios_extremos.xlsx",
              "⬇️ Descargar cambios extremos (.xlsx)", key="dl_ext")


# ── TAB 7: INCONSISTENCIAS ────────────────────────────────────────────────────
with tabs[6]:
    st.markdown('<div class="sec-tit">⚠️ Inconsistencias entre la base del municipio y el gestor</div>',
                unsafe_allow_html=True)

    if not df_incons.empty and "TIPO" in df_incons.columns:
        cnt = df_incons["TIPO"].value_counts()
        ki = st.columns(min(len(cnt), 4) or 1)
        for i, (t, n) in enumerate(cnt.items()):
            if i < len(ki):
                kpi(ki[i], f"{n:,}", t, "Casos detectados", "rojo")
        st.markdown("<br>", unsafe_allow_html=True)

        fig = go.Figure(go.Bar(
            x=cnt.values, y=cnt.index, orientation="h", marker_color=ROJO,
            text=cnt.values, textposition="outside",
        ))
        fig.update_layout(
            title=dict(text="Inconsistencias por tipo", font=dict(size=13, color=AZUL_OSC)),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(title="Casos", showgrid=True, gridcolor="#eee",
                       tickfont=dict(color="black")),
            yaxis=dict(tickfont=dict(color="black")),
            margin=dict(t=50, b=20, l=10, r=60), height=280,
        )
        st.plotly_chart(fig, use_container_width=True)

        sel_tipo = st.multiselect("Filtrar por tipo:", sorted(cnt.index), placeholder="Todos")
        di = df_incons[df_incons["TIPO"].isin(sel_tipo)] if sel_tipo else df_incons
        st.dataframe(di, use_container_width=True, height=400, hide_index=True)
        descargar(di, "girardota_inconsistencias.xlsx",
                  "⬇️ Descargar inconsistencias (.xlsx)", key="dl_inc")
    else:
        st.success("No se detectaron inconsistencias entre las bases.")

    st.markdown('<div class="sec-tit">📂 Trazabilidad de los orígenes de datos</div>',
                unsafe_allow_html=True)
    st.markdown(
        "<div class='nota-legal'>⚠️ <b>Advertencia metodológica.</b> Ninguna de las bases "
        "catastrales nuevas contiene tarifa ni impuesto. Como los límites legales se expresan "
        "sobre el impuesto y no sobre el avalúo, la tarifa por mil y el impuesto facturado de "
        "2025 se cruzan por ficha con la base de facturación del municipio. El origen de la "
        "tarifa de cada predio queda registrado en la columna <b>ORIGEN_TARIFA</b>.</div>",
        unsafe_allow_html=True)

    if not df_origenes.empty:
        st.dataframe(df_origenes, use_container_width=True, height=340, hide_index=True)

    orig = df_all["ORIGEN_TARIFA"].astype(str).value_counts().reset_index()
    orig.columns = ["Origen de la tarifa", "Predios"]
    orig["% del total"] = (orig["Predios"] / len(df_all) * 100).round(2)
    c1, c2 = st.columns([1.4, 1.6])
    with c1:
        st.dataframe(orig, use_container_width=True, hide_index=True,
                     column_config={"% del total": st.column_config.NumberColumn(format="%.2f %%")})
    with c2:
        fig = go.Figure(go.Pie(
            labels=orig["Origen de la tarifa"], values=orig["Predios"], hole=0.5,
            marker_colors=[VERDE, AMBAR, NARANJA, ROJO][:len(orig)],
            textinfo="label+percent", textfont=dict(size=10),
        ))
        fig.update_layout(
            title=dict(text="Cobertura del origen de la tarifa",
                       font=dict(size=13, color=AZUL_OSC)),
            showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=300,
            paper_bgcolor="white",
        )
        st.plotly_chart(fig, use_container_width=True)


# ── TAB 8: DETALLE ────────────────────────────────────────────────────────────
with tabs[7]:
    st.markdown('<div class="sec-tit">📋 Detalle de predios</div>', unsafe_allow_html=True)

    cols_vis = [c for c in [
        "FICHA", "MATRICULA", "CLASIFICACION", "ZONA", "DESTINO", "DIRECCION",
        "CEDULA_CATASTRAL", "NPN", "N_PROPIETARIOS",
        "DEST_COD_2025", "DEST_NOM_2025", "DEST_COD_2026", "DEST_NOM_2026",
        "AREA_TERR_2025_M2", "AREA_TERR_2026_M2",
        "AREA_CONS_2025_M2", "AREA_CONS_2026_M2", "N_UNIDADES_2026",
        "AVALUO_2025", "AVALUO_2026", "VAR_AVALUO_%",
        "TARIFA_MIL", "ORIGEN_TARIFA",
        "IMPTO_2025", "IMPTO_2026_PLENO", "APLICA_LIMITE_LEY44", "LIMITE_LEY44",
        "EXCEDE_LEY44", "IMPTO_CORRECTO_2026", "EXCESO_LEY44", "VAR_IMPTO_%",
        "APLICA_LIMITE_LOCAL", "TIPO_LIMITE_LOCAL", "LIMITE_LOCAL",
        "IMPTO_CON_LOCAL", "EXCESO_LOCAL",
        "RANGO_AVALUO_2026", "NOVEDADES",
    ] if c in df.columns]

    busq = st.text_input("🔎 Buscar por ficha, matrícula, dirección o NPN:",
                         placeholder="Escriba para filtrar…")
    d_vis = df[cols_vis].copy()
    if busq:
        m = d_vis.apply(lambda c: c.astype(str).str.contains(busq, case=False, na=False)).any(axis=1)
        d_vis = d_vis[m]

    st.markdown(f"**{len(d_vis):,}** registros")

    cfg = {}
    for c in ["AVALUO_2025", "AVALUO_2026", "IMPTO_2025", "IMPTO_2026_PLENO",
              "LIMITE_LEY44", "IMPTO_CORRECTO_2026", "EXCESO_LEY44",
              "LIMITE_LOCAL", "IMPTO_CON_LOCAL", "EXCESO_LOCAL"]:
        if c in cols_vis:
            cfg[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
    for c in ["VAR_AVALUO_%", "VAR_IMPTO_%"]:
        if c in cols_vis:
            cfg[c] = st.column_config.NumberColumn(c, format="%.2f %%")
    if "TARIFA_MIL" in cols_vis:
        cfg["TARIFA_MIL"] = st.column_config.NumberColumn("TARIFA_MIL", format="%.2f ‰")
    for c in ["AREA_TERR_2025_M2", "AREA_TERR_2026_M2",
              "AREA_CONS_2025_M2", "AREA_CONS_2026_M2"]:
        if c in cols_vis:
            cfg[c] = st.column_config.NumberColumn(c, format="%.2f m²")

    st.dataframe(d_vis, use_container_width=True, height=520,
                 hide_index=True, column_config=cfg)

    csv = d_vis.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Descargar tabla filtrada (.csv)", data=csv,
                       file_name="girardota_predios_2026.csv", mime="text/csv")

    if os.path.exists(RUTA_WORD):
        with open(RUTA_WORD, "rb") as fh:
            st.download_button(
                "📄 Descargar informe en Word", data=fh.read(),
                file_name="Informe_Predial_Girardota_2026.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ── NOTA LEGAL ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="nota-legal">
⚖️ <strong>Nota legal y metodológica.</strong>
Límite Ley 44/1990 Art. 6 = impuesto 2025 × 2: el impuesto no puede superar el doble del año
anterior. No aplica a lotes sin construir (destinos 12 y 13), uso público o exento (19) —excluido
del análisis— ni a parcelas rurales (21, 22, 23), que van a tarifa plena. Los destinos 14 (lote no
urbanizable) y 31 (lote rural) sí aplican. Los predios que ingresan por primera vez al catastro
están excluidos del tope.
El <b>impuesto 2026</b> corresponde al avalúo reportado por el gestor liquidado a tarifa plena, y
el <b>impuesto 2025</b> al facturado por el municipio; las bases catastrales no contienen tarifa,
por lo que ésta proviene del cruce con la base de facturación (ver columna ORIGEN_TARIFA).
Los datos están agregados por ficha. Áreas en m². UVT 2026: $52.374.
Verificar con el Acuerdo Municipal vigente de Girardota.
</div>""", unsafe_allow_html=True)

st.markdown(
    "<br><center style='color:#aaa; font-size:0.76rem;'>"
    "Municipio de Girardota · Predial 2026 · INFORME CATASTRO_DELTA (2025) vs Registro Básico y "
    "Complementario del gestor (2026) · Ley 44/1990 Art. 6 · Acuerdo 49"
    "</center>", unsafe_allow_html=True)


