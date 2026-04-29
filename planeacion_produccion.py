"""
============================================================
  PLANEACIÓN DE PRODUCCIÓN — ESTRA
  Sube el PDF de Epicor y ve las OT por máquina y por hora
============================================================
  Instalar:
    pip install streamlit pdfplumber pandas
  Correr:
    streamlit run planeacion_produccion.py
============================================================
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime, date

# ── CONFIG ────────────────────────────────────────────────
st.set_page_config(
    page_title="Planeación Producción ESTRA",
    page_icon="🏭",
    layout="wide"
)

# ── ESTILOS ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Header principal */
.main-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%);
    border-radius: 12px;
    padding: 28px 36px;
    margin-bottom: 28px;
    border-left: 5px solid #22d3ee;
}
.main-header h1 {
    color: #f0f9ff;
    font-size: 28px;
    font-weight: 700;
    margin: 0 0 4px 0;
    letter-spacing: -0.5px;
}
.main-header p {
    color: #94a3b8;
    margin: 0;
    font-size: 14px;
}

/* Tarjeta de máquina */
.machine-card {
    background: #0f172a;
    border: 1px solid #1e3a5f;
    border-radius: 10px;
    padding: 0;
    margin-bottom: 20px;
    overflow: hidden;
}
.machine-header {
    background: linear-gradient(90deg, #1e3a5f, #0f172a);
    padding: 14px 20px;
    border-bottom: 1px solid #22d3ee33;
    display: flex;
    align-items: center;
    gap: 12px;
}
.machine-name {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 18px;
    font-weight: 600;
    color: #22d3ee;
    letter-spacing: 1px;
}
.machine-count {
    background: #22d3ee22;
    color: #22d3ee;
    border: 1px solid #22d3ee44;
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 12px;
    font-family: 'IBM Plex Mono', monospace;
}

/* Fila de OT */
.ot-row {
    padding: 12px 20px;
    border-bottom: 1px solid #1e293b;
    display: grid;
    grid-template-columns: 90px 100px 130px 1fr 80px 90px;
    gap: 12px;
    align-items: center;
    transition: background 0.15s;
}
.ot-row:hover { background: #1e293b; }
.ot-row:last-child { border-bottom: none; }

.ot-job { font-family: 'IBM Plex Mono', monospace; color: #64748b; font-size: 12px; }
.ot-mol { font-family: 'IBM Plex Mono', monospace; color: #f59e0b; font-size: 13px; font-weight: 600; }
.ot-hora { font-family: 'IBM Plex Mono', monospace; color: #22d3ee; font-size: 13px; }
.ot-desc { color: #cbd5e1; font-size: 13px; }
.ot-status-RUN  { background:#166534; color:#4ade80; border-radius:4px; padding:2px 8px; font-size:11px; font-weight:600; text-align:center; }
.ot-status-PEND { background:#1e3a5f; color:#93c5fd; border-radius:4px; padding:2px 8px; font-size:11px; font-weight:600; text-align:center; }
.ot-status-SUSP { background:#451a03; color:#fb923c; border-radius:4px; padding:2px 8px; font-size:11px; font-weight:600; text-align:center; }
.ot-hours { color: #475569; font-size: 12px; font-family: 'IBM Plex Mono', monospace; }

/* Vista cronológica */
.crono-row {
    background: #0f172a;
    border: 1px solid #1e3a5f;
    border-radius: 8px;
    padding: 12px 18px;
    margin-bottom: 8px;
    display: grid;
    grid-template-columns: 110px 100px 120px 130px 1fr 80px;
    gap: 12px;
    align-items: center;
}
.crono-row:hover { border-color: #22d3ee55; background: #1e293b; }

/* Badges de fecha */
.date-badge {
    background: #1e3a5f;
    color: #7dd3fc;
    border-radius: 6px;
    padding: 4px 12px;
    font-size: 12px;
    font-family: 'IBM Plex Mono', monospace;
    display: inline-block;
    margin: 16px 0 8px 0;
}

/* Stats */
.stat-card {
    background: #0f172a;
    border: 1px solid #1e3a5f;
    border-radius: 10px;
    padding: 20px;
    text-align: center;
}
.stat-num { font-size: 36px; font-weight: 700; color: #22d3ee; font-family: 'IBM Plex Mono', monospace; }
.stat-label { color: #64748b; font-size: 12px; margin-top: 4px; text-transform: uppercase; letter-spacing: 1px; }

/* Upload zone */
.upload-zone {
    border: 2px dashed #1e3a5f;
    border-radius: 12px;
    padding: 60px;
    text-align: center;
    background: #0f172a;
}

/* Filtro de fecha */
.fecha-hoy {
    background: #064e3b;
    color: #34d399;
    border: 1px solid #059669;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 600;
    display: inline-block;
    margin-bottom: 16px;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #1e3a5f !important;
    border-radius: 12px !important;
    background: #0f172a !important;
}
</style>
""", unsafe_allow_html=True)


# ── PARSER DEL PDF ────────────────────────────────────────

def parse_epicor_pdf(file) -> pd.DataFrame:
    """Extrae todas las OT del PDF de Epicor Job Schedule."""
    
    # Regex para una línea de OT:
    # Job#  Part#  TotParts  HH:MMToolNumber  DD/MM/YYYY HH:MM  DD/MM/YYYY HH:MM  STATUS  MACHINE  Descripcion
    LINEA = re.compile(
        r'^(\d+)\s+'                          # Job Number
        r'(\S+)\s+'                           # Part Number
        r'([\d.,]+)\s+'                       # Tot Parts To Go
        r'(\d+:\d+)'                          # Hours To Go (pegado al tool a veces)
        r'(MOL\w+|Z\w+|\w+MOL\w*)\s+'        # Tool Number
        r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})\s+'  # Forecasted Start
        r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})\s+'  # Forecasted End
        r'(RUN|PEND|SUSP|COMP)\s+'            # Status
        r'(\w+MED)\s+'                        # Machine
        r'(.+)$'                              # Description
    )
    
    registros = []
    maquina_actual = None

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue
            for linea in texto.split('\n'):
                linea = linea.strip()
                
                # Detectar encabezado de máquina (ej: "H64MED")
                m_maq = re.match(r'^(H\d+MED|S\d+MED)$', linea)
                if m_maq:
                    maquina_actual = m_maq.group(1)
                    continue
                
                # Intentar parsear línea de OT
                m = LINEA.match(linea)
                if m:
                    try:
                        inicio = datetime.strptime(m.group(6).strip(), "%d/%m/%Y %H:%M")
                        fin    = datetime.strptime(m.group(7).strip(), "%d/%m/%Y %H:%M")
                        registros.append({
                            "Job Number":       m.group(1),
                            "Part Number":      m.group(2),
                            "Tot Parts":        m.group(3),
                            "Horas":            m.group(4),
                            "Molde":            m.group(5),
                            "Inicio":           inicio,
                            "Fin":              fin,
                            "Status":           m.group(8),
                            "Máquina":          m.group(9) if m.group(9) else maquina_actual,
                            "Descripción":      m.group(10).replace('\\', ' ').strip(),
                            "Inicio_str":       inicio.strftime("%d/%m %H:%M"),
                            "Fin_str":          fin.strftime("%d/%m %H:%M"),
                            "Fecha":            inicio.date(),
                        })
                    except Exception:
                        pass

    df = pd.DataFrame(registros)
    return df


def filtrar_por_fecha(df: pd.DataFrame, fecha: date) -> pd.DataFrame:
    """Filtra OT que están activas durante la fecha indicada."""
    mask = (df["Inicio"].dt.date <= fecha) & (df["Fin"].dt.date >= fecha)
    return df[mask].copy()


def badge_status(status):
    return f'<span class="ot-status-{status}">{status}</span>'


# ── INTERFAZ ──────────────────────────────────────────────

st.markdown("""
<div class="main-header">
  <h1>🏭 Planeación de Producción</h1>
  <p>Sube el PDF del Job Schedule de Epicor — visualización automática por máquina y cronológica</p>
</div>
""", unsafe_allow_html=True)

# Upload PDF
pdf_file = st.file_uploader(
    "Arrastra el PDF de Epicor aquí",
    type=["pdf"],
    label_visibility="collapsed"
)

if pdf_file is None:
    st.markdown("""
    <div class="upload-zone">
      <div style="font-size:48px;margin-bottom:16px;">📄</div>
      <div style="color:#475569;font-size:16px;font-weight:600;">Sube el PDF del Job Schedule de Epicor</div>
      <div style="color:#334155;font-size:13px;margin-top:8px;">Exporta desde Epicor → Job Schedule → Print/Export PDF</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Parsear PDF
with st.spinner("Leyendo PDF y extrayendo órdenes de trabajo..."):
    df_total = parse_epicor_pdf(pdf_file)

if df_total.empty:
    st.error("❌ No se pudieron extraer órdenes del PDF. Verifica que sea el Job Schedule de Epicor.")
    st.stop()

# ── FILTRO DE FECHA ───────────────────────────────────────
st.markdown("---")
col_fecha, col_maquina, col_status = st.columns([2, 2, 2])

with col_fecha:
    fecha_sel = st.date_input(
        "📅 Fecha a consultar",
        value=date.today(),
        format="DD/MM/YYYY"
    )

with col_maquina:
    maquinas_disponibles = sorted(df_total["Máquina"].dropna().unique().tolist())
    maquina_sel = st.multiselect(
        "Maquinas a visualizar",
        options=maquinas_disponibles,
        default=maquinas_disponibles,
        placeholder="Selecciona una o mas maquinas..."
    )

with col_status:
    status_sel = st.multiselect(
        "Estado",
        options=["RUN", "PEND", "SUSP"],
        default=["RUN", "PEND", "SUSP"]
    )

# Aplicar filtros
df_dia = filtrar_por_fecha(df_total, fecha_sel)
if maquina_sel:
    df_dia = df_dia[df_dia["Máquina"].isin(maquina_sel)]
if status_sel:
    df_dia = df_dia[df_dia["Status"].isin(status_sel)]

df_dia = df_dia.sort_values(["Máquina", "Inicio"]).reset_index(drop=True)

# ── STATS ─────────────────────────────────────────────────
st.markdown("---")
maquinas_activas = df_dia["Máquina"].nunique()
ot_total         = len(df_dia)
ot_run           = len(df_dia[df_dia["Status"] == "RUN"])
moldes_distintos = df_dia["Molde"].nunique()

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f'<div class="stat-card"><div class="stat-num">{maquinas_activas}</div><div class="stat-label">Máquinas activas</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="stat-card"><div class="stat-num">{ot_total}</div><div class="stat-label">Órdenes de trabajo</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="stat-card"><div class="stat-num">{ot_run}</div><div class="stat-label">En producción (RUN)</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="stat-card"><div class="stat-num">{moldes_distintos}</div><div class="stat-label">Moldes diferentes</div></div>', unsafe_allow_html=True)

st.markdown("---")

if df_dia.empty:
    st.warning(f"No hay órdenes activas para el {fecha_sel.strftime('%d/%m/%Y')} con los filtros seleccionados.")
    st.stop()

# ── TABS: POR MÁQUINA / CRONOLÓGICO ───────────────────────
tab1, tab2 = st.tabs(["🏭  Por Máquina", "🕐  Cronológico"])


# ────────────────────────────────────────────────────────
# TAB 1: POR MÁQUINA
# ────────────────────────────────────────────────────────
with tab1:
    st.markdown(f'<div class="fecha-hoy">📅 {fecha_sel.strftime("%A %d de %B de %Y").capitalize()}</div>', unsafe_allow_html=True)

    STATUS_ICON = {"RUN": "🟢", "PEND": "🔵", "SUSP": "🟠"}
    STATUS_COLOR = {"RUN": "normal", "PEND": "normal", "SUSP": "normal"}

    for maquina, grupo in df_dia.groupby("Máquina"):
        grupo = grupo.sort_values("Inicio").reset_index(drop=True)
        n = len(grupo)

        with st.expander(f"🏭 **{maquina}** — {n} OT", expanded=True):
            # Encabezados de columna
            h1, h2, h3, h4, h5, h6 = st.columns([1.2, 1.4, 1.8, 2.8, 1, 1])
            h1.markdown("**Job #**")
            h2.markdown("**Molde**")
            h3.markdown("**Inicio → Fin**")
            h4.markdown("**Descripción**")
            h5.markdown("**Estado**")
            h6.markdown("**Horas**")
            st.divider()

            for _, row in grupo.iterrows():
                es_inicio_hoy = row["Inicio"].date() == fecha_sel
                c1, c2, c3, c4, c5, c6 = st.columns([1.2, 1.4, 1.8, 2.8, 1, 1])

                # Resaltar cambio de molde con ícono
                prefix = "🔄 " if es_inicio_hoy else ""

                c1.markdown(f"`{row['Job Number']}`")
                c2.markdown(f"**{row['Molde']}**")
                c3.markdown(f"{prefix}▶ `{row['Inicio_str']}`  \n⏹ `{row['Fin_str']}`")
                c4.markdown(row['Descripción'][:60])
                icon = STATUS_ICON.get(row['Status'], '⚪')
                c5.markdown(f"{icon} **{row['Status']}**")
                c6.markdown(f"`{row['Horas']}`")


# ────────────────────────────────────────────────────────
# TAB 2: CRONOLÓGICO
# ────────────────────────────────────────────────────────
with tab2:
    st.markdown(f'<div class="fecha-hoy">📅 Orden cronológico — {fecha_sel.strftime("%d/%m/%Y")}</div>', unsafe_allow_html=True)

    df_crono = df_dia.sort_values("Inicio").reset_index(drop=True)

    # Encabezados
    h1, h2, h3, h4, h5, h6 = st.columns([1.5, 1.3, 1.5, 1.2, 3, 1])
    h1.markdown("**Inicio**")
    h2.markdown("**Máquina**")
    h3.markdown("**Molde**")
    h4.markdown("**Job #**")
    h5.markdown("**Descripción**")
    h6.markdown("**Estado**")
    st.divider()

    fecha_anterior = None
    STATUS_ICON = {"RUN": "🟢", "PEND": "🔵", "SUSP": "🟠"}

    for _, row in df_crono.iterrows():
        if row["Fecha"] != fecha_anterior:
            st.markdown(f'<div class="date-badge">📆 {row["Inicio"].strftime("%A %d de %B").capitalize()}</div>', unsafe_allow_html=True)
            fecha_anterior = row["Fecha"]

        es_inicio_hoy = row["Inicio"].date() == fecha_sel
        prefix = "🔄 " if es_inicio_hoy else ""

        c1, c2, c3, c4, c5, c6 = st.columns([1.5, 1.3, 1.5, 1.2, 3, 1])
        c1.markdown(f"`{row['Inicio_str']}`")
        c2.markdown(f"**{row['Máquina']}**")
        c3.markdown(f"{prefix}**{row['Molde']}**")
        c4.markdown(f"`{row['Job Number']}`")
        c5.markdown(row['Descripción'][:65])
        icon = STATUS_ICON.get(row['Status'], '⚪')
        c6.markdown(f"{icon} **{row['Status']}**")

# ── DESCARGA EXCEL ────────────────────────────────────────
st.markdown("---")
import io
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    df_dia[["Job Number","Máquina","Molde","Inicio_str","Fin_str","Horas","Status","Descripción","Part Number"]]\
        .rename(columns={"Inicio_str":"Inicio","Fin_str":"Fin"})\
        .sort_values(["Máquina","Inicio"])\
        .to_excel(writer, index=False, sheet_name="OT del día")

st.download_button(
    label=f"⬇️  Descargar OT del {fecha_sel.strftime('%d/%m/%Y')} en Excel",
    data=output.getvalue(),
    file_name=f"OT_{fecha_sel.strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
