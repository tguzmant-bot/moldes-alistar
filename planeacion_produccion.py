"""
============================================================
  PLANEACIÓN DE PRODUCCIÓN — ESTRA
  Sube el PDF de Epicor → visualización por máquina,
  orden secuencial multi-día y checklist exportable a PDF
============================================================
  pip install streamlit pdfplumber pandas reportlab openpyxl
  streamlit run planeacion_produccion.py
============================================================
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime, date, timedelta
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

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
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
.main-header {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%);
    border-radius: 12px; padding: 28px 36px; margin-bottom: 28px;
    border-left: 5px solid #22d3ee;
}
.main-header h1 { color: #f0f9ff; font-size: 28px; font-weight: 700; margin: 0 0 4px 0; }
.main-header p  { color: #94a3b8; margin: 0; font-size: 14px; }
.stat-card { background:#0f172a; border:1px solid #1e3a5f; border-radius:10px; padding:20px; text-align:center; }
.stat-num  { font-size:36px; font-weight:700; color:#22d3ee; font-family:'IBM Plex Mono',monospace; }
.stat-label{ color:#64748b; font-size:12px; margin-top:4px; text-transform:uppercase; letter-spacing:1px; }
.fecha-hoy { background:#064e3b; color:#34d399; border:1px solid #059669;
             border-radius:6px; padding:6px 14px; font-size:13px; font-weight:600;
             display:inline-block; margin-bottom:16px; }
.date-badge{ background:#1e3a5f; color:#7dd3fc; border-radius:6px; padding:4px 12px;
             font-size:12px; font-family:'IBM Plex Mono',monospace;
             display:inline-block; margin:16px 0 8px 0; }
.upload-zone{ border:2px dashed #1e3a5f; border-radius:12px; padding:60px;
              text-align:center; background:#0f172a; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
#  PARSER PDF
# ══════════════════════════════════════════════════════════
def parse_epicor_pdf(file) -> pd.DataFrame:
    LINEA = re.compile(
        r'^(\d+)\s+'
        r'(\S+)\s+'
        r'([\d.,]+)\s+'
        r'(\d+:\d+)'
        r'(MOL\w+|Z\w+)\s+'
        r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})\s+'
        r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})\s+'
        r'(RUN|PEND|SUSP|COMP)\s+'
        r'(\w+MED)\s+'
        r'(.+)$'
    )
    registros = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue
            for linea in texto.split('\n'):
                m = LINEA.match(linea.strip())
                if m:
                    try:
                        inicio = datetime.strptime(m.group(6).strip(), "%d/%m/%Y %H:%M")
                        fin    = datetime.strptime(m.group(7).strip(), "%d/%m/%Y %H:%M")
                        registros.append({
                            "Job Number":  m.group(1),
                            "Part Number": m.group(2),
                            "Tot Parts":   m.group(3),
                            "Horas":       m.group(4),
                            "Molde":       m.group(5),
                            "Inicio":      inicio,
                            "Fin":         fin,
                            "Status":      m.group(8),
                            "Máquina":     m.group(9),
                            "Descripción": m.group(10).replace('\\', ' ').strip(),
                            "Inicio_str":  inicio.strftime("%d/%m %H:%M"),
                            "Fin_str":     fin.strftime("%d/%m %H:%M"),
                            "Fecha":       inicio.date(),
                        })
                    except Exception:
                        pass
    return pd.DataFrame(registros)


def filtrar_rango(df: pd.DataFrame, fechas: list) -> pd.DataFrame:
    """Filtra OT activas en cualquiera de las fechas del rango."""
    fechas_set = set(fechas)
    mask = df.apply(
        lambda r: any(
            r["Inicio"].date() <= f <= r["Fin"].date()
            for f in fechas_set
        ), axis=1
    )
    return df[mask].copy()


# ══════════════════════════════════════════════════════════
#  GENERADOR DE PDF CHECKLIST
# ══════════════════════════════════════════════════════════
def generar_pdf_checklist(df_vista: pd.DataFrame, fechas: list) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=2*cm,    bottomMargin=2*cm,
        title="Checklist Alistamiento de Moldes — ESTRA"
    )

    # ── Paleta de colores ──
    C_DARK    = colors.HexColor("#0f172a")
    C_BLUE    = colors.HexColor("#1e3a5f")
    C_CYAN    = colors.HexColor("#22d3ee")
    C_WHITE   = colors.white
    C_LIGHT   = colors.HexColor("#f0f9ff")
    C_GRAY    = colors.HexColor("#cbd5e1")
    C_ROW_ALT = colors.HexColor("#e8f4fd")
    C_GREEN   = colors.HexColor("#dcfce7")
    C_AMBER   = colors.HexColor("#fef9c3")
    C_RED     = colors.HexColor("#fee2e2")
    C_RUN     = colors.HexColor("#166534")

    styles = getSampleStyleSheet()
    desc_style = ParagraphStyle(
        'desc', parent=styles['Normal'],
        fontSize=8, leading=10, textColor=colors.HexColor("#1e293b")
    )
    cell_style = ParagraphStyle(
        'cell', parent=styles['Normal'],
        fontSize=8, leading=10, alignment=TA_CENTER,
        textColor=colors.HexColor("#1e293b")
    )

    story = []

    # ── Portada / Encabezado del documento ──
    fecha_rango = (
        f"{min(fechas).strftime('%d/%m/%Y')}"
        if len(fechas) == 1
        else f"{min(fechas).strftime('%d/%m/%Y')}  →  {max(fechas).strftime('%d/%m/%Y')}"
    )
    header_data = [[
        Paragraph(
            f"<font color='#22d3ee' size='16'><b>🏭 ESTRA — Checklist Alistamiento de Moldes</b></font><br/>"
            f"<font color='#94a3b8' size='9'>Período: {fecha_rango}  ·  "
            f"Total OT: {len(df_vista)}  ·  "
            f"Máquinas: {df_vista['Máquina'].nunique()}  ·  "
            f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</font>",
            ParagraphStyle('hdr', fontSize=14, textColor=C_WHITE,
                           backColor=C_DARK, leading=20)
        )
    ]]
    t_header = Table(header_data, colWidths=[25*cm])
    t_header.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), C_DARK),
        ('TOPPADDING',   (0,0), (-1,-1), 12),
        ('BOTTOMPADDING',(0,0), (-1,-1), 12),
        ('LEFTPADDING',  (0,0), (-1,-1), 16),
        ('RIGHTPADDING', (0,0), (-1,-1), 16),
        ('ROUNDEDCORNERS', [8]),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 0.4*cm))

    # Leyenda de estados checklist
    leyenda_data = [[
        Paragraph("<b>LEYENDA:</b>", cell_style),
        Paragraph("□ <b>ALISTADO</b>", ParagraphStyle('l1', fontSize=8, backColor=C_GREEN,  leading=12, leftIndent=4, rightIndent=4)),
        Paragraph("□ <b>NO ALISTADO</b>", ParagraphStyle('l2', fontSize=8, backColor=C_RED,   leading=12, leftIndent=4, rightIndent=4)),
        Paragraph("□ <b>CAMBIO DE PLANEACIÓN</b>", ParagraphStyle('l3', fontSize=8, backColor=C_AMBER, leading=12, leftIndent=4, rightIndent=4)),
        Paragraph("<b>🔄</b> = Inicio de montaje en esa fecha", cell_style),
    ]]
    t_ley = Table(leyenda_data, colWidths=[3*cm, 3*cm, 3.5*cm, 5*cm, 6*cm])
    t_ley.setStyle(TableStyle([
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN',        (0,0), (-1,-1), 'CENTER'),
        ('GRID',         (0,0), (-1,-1), 0.3, C_GRAY),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('BACKGROUND',   (1,0), (1,0), C_GREEN),
        ('BACKGROUND',   (2,0), (2,0), C_RED),
        ('BACKGROUND',   (3,0), (3,0), C_AMBER),
    ]))
    story.append(t_ley)
    story.append(Spacer(1, 0.5*cm))

    # ── Tabla por día y máquina ──
    df_ord = df_vista.sort_values(["Fecha", "Máquina", "Inicio"]).reset_index(drop=True)

    for fecha in sorted(fechas):
        df_fecha = df_ord[
            (df_ord["Inicio"].dt.date <= fecha) &
            (df_ord["Fin"].dt.date >= fecha)
        ].copy()

        if df_fecha.empty:
            continue

        # Encabezado de fecha
        fecha_titulo = [[
            Paragraph(
                f"<font color='white' size='11'><b>📅  {fecha.strftime('%A %d de %B de %Y').upper()}</b></font>"
                f"<font color='#94a3b8' size='9'>   —   {len(df_fecha)} órdenes en {df_fecha['Máquina'].nunique()} máquinas</font>",
                ParagraphStyle('ft', fontSize=11, textColor=C_WHITE, leading=14)
            )
        ]]
        t_fecha = Table(fecha_titulo, colWidths=[25*cm])
        t_fecha.setStyle(TableStyle([
            ('BACKGROUND',   (0,0), (-1,-1), C_BLUE),
            ('TOPPADDING',   (0,0), (-1,-1), 8),
            ('BOTTOMPADDING',(0,0), (-1,-1), 8),
            ('LEFTPADDING',  (0,0), (-1,-1), 12),
        ]))
        story.append(KeepTogether([t_fecha, Spacer(1, 0.2*cm)]))

        # Encabezados columnas
        col_headers = [
            Paragraph("<b>#</b>",                   cell_style),
            Paragraph("<b>MÁQUINA</b>",              cell_style),
            Paragraph("<b>MOLDE</b>",                cell_style),
            Paragraph("<b>JOB #</b>",                cell_style),
            Paragraph("<b>HORA INICIO</b>",          cell_style),
            Paragraph("<b>HORA FIN</b>",             cell_style),
            Paragraph("<b>HORAS</b>",                cell_style),
            Paragraph("<b>DESCRIPCIÓN</b>",          cell_style),
            Paragraph("<b>ESTADO OT</b>",            cell_style),
            Paragraph("<b>ALISTADO</b>",             cell_style),
            Paragraph("<b>NO ALISTADO</b>",          cell_style),
            Paragraph("<b>CAMBIO PLAN.</b>",         cell_style),
            Paragraph("<b>OBSERVACIONES</b>",        cell_style),
        ]
        col_widths = [0.7, 2.0, 2.2, 1.8, 2.0, 2.0, 1.4, 4.5, 1.5, 1.5, 1.8, 1.8, 3.0]
        col_widths_cm = [w*cm for w in col_widths]

        table_data = [col_headers]
        row_styles = []
        seq = 0

        for maq, g_maq in df_fecha.groupby("Máquina"):
            g_maq = g_maq.sort_values("Inicio")

            # Sub-encabezado de máquina
            maq_row = [
                Paragraph(f"<b>{maq}</b> — {len(g_maq)} OT",
                          ParagraphStyle('maq', fontSize=9, textColor=C_CYAN,
                                         backColor=C_DARK, leading=12)),
                "", "", "", "", "", "", "", "", "", "", "", ""
            ]
            r_idx = len(table_data)
            table_data.append(maq_row)
            row_styles.append(('BACKGROUND',   (0, r_idx), (-1, r_idx), C_DARK))
            row_styles.append(('SPAN',         (0, r_idx), (-1, r_idx)))
            row_styles.append(('LEFTPADDING',  (0, r_idx), (-1, r_idx), 8))
            row_styles.append(('TOPPADDING',   (0, r_idx), (-1, r_idx), 4))
            row_styles.append(('BOTTOMPADDING',(0, r_idx), (-1, r_idx), 4))

            for i, (_, row) in enumerate(g_maq.iterrows()):
                seq += 1
                es_inicio = row["Inicio"].date() == fecha
                prefix = "🔄 " if es_inicio else "   "
                bg = C_ROW_ALT if i % 2 == 0 else C_WHITE

                # Color de estado
                if row["Status"] == "RUN":
                    st_color = colors.HexColor("#166534")
                    st_txt = "🟢 RUN"
                elif row["Status"] == "SUSP":
                    st_color = colors.HexColor("#9a3412")
                    st_txt = "🟠 SUSP"
                else:
                    st_color = colors.HexColor("#1e40af")
                    st_txt = "🔵 PEND"

                fila = [
                    Paragraph(str(seq), cell_style),
                    Paragraph(f"<b>{maq}</b>", cell_style),
                    Paragraph(f"<b>{row['Molde']}</b>", cell_style),
                    Paragraph(row['Job Number'], cell_style),
                    Paragraph(f"{prefix}<b>{row['Inicio_str']}</b>", cell_style),
                    Paragraph(row['Fin_str'], cell_style),
                    Paragraph(row['Horas'], cell_style),
                    Paragraph(row['Descripción'][:60], desc_style),
                    Paragraph(st_txt, cell_style),
                    Paragraph("□", ParagraphStyle('chk', fontSize=14, alignment=TA_CENTER, textColor=colors.HexColor("#166534"))),
                    Paragraph("□", ParagraphStyle('chk', fontSize=14, alignment=TA_CENTER, textColor=colors.HexColor("#dc2626"))),
                    Paragraph("□", ParagraphStyle('chk', fontSize=14, alignment=TA_CENTER, textColor=colors.HexColor("#d97706"))),
                    Paragraph("", cell_style),
                ]
                r_idx = len(table_data)
                table_data.append(fila)
                row_styles.append(('BACKGROUND', (0, r_idx), (-1, r_idx), bg))

        # Construir tabla
        t = Table(table_data, colWidths=col_widths_cm, repeatRows=1)
        base_style = [
            # Encabezado
            ('BACKGROUND',   (0, 0), (-1, 0), C_BLUE),
            ('TEXTCOLOR',    (0, 0), (-1, 0), C_WHITE),
            ('FONTNAME',     (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE',     (0, 0), (-1, 0), 8),
            ('ALIGN',        (0, 0), (-1, 0), 'CENTER'),
            ('VALIGN',       (0, 0), (-1,-1), 'MIDDLE'),
            # Bordes
            ('GRID',         (0, 0), (-1,-1), 0.3, C_GRAY),
            ('LINEBELOW',    (0, 0), (-1, 0), 1,   C_CYAN),
            # Padding
            ('TOPPADDING',   (0, 0), (-1,-1), 4),
            ('BOTTOMPADDING',(0, 0), (-1,-1), 4),
            ('LEFTPADDING',  (0, 0), (-1,-1), 4),
            ('RIGHTPADDING', (0, 0), (-1,-1), 4),
            # Columnas de checklist con fondo de color suave
            ('BACKGROUND',   (9, 1), (9, -1), colors.HexColor("#f0fdf4")),
            ('BACKGROUND',   (10,1), (10,-1), colors.HexColor("#fff1f2")),
            ('BACKGROUND',   (11,1), (11,-1), colors.HexColor("#fefce8")),
            # Línea de observaciones punteada
            ('LINEAFTER',    (11,1), (11,-1), 0.5, C_GRAY),
        ]
        t.setStyle(TableStyle(base_style + row_styles))
        story.append(KeepTogether([t]))
        story.append(Spacer(1, 0.6*cm))

    # Pie de página
    story.append(HRFlowable(width="100%", thickness=0.5, color=C_GRAY))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        f"<font color='#94a3b8' size='8'>Documento generado automáticamente — "
        f"Sistema de Planeación ESTRA — {datetime.now().strftime('%d/%m/%Y %H:%M')} — "
        f"CONFIDENCIAL</font>",
        ParagraphStyle('footer', alignment=TA_CENTER, fontSize=8)
    ))

    doc.build(story)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════
#  INTERFAZ PRINCIPAL
# ══════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
  <h1>🏭 Planeación de Producción</h1>
  <p>Sube el PDF del Job Schedule de Epicor — visualización secuencial por máquina, multi-día y checklist PDF</p>
</div>
""", unsafe_allow_html=True)

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

with st.spinner("Leyendo PDF..."):
    df_total = parse_epicor_pdf(pdf_file)

if df_total.empty:
    st.error("❌ No se pudieron extraer órdenes del PDF.")
    st.stop()

# Rango de fechas disponibles
fecha_min = df_total["Inicio"].dt.date.min()
fecha_max = df_total["Fin"].dt.date.max()

# ── FILTROS ───────────────────────────────────────────────
st.markdown("---")
col_f1, col_f2, col_f3, col_f4 = st.columns([2, 2, 2, 2])

with col_f1:
    fecha_inicio_sel = st.date_input(
        "📅 Desde",
        value=date.today(),
        min_value=fecha_min,
        max_value=fecha_max,
        format="DD/MM/YYYY"
    )

with col_f2:
    fecha_fin_sel = st.date_input(
        "📅 Hasta",
        value=date.today(),
        min_value=fecha_min,
        max_value=fecha_max,
        format="DD/MM/YYYY"
    )

with col_f3:
    maquinas_disponibles = sorted(df_total["Máquina"].dropna().unique().tolist())
    maquina_sel = st.multiselect(
        "🏭 Máquinas",
        options=maquinas_disponibles,
        default=maquinas_disponibles,
        placeholder="Selecciona máquinas..."
    )

with col_f4:
    status_sel = st.multiselect(
        "🔘 Estado",
        options=["RUN", "PEND", "SUSP"],
        default=["RUN", "PEND", "SUSP"]
    )

# Construir lista de fechas del rango
if fecha_inicio_sel > fecha_fin_sel:
    st.warning("⚠️ La fecha de inicio no puede ser mayor a la fecha fin.")
    st.stop()

fechas_rango = [
    fecha_inicio_sel + timedelta(days=i)
    for i in range((fecha_fin_sel - fecha_inicio_sel).days + 1)
]

# Aplicar filtros
df_vista = filtrar_rango(df_total, fechas_rango)
if maquina_sel:
    df_vista = df_vista[df_vista["Máquina"].isin(maquina_sel)]
if status_sel:
    df_vista = df_vista[df_vista["Status"].isin(status_sel)]

df_vista = df_vista.sort_values(["Máquina", "Inicio"]).reset_index(drop=True)

# ── STATS ─────────────────────────────────────────────────
st.markdown("---")
c1, c2, c3, c4, c5 = st.columns(5)
stats = [
    (len(fechas_rango),                   "Días seleccionados"),
    (df_vista["Máquina"].nunique(),        "Máquinas activas"),
    (len(df_vista),                        "Órdenes de trabajo"),
    (len(df_vista[df_vista["Status"]=="RUN"]), "En producción (RUN)"),
    (df_vista["Molde"].nunique(),          "Moldes distintos"),
]
for col, (num, lbl) in zip([c1,c2,c3,c4,c5], stats):
    col.markdown(f'<div class="stat-card"><div class="stat-num">{num}</div><div class="stat-label">{lbl}</div></div>', unsafe_allow_html=True)

st.markdown("---")

if df_vista.empty:
    st.warning("No hay órdenes activas para el rango y filtros seleccionados.")
    st.stop()

# ── BOTÓN DE DESCARGA PDF ─────────────────────────────────
col_pdf, col_excel, _ = st.columns([2, 2, 4])

with col_pdf:
    with st.spinner("Preparando PDF..."):
        pdf_bytes = generar_pdf_checklist(df_vista, fechas_rango)
    rango_str = (
        fecha_inicio_sel.strftime('%Y%m%d')
        if fecha_inicio_sel == fecha_fin_sel
        else f"{fecha_inicio_sel.strftime('%Y%m%d')}_{fecha_fin_sel.strftime('%Y%m%d')}"
    )
    st.download_button(
        label="📄 Descargar Checklist PDF",
        data=pdf_bytes,
        file_name=f"Checklist_Alistamiento_{rango_str}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

with col_excel:
    buf_xl = io.BytesIO()
    with pd.ExcelWriter(buf_xl, engine='openpyxl') as writer:
        df_vista[["Job Number","Máquina","Molde","Inicio_str","Fin_str",
                  "Horas","Status","Descripción","Part Number"]]\
            .rename(columns={"Inicio_str":"Inicio","Fin_str":"Fin"})\
            .sort_values(["Máquina","Inicio"])\
            .to_excel(writer, index=False, sheet_name="OT del rango")
    st.download_button(
        label="⬇️ Descargar Excel",
        data=buf_xl.getvalue(),
        file_name=f"OT_{rango_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")

# ── TABS ──────────────────────────────────────────────────
tab1, tab2 = st.tabs(["🏭  Por Máquina  (secuencial)", "🕐  Cronológico  (todos los días)"])

STATUS_ICON = {"RUN": "🟢", "PEND": "🔵", "SUSP": "🟠"}

# ────────────────────────────────────────────────────────
# TAB 1: POR MÁQUINA — secuencial por día
# ────────────────────────────────────────────────────────
with tab1:
    for fecha in fechas_rango:
        df_fecha = df_vista[
            (df_vista["Inicio"].dt.date <= fecha) &
            (df_vista["Fin"].dt.date   >= fecha)
        ].copy()

        if df_fecha.empty:
            continue

        st.markdown(
            f'<div class="date-badge">📅 {fecha.strftime("%A %d de %B de %Y").capitalize()} — {len(df_fecha)} OT</div>',
            unsafe_allow_html=True
        )

        for maquina, grupo in df_fecha.groupby("Máquina"):
            grupo = grupo.sort_values("Inicio").reset_index(drop=True)
            n = len(grupo)

            with st.expander(f"🏭 **{maquina}** — {n} OT", expanded=True):
                h1,h2,h3,h4,h5,h6,h7 = st.columns([0.5, 1.2, 1.4, 1.8, 2.8, 1, 1])
                h1.markdown("**#**")
                h2.markdown("**Job #**")
                h3.markdown("**Molde**")
                h4.markdown("**Inicio → Fin**")
                h5.markdown("**Descripción**")
                h6.markdown("**Estado**")
                h7.markdown("**Horas**")
                st.divider()

                for idx, (_, row) in enumerate(grupo.iterrows(), 1):
                    es_inicio = row["Inicio"].date() == fecha
                    c1,c2,c3,c4,c5,c6,c7 = st.columns([0.5,1.2,1.4,1.8,2.8,1,1])
                    c1.markdown(f"`{idx}`")
                    c2.markdown(f"`{row['Job Number']}`")
                    c3.markdown(f"**{row['Molde']}**")
                    prefx = "🔄 " if es_inicio else ""
                    c4.markdown(f"{prefx}▶ `{row['Inicio_str']}`  \n⏹ `{row['Fin_str']}`")
                    c5.markdown(row['Descripción'][:65])
                    c6.markdown(f"{STATUS_ICON.get(row['Status'],'⚪')} **{row['Status']}**")
                    c7.markdown(f"`{row['Horas']}`")

# ────────────────────────────────────────────────────────
# TAB 2: CRONOLÓGICO MULTI-DÍA
# ────────────────────────────────────────────────────────
with tab2:
    df_crono = df_vista.sort_values(["Inicio", "Máquina"]).reset_index(drop=True)

    h1,h2,h3,h4,h5,h6 = st.columns([1.5,1.3,1.5,1.2,3,1])
    for h, lbl in zip([h1,h2,h3,h4,h5,h6],
                      ["**Inicio**","**Máquina**","**Molde**","**Job #**","**Descripción**","**Estado**"]):
        h.markdown(lbl)
    st.divider()

    fecha_anterior = None
    for _, row in df_crono.iterrows():
        if row["Fecha"] != fecha_anterior:
            st.markdown(
                f'<div class="date-badge">📆 {row["Inicio"].strftime("%A %d de %B").capitalize()}</div>',
                unsafe_allow_html=True
            )
            fecha_anterior = row["Fecha"]

        es_inicio = row["Inicio"].date() == row["Fecha"]
        prefix = "🔄 " if es_inicio else ""
        c1,c2,c3,c4,c5,c6 = st.columns([1.5,1.3,1.5,1.2,3,1])
        c1.markdown(f"`{row['Inicio_str']}`")
        c2.markdown(f"**{row['Máquina']}**")
        c3.markdown(f"{prefix}**{row['Molde']}**")
        c4.markdown(f"`{row['Job Number']}`")
        c5.markdown(row['Descripción'][:65])
        c6.markdown(f"{STATUS_ICON.get(row['Status'],'⚪')} **{row['Status']}**")
