import streamlit as st
import pandas as pd
import urllib.parse
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as r_colors

st.set_page_config(
    page_title="Analisi Feedback NPS - SDA",
    page_icon="📦",
    layout="wide"
)

# --- CSS PERSONALIZZATO: SCHERMATA PRINCIPALE NERA + SIDEBAR GIALLA ---
st.markdown("""
   <style>
        /* ========================================================
           1. SFONDO PRINCIPALE TOTAL BLACK E TESTI CHIARI
           ======================================================== */
        .stApp, [data-testid="stAppViewContainer"] {
            background-color: #0A0A0A !important;
            color: #F8FAFC !important;
        }

        .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp p, .stApp span, .stApp label {
            color: #F8FAFC !important;
        }

        /* Menu a tendina nella pagina nera */
        .stApp [data-baseweb="select"] > div {
            background-color: #1E293B !important;
            color: #FFFFFF !important;
            border: 1px solid #334155 !important;
        }

        /* ========================================================
           2. PULSANTE UPLOADER: "ALLEGA FILE EXCEL" AD ALTA VISIBILITÀ
           ======================================================== */
        /* Riquadro esterno dell'uploader */
        .stApp [data-testid="stFileUploader"] section {
            background-color: #111827 !important;
            border: 2px dashed #CA8A04 !important;
            border-radius: 12px !important;
            padding: 20px !important;
        }

        /* Testo informativo (trascina file / limite formato) */
        .stApp [data-testid="stFileUploader"] section div,
        .stApp [data-testid="stFileUploader"] section span,
        .stApp [data-testid="stFileUploader"] section small {
            color: #E2E8F0 !important;
            font-size: 13.5px !important;
        }

        /* Pulsante "Browse files" trasformato in "Allega File Excel" */
        .stApp [data-testid="stFileUploader"] button {
            background-color: #FACC15 !important; /* Giallo brillante */
            border: 2px solid #EAB308 !important;
            border-radius: 8px !important;
            padding: 8px 18px !important;
            box-shadow: 0 3px 8px rgba(0,0,0,0.4) !important;
            font-size: 0px !important; /* Nasconde il testo originale 'Browse files' */
            visibility: visible !important;
        }

        /* Nuovo testo in italiano ad alto contrasto */
        .stApp [data-testid="stFileUploader"] button::after {
            content: "📁 Allega File Excel" !important;
            font-size: 14.5px !important;
            font-weight: 800 !important;
            color: #0F172A !important; /* Testo nero scuro ben contrastato */
            display: inline-block !important;
            visibility: visible !important;
        }

        .stApp [data-testid="stFileUploader"] button:hover {
            background-color: #FDE047 !important;
            border-color: #CA8A04 !important;
        }

        /* ========================================================
           3. TABELLE SU SFONDO SCURO
           ======================================================== */
        .stApp table {
            background-color: #0F172A !important;
            color: #F8FAFC !important;
            border: 1px solid #334155 !important;
        }
        .stApp th {
            background-color: #1E293B !important;
            color: #FDE047 !important;
        }
        .stApp td {
            border-bottom: 1px solid #1E293B !important;
        }

        /* ========================================================
           4. SIDEBAR SINISTRA GIALLA CON FINESTRELLA BLU
           ======================================================== */
        [data-testid="stSidebar"], 
        [data-testid="stSidebar"] > div:first-child {
            background-color: #FEF08A !important;
        }

        [data-testid="stSidebar"] h1, 
        [data-testid="stSidebar"] h2, 
        [data-testid="stSidebar"] h3, 
        [data-testid="stSidebar"] p, 
        [data-testid="stSidebar"] span, 
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] .stCaption {
            color: #713F12 !important;
            font-weight: 600 !important;
        }

        [data-testid="stSidebar"] hr {
            border-color: #EAB308 !important;
        }

        [data-testid="stSidebar"] [data-testid="stRadio"] > div {
            background-color: #0F172A !important;
            padding: 10px !important;
            border-radius: 12px !important;
            border: 2px solid #1E3A8A !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.20) !important;
            gap: 8px !important;
        }

        [data-testid="stSidebar"] [data-testid="stRadio"] label {
            background-color: #1E293B !important;
            padding: 8px 12px !important;
            border-radius: 8px !important;
            border: 1px solid #334155 !important;
            margin-bottom: 4px !important;
            transition: all 0.2s ease !important;
        }

        [data-testid="stSidebar"] [data-testid="stRadio"] label p,
        [data-testid="stSidebar"] [data-testid="stRadio"] label span {
            color: #FEF08A !important;
            font-weight: 600 !important;
            font-size: 14.5px !important;
        }

        [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
            background-color: #FACC15 !important;
            border: 2px solid #EAB308 !important;
        }

        [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) p,
        [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) span {
            color: #0F172A !important;
            font-weight: 800 !important;
            font-size: 15px !important;
        }

        [data-testid="stSidebar"] [data-testid="stRadio"] [data-baseweb="radio"] div {
            border-color: #0F172A !important;
        }
        [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) [data-baseweb="radio"] div {
            border-color: #0F172A !important;
            background-color: #0F172A !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- FUNZIONI DI UTILITÀ COMUNI ---
def formatta_media(val):
    if pd.isna(val) or val is None:
        return "-"
    return f"{float(val):.2f}".rstrip('0').rstrip('.')

def clean_courier_code(val):
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    s_clean = s.lstrip('0')
    if s_clean == '':
        return '0' if '0' in s else None
    return s_clean

def get_color_by_score(media):
    if pd.isna(media): return '#94a3b8'
    if media >= 9.0: return '#22c55e'
    elif media >= 7.0: return '#eab308'
    elif media >= 4.0: return '#f97316'
    else: return '#ef4444'

def carica_foglio_principale(file_caricato):
    xls = pd.ExcelFile(file_caricato)
    miglior_foglio = xls.sheet_names[0]
    max_cols = 0
    for sheet in xls.sheet_names:
        df_temp = pd.read_excel(file_caricato, sheet_name=sheet, nrows=5)
        if len(df_temp.columns) > max_cols:
            max_cols = len(df_temp.columns)
            miglior_foglio = sheet
    return pd.read_excel(file_caricato, sheet_name=miglior_foglio)

def genera_grafico_orizzontale(df_report, piu_votati):
    df_plot = df_report.sort_values(by="Voti", ascending=True).copy()
    num_elementi = len(df_plot)
    fig_height = max(4.5, num_elementi * 0.38)
    fig, ax = plt.subplots(figsize=(9, fig_height), dpi=100)
    
    # Sfondo del grafico armonizzato con la schermata scura
    fig.patch.set_facecolor('#111827')
    ax.set_facecolor('#111827')
    
    y_pos = range(num_elementi)
    colori = [get_color_by_score(m) for m in df_plot['Media_Num']]
    ax.barh(y_pos, df_plot['Voti'], color=colori, height=0.60, edgecolor='none', zorder=3)
    ax.set_yticks(y_pos)
    ax.set_yticklabels([f"Corriere {c}" for c in df_plot['Elemento']], fontsize=10.5, fontweight='bold', color='#F1F5F9')
    
    max_val = df_plot['Voti'].max() if len(df_plot) > 0 else 10
    for y, c, v in zip(y_pos, df_plot['Elemento'], df_plot['Voti']):
        sing_plur = "voto" if v == 1 else "voti"
        if c in piu_votati:
            ax.text(v + (max_val * 0.02), y, f"{v} {sing_plur} ★ 1°", va='center', ha='left',
                    fontsize=10.5, fontweight='bold', color='#FEF08A',
                    bbox=dict(boxstyle="round,pad=0.25", facecolor="#854D0E", edgecolor="#EAB308", linewidth=1.2))
        else:
            ax.text(v + (max_val * 0.02), y, f"{v} {sing_plur}", va='center', ha='left',
                    fontsize=10, fontweight='bold', color='#CBD5E1')

    ax.set_xlim(0, max_val * 1.30)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#475569')
    ax.spines['bottom'].set_color('#475569')
    ax.xaxis.grid(True, linestyle='--', alpha=0.3, color='#334155', zorder=0)
    ax.tick_params(axis='x', colors='#94A3B8', labelsize=9.5)
    ax.set_title("CLASSIFICA CORRIERI PER NUMERO DI VALUTAZIONI", fontsize=12.5, fontweight='bold', color='#FDE047', pad=15)
    
    p_green = mpatches.Patch(color='#22c55e', label='Fascia 9 - 10 (Verde)')
    p_yellow = mpatches.Patch(color='#eab308', label='Fascia 7 - 8 (Giallo)')
    p_orange = mpatches.Patch(color='#f97316', label='Fascia 4 - 6 (Arancione)')
    p_red = mpatches.Patch(color='#ef4444', label='Fascia 0 - 3 (Rosso)')
    
    ax.legend(handles=[p_green, p_yellow, p_orange, p_red],
              loc='lower right', facecolor='#1E293B', edgecolor='#334155',
              fontsize=8.5, labelcolor='#E2E8F0', framealpha=0.95)
              
    plt.tight_layout()
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

# --- GENERAZIONE PDF UNIFICATA ---
def crea_pdf_singolo_corriere(corriere_id, filiale_nome, voti, media_formatted, p_list, n_list, d_list, tipo_servizio="CONSEGNE"):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    t_style = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=18, leading=22, textColor=r_colors.HexColor('#0F172A'))
    sub_style = ParagraphStyle('S1', parent=styles['Normal'], fontSize=10.5, leading=14, textColor=r_colors.HexColor('#475569'))
    sec_p = ParagraphStyle('SP', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#16A34A'))
    sec_n = ParagraphStyle('SN', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#D97706'))
    sec_d = ParagraphStyle('SD', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#DC2626'))
    txt_style = ParagraphStyle('TX', parent=styles['Normal'], fontSize=9.5, leading=13.5, textColor=r_colors.HexColor('#1E293B'))
    empty_style = ParagraphStyle('EM', parent=styles['Italic'], fontSize=9, leading=13, textColor=r_colors.HexColor('#94A3B8'))
    
    filiale_txt = f" &nbsp;|&nbsp; <b>Filiale:</b> {filiale_nome}" if filiale_nome else ""
    story = [
        Paragraph(f"<b>SCHEDA {tipo_servizio} &bull; CORRIERE {corriere_id}</b>", t_style),
        Spacer(1, 4),
        Paragraph(f"<b>Totale Valutazioni:</b> {voti} &nbsp;|&nbsp; <b>Media NPS:</b> {media_formatted}{filiale_txt}", sub_style),
        Spacer(1, 10),
        HRFlowable(width="100%", thickness=1, color=r_colors.HexColor('#CBD5E1'), spaceAfter=14),
        Paragraph(f"<b>🟢 PROMOTORI / POSITIVI ({len(p_list)})</b>", sec_p),
        Spacer(1, 4)
    ]
    if p_list:
        for t in p_list:
            story.append(Paragraph(f"&bull; &ldquo;{t}&rdquo;", txt_style))
            story.append(Spacer(1, 3))
    else:
        story.append(Paragraph("Nessun commento scritto dai clienti.", empty_style))
        
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>🟡 NEUTRALI ({len(n_list)})</b>", sec_n))
    story.append(Spacer(1, 4))
    if n_list:
        for t in n_list:
            story.append(Paragraph(f"&bull; &ldquo;{t}&rdquo;", txt_style))
            story.append(Spacer(1, 3))
    else:
        story.append(Paragraph("Nessun commento scritto dai clienti.", empty_style))
        
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>🔴 DETRATTORI / RECLAMI ({len(d_list)})</b>", sec_d))
    story.append(Spacer(1, 4))
    if d_list:
        for t in d_list:
            story.append(Paragraph(f"&bull; &ldquo;{t}&rdquo;", txt_style))
            story.append(Spacer(1, 3))
    else:
        story.append(Paragraph("Nessun commento scritto dai clienti.", empty_style))
        
    doc.build(story)
    buf.seek(0)
    return buf

def crea_pdf_tutti_corrieri(filiale_nome, df_report, verbatim_dict, tipo_servizio="CONSEGNE"):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    t_style = ParagraphStyle('TAll', parent=styles['Heading1'], fontSize=17, leading=21, textColor=r_colors.HexColor('#0F172A'))
    sub_style = ParagraphStyle('SAll', parent=styles['Normal'], fontSize=10, textColor=r_colors.HexColor('#475569'))
    corr_head = ParagraphStyle('CHead', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#1E3A8A'))
    text_p = ParagraphStyle('PStyle', parent=styles['Normal'], fontSize=9.5, leading=13, textColor=r_colors.HexColor('#166534'))
    text_n = ParagraphStyle('NStyle', parent=styles['Normal'], fontSize=9.5, leading=13, textColor=r_colors.HexColor('#854D0E'))
    text_d = ParagraphStyle('DStyle', parent=styles['Normal'], fontSize=9.5, leading=13, textColor=r_colors.HexColor('#991B1B'))
    empty_style = ParagraphStyle('EMAll', parent=styles['Italic'], fontSize=9, textColor=r_colors.HexColor('#94A3B8'))
    
    filiale_txt = f" &bull; {filiale_nome}" if filiale_nome else ""
    story = [
        Paragraph(f"<b>REPORT GENERALE {tipo_servizio}{filiale_txt}</b>", t_style),
        Paragraph(f"Totale corrieri monitorati: {len(df_report)}", sub_style),
        Spacer(1, 15)
    ]
    for _, row in df_report.iterrows():
        c_id = row['Elemento']
        v_data = verbatim_dict.get(c_id, {"prom": [], "neut": [], "detr": []})
        m_txt = f"Media NPS: {row['Media NPS']}"
        story.append(Paragraph(f"<b>CORRIERE {c_id}</b> — {row['Voti']} valutazioni | {m_txt}", corr_head))
        story.append(Spacer(1, 4))
        has_any = False
        if v_data['prom']:
            has_any = True
            story.append(Paragraph("<b>🟢 Promotori / Positivi:</b>", text_p))
            for txt in v_data['prom']:
                story.append(Paragraph(f"&bull; &ldquo;{txt}&rdquo;", text_p))
                story.append(Spacer(1, 2))
        if v_data['neut']:
            has_any = True
            story.append(Paragraph("<b>🟡 Neutrali:</b>", text_n))
            for txt in v_data['neut']:
                story.append(Paragraph(f"&bull; &ldquo;{txt}&rdquo;", text_n))
                story.append(Spacer(1, 2))
        if v_data['detr']:
            has_any = True
            story.append(Paragraph("<b>🔴 Detrattori / Criticità:</b>", text_d))
            for txt in v_data['detr']:
                story.append(Paragraph(f"&bull; &ldquo;{txt}&rdquo;", text_d))
                story.append(Spacer(1, 2))
        if not has_any:
            story.append(Paragraph("<i>Nessuna recensione testuale lasciata dai clienti.</i>", empty_style))
        story.append(Spacer(1, 8))
        story.append(HRFlowable(width="100%", thickness=0.5, color=r_colors.HexColor('#CBD5E1'), spaceBefore=4, spaceAfter=10))
        
    doc.build(story)
    buf.seek(0)
    return buf

# --- MENU LATERALE ---
st.sidebar.title("🎛️ Modalità Operativa")
tipo_app = st.sidebar.radio(
    "Scegli il tipo di analisi:",
    ["📦 NPS Consegne", "🚚 NPS Ritiri"],
    index=0
)
st.sidebar.markdown("---")

# BADGE FIRMA NELLA SIDEBAR
st.sidebar.markdown("""
<div style="
    background: #0F172A; 
    border: 1.5px solid #CA8A04; 
    border-radius: 8px; 
    padding: 10px 12px; 
    text-align: center;
">
    <div style="font-size: 11px; color: #94A3B8; text-transform: uppercase; letter-spacing: 0.5px;">Sviluppato per SDA</div>
    <div style="font-size: 13.5px; font-weight: 800; color: #FACC15; margin-top: 3px;">
        ⚡ Ideato da Angelo Grifasi
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 1. MODULO NPS CONSEGNE
# ==========================================
if tipo_app == "📦 NPS Consegne":
    # Logo e Intestazione
    c_title1, c_title2 = st.columns([3, 1])
    with c_title1:
        st.title("📦 Analisi Feedback Consegne")
        st.markdown("<p style='color: #94A3B8; margin-top:-10px;'>Elaborazione metrica NPS, statistiche corrieri e verbatim clienti.</p>", unsafe_allow_html=True)
    with c_title2:
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, #1E293B, #0F172A);
            border: 1.5px solid #EAB308;
            border-radius: 10px;
            padding: 8px 12px;
            text-align: center;
            margin-top: 10px;
        ">
            <span style="font-size: 11px; color: #94A3B8;">DESIGN & DEV</span><br>
            <b style="color: #FACC15; font-size: 13px;">Ideato da Angelo Grifasi</b>
        </div>
        """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Trascina qui il file Excel delle CONSEGNE", type=["xlsx", "xls"], key="upl_consegne")

    if uploaded_file is not None:
        try:
            df = carica_foglio_principale(uploaded_file)
            df.columns = [str(c).strip() for c in df.columns]

            col_ufficio_c = next((c for c in df.columns if any(k in c.upper() for k in ["NOME UFFICIO", "FILIALE", "SEDE", "CENTRO"])), None)
            
            filiale_sel_c = "TUTTE LE FILIALI"
            if col_ufficio_c and col_ufficio_c in df.columns:
                st.markdown("### 🏢 Selezione Filiale")
                uffici_c = sorted([str(x).strip() for x in df[col_ufficio_c].dropna().unique()])
                filiali_disp_c = ["TUTTE LE FILIALI"] + uffici_c
                
                idx_def_c = 0
                for i, nome_f in enumerate(filiali_disp_c):
                    if "CALTANISSETTA" in nome_f.upper():
                        idx_def_c = i
                        break
                        
                filiale_sel_c = st.selectbox("Seleziona la filiale da analizzare:", filiali_disp_c, index=idx_def_c, key="sel_filiale_consegne")
                if filiale_sel_c != "TUTTE LE FILIALI":
                    df = df[df[col_ufficio_c].astype(str).str.strip() == filiale_sel_c].copy()

            # PRIORITÀ A ZONA_PTL
            auto_corriere = next((c for c in df.columns if "ZONA_PTL" in c.upper() or "ZONA PTL" in c.upper()), None)
            if not auto_corriere:
                auto_corriere = next((c for c in df.columns if any(k in c.upper() for k in ["CORRIERE", "AUTISTA", "RISORSA", "VETTORE"])), None)

            auto_score = next((c for c in df.columns if "SCORE" in c.upper() or "VOTO" in c.upper() or "NPS" in c.upper()), None)
            
            auto_detr, auto_neut, auto_prom = None, None, None
            for c in df.columns:
                cu = c.upper()
                if "VERB" in cu:
                    if "DETR" in cu: auto_detr = c
                    elif "NEUT" in cu or "PASSIV" in cu: auto_neut = c
                    elif "PROM" in cu: auto_prom = c

            cols_list = ["-- Seleziona --"] + list(df.columns)
            st.markdown("### ⚙️ Selezione Colonne")
            c_sel1, c_sel2 = st.columns(2)
            with c_sel1:
                idx_corr = cols_list.index(auto_corriere) if auto_corriere in cols_list else 0
                col_corriere = st.selectbox("📌 Colonna CORRIERE (ZONA_PTL):", cols_list, index=idx_corr)
            with c_sel2:
                idx_score = cols_list.index(auto_score) if auto_score in cols_list else 0
                col_score = st.selectbox("⭐ Colonna VOTO / NPS SCORE:", cols_list, index=idx_score)

            if col_corriere != "-- Seleziona --":
                df['Elemento'] = df[col_corriere].apply(clean_courier_code)
                df_valido = df.dropna(subset=['Elemento']).copy()
                df_valido = df_valido[~df_valido['Elemento'].isin(['nan', 'None', ''])]
                
                if col_score != "-- Seleziona --":
                    df_valido['Score_Num'] = pd.to_numeric(df_valido[col_score], errors='coerce')
                else:
                    df_valido['Score_Num'] = None
                    
                report_list = []
                verbatim_dict = {}
                verbatim_filiale_all = []
                
                for corriere, gruppo in df_valido.groupby('Elemento'):
                    voti_totali = len(gruppo)
                    punteggi_validi = gruppo['Score_Num'].dropna()
                    if len(punteggi_validi) > 0:
                        media_val = float(punteggi_validi.mean())
                        promotori = int((punteggi_validi >= 9).sum())
                        detrattori = int((punteggi_validi <= 6).sum())
                        neutrali = int(((punteggi_validi >= 7) & (punteggi_validi <= 8)).sum())
                    else:
                        media_val, promotori, neutrali, detrattori = None, 0, 0, 0
                        
                    report_list.append({
                        "Elemento": corriere,
                        "Voti": voti_totali,
                        "Media_Num": media_val,
                        "Media NPS": formatta_media(media_val),
                        "Promotori (9-10)": promotori,
                        "Neutrali (7-8)": neutrali,
                        "Detrattori (0-6)": detrattori
                    })
                    
                    p_txts, n_txts, d_txts = [], [], []
                    if auto_prom and auto_prom in gruppo:
                        p_txts = [str(x).strip() for x in gruppo[auto_prom].dropna() if str(x).strip() not in ['', '.', '?', 'nan']]
                        for t in p_txts:
                            verbatim_filiale_all.append({"Corriere": corriere, "Tipo": "Promotore (9-10)", "Icona": "🟢", "Testo": t})
                            
                    if auto_neut and auto_neut in gruppo:
                        n_txts = [str(x).strip() for x in gruppo[auto_neut].dropna() if str(x).strip() not in ['', '.', '?', 'nan']]
                        for t in n_txts:
                            verbatim_filiale_all.append({"Corriere": corriere, "Tipo": "Neutrale (7-8)", "Icona": "🟡", "Testo": t})
                            
                    if auto_detr and auto_detr in gruppo:
                        d_txts = [str(x).strip() for x in gruppo[auto_detr].dropna() if str(x).strip() not in ['', '.', '?', 'nan']]
                        for t in d_txts:
                            verbatim_filiale_all.append({"Corriere": corriere, "Tipo": "Detrattore (0-6)", "Icona": "🔴", "Testo": t})
                            
                    verbatim_dict[corriere] = {"prom": p_txts, "neut": n_txts, "detr": d_txts}
                
                if len(report_list) > 0:
                    df_report = pd.DataFrame(report_list).sort_values(by="Voti", ascending=False).reset_index(drop=True)
                    max_voti = df_report['Voti'].max()
                    piu_votati = df_report[df_report['Voti'] == max_voti]['Elemento'].tolist()
                    
                    tot_prom = df_report['Promotori (9-10)'].sum()
                    tot_neut = df_report['Neutrali (7-8)'].sum()
                    tot_detr = df_report['Detrattori (0-6)'].sum()

                    st.markdown("---")
                    titolo_scheda = f"### 📊 Scheda Riepilogativa • {filiale_sel_c}" if filiale_sel_c != "TUTTE LE FILIALI" else "### 📊 Scheda Riepilogativa"
                    st.markdown(titolo_scheda)
                    
                    card_html = f"""
                    <div style="display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap;">
                        <div style="flex: 1; min-width: 140px; background: #1E293B; border: 1px solid #334155; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #94A3B8; font-weight: 700;">TOTALE VOTI</div>
                            <div style="font-size: 24px; font-weight: 800; color: #F8FAFC; margin-top: 2px;">{len(df_valido)}</div>
                        </div>
                        <div style="flex: 1; min-width: 140px; background: #2A2006; border: 1px solid #854D0E; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #FDE047; font-weight: 700;">🏆 PIÙ VOTATO</div>
                            <div style="font-size: 18px; font-weight: 800; color: #FEF08A; margin-top: 2px;">Corr. {', '.join(piu_votati)} <span style="font-size: 13px; color: #FACC15;">({max_voti})</span></div>
                        </div>
                        <div style="flex: 1; min-width: 140px; background: #052E16; border: 1px solid #166534; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #4ADE80; font-weight: 700;">PROMOTORI (9-10)</div>
                            <div style="font-size: 24px; font-weight: 800; color: #22C55E; margin-top: 2px;">{tot_prom} 🟢</div>
                        </div>
                        <div style="flex: 1; min-width: 140px; background: #450A0A; border: 1px solid #991B1B; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #F87171; font-weight: 700;">DETRATTORI (0-6)</div>
                            <div style="font-size: 24px; font-weight: 800; color: #EF4444; margin-top: 2px;">{tot_detr} 🔴</div>
                        </div>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)

                    st.markdown("---")
                    st.markdown(f"### 📈 Classifica Voti Corrieri ({filiale_sel_c})")
                    img_grafico_1 = genera_grafico_orizzontale(df_report, piu_votati)
                    st.image(img_grafico_1.getvalue(), use_container_width=True)
                    st.download_button("💾 Scarica Grafico (PNG)", img_grafico_1.getvalue(), "classifica_voti.png", "image/png", use_container_width=True)

                    # --- RECENSIONI FILIALE CONSEGNE ---
                    st.markdown("---")
                    st.markdown(f"### 💬 Recensioni & Verbatim di {filiale_sel_c}")
                    
                    c_frec1, c_frec2 = st.columns([2, 1])
                    with c_frec1:
                        filtro_tipo_rec = st.selectbox(
                            "Filtra per tipologia di recensione:",
                            ["Tutte le recensioni", "🟢 Solo Promotori (9-10)", "🟡 Solo Neutrali (7-8)", "🔴 Solo Detrattori (0-6)"],
                            key="sel_tipo_rec_c"
                        )
                    
                    recensioni_mostrate = []
                    for r in verbatim_filiale_all:
                        if filtro_tipo_rec == "Tutte le recensioni":
                            recensioni_mostrate.append(r)
                        elif "Promotori" in filtro_tipo_rec and "Promotore" in r['Tipo']:
                            recensioni_mostrate.append(r)
                        elif "Neutrali" in filtro_tipo_rec and "Neutrale" in r['Tipo']:
                            recensioni_mostrate.append(r)
                        elif "Detrattori" in filtro_tipo_rec and "Detrattore" in r['Tipo']:
                            recensioni_mostrate.append(r)
                    
                    with c_frec2:
                        st.metric("Totale Recensioni Trovate", len(recensioni_mostrate))
                    
                    if recensioni_mostrate:
                        for rec in recensioni_mostrate:
                            bg_c = "#064E3B" if "🟢" in rec['Icona'] else ("#78350F" if "🟡" in rec['Icona'] else "#7F1D1D")
                            bd_c = "#059669" if "🟢" in rec['Icona'] else ("#D97706" if "🟡" in rec['Icona'] else "#DC2626")
                            st.markdown(f"""
                            <div style="background:{bg_c}; border:1px solid {bd_c}; border-radius:6px; padding:10px 14px; margin-bottom:8px;">
                                <div style="font-size:12px; font-weight:bold; color:#F8FAFC;">
                                    {rec['Icona']} <b>CORRIERE {rec['Corriere']}</b> &bull; <span style="font-weight:normal;">{rec['Tipo']}</span>
                                </div>
                                <div style="font-size:14px; color:#FFFFFF; margin-top:4px; font-weight:500;">
                                    &ldquo;{rec['Testo']}&rdquo;
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        testo_wa_fil = f"*📦 RECENSIONI {filiale_sel_c.upper()} ({filtro_tipo_rec})*\n\n"
                        for rec in recensioni_mostrate:
                            testo_wa_fil += f"{rec['Icona']} *Corr. {rec['Corriere']}:* “{rec['Testo']}”\n\n"
                        wa_fil_url = f"https://wa.me/?text={urllib.parse.quote(testo_wa_fil)}"
                        st.markdown(f'''
                            <a href="{wa_fil_url}" target="_blank" style="
                                display: inline-block; background-color: #25D366; color: white; padding: 9px 15px;
                                text-align: center; text-decoration: none; font-size: 13.5px; font-weight: bold;
                                border-radius: 6px; margin-top: 6px; margin-bottom: 12px;
                            ">💬 Condividi Recensioni Filtrate su WhatsApp</a>
                        ''', unsafe_allow_html=True)
                    else:
                        st.info("Nessuna recensione scritta trovata per i criteri selezionati.")

                    # --- SEZIONE EXPORT PDF ---
                    st.markdown("---")
                    st.markdown("### 📄 Generazione File PDF")
                    col_pdf_single, col_pdf_all = st.columns(2)
                    
                    lista_corrieri_c = df_report['Elemento'].tolist()
                    with col_pdf_single:
                        st.markdown("#### 👤 PDF Singolo Corriere")
                        corriere_scelto = st.selectbox("Seleziona corriere:", lista_corrieri_c, key=f"sel_corr_c_{filiale_sel_c}")
                        dati_c = verbatim_dict.get(corriere_scelto, {"prom": [], "neut": [], "detr": []})
                        r_c = df_report[df_report['Elemento'] == corriere_scelto].iloc[0]
                        
                        fil_label = filiale_sel_c if filiale_sel_c != "TUTTE LE FILIALI" else ""
                        if st.button(f"⚙️ Prepara PDF Corriere {corriere_scelto}", key=f"btn_prep_c_{corriere_scelto}", use_container_width=True):
                            with st.spinner("Generazione PDF in corso..."):
                                pdf_singolo = crea_pdf_singolo_corriere(
                                    corriere_scelto, fil_label, r_c['Voti'], r_c['Media NPS'],
                                    dati_c['prom'], dati_c['neut'], dati_c['detr'], tipo_servizio="CONSEGNE"
                                )
                                st.download_button(
                                    label="📥 Clicca qui per scaricare il PDF",
                                    data=pdf_singolo.getvalue(),
                                    file_name=f"Report_Consegne_Corriere_{corriere_scelto}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                        
                        testo_c = f"*📦 RECENSIONI CONSEGNE CORRIERE {corriere_scelto}*\n"
                        if fil_label: testo_c += f"*Filiale:* {fil_label}\n"
                        testo_c += f"• *Totale Voti:* {r_c['Voti']} | *Media:* {r_c['Media NPS']}\n\n"
                        if dati_c['prom']: testo_c += "🟢 *Promotori:*\n" + "\n".join([f"• “{t}”" for t in dati_c['prom']]) + "\n\n"
                        if dati_c['neut']: testo_c += "🟡 *Neutrali:*\n" + "\n".join([f"• “{t}”" for t in dati_c['neut']]) + "\n\n"
                        if dati_c['detr']: testo_c += "🔴 *Detrattori:*\n" + "\n".join([f"• “{t}”" for t in dati_c['detr']]) + "\n\n"
                        if not (dati_c['prom'] or dati_c['neut'] or dati_c['detr']): testo_c += "_Nessun commento scritto._\n"
                        
                        wa_single_url = f"https://wa.me/?text={urllib.parse.quote(testo_c)}"
                        st.markdown(f'''
                            <a href="{wa_single_url}" target="_blank" style="
                                display: block; background-color: #25D366; color: white; padding: 9px 12px;
                                text-align: center; text-decoration: none; font-size: 13.5px; font-weight: bold;
                                border-radius: 6px; margin-top: 8px;
                            ">💬 Invia Testo Corriere {corriere_scelto} su WhatsApp</a>
                        ''', unsafe_allow_html=True)
                        
                    with col_pdf_all:
                        st.markdown("#### 📚 PDF Completo (Tutti i Corrieri)")
                        st.caption("Fascicolo completo con tutte le schede e recensioni.")
                        if st.button("⚙️ Prepara PDF Completo di Tutti i Corrieri", key=f"btn_prep_all_c_{filiale_sel_c}", use_container_width=True):
                            with st.spinner("Compilazione fascicolo completo in corso..."):
                                pdf_tutti = crea_pdf_tutti_corrieri(fil_label, df_report, verbatim_dict, tipo_servizio="CONSEGNE")
                                st.download_button(
                                    label="📥 Clicca qui per scaricare il Fascicolo Completo",
                                    data=pdf_tutti.getvalue(),
                                    file_name="Report_Generale_Consegne_Corrieri.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )

                    st.markdown("---")
                    st.markdown(f"#### 📋 Tabella Riepilogo Corrieri • {filiale_sel_c}")
                    df_tabella = df_report.drop(columns=['Media_Num']).rename(columns={'Elemento': 'Corriere'})
                    st.table(df_tabella)
                else:
                    st.warning("Nessun dato trovato per la filiale selezionata.")

        except Exception as e:
            st.error(f"Errore durante l'elaborazione del file Consegne: {e}")

# ==========================================
# 2. MODULO NPS RITIRI
# ==========================================
else:
    # Logo e Intestazione
    c_title1, c_title2 = st.columns([3, 1])
    with c_title1:
        st.title("🚚 Analisi Feedback Ritiri")
        st.markdown("<p style='color: #94A3B8; margin-top:-10px;'>Monitoraggio esiti prese, zone autisti e segnalazioni clienti.</p>", unsafe_allow_html=True)
    with c_title2:
        st.markdown("""
        <div style="
            background: linear-gradient(135deg, #1E293B, #0F172A);
            border: 1.5px solid #EAB308;
            border-radius: 10px;
            padding: 8px 12px;
            text-align: center;
            margin-top: 10px;
        ">
            <span style="font-size: 11px; color: #94A3B8;">DESIGN & DEV</span><br>
            <b style="color: #FACC15; font-size: 13px;">Ideato da Angelo Grifasi</b>
        </div>
        """, unsafe_allow_html=True)

    uploaded_ritiri = st.file_uploader("Trascina qui il file Excel dei RITIRI", type=["xlsx", "xls"], key="upl_ritiri")

    if uploaded_ritiri is not None:
        try:
            dfr_master = carica_foglio_principale(uploaded_ritiri)
            dfr_master.columns = [str(c).strip() for c in dfr_master.columns]
            
            col_ufficio = next((c for c in dfr_master.columns if any(k in c.upper() for k in ["NOME UFFICIO", "FILIALE", "SEDE", "CENTRO"])), None)
            
            st.markdown("### 🏢 Selezione Filiale")
            filiali_disponibili = ["TUTTE LE FILIALI"]
            if col_ufficio and col_ufficio in dfr_master.columns:
                uffici_unici = sorted([str(x).strip() for x in dfr_master[col_ufficio].dropna().unique()])
                filiali_disponibili.extend(uffici_unici)
            
            default_idx = 0
            for i, f_nome in enumerate(filiali_disponibili):
                if "CALTANISSETTA" in f_nome.upper():
                    default_idx = i
                    break
                    
            filiale_selezionata = st.selectbox("Seleziona la filiale da analizzare:", filiali_disponibili, index=default_idx, key="sel_filiale_ritiri_main")
            
            if filiale_selezionata != "TUTTE LE FILIALI" and col_ufficio:
                dfr = dfr_master[dfr_master[col_ufficio].astype(str).str.strip() == filiale_selezionata].copy()
            else:
                dfr = dfr_master.copy()

            col_zona = next((c for c in dfr.columns if "ZONA_PTL" in c.upper() or "ZONA PTL" in c.upper()), None)
            if not col_zona:
                col_zona = next((c for c in dfr.columns if any(k in c.upper() for k in ["CORRIERE", "AUTISTA", "RISORSA", "VETTORE"])), None)

            col_score = next((c for c in dfr.columns if "SCORE" in c.upper() or "VOTO" in c.upper() or "NPS" in c.upper()), None)
            col_v_ok = next((c for c in dfr.columns if "VERBATIM" in c.upper() and "OK" in c.upper()), None)
            col_v_ko = next((c for c in dfr.columns if "VERBATIM" in c.upper() and "KO" in c.upper()), None)
            col_cliente = next((c for c in dfr.columns if any(k in c.upper() for k in ["CLIENTE", "RAGIONE SOCIALE", "MITTENTE"])), None)

            cols_r_list = ["-- Seleziona --"] + list(dfr.columns)
            st.markdown("### ⚙️ Selezione Colonne")
            cr1, cr2 = st.columns(2)
            with cr1:
                idx_zr = cols_r_list.index(col_zona) if col_zona in cols_r_list else 0
                sel_col_corriere_r = st.selectbox("📌 Colonna CORRIERE (ZONA_PTL):", cols_r_list, index=idx_zr, key="sel_col_corr_r_dyn")
            with cr2:
                idx_sr = cols_r_list.index(col_score) if col_score in cols_r_list else 0
                sel_col_score_r = st.selectbox("⭐ Colonna VOTO / NPS SCORE:", cols_r_list, index=idx_sr, key="sel_col_score_r_dyn")

            if sel_col_corriere_r != "-- Seleziona --":
                dfr['Elemento'] = dfr[sel_col_corriere_r].apply(clean_courier_code)
                dfr_valido = dfr.dropna(subset=['Elemento']).copy()
                dfr_valido = dfr_valido[~dfr_valido['Elemento'].isin(['nan', 'None', ''])]
                
                if sel_col_score_r != "-- Seleziona --":
                    dfr_valido['Score_Num'] = pd.to_numeric(dfr_valido[sel_col_score_r], errors='coerce')
                else:
                    dfr_valido['Score_Num'] = None
                
                report_ritiri = []
                verbatim_dict_ritiri = {}
                verbatim_filiale_ritiri_all = []
                
                for corriere, gruppo in dfr_valido.groupby('Elemento'):
                    voti_totali = len(gruppo)
                    punteggi_validi = gruppo['Score_Num'].dropna()
                    
                    if len(punteggi_validi) > 0:
                        media_val = float(punteggi_validi.mean())
                        promotori = int((punteggi_validi >= 9).sum())
                        detrattori = int((punteggi_validi <= 6).sum())
                        neutrali = int(((punteggi_validi >= 7) & (punteggi_validi <= 8)).sum())
                    else:
                        media_val, promotori, neutrali, detrattori = None, 0, 0, 0
                    
                    report_ritiri.append({
                        "Elemento": corriere,
                        "Voti": voti_totali,
                        "Media_Num": media_val,
                        "Media NPS": formatta_media(media_val),
                        "Promotori (9-10)": promotori,
                        "Neutrali (7-8)": neutrali,
                        "Detrattori (0-6)": detrattori
                    })
                    
                    p_txts, n_txts, d_txts = [], [], []
                    
                    for _, riga in gruppo.iterrows():
                        cl = str(riga[col_cliente]).strip() if col_cliente and pd.notna(riga[col_cliente]) else ""
                        voto_riga = riga['Score_Num']
                        
                        if col_v_ok and pd.notna(riga[col_v_ok]):
                            txt_ok = str(riga[col_v_ok]).strip()
                            if txt_ok not in ['', '.', 'nan', '?']:
                                if pd.notna(voto_riga) and voto_riga <= 6:
                                    d_txts.append(txt_ok)
                                    verbatim_filiale_ritiri_all.append({"Corriere": corriere, "Cliente": cl, "Tipo": "Detrattore (0-6)", "Icona": "🔴", "Testo": txt_ok})
                                elif pd.notna(voto_riga) and voto_riga in [7, 8]:
                                    n_txts.append(txt_ok)
                                    verbatim_filiale_ritiri_all.append({"Corriere": corriere, "Cliente": cl, "Tipo": "Neutrale (7-8)", "Icona": "🟡", "Testo": txt_ok})
                                else:
                                    p_txts.append(txt_ok)
                                    verbatim_filiale_ritiri_all.append({"Corriere": corriere, "Cliente": cl, "Tipo": "Promotore (9-10)", "Icona": "🟢", "Testo": txt_ok})
                        
                        if col_v_ko and pd.notna(riga[col_v_ko]):
                            txt_ko = str(riga[col_v_ko]).strip()
                            if txt_ko not in ['', '.', 'nan', '?']:
                                d_txts.append(txt_ko)
                                verbatim_filiale_ritiri_all.append({"Corriere": corriere, "Cliente": cl, "Tipo": "Criticità / Reclamo KO", "Icona": "🔴", "Testo": txt_ko})
                                
                    verbatim_dict_ritiri[corriere] = {"prom": p_txts, "neut": n_txts, "detr": d_txts}

                if len(report_ritiri) > 0:
                    df_report_r = pd.DataFrame(report_ritiri).sort_values(by="Voti", ascending=False).reset_index(drop=True)
                    max_voti_r = df_report_r['Voti'].max()
                    piu_votati_r = df_report_r[df_report_r['Voti'] == max_voti_r]['Elemento'].tolist()
                    
                    tot_prom_r = df_report_r['Promotori (9-10)'].sum()
                    tot_detr_r = df_report_r['Detrattori (0-6)'].sum()

                    st.markdown("---")
                    titolo_scheda_r = f"### 📊 Scheda Riepilogativa Ritiri • {filiale_selezionata}" if filiale_selezionata != "TUTTE LE FILIALI" else "### 📊 Scheda Riepilogativa Ritiri"
                    st.markdown(titolo_scheda_r)
                    
                    card_html_r = f"""
                    <div style="display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap;">
                        <div style="flex: 1; min-width: 140px; background: #1E293B; border: 1px solid #334155; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #94A3B8; font-weight: 700;">TOTALE RITIRI</div>
                            <div style="font-size: 24px; font-weight: 800; color: #F8FAFC; margin-top: 2px;">{len(dfr_valido)}</div>
                        </div>
                        <div style="flex: 1; min-width: 140px; background: #2A2006; border: 1px solid #854D0E; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #FDE047; font-weight: 700;">🏆 PIÙ VOTATO</div>
                            <div style="font-size: 18px; font-weight: 800; color: #FEF08A; margin-top: 2px;">Corr. {', '.join(piu_votati_r)} <span style="font-size: 13px; color: #FACC15;">({max_voti_r})</span></div>
                        </div>
                        <div style="flex: 1; min-width: 140px; background: #052E16; border: 1px solid #166534; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #4ADE80; font-weight: 700;">PROMOTORI (9-10)</div>
                            <div style="font-size: 24px; font-weight: 800; color: #22C55E; margin-top: 2px;">{tot_prom_r} 🟢</div>
                        </div>
                        <div style="flex: 1; min-width: 140px; background: #450A0A; border: 1px solid #991B1B; border-radius: 8px; padding: 12px 14px;">
                            <div style="font-size: 11.5px; color: #F87171; font-weight: 700;">DETRATTORI (0-6)</div>
                            <div style="font-size: 24px; font-weight: 800; color: #EF4444; margin-top: 2px;">{tot_detr_r} 🔴</div>
                        </div>
                    </div>
                    """
                    st.markdown(card_html_r, unsafe_allow_html=True)

                    st.markdown("---")
                    st.markdown(f"### 📈 Classifica Valutazioni Corrieri ({filiale_selezionata})")
                    img_grafico_r = genera_grafico_orizzontale(df_report_r, piu_votati_r)
                    st.image(img_grafico_r.getvalue(), use_container_width=True)
                    st.download_button("💾 Scarica Grafico (PNG)", img_grafico_r.getvalue(), "classifica_ritiri.png", "image/png", use_container_width=True)

                    # --- RECENSIONI FILIALE RITIRI ---
                    st.markdown("---")
                    st.markdown(f"### 💬 Recensioni & Verbatim Ritiri di {filiale_selezionata}")
                    c_frec_r1, c_frec_r2 = st.columns([2, 1])
                    with c_frec_r1:
                        filtro_tipo_rec_r = st.selectbox(
                            "Filtra tipologia recensioni:",
                            ["Tutte le recensioni", "🟢 Solo Promotori / Soddisfatti", "🟡 Solo Neutrali (7-8)", "🔴 Solo Detrattori / Reclami KO"],
                            key=f"sel_tipo_rec_r_{filiale_selezionata}"
                        )
                    
                    recensioni_r_mostrate = []
                    for r in verbatim_filiale_ritiri_all:
                        if filtro_tipo_rec_r == "Tutte le recensioni":
                            recensioni_r_mostrate.append(r)
                        elif "Promotori" in filtro_tipo_rec_r and "Promotore" in r['Tipo']:
                            recensioni_r_mostrate.append(r)
                        elif "Neutrali" in filtro_tipo_rec_r and "Neutrale" in r['Tipo']:
                            recensioni_r_mostrate.append(r)
                        elif "Detrattori" in filtro_tipo_rec_r and ("Detrattore" in r['Tipo'] or "KO" in r['Tipo']):
                            recensioni_r_mostrate.append(r)

                    with c_frec_r2:
                        st.metric("Totale Recensioni Trovate", len(recensioni_r_mostrate))

                    if recensioni_r_mostrate:
                        for rec in recensioni_r_mostrate:
                            bg_c = "#064E3B" if "🟢" in rec['Icona'] else ("#78350F" if "🟡" in rec['Icona'] else "#7F1D1D")
                            bd_c = "#059669" if "🟢" in rec['Icona'] else ("#D97706" if "🟡" in rec['Icona'] else "#DC2626")
                            cl_info = f" &bull; <i>Cliente: {rec['Cliente']}</i>" if rec['Cliente'] else ""
                            st.markdown(f"""
                            <div style="background:{bg_c}; border:1px solid {bd_c}; border-radius:6px; padding:10px 14px; margin-bottom:8px;">
                                <div style="font-size:12px; font-weight:bold; color:#F8FAFC;">
                                    {rec['Icona']} <b>CORRIERE {rec['Corriere']}</b> &bull; {rec['Tipo']}{cl_info}
                                </div>
                                <div style="font-size:14px; color:#FFFFFF; margin-top:4px; font-weight:500;">
                                    &ldquo;{rec['Testo']}&rdquo;
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                            
                        testo_wa_fil_r = f"*🚚 FEEDBACK RITIRI {filiale_selezionata.upper()} ({filtro_tipo_rec_r})*\n\n"
                        for rec in recensioni_r_mostrate:
                            cl_p = f" ({rec['Cliente']})" if rec['Cliente'] else ""
                            testo_wa_fil_r += f"{rec['Icona']} *Corr. {rec['Corriere']}*{cl_p}: “{rec['Testo']}”\n\n"
                        wa_fil_r_url = f"https://wa.me/?text={urllib.parse.quote(testo_wa_fil_r)}"
                        st.markdown(f'''
                            <a href="{wa_fil_r_url}" target="_blank" style="
                                display: inline-block; background-color: #25D366; color: white; padding: 9px 15px;
                                text-align: center; text-decoration: none; font-size: 13.5px; font-weight: bold;
                                border-radius: 6px; margin-top: 6px; margin-bottom: 12px;
                            ">💬 Condividi Recensioni Filtrate su WhatsApp</a>
                        ''', unsafe_allow_html=True)
                    else:
                        st.info("Nessuna recensione scritta trovata per i filtri selezionati.")

                    # --- SEZIONE EXPORT PDF RITIRI ---
                    st.markdown("---")
                    st.markdown("### 📄 Esportazione Report PDF Ritiri")
                    col_pdf_z, col_pdf_ftot = st.columns(2)
                    
                    lista_corrieri_r = df_report_r['Elemento'].tolist()
                    with col_pdf_z:
                        st.markdown("#### 👤 PDF Singolo Corriere")
                        corriere_scelto = st.selectbox("Seleziona Corriere:", lista_corrieri_r, key=f"sel_corr_r_{filiale_selezionata}")
                        dati_z = verbatim_dict_ritiri.get(corriere_scelto, {"prom": [], "neut": [], "detr": []})
                        r_z = df_report_r[df_report_r['Elemento'] == corriere_scelto].iloc[0]
                        
                        fil_label_r = filiale_selezionata if filiale_selezionata != "TUTTE LE FILIALI" else ""
                        if st.button(f"⚙️ Prepara PDF Corriere {corriere_scelto}", key=f"btn_prep_r_{corriere_scelto}", use_container_width=True):
                            with st.spinner("Generazione PDF Corriere in corso..."):
                                pdf_corriere = crea_pdf_singolo_corriere(
                                    corriere_scelto, fil_label_r, r_z['Voti'], r_z['Media NPS'],
                                    dati_z['prom'], dati_z['neut'], dati_z['detr'], tipo_servizio="RITIRI"
                                )
                                st.download_button(
                                    label="📥 Clicca qui per scaricare il PDF",
                                    data=pdf_corriere.getvalue(),
                                    file_name=f"Report_Ritiri_Corriere_{corriere_scelto}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                        
                        testo_z = f"*🚚 FEEDBACK RITIRI CORRIERE {corriere_scelto}*\n"
                        if fil_label_r: testo_z += f"*Filiale:* {fil_label_r}\n"
                        testo_z += f"• *Valutazioni:* {r_z['Voti']} | *Media NPS:* {r_z['Media NPS']}\n\n"
                        if dati_z['prom']:
                            testo_z += "🟢 *Promotori / Positivi:*\n" + "\n".join([f"• “{t}”" for t in dati_z['prom']]) + "\n\n"
                        if dati_z['neut']:
                            testo_z += "🟡 *Neutrali:*\n" + "\n".join([f"• “{t}”" for t in dati_z['neut']]) + "\n\n"
                        if dati_z['detr']:
                            testo_z += "🔴 *Detrattori / Reclami:*\n" + "\n".join([f"• “{t}”" for t in dati_z['detr']]) + "\n\n"
                        if not (dati_z['prom'] or dati_z['neut'] or dati_z['detr']):
                            testo_z += "_Nessuna nota scritta lasciata dai clienti._\n"
                            
                        wa_z_url = f"https://wa.me/?text={urllib.parse.quote(testo_z)}"
                        st.markdown(f'''
                            <a href="{wa_z_url}" target="_blank" style="
                                display: block; background-color: #25D366; color: white; padding: 9px 12px;
                                text-align: center; text-decoration: none; font-size: 13.5px; font-weight: bold;
                                border-radius: 6px; margin-top: 8px;
                            ">💬 Invia Testo Corriere {corriere_scelto} su WhatsApp</a>
                        ''', unsafe_allow_html=True)

                    with col_pdf_ftot:
                        st.markdown("#### 📚 Fascicolo Filiale Completo")
                        st.caption(f"Tutte le schede e le note dei ritiri per {filiale_selezionata}.")
                        if st.button("⚙️ Prepara Fascicolo Completo Filiale", key=f"btn_prep_all_r_{filiale_selezionata}", use_container_width=True):
                            with st.spinner("Compilazione fascicolo filiale in corso..."):
                                pdf_f_tot = crea_pdf_tutti_corrieri(fil_label_r, df_report_r, verbatim_dict_ritiri, tipo_servizio="RITIRI")
                                st.download_button(
                                    label="📥 Clicca qui per scaricare il Fascicolo Completo",
                                    data=pdf_f_tot.getvalue(),
                                    file_name=f"Report_Ritiri_{filiale_selezionata.replace(' ', '_')}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )

                    st.markdown("---")
                    st.markdown(f"#### 📋 Tabella Riepilogo Corrieri • {filiale_selezionata}")
                    df_tabella_r = df_report_r.drop(columns=['Media_Num']).rename(columns={'Elemento': 'Corriere'})
                    st.table(df_tabella_r)
                else:
                    st.warning(f"Nessun dato valido trovato per {filiale_selezionata}.")
            else:
                st.error("Colonna di riferimento (`ZONA_PTL`) non rilevata nel file caricato.")

        except Exception as e:
            st.error(f"Errore durante l'elaborazione del file Ritiri: {e}")
