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
    page_title="Analisi Feedback Corrieri",
    page_icon="📦",
    layout="wide"
)

st.title("📦 Analisi Feedback & Verbatim Corrieri")
st.markdown("Elaborazione dati, calcolo metriche NPS, grafici e generazione report PDF per singolo corriere o generale.")

uploaded_file = st.file_uploader("Trascina qui il file Excel", type=["xlsx", "xls"])

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

def genera_grafico_orizzontale(df_report, piu_votati):
    df_plot = df_report.sort_values(by="Voti", ascending=True).copy()
    num_corrieri = len(df_plot)
    fig_height = max(5.0, num_corrieri * 0.42)
    fig, ax = plt.subplots(figsize=(10, fig_height), dpi=200)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    
    y_pos = range(num_corrieri)
    colori = [get_color_by_score(m) for m in df_plot['Media NPS']]
    ax.barh(y_pos, df_plot['Voti'], color=colori, height=0.62, edgecolor='none', zorder=3)
    ax.set_yticks(y_pos)
    ax.set_yticklabels([f"Corriere {c}" for c in df_plot['Corriere']], fontsize=11, fontweight='bold', color='#1e293b')
    
    max_val = df_plot['Voti'].max() if len(df_plot) > 0 else 10
    for y, c, v in zip(y_pos, df_plot['Corriere'], df_plot['Voti']):
        sing_plur = "voto" if v == 1 else "voti"
        if c in piu_votati:
            ax.text(v + (max_val * 0.02), y, f"{v} {sing_plur}  ★ 1°", va='center', ha='left',
                    fontsize=11, fontweight='bold', color='#b45309',
                    bbox=dict(boxstyle="round,pad=0.25", facecolor="#fef3c7", edgecolor="#f59e0b", linewidth=1.2))
        else:
            ax.text(v + (max_val * 0.02), y, f"{v} {sing_plur}", va='center', ha='left',
                    fontsize=10.5, fontweight='bold', color='#334155')

    ax.set_xlim(0, max_val * 1.30)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#cbd5e1')
    ax.spines['bottom'].set_color('#cbd5e1')
    ax.xaxis.grid(True, linestyle='--', alpha=0.5, color='#e2e8f0', zorder=0)
    ax.tick_params(axis='x', colors='#64748b', labelsize=10)
    ax.set_title("CLASSIFICA CORRIERI PER NUMERO DI VOTI", fontsize=13, fontweight='bold', color='#0f172a', pad=18)
    
    p_green = mpatches.Patch(color='#22c55e', label='Fascia 9 - 10 (Verde)')
    p_yellow = mpatches.Patch(color='#eab308', label='Fascia 7 - 8 (Giallo)')
    p_orange = mpatches.Patch(color='#f97316', label='Fascia 4 - 6 (Arancione)')
    p_red = mpatches.Patch(color='#ef4444', label='Fascia 0 - 3 (Rosso)')
    
    ax.legend(handles=[p_green, p_yellow, p_orange, p_red],
              loc='lower right', facecolor='#f8fafc', edgecolor='#e2e8f0',
              fontsize=9, labelcolor='#334155', framealpha=0.95)
              
    plt.tight_layout()
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

def crea_pdf_singolo_corriere(corriere_id, voti, media, p_list, n_list, d_list):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    
    t_style = ParagraphStyle('T1', parent=styles['Heading1'], fontSize=18, leading=22, textColor=r_colors.HexColor('#0F172A'))
    sub_style = ParagraphStyle('S1', parent=styles['Normal'], fontSize=11, leading=15, textColor=r_colors.HexColor('#475569'))
    sec_p = ParagraphStyle('SP', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#16A34A'))
    sec_n = ParagraphStyle('SN', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#D97706'))
    sec_d = ParagraphStyle('SD', parent=styles['Heading2'], fontSize=12, leading=16, textColor=r_colors.HexColor('#DC2626'))
    txt_style = ParagraphStyle('TX', parent=styles['Normal'], fontSize=10, leading=14, textColor=r_colors.HexColor('#1E293B'))
    empty_style = ParagraphStyle('EM', parent=styles['Italic'], fontSize=9.5, leading=13, textColor=r_colors.HexColor('#94A3B8'))
    
    media_txt = f"{media}" if pd.notna(media) else "-"
    story = [
        Paragraph(f"<b>SCHEDA FEEDBACK &bull; CORRIERE {corriere_id}</b>", t_style),
        Spacer(1, 4),
        Paragraph(f"<b>Totale Voti:</b> {voti} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Media NPS:</b> {media_txt}", sub_style),
        Spacer(1, 10),
        HRFlowable(width="100%", thickness=1, color=r_colors.HexColor('#CBD5E1'), spaceAfter=14),
        
        Paragraph(f"<b>🟢 PROMOTORI ({len(p_list)})</b>", sec_p),
        Spacer(1, 4)
    ]
    if p_list:
        for t in p_list:
            story.append(Paragraph(f"&bull; &ldquo;{t}&rdquo;", txt_style))
            story.append(Spacer(1, 4))
    else:
        story.append(Paragraph("Nessun commento scritto dai clienti.", empty_style))
        
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>🟡 NEUTRALI ({len(n_list)})</b>", sec_n))
    story.append(Spacer(1, 4))
    if n_list:
        for t in n_list:
            story.append(Paragraph(f"&bull; &ldquo;{t}&rdquo;", txt_style))
            story.append(Spacer(1, 4))
    else:
        story.append(Paragraph("Nessun commento scritto dai clienti.", empty_style))
        
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>🔴 DETRATTORI ({len(d_list)})</b>", sec_d))
    story.append(Spacer(1, 4))
    if d_list:
        for t in d_list:
            story.append(Paragraph(f"&bull; &ldquo;{t}&rdquo;", txt_style))
            story.append(Spacer(1, 4))
    else:
        story.append(Paragraph("Nessun commento scritto dai clienti.", empty_style))
        
    doc.build(story)
    buf.seek(0)
    return buf

def crea_pdf_tutti_corrieri(df_report, verbatim_dict):
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
    
    story = [
        Paragraph("<b>REPORT GENERALE FEEDBACK CORRIERI</b>", t_style),
        Paragraph(f"Totale corrieri monitorati: {len(df_report)}", sub_style),
        Spacer(1, 15)
    ]
    
    for _, row in df_report.iterrows():
        c_id = row['Corriere']
        v_data = verbatim_dict.get(c_id, {"prom": [], "neut": [], "detr": []})
        m_txt = f"Media NPS: {row['Media NPS']}" if pd.notna(row['Media NPS']) else "Media: -"
        
        story.append(Paragraph(f"<b>CORRIERE {c_id}</b> — {row['Voti']} voti | {m_txt}", corr_head))
        story.append(Spacer(1, 4))
        
        has_any = False
        if v_data['prom']:
            has_any = True
            story.append(Paragraph("<b>🟢 Promotori:</b>", text_p))
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
            story.append(Paragraph("<b>🔴 Detrattori:</b>", text_d))
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

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        
        auto_corriere, auto_score = None, None
        auto_detr, auto_neut, auto_prom = None, None, None

        for c in df.columns:
            cu = c.upper()
            if not auto_corriere and any(k in cu for k in ["CORRIERE", "AUTISTA", "RISORSA", "VETTORE"]): auto_corriere = c
            elif not auto_score and any(k in cu for k in ["SCORE", "VOTO"]): auto_score = c
            elif "VERB" in cu:
                if "DETR" in cu: auto_detr = c
                elif "NEUT" in cu or "PASSIV" in cu: auto_neut = c
                elif "PROM" in cu: auto_prom = c

        cols_list = ["-- Seleziona --"] + list(df.columns)
        st.markdown("### ⚙️ Selezione Colonne")
        c_sel1, c_sel2 = st.columns(2)
        with c_sel1:
            idx_corr = cols_list.index(auto_corriere) if auto_corriere in cols_list else 0
            col_corriere = st.selectbox("📌 Colonna CORRIERE:", cols_list, index=idx_corr)
        with c_sel2:
            idx_score = cols_list.index(auto_score) if auto_score in cols_list else 0
            col_score = st.selectbox("⭐ Colonna VOTO / NPS SCORE:", cols_list, index=idx_score)

        if col_corriere != "-- Seleziona --":
            df['Corriere_Pulito'] = df[col_corriere].apply(clean_courier_code)
            df_valido = df.dropna(subset=['Corriere_Pulito']).copy()
            df_valido = df_valido[~df_valido['Corriere_Pulito'].isin(['nan', 'None', ''])]
            
            if col_score != "-- Seleziona --":
                df_valido['Score_Num'] = pd.to_numeric(df_valido[col_score], errors='coerce')
            else:
                df_valido['Score_Num'] = None
                
            report_list = []
            verbatim_dict = {}
            
            for corriere, gruppo in df_valido.groupby('Corriere_Pulito'):
                voti_totali = len(gruppo)
                punteggi_validi = gruppo['Score_Num'].dropna()
                
                if len(punteggi_validi) > 0:
                    media = round(punteggi_validi.mean(), 1)
                    promotori = int((punteggi_validi >= 9).sum())
                    detrattori = int((punteggi_validi <= 6).sum())
                    neutrali = int(((punteggi_validi >= 7) & (punteggi_validi <= 8)).sum())
                else:
                    media, promotori, neutrali, detrattori = None, 0, 0, 0
                    
                report_list.append({
                    "Corriere": corriere,
                    "Voti": voti_totali,
                    "Media NPS": media,
                    "Promotori (9-10)": promotori,
                    "Neutrali (7-8)": neutrali,
                    "Detrattori (0-6)": detrattori
                })
                
                p_txts, n_txts, d_txts = [], [], []
                if auto_prom and auto_prom in gruppo:
                    p_txts = [str(x).strip() for x in gruppo[auto_prom].dropna() if str(x).strip() not in ['', '.', '?', 'nan']]
                if auto_neut and auto_neut in gruppo:
                    n_txts = [str(x).strip() for x in gruppo[auto_neut].dropna() if str(x).strip() not in ['', '.', '?', 'nan']]
                if auto_detr and auto_detr in gruppo:
                    d_txts = [str(x).strip() for x in gruppo[auto_detr].dropna() if str(x).strip() not in ['', '.', '?', 'nan']]
                verbatim_dict[corriere] = {"prom": p_txts, "neut": n_txts, "detr": d_txts}
            
            df_report = pd.DataFrame(report_list).sort_values(by="Voti", ascending=False).reset_index(drop=True)
            max_voti = df_report['Voti'].max()
            piu_votati = df_report[df_report['Voti'] == max_voti]['Corriere'].tolist()
            
            tot_prom = df_report['Promotori (9-10)'].sum()
            tot_neut = df_report['Neutrali (7-8)'].sum()
            tot_detr = df_report['Detrattori (0-6)'].sum()

            # --- RIEPILOGO METRICHE IN HTML/CSS (Zero errori script) ---
            st.markdown("---")
            st.markdown("### 📊 Scheda Riepilogativa")
            
            card_html = f"""
            <div style="display: flex; gap: 15px; margin-bottom: 20px; flex-wrap: wrap;">
                <div style="flex: 1; min-width: 150px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 18px;">
                    <div style="font-size: 13px; color: #64748b; font-weight: 600; text-transform: uppercase;">Totale Voti</div>
                    <div style="font-size: 26px; font-weight: bold; color: #0f172a; margin-top: 4px;">{len(df_valido)}</div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #fefce8; border: 1px solid #fef08a; border-radius: 8px; padding: 14px 18px;">
                    <div style="font-size: 13px; color: #854d0e; font-weight: 600; text-transform: uppercase;">🏆 Più Votato</div>
                    <div style="font-size: 22px; font-weight: bold; color: #713f12; margin-top: 4px;">Corr. {', '.join(piu_votati)} <span style="font-size: 14px; color: #a16207;">({max_voti} voti)</span></div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 14px 18px;">
                    <div style="font-size: 13px; color: #166534; font-weight: 600; text-transform: uppercase;">Promotori (9-10)</div>
                    <div style="font-size: 26px; font-weight: bold; color: #15803d; margin-top: 4px;">{tot_prom} 🟢</div>
                </div>
                <div style="flex: 1; min-width: 150px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; padding: 14px 18px;">
                    <div style="font-size: 13px; color: #991b1b; font-weight: 600; text-transform: uppercase;">Detrattori (0-6)</div>
                    <div style="font-size: 26px; font-weight: bold; color: #b91c1c; margin-top: 4px;">{tot_detr} 🔴</div>
                </div>
            </div>
            """
            st.markdown(card_html, unsafe_allow_html=True)

            # --- GRAFICO CLASSIFICA ---
            st.markdown("---")
            st.markdown("### 📈 Classifica Voti Corrieri")
            img_grafico_1 = genera_grafico_orizzontale(df_report, piu_votati)
            cg1, cg2 = st.columns([2.5, 1])
            with cg1: st.image(img_grafico_1.getvalue(), use_container_width=True)
            with cg2:
                st.download_button("💾 Scarica Grafico (PNG)", img_grafico_1.getvalue(), "classifica_voti.png", "image/png", use_container_width=True)
                st.info("💡 Fai clic destro sul grafico per copiarlo e incollarlo su WhatsApp.")

            # --- SEZIONE EXPORT PDF ---
            st.markdown("---")
            st.markdown("### 📄 Generazione File PDF")
            
            col_pdf_single, col_pdf_all = st.columns(2)
            
            # Opzione 1: Singolo Corriere
            with col_pdf_single:
                st.markdown("#### 👤 PDF Singolo Corriere")
                corriere_scelto = st.selectbox("Seleziona il corriere da esportare in PDF:", df_report['Corriere'].tolist())
                dati_c = verbatim_dict.get(corriere_scelto, {"prom": [], "neut": [], "detr": []})
                r_c = df_report[df_report['Corriere'] == corriere_scelto].iloc[0]
                
                pdf_singolo = crea_pdf_singolo_corriere(
                    corriere_scelto, r_c['Voti'], r_c['Media NPS'],
                    dati_c['prom'], dati_c['neut'], dati_c['detr']
                )
                
                st.download_button(
                    label=f"📥 Scarica PDF Corriere {corriere_scelto}",
                    data=pdf_singolo.getvalue(),
                    file_name=f"Report_Corriere_{corriere_scelto}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
                
                # Testo WhatsApp rapido del singolo corriere
                testo_c = f"*📦 RECENSIONI CORRIERE {corriere_scelto}*\n"
                testo_c += f"• *Totale Voti:* {r_c['Voti']} | *Media:* {r_c['Media NPS'] if pd.notna(r_c['Media NPS']) else '-'}\n\n"
                if dati_c['prom']: testo_c += "🟢 *Promotori:*\n" + "\n".join([f"• “{t}”" for t in dati_c['prom']]) + "\n\n"
                if dati_c['neut']: testo_c += "🟡 *Neutrali:*\n" + "\n".join([f"• “{t}”" for t in dati_c['neut']]) + "\n\n"
                if dati_c['detr']: testo_c += "🔴 *Detrattori:*\n" + "\n".join([f"• “{t}”" for t in dati_c['detr']]) + "\n\n"
                if not (dati_c['prom'] or dati_c['neut'] or dati_c['detr']): testo_c += "_Nessun commento scritto._\n"
                
                wa_single_url = f"https://wa.me/?text={urllib.parse.quote(testo_c)}"
                st.markdown(f'''
                    <a href="{wa_single_url}" target="_blank" style="
                        display: block;
                        background-color: #25D366;
                        color: white;
                        padding: 10px 14px;
                        text-align: center;
                        text-decoration: none;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 6px;
                        margin-top: 10px;
                    ">💬 Invia Testo Corriere {corriere_scelto} su WhatsApp</a>
                ''', unsafe_allow_html=True)
                
            # Opzione 2: Tutti i Corrieri
            with col_pdf_all:
                st.markdown("#### 📚 PDF Completo (Tutti i Corrieri)")
                st.caption("Crea un documento PDF con tutte le schede e tutte le recensioni di ogni singolo corriere.")
                
                pdf_tutti = crea_pdf_tutti_corrieri(df_report, verbatim_dict)
                st.download_button(
                    label="📥 Scarica PDF Completo con Tutti i Corrieri",
                    data=pdf_tutti.getvalue(),
                    file_name="Report_Generale_Recensioni_Corrieri.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            # --- TABELLA STATICA PULITA (Zero errori script) ---
            st.markdown("---")
            st.markdown("#### 📋 Tabella Riepilogo Corrieri")
            st.table(df_report)

    except Exception as e:
        st.error(f"Errore: {e}")
