import streamlit as st
import pandas as pd
import datetime
import os
import json
import urllib.parse
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# Configurazione Pagina Streamlit
st.set_page_config(
    page_title="Gestione Ordini",
    page_icon="📦",
    layout="centered"
)

DB_FILE = "inventario_ottimizzato_app.xlsx"
STORICO_FILE = "storico_ordini.json"

# Inizializzazione dati
@st.cache_data
def load_data():
    if os.path.exists(DB_FILE):
        df = pd.read_excel(DB_FILE, sheet_name="DATABASE PRODOTTI", skiprows=2)
        return df
    else:
        st.error(f"File {DB_FILE} non trovato!")
        return pd.DataFrame()

def save_data(df):
    # Salvataggio nel file excel
    with pd.ExcelWriter(DB_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name="DATABASE PRODOTTI", index=False, startrow=2)

def load_storico():
    if os.path.exists(STORICO_FILE):
        with open(STORICO_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_storico(ordine_data):
    storico = load_storico()
    storico.insert(0, ordine_data)
    with open(STORICO_FILE, "w", encoding="utf-8") as f:
        json.dump(storico, f, ensure_ascii=False, indent=2)

def genera_pdf_ordine(fornitore, data_str, articoli):
    pdf_filename = f"Ordine_{fornitore.replace(' ', '_')}_{data_str}.pdf"
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitle', parent=styles['Heading1'],
        fontName='Helvetica-Bold', fontSize=18,
        textColor=HexColor('#1E3A8A'), spaceAfter=6
    )
    subtitle_style = ParagraphStyle(
        'SubTitle', parent=styles['Normal'],
        fontName='Helvetica', fontSize=10,
        textColor=HexColor('#4B5563'), spaceAfter=15
    )
    cell_style = ParagraphStyle(
        'TableCell', parent=styles['Normal'],
        fontName='Helvetica', fontSize=9, leading=11
    )
    cell_bold = ParagraphStyle(
        'TableBold', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=9, leading=11
    )
    cell_head = ParagraphStyle(
        'TableHead', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=10, leading=12,
        textColor=HexColor('#FFFFFF')
    )
    
    story = []
    story.append(Paragraph(f"ORDINE FORNITORE: {fornitore.upper()}", title_style))
    story.append(Paragraph(f"Data Ordine: {data_str} | Generato da App Gestione Ordini", subtitle_style))
    story.append(Spacer(1, 10))
    
    headers = ["#", "Codice", "Descrizione Prodotto", "Quantità", "Note"]
    rows = [[Paragraph(h, cell_head) for h in headers]]
    
    for idx, item in enumerate(articoli, start=1):
        rows.append([
            Paragraph(str(idx), cell_style),
            Paragraph(str(item.get('id', '-')), cell_style),
            Paragraph(str(item.get('nome', '')), cell_bold),
            Paragraph(f"{item.get('qty', '')} {item.get('udm', '')}", cell_bold),
            Paragraph(str(item.get('note', '')), cell_style)
        ])
    
    col_widths = [1*cm, 2*cm, 8*cm, 3*cm, 4*cm]
    table = Table(rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor('#1E3A8A')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#CBD5E1')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [HexColor('#FFFFFF'), HexColor('#F8FAFC')])
    ]))
    
    story.append(table)
    story.append(Spacer(1, 20))
    story.append(Paragraph("Note di consegna: Si prega di verificare la merce allo scarico.", subtitle_style))
    
    doc.build(story)
    return pdf_filename

# --- INTERFACCIA STREAMLIT ---
st.title("📦 Gestione Ordini")

menu = st.sidebar.radio("Navigazione", ["🛒 Compila Ordine", "📋 Storico Ordini", "➕ Aggiungi Prodotto"])

df_prodotti = load_data()

if menu == "🛒 Compila Ordine":
    if df_prodotti.empty:
        st.warning("Carica il file Excel per iniziare.")
    else:
        fornitori = sorted(df_prodotti["Fornitore"].dropna().unique().tolist())
        fornitore_scelto = st.selectbox("Seleziona Fornitore:", fornitori)
        
        df_fornitore = df_prodotti[df_prodotti["Fornitore"] == fornitore_scelto].copy()
        
        st.markdown(f"### Prodotti per: **{fornitore_scelto}**")
        st.caption("🟢 Verde = In Ordine | 🔴 Rosso = Non in Ordine")
        
        ordini_selezionati = []
        
        for idx, row in df_fornitore.iterrows():
            col_status, col_nome, col_qty, col_note = st.columns([1, 4, 2, 3])
            
            p_id = row.get("ID Prodotto", f"P{idx}")
            p_nome = row.get("Nome Prodotto", "")
            p_udm = row.get("Unità di Misura", "pz")
            p_note_def = str(row.get("Note Ordine", "")) if pd.notna(row.get("Note Ordine")) else ""
            p_qty_def = str(row.get("Quantità", 1)) if pd.notna(row.get("Quantità")) else "1"
            
            with col_status:
                in_ordine = st.checkbox("🟢", key=f"chk_{idx}", help="Spunta per inserire nell'ordine")
            
            with col_nome:
                if in_ordine:
                    st.markdown(f"🟢 **{p_nome}**")
                else:
                    st.markdown(f"🔴 {p_nome}")
            
            with col_qty:
                qty_val = st.text_input("Qtà", value=p_qty_def if in_ordine else "", key=f"qty_{idx}", placeholder=p_udm)
            
            with col_note:
                note_val = st.text_input("Note", value=p_note_def if in_ordine else "", key=f"note_{idx}", placeholder="Note...")
            
            if in_ordine:
                ordini_selezionati.append({
                    "id": p_id,
                    "nome": p_nome,
                    "qty": qty_val if qty_val else "1",
                    "udm": p_udm,
                    "note": note_val
                })
        
        st.markdown("---")
        st.subheader("🚀 Azioni Ordine")
        
        if ordini_selezionati:
            st.success(f"Articoli selezionati in ordine: {len(ordini_selezionati)}")
            
            # Formattazione Testo WhatsApp
            oggi_str = datetime.date.today().strftime("%d/%m/%Y")
            testo_wa = f"🛒 *ORDINE {fornitore_scelto.upper()}* - {oggi_str}\n"
            testo_wa += "────────────────────────\n"
            for item in ordini_selezionati:
                note_str = f" ({item['note']})" if item['note'] else ""
                testo_wa += f"• *{item['nome']}*: x {item['qty']} {item['udm']}{note_str}\n"
            testo_wa += "────────────────────────\n"
            testo_wa += "Si prega di confermare la ricezione. Grazie!"
            
            col_wa, col_pdf = st.columns(2)
            
            with col_wa:
                st.text_area("Testo per WhatsApp:", value=testo_wa, height=180)
                wa_url = f"https://wa.me/?text={urllib.parse.quote(testo_wa)}"
                st.markdown(f"[📲 Invia direttamente su WhatsApp]({wa_url})", unsafe_allow_html=True)
            
            with col_pdf:
                if st.button("📄 Genera e Salva PDF"):
                    pdf_file = genera_pdf_ordine(fornitore_scelto, oggi_str.replace('/', '-'), ordini_selezionati)
                    
                    # Salva nello storico
                    save_storico({
                        "data": oggi_str,
                        "fornitore": fornitore_scelto,
                        "articoli_count": len(ordini_selezionati),
                        "dettagli": ordini_selezionati
                    })
                    
                    with open(pdf_file, "rb") as f:
                        st.download_button("⬇️ Scarica PDF Ordine", data=f, file_name=pdf_file, mime="application/pdf")
                    st.success("Ordine salvato nello storico!")
        else:
            st.info("Nessun articolo selezionato col pallino verde 🟢.")

elif menu == "📋 Storico Ordini":
    st.subheader("📋 Storico degli Ordini")
    storico = load_storico()
    
    if not storico:
        st.info("Nessun ordine presente nello storico.")
    else:
        for order in storico:
            with st.expander(f"• {order['data']} - {order['fornitore']} ({order['articoli_count']} articoli)"):
                for item in order['dettagli']:
                    note_str = f" - Note: {item['note']}" if item['note'] else ""
                    st.write(f"- **{item['nome']}**: {item['qty']} {item['udm']}{note_str}")

elif menu == "➕ Aggiungi Prodotto":
    st.subheader("➕ Aggiungi Nuovo Prodotto al Listino")
    
    with st.form("form_nuovo_prodotto"):
        fornitori_list = sorted(df_prodotti["Fornitore"].dropna().unique().tolist()) if not df_prodotti.empty else []
        fornitore_nuovo = st.selectbox("Fornitore:", fornitori_list)
        nome_nuovo = st.text_input("Nome Prodotto:")
        udm_nuovo = st.selectbox("Unità di Misura:", ["ct", "confezione", "flacone", "kg", "pezzo", "bottiglia"])
        
        submitted = st.form_submit_button("Salva Prodotto")
        if submitted and nome_nuovo:
            nuovo_id = f"P{len(df_prodotti)+1:03d}"
            nuova_riga = {
                "ID Prodotto": nuovo_id,
                "Fornitore": fornitore_nuovo,
                "Nome Prodotto": nome_nuovo,
                "Da Ordinare": "FALSO",
                "Quantità": 1,
                "Unità di Misura": udm_nuovo,
                "Note Ordine": ""
            }
            df_prodotti = pd.concat([df_prodotti, pd.DataFrame([nuova_riga])], ignore_index=True)
            save_data(df_prodotti)
            st.success(f"Prodotto '{nome_nuovo}' aggiunto con successo!")
            st.cache_data.clear()
