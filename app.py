import streamlit as st
import pandas as pd
from datetime import datetime
import io
from st_supabase_connection import SupabaseConnection
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# Import pour la génération de PDF avec ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ---------------------------------------------------------
# CONFIGURATION DE LA PAGE
# ---------------------------------------------------------
st.set_page_config(
    page_title="NOMATIS - Gestion de Stock Multi-Clients",
    page_icon="📦",
    layout="wide"
)

# ---------------------------------------------------------
# CONNEXION SUPABASE
# ---------------------------------------------------------
@st.cache_resource
def get_supabase_conn():
    return st.connection("supabase", type=SupabaseConnection)

try:
    conn = get_supabase_conn()
    supabase = conn.client
except Exception:
    st.error("Erreur de connexion à Supabase. Vérifiez vos identifiants.")

# ---------------------------------------------------------
# CLIENTS & THÈMES
# ---------------------------------------------------------
CLIENTS = {
    "Orange": {"color": "#FF6600"},
    "Inwi": {"color": "#A1006B"},
    "ZTE": {"color": "#005BAC"}
}

# ---------------------------------------------------------
# SESSIONS ET AUTHENTIFICATION
# ---------------------------------------------------------
if "user" not in st.session_state:
    st.session_state.user = None
if "role" not in st.session_state:
    st.session_state.role = "MAGASINIER"
if "selected_client" not in st.session_state:
    st.session_state.selected_client = None

active_color = CLIENTS[st.session_state.selected_client]["color"] if st.session_state.selected_client else "#0284C7"

st.markdown(f"""
    <style>
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header {{visibility: hidden;}}
    .stButton > button {{
        background-color: {active_color} !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
        font-weight: bold !important;
    }}
    .client-card {{
        border: 2px solid #e0e0e0;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        background-color: #ffffff;
    }}
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# BDD UTILS
# ---------------------------------------------------------
def fetch_users():
    try:
        res = supabase.table("utilisateurs").select("*").execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def fetch_mouvements(client):
    try:
        res = supabase.table("mouvements").select("*").eq("client", client).order("created_at", desc=True).execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def fetch_articles(client):
    try:
        res = supabase.table("articles").select("*").eq("client", client).execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def fetch_fournisseurs():
    try:
        res = supabase.table("fournisseurs").select("*").execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def generate_n_bon(type_bon, client):
    prefix = "BE" if type_bon == "ENTREE" else "BS"
    df = fetch_mouvements(client)
    year = datetime.now().strftime("%Y")
    count = (len(df[df["type_mouvement"] == type_bon]) + 1) if not df.empty else 1
    return f"{prefix}-{year}-{client[:3].upper()}-{count:04d}"

# ---------------------------------------------------------
# LOGIN PAGE
# ---------------------------------------------------------
if not st.session_state.user:
    st.title("🔐 NOMATIS - Connexion Gestion de Stock")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Se connecter", use_container_width=True)
            if submit:
                df_u = fetch_users()
                if not df_u.empty:
                    match = df_u[(df_u["username"] == username) & (df_u["password"] == password)]
                    if not match.empty:
                        st.session_state.user = username
                        st.session_state.role = match.iloc[0]["role"]
                        st.rerun()
                    else:
                        st.error("Identifiants incorrects.")
                else:
                    if username == "admin" and password == "admin123":
                        st.session_state.user = "admin"
                        st.session_state.role = "ADMIN"
                        st.rerun()
                    else:
                        st.error("Identifiants incorrects.")
    st.stop()

# ---------------------------------------------------------
# SÉLECTION CLIENT
# ---------------------------------------------------------
if not st.session_state.selected_client:
    st.title(f"👋 Bienvenue, {st.session_state.user}")
    st.subheader("Sélectionnez l'espace Client :")
    cols = st.columns(3)
    for idx, (client_name, info) in enumerate(CLIENTS.items()):
        with cols[idx]:
            st.markdown(f"""
                <div class="client-card">
                    <h2 style="color: {info['color']};">{client_name}</h2>
                </div>
            """, unsafe_allow_html=True)
            st.write("")
            if st.button(f"Accéder au Stock {client_name}", key=f"btn_{client_name}", use_container_width=True):
                st.session_state.selected_client = client_name
                st.rerun()
    st.stop()

# ---------------------------------------------------------
# EN-TÊTE APPLI
# ---------------------------------------------------------
c_head1, c_head2, c_head3 = st.columns([2, 4, 2])
with c_head1:
    st.markdown("### 🏢 **NOMATIS**")
    st.caption("Gestionnaire de Stock")
with c_head2:
    st.markdown(f"<h2 style='color:{active_color}; text-align:center;'>ESPACE STOCK : {st.session_state.selected_client}</h2>", unsafe_allow_html=True)
with c_head3:
    st.write(f"👤 **{st.session_state.user}** (`{st.session_state.role}`)")
    b1, b2 = st.columns(2)
    with b1:
        if st.button("Changer Client"):
            st.session_state.selected_client = None
            st.rerun()
    with b2:
        if st.button("Déconnexion"):
            st.session_state.user = None
            st.session_state.selected_client = None
            st.rerun()

st.divider()

# ---------------------------------------------------------
# GENERATION EXCEL
# ---------------------------------------------------------
def export_excel_modele(type_bon, n_bon, date_bl, n_bl, tiers_nom, lieu_livraison, receptionne_par, client, articles_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = type_bon

    font_bold = Font(name="Arial", size=10, bold=True)
    font_title = Font(name="Arial", size=16, bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal="center", vertical="center")

    ws["A1"] = "NOMATIS"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = "NOMATIS"
    ws["A3"] = "32 Rue Al Hatim"
    ws["A4"] = "les Orangers"
    ws["A5"] = "10000"

    ws["F1"] = f"Client : {client}"
    ws["F1"].font = font_bold
    ws["F1"].alignment = Alignment(horizontal="right")

    ws.merge_cells("A7:F7")
    ws["A7"] = "Bon d'entrée" if type_bon == "ENTREE" else "Bon de sortie"
    ws["A7"].font = font_title
    ws["A7"].alignment = center_align

    col_tiers = "Fournisseur" if type_bon == "ENTREE" else "Demandeur/Équipe"
    col_bl = "Bon de Livraison" if type_bon == "ENTREE" else "N° Order / Ordre"

    headers_top = [col_bl, "Date", col_tiers, "Lieu de livraison", "réceptionné par", "Stock"]
    for col_num, h_text in enumerate(headers_top, 1):
        cell = ws.cell(row=9, column=col_num, value=h_text)
        cell.font = font_bold
        cell.alignment = center_align
        cell.border = thin_border

    values_top = [n_bl, str(date_bl), tiers_nom, lieu_livraison, receptionne_par, client]
    for col_num, val in enumerate(values_top, 1):
        cell = ws.cell(row=10, column=col_num, value=val)
        cell.alignment = center_align
        cell.border = thin_border

    ws["A12"] = "Référence"
    ws["A12"].font = font_bold
    ws["A12"].alignment = center_align
    ws["A12"].border = thin_border
    
    ws.merge_cells("B12:E12")
    ws["B12"] = "Désignation"
    ws["B12"].font = font_bold
    ws["B12"].alignment = center_align
    for c in range(2, 6):
        ws.cell(row=12, column=c).border = thin_border

    ws["F12"] = "Qté"
    ws["F12"].font = font_bold
    ws["F12"].alignment = center_align
    ws["F12"].border = thin_border

    current_row = 13
    for _, row in articles_df.iterrows():
        ws.cell(row=current_row, column=1, value=str(row.get("Référence", ""))).border = thin_border
        
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
        ws.cell(row=current_row, column=2, value=str(row.get("Désignation", ""))).border = thin_border
        for c in range(2, 6):
            ws.cell(row=current_row, column=c).border = thin_border
            
        ws.cell(row=current_row, column=6, value=int(row.get("Quantité", 1))).border = thin_border
        ws.cell(row=current_row, column=6).alignment = center_align
        current_row += 1

    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 15

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# GENERATION PDF
# ---------------------------------------------------------
def export_pdf_modele(type_bon, n_bon, date_bl, n_bl, tiers_nom, lieu_livraison, receptionne_par, client, articles_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, alignment=1, spaceAfter=15)
    normal_style = styles['Normal']
    
    header_data = [
        [Paragraph("<b>NOMATIS</b><br/>32 Rue Al Hatim<br/>les Orangers<br/>10000", normal_style),
         Paragraph(f"<b>Client : {client}</b><br/><b>{n_bon}</b>", ParagraphStyle('Right', parent=normal_style, alignment=2))]
    ]
    t_header = Table(header_data, colWidths=[250, 255])
    elements.append(t_header)
    elements.append(Spacer(1, 15))
    
    title_text = "Bon d'entrée" if type_bon == "ENTREE" else "Bon de sortie"
    elements.append(Paragraph(f"<b>{title_text}</b>", title_style))
    elements.append(Spacer(1, 10))
    
    col_tiers = "Fournisseur" if type_bon == "ENTREE" else "Demandeur/Équipe"
    col_bl = "Bon de Livraison" if type_bon == "ENTREE" else "N° Order / Ordre"
    
    meta_headers = [col_bl, "Date", col_tiers, "Lieu de livraison", "Réceptionné par", "Stock"]
    meta_values = [str(n_bl), str(date_bl), str(tiers_nom), str(lieu_livraison), str(receptionne_par), str(client)]
    
    t_meta = Table([meta_headers, meta_values], colWidths=[85, 70, 90, 110, 90, 60])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F0F0F0')),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.black),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 8),
    ]))
    elements.append(t_meta)
    elements.append(Spacer(1, 20))
    
    art_headers = ["Référence", "Désignation", "Qté"]
    art_data = [art_headers]
    for _, row in articles_df.iterrows():
        art_data.append([
            str(row.get("Référence", "")),
            str(row.get("Désignation", "")),
            str(row.get("Quantité", 1))
        ])
        
    t_art = Table(art_data, colWidths=[120, 315, 70])
    t_art.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E0E0E0')),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('ALIGN', (2,0), (2,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
    ]))
    elements.append(t_art)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------------------------------------
# INITIALISATION DICT DÉSIGNATION -> RÉFÉRENCE
# ---------------------------------------------------------
df_arts = fetch_articles(st.session_state.selected_client)
if not df_arts.empty:
    art_map = dict(zip(df_arts["designation"], df_arts["reference"]))
    list_articles = list(art_map.keys())
else:
    art_map = {"Câble RJ45": "REF-RJ45", "Câble IF": "REF-IF"}
    list_articles = list(art_map.keys())

df_fourn = fetch_fournisseurs()
list_fournisseurs = df_fourn["nom"].tolist() if not df_fourn.empty else ["NOMATIS"]

today_date = datetime.today().date()

# ---------------------------------------------------------
# ONGLETS PRINCIPAUX
# ---------------------------------------------------------
tab_names = ["📥 Bon d'Entrée (BE)", "📤 Bon de Sortie (BS)", "📊 État du Stock", "📜 Historique", "✏️ Modifier un Bon"]
if st.session_state.role == "ADMIN":
    tab_names.append("⚙️ Administration (Config)")

tabs = st.tabs(tab_names)

# ---------------------------------------------------------
# 1. BON D'ENTRÉE (BE)
# ---------------------------------------------------------
with tabs[0]:
    st.subheader("📥 Créer un Bon d'Entrée (BE)")
    n_bon_be = generate_n_bon("ENTREE", st.session_state.selected_client)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.text_input("N° BON BE", value=n_bon_be, disabled=True)
        date_be = st.date_input("Date", value=today_date, max_value=today_date, key="d_be")
        n_bl_be = st.text_input("Bon de Livraison (N° BL)", key="nbl_be")

    with col2:
        fournisseur_be = st.selectbox("Fournisseur", options=list_fournisseurs, key="f_be")
        lieu_livraison_be = st.text_input("Lieu de livraison", value="Dépôt NOMATIS", key="ll_be")

    with col3:
        receptionne_par_be = st.text_input("Réceptionné par", value=st.session_state.user, key="rec_be")
        stock_be = st.text_input("Stock", value=st.session_state.selected_client, disabled=True, key="s_be")

    st.markdown("#### 📦 Articles répertoriés :")
    
    df_be_input = st.data_editor(
        pd.DataFrame([{"Désignation": list_articles[0], "Quantité": 1}]),
        num_rows="dynamic",
        column_config={
            "Désignation": st.column_config.SelectboxColumn("Désignation", options=list_articles, required=True),
            "Quantité": st.column_config.NumberColumn("Qté", min_value=1, step=1, default=1, required=True)
        },
        use_container_width=True,
        key="be_editor"
    )

    df_be_input["Référence"] = df_be_input["Désignation"].map(lambda x: art_map.get(x, "N/A"))
    
    st.caption("Aperçu du tableau final avec références automatiques :")
    st.dataframe(df_be_input[["Référence", "Désignation", "Quantité"]], use_container_width=True)

    format_be = st.radio("Format de document souhaité :", ["Excel (.xlsx)", "PDF (.pdf)"], horizontal=True, key="fmt_be")

    c_btn1, c_btn2 = st.columns(2)
    
    with c_btn1:
        valid_be = st.button("💾 Enregistrer le Bon d'Entrée dans la BDD", use_container_width=True)

    if valid_be:
        if not n_bl_be or df_be_input.empty:
            st.error("Veuillez renseigner le N° BL et ajouter des articles.")
        else:
            for _, row in df_be_input.iterrows():
                supabase.table("mouvements").insert({
                    "client": st.session_state.selected_client,
                    "type_mouvement": "ENTREE",
                    "n_bon": n_bon_be,
                    "date_creation": str(date_be),
                    "date_bl": str(date_be),
                    "n_bl": n_bl_be,
                    "fournisseur": fournisseur_be,
                    "destination": lieu_livraison_be,
                    "equipe_recuperation": receptionne_par_be,
                    "article": row["Désignation"],
                    "quantite": int(row["Quantité"]),
                    "utilisateur": st.session_state.user
                }).execute()
            st.session_state["be_saved"] = True
            st.success(f"Bon d'Entrée {n_bon_be} enregistré en base de données avec succès !")

    with c_btn2:
        if st.session_state.get("be_saved", False):
            if "Excel" in format_be:
                file_bytes_be = export_excel_modele(
                    type_bon="ENTREE",
                    n_bon=n_bon_be,
                    date_bl=date_be,
                    n_bl=n_bl_be,
                    tiers_nom=fournisseur_be,
                    lieu_livraison=lieu_livraison_be,
                    receptionne_par=receptionne_par_be,
                    client=st.session_state.selected_client,
                    articles_df=df_be_input
                )
                file_ext = "xlsx"
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                file_bytes_be = export_pdf_modele(
                    type_bon="ENTREE",
                    n_bon=n_bon_be,
                    date_bl=date_be,
                    n_bl=n_bl_be,
                    tiers_nom=fournisseur_be,
                    lieu_livraison=lieu_livraison_be,
                    receptionne_par=receptionne_par_be,
                    client=st.session_state.selected_client,
                    articles_df=df_be_input
                )
                file_ext = "pdf"
                mime_type = "application/pdf"

            st.download_button(
                f"🖨️ Télécharger le Bon BE en {format_be}",
                data=file_bytes_be,
                file_name=f"{n_bon_be}.{file_ext}",
                mime=mime_type,
                use_container_width=True
            )

# ---------------------------------------------------------
# 2. BON DE SORTIE (BS)
# ---------------------------------------------------------
with tabs[1]:
    st.subheader("📤 Créer un Bon de Sortie (BS)")
    n_bon_bs = generate_n_bon("SORTIE", st.session_state.selected_client)
    
    col_bs1, col_bs2, col_bs3 = st.columns(3)
    with col_bs1:
        st.text_input("N° BON BS", value=n_bon_bs, disabled=True)
        date_bs = st.date_input("Date", value=today_date, max_value=today_date, key="d_bs")
        n_bl_bs = st.text_input("N° Order / Ordre", key="nbl_bs")

    with col_bs2:
        equipe_bs = st.text_input("Demandeur / Équipe", value="Équipe Technique", key="eq_bs")
        destination_bs = st.text_input("Lieu de livraison / Destination", key="dest_bs")

    with col_bs3:
        receptionne_par_bs = st.text_input("Réceptionné par (Nom du récepteur)", key="rec_bs")
        stock_bs = st.text_input("Stock", value=st.session_state.selected_client, disabled=True, key="s_bs")

    st.markdown("#### 📦 Articles à sortir :")
    df_bs_input = st.data_editor(
        pd.DataFrame([{"Désignation": list_articles[0], "Quantité": 1}]),
        num_rows="dynamic",
        column_config={
            "Désignation": st.column_config.SelectboxColumn("Désignation", options=list_articles, required=True),
            "Quantité": st.column_config.NumberColumn("Qté", min_value=1, step=1, default=1, required=True)
        },
        use_container_width=True,
        key="bs_editor"
    )

    df_bs_input["Référence"] = df_bs_input["Désignation"].map(lambda x: art_map.get(x, "N/A"))

    st.caption("Aperçu du tableau final avec références automatiques :")
    st.dataframe(df_bs_input[["Référence", "Désignation", "Quantité"]], use_container_width=True)

    format_bs = st.radio("Format de document souhaité :", ["Excel (.xlsx)", "PDF (.pdf)"], horizontal=True, key="fmt_bs")

    c_bs_btn1, c_bs_btn2 = st.columns(2)

    with c_bs_btn1:
        valid_bs = st.button("💾 Enregistrer le Bon de Sortie dans la BDD", use_container_width=True)

    if valid_bs:
        if not n_bl_bs or not destination_bs or not receptionne_par_bs or df_bs_input.empty:
            st.error("Veuillez remplir le N° Order, le Lieu de livraison et le Récepteur.")
        else:
            for _, row in df_bs_input.iterrows():
                supabase.table("mouvements").insert({
                    "client": st.session_state.selected_client,
                    "type_mouvement": "SORTIE",
                    "n_bon": n_bon_bs,
                    "date_creation": str(date_bs),
                    "date_bl": str(date_bs),
                    "n_bl": n_bl_bs,
                    "equipe_recuperation": equipe_bs,
                    "destination": destination_bs,
                    "article": row["Désignation"],
                    "quantite": int(row["Quantité"]),
                    "utilisateur": st.session_state.user
                }).execute()
            st.session_state["bs_saved"] = True
            st.success(f"Bon de Sortie {n_bon_bs} enregistré en base de données avec succès !")

    with c_bs_btn2:
        if st.session_state.get("bs_saved", False):
            if "Excel" in format_bs:
                file_bytes_bs = export_excel_modele(
                    type_bon="SORTIE",
                    n_bon=n_bon_bs,
                    date_bl=date_bs,
                    n_bl=n_bl_bs,
                    tiers_nom=equipe_bs,
                    lieu_livraison=destination_bs,
                    receptionne_par=receptionne_par_bs,
                    client=st.session_state.selected_client,
                    articles_df=df_bs_input
                )
                file_ext = "xlsx"
                mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                file_bytes_bs = export_pdf_modele(
                    type_bon="SORTIE",
                    n_bon=n_bon_bs,
                    date_bl=date_bs,
                    n_bl=n_bl_bs,
                    tiers_nom=equipe_bs,
                    lieu_livraison=destination_bs,
                    receptionne_par=receptionne_par_bs,
                    client=st.session_state.selected_client,
                    articles_df=df_bs_input
                )
                file_ext = "pdf"
                mime_type = "application/pdf"

            st.download_button(
                f"🖨️ Télécharger le Bon BS en {format_bs}",
                data=file_bytes_bs,
                file_name=f"{n_bon_bs}.{file_ext}",
                mime=mime_type,
                use_container_width=True
            )

# ---------------------------------------------------------
# 3. ÉTAT DU STOCK
# ---------------------------------------------------------
with tabs[2]:
    st.subheader(f"📊 État du Stock - Client {st.session_state.selected_client}")
    df_mvt = fetch_mouvements(st.session_state.selected_client)
    if not df_mvt.empty:
        df_mvt['q_signe'] = df_mvt.apply(lambda r: r['quantite'] if r['type_mouvement'] == 'ENTREE' else -r['quantite'], axis=1)
        stock_summary = df_mvt.groupby('article').agg(
            Entrees=('quantite', lambda x: x[df_mvt.loc[x.index, 'type_mouvement'] == 'ENTREE'].sum()),
            Sorties=('quantite', lambda x: x[df_mvt.loc[x.index, 'type_mouvement'] == 'SORTIE'].sum()),
            Stock_Disponible=('q_signe', 'sum')
        ).reset_index()
        st.dataframe(stock_summary, use_container_width=True)
    else:
        st.info("Aucun mouvement enregistré.")

# ---------------------------------------------------------
# 4. HISTORIQUE
# ---------------------------------------------------------
with tabs[3]:
    st.subheader("📜 Historique des Bons")
    df_all = fetch_mouvements(st.session_state.selected_client)
    if not df_all.empty:
        st.dataframe(df_all, use_container_width=True)

# ---------------------------------------------------------
# 5. MODIFICATION DE BON
# ---------------------------------------------------------
with tabs[4]:
    st.subheader("✏️ Modifier un Bon")
    df_mod = fetch_mouvements(st.session_state.selected_client)
    if not df_mod.empty:
        bon_id_mod = st.selectbox("Choisir le Bon à modifier :", df_mod["n_bon"].unique())
        item_mod = df_mod[df_mod["n_bon"] == bon_id_mod].iloc[0]
        
        with st.form("form_edit"):
            e_nbl = st.text_input("N° BL / Order", value=str(item_mod["n_bl"]))
            e_qty = st.number_input("Quantité", value=int(item_mod["quantite"]), min_value=1)
            if st.form_submit_button("Sauvegarder"):
                supabase.table("mouvements").update({"n_bl": e_nbl, "quantite": e_qty}).eq("id", int(item_mod["id"])).execute()
                st.success("Bon mis à jour !")
                st.rerun()

# ---------------------------------------------------------
# 6. ADMINISTRATION (CONFIG - RÉSERVÉ ADMIN)
# ---------------------------------------------------------
if st.session_state.role == "ADMIN":
    with tabs[5]:
        st.title("⚙️ Espace Administration & Configuration")
        
        sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📦 Configuration Articles", "🏭 Fournisseurs", "👥 Gestion Utilisateurs"])

        with sub_tab1:
            st.subheader(f"Ajouter un Article pour le Client : {st.session_state.selected_client}")
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                new_ref = st.text_input("Référence Article (ex: REF-001)")
            with col_a2:
                new_des = st.text_input("Désignation Article (ex: Câble Fibre 50m)")

            if st.button("➕ Ajouter l'Article"):
                if new_ref and new_des:
                    supabase.table("articles").insert({
                        "client": st.session_state.selected_client,
                        "reference": new_ref,
                        "designation": new_des
                    }).execute()
                    st.success("Article ajouté avec succès !")
                    st.rerun()
                else:
                    st.error("Veuillez remplir les deux champs.")

            st.divider()
            st.markdown("#### Articles Existants :")
            df_cur_arts = fetch_articles(st.session_state.selected_client)
            if not df_cur_arts.empty:
                st.dataframe(df_cur_arts[["reference", "designation"]], use_container_width=True)
                art_to_del = st.selectbox("Supprimer un article :", df_cur_arts["designation"].tolist())
                if st.button("❌ Supprimer Article"):
                    supabase.table("articles").delete().eq("designation", art_to_del).eq("client", st.session_state.selected_client).execute()
                    st.success("Article supprimé !")
                    st.rerun()

        with sub_tab2:
            st.subheader("Ajouter un Fournisseur")
            new_fourn = st.text_input("Nom du Fournisseur")
            if st.button("➕ Ajouter Fournisseur"):
                if new_fourn:
                    supabase.table("fournisseurs").insert({"nom": new_fourn}).execute()
                    st.success("Fournisseur ajouté !")
                    st.rerun()

            st.divider()
            df_f = fetch_fournisseurs()
            if not df_f.empty:
                st.dataframe(df_f[["nom"]], use_container_width=True)
                f_to_del = st.selectbox("Supprimer un fournisseur :", df_f["nom"].tolist())
                if st.button("❌ Supprimer Fournisseur"):
                    supabase.table("fournisseurs").delete().eq("nom", f_to_del).execute()
                    st.success("Fournisseur supprimé !")
                    st.rerun()

        with sub_tab3:
            st.subheader("Gestion des Utilisateurs")
            u_col1, u_col2, u_col3 = st.columns(3)
            with u_col1:
                new_u = st.text_input("Nom Utilisateur")
            with u_col2:
                new_p = st.text_input("Mot de Passe", type="password")
            with u_col3:
                new_r = st.selectbox("Rôle", ["MAGASINIER", "ADMIN"])

            if st.button("➕ Ajouter / Mettre à jour Utilisateur"):
                if new_u and new_p:
                    supabase.table("utilisateurs").upsert({"username": new_u, "password": new_p, "role": new_r}, on_conflict="username").execute()
                    st.success(f"Utilisateur {new_u} créé ou mis à jour !")
                    st.rerun()

            st.divider()
            df_users_list = fetch_users()
            if not df_users_list.empty:
                st.dataframe(df_users_list[["username", "role"]], use_container_width=True)
                u_to_del = st.selectbox("Supprimer un utilisateur :", df_users_list["username"].tolist())
                if st.button("❌ Supprimer Utilisateur"):
                    if u_to_del != "admin":
                        supabase.table("utilisateurs").delete().eq("username", u_to_del).execute()
                        st.success("Utilisateur supprimé !")
                        st.rerun()
                    else:
                        st.error("Impossible de supprimer le compte Admin principal.")
