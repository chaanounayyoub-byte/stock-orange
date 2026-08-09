import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
from st_supabase_connection import SupabaseConnection
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

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
    st.error("Erreur de connexion à Supabase. Vérifiez la configuration de vos secrets.")

# ---------------------------------------------------------
# CLIENTS & THÈMES
# ---------------------------------------------------------
CLIENTS = {
    "Orange": {"color": "#FF6600", "bg_light": "#FFF5EF"},
    "Inwi": {"color": "#A1006B", "bg_light": "#FDF2F9"},
    "ZTE": {"color": "#005BAC", "bg_light": "#F0F7FF"}
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
                if username == "admin" and password == "admin123":
                    st.session_state.user = "admin"
                    st.session_state.role = "ADMIN"
                    st.rerun()
                elif username == "magasinier" and password == "123456":
                    st.session_state.user = "magasinier"
                    st.session_state.role = "MAGASINIER"
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
# BDD UTILS
# ---------------------------------------------------------
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
# GÉNÉRATION EXCEL UNIFIÉE POUR BE ET BS (MODÈLE EXACT)
# ---------------------------------------------------------
def export_excel_modele(type_bon, n_bon, date_bl, n_bl, tiers_nom, lieu_livraison, receptionne_par, client, articles_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = type_bon

    # Styles
    font_bold = Font(name="Arial", size=10, bold=True)
    font_title = Font(name="Arial", size=16, bold=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal="center", vertical="center")

    # En-tête Nomatis
    ws["A1"] = "NOMATIS"
    ws["A1"].font = Font(name="Arial", size=14, bold=True)
    ws["A2"] = "NOMATIS"
    ws["A3"] = "32 Rue Al Hatim"
    ws["A4"] = "les Orangers"
    ws["A5"] = "10000"

    # Logo / Identifiant Client
    ws["F1"] = f"Client : {client}"
    ws["F1"].font = font_bold
    ws["F1"].alignment = Alignment(horizontal="right")

    # Titre du Bon
    ws.merge_cells("A7:F7")
    ws["A7"] = "Bon d'entrée" if type_bon == "ENTREE" else "Bon de sortie"
    ws["A7"].font = font_title
    ws["A7"].alignment = center_align

    # Libellé dynamique de la colonne (Fournisseur pour BE, Demandeur pour BS)
    col_tiers = "Fournisseur" if type_bon == "ENTREE" else "Demandeur/Équipe"
    col_bl = "Bon de Livraison" if type_bon == "ENTREE" else "N° Order / Ordre"

    # En-tête du tableau d'information
    headers_top = [col_bl, "Date", col_tiers, "Lieu de livraison", "réceptionné par", "Stock"]
    for col_num, h_text in enumerate(headers_top, 1):
        cell = ws.cell(row=9, column=col_num)
        cell.value = h_text
        cell.font = font_bold
        cell.alignment = center_align
        cell.border = thin_border

    values_top = [n_bl, str(date_bl), tiers_nom, lieu_livraison, receptionne_par, client]
    for col_num, val in enumerate(values_top, 1):
        cell = ws.cell(row=10, column=col_num)
        cell.value = val
        cell.alignment = center_align
        cell.border = thin_border

    # En-tête du tableau des articles
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

    # Remplissage des articles
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

    # Format de largeur des colonnes
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
# ONGLETS
# ---------------------------------------------------------
tabs = st.tabs([
    "📥 Bon d'Entrée (BE)", 
    "📤 Bon de Sortie (BS)", 
    "📊 État du Stock", 
    "📜 Historique & Imprimer", 
    "✏️ Modifier un Bon",
    "👤 Mon Profil"
])

df_arts = fetch_articles(st.session_state.selected_client)
list_articles = df_arts["designation"].tolist() if not df_arts.empty else ["Article Exemple"]

df_fourn = fetch_fournisseurs()
list_fournisseurs = df_fourn["nom"].tolist() if not df_fourn.empty else ["Fournisseur 1"]

today_date = datetime.today().date()

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
        pd.DataFrame([{"Référence": "REF-001", "Désignation": list_articles[0], "Quantité": 1}]),
        num_rows="dynamic",
        column_config={
            "Référence": st.column_config.TextColumn("Référence", required=True),
            "Désignation": st.column_config.SelectboxColumn("Désignation", options=list_articles, required=True),
            "Quantité": st.column_config.NumberColumn("Qté", min_value=1, step=1, default=1, required=True)
        },
        use_container_width=True,
        key="be_editor"
    )

    if st.button("✅ Valider & Imprimer le Bon d'Entrée", use_container_width=True):
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
            
            st.success(f"Bon d'Entrée {n_bon_be} enregistré !")
            
            excel_bytes = export_excel_modele(
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

            st.download_button(
                "🖨️ Télécharger le Bon d'Entrée (Excel)",
                data=excel_bytes,
                file_name=f"{n_bon_be}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ---------------------------------------------------------
# 2. BON DE SORTIE (BS) - IDENTIQUE AU BE
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
        pd.DataFrame([{"Référence": "REF-001", "Désignation": list_articles[0], "Quantité": 1}]),
        num_rows="dynamic",
        column_config={
            "Référence": st.column_config.TextColumn("Référence", required=True),
            "Désignation": st.column_config.SelectboxColumn("Désignation", options=list_articles, required=True),
            "Quantité": st.column_config.NumberColumn("Qté", min_value=1, step=1, default=1, required=True)
        },
        use_container_width=True,
        key="bs_editor"
    )

    if st.button("✅ Valider & Imprimer le Bon de Sortie", use_container_width=True):
        if not n_bl_bs or not destination_bs or not receptionne_par_bs or df_bs_input.empty:
            st.error("Veuillez renseigner le N° Order, le Lieu de livraison et le Récepteur.")
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
            
            st.success(f"Bon de Sortie {n_bon_bs} enregistré !")

            excel_bytes_bs = export_excel_modele(
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

            st.download_button(
                "🖨️ Télécharger le Bon de Sortie (Excel)",
                data=excel_bytes_bs,
                file_name=f"{n_bon_bs}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
# 6. MON PROFIL
# ---------------------------------------------------------
with tabs[5]:
    st.subheader("👤 Mon Profil")
    st.write(f"Utilisateur : **{st.session_state.user}**")
