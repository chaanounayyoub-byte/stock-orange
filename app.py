import streamlit as st
import pandas as pd
from datetime import datetime
import io
import base64
from st_supabase_connection import SupabaseConnection, execute_query

# ---------------------------------------------------------
# CONFIGURATION DE LA PAGE & CSS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Gestion de Stock Multi-Clients",
    page_icon="📦",
    layout="wide"
)

# Masquer la barre de navigation et le footer Streamlit / GitHub
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .client-card {
        border: 2px solid #e0e0e0;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        background-color: #ffffff;
        transition: transform 0.2s, border-color 0.2s;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    .client-card:hover {
        transform: translateY(-5px);
        border-color: #ff6600;
        box-shadow: 0 8px 15px rgba(0,0,0,0.1);
    }
    .stButton > button {
        border-radius: 8px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# CONNEXION SUPABASE
# ---------------------------------------------------------
@st.cache_resource
def get_supabase_client():
    return st.connection("supabase", type=SupabaseConnection)

try:
    conn = get_supabase_client()
except Exception as e:
    st.error("Erreur de connexion à Supabase. Vérifiez la configuration des secrets.")

# Logos hébergés fiables
CLIENTS = {
    "Orange": {
        "logo": "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c8/Orange_logo.svg/512px-Orange_logo.svg.png",
        "color": "#FF6600"
    },
    "Inwi": {
        "logo": "https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/Inwi_Logo.svg/512px-Inwi_Logo.svg.png",
        "color": "#A1006B"
    },
    "ZTE": {
        "logo": "https://upload.wikimedia.org/wikipedia/commons/thumb/d/df/ZTE_logo.svg/512px-ZTE_logo.svg.png",
        "color": "#005BAC"
    }
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

if not st.session_state.user:
    st.title("🔐 Connexion - Gestion de Stock")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Se connecter", use_container_width=True)
            
            if submit:
                if username == "admin" and password == "admin123":
                    st.session_state.user = username
                    st.session_state.role = "ADMIN"
                    st.rerun()
                elif username == "magasinier" and password == "123456":
                    st.session_state.user = username
                    st.session_state.role = "MAGASINIER"
                    st.rerun()
                else:
                    st.error("Identifiants incorrects.")
    st.stop()

# ---------------------------------------------------------
# SÉLECTION CLIENT
# ---------------------------------------------------------
if not st.session_state.selected_client:
    st.title(f"👋 Bienvenue, {st.session_state.user} ({st.session_state.role})")
    st.subheader("Choisissez le client :")
    
    cols = st.columns(3)
    for idx, (client_name, info) in enumerate(CLIENTS.items()):
        with cols[idx]:
            st.markdown(f"""
                <div class="client-card">
                    <img src="{info['logo']}" height="80" style="object-fit: contain; margin-bottom: 15px;">
                    <h3 style="color: {info['color']};">{client_name}</h3>
                </div>
            """, unsafe_allow_html=True)
            st.write("")
            if st.button(f"Accéder au Stock {client_name}", key=f"btn_{client_name}", use_container_width=True):
                st.session_state.selected_client = client_name
                st.rerun()
    st.stop()

# ---------------------------------------------------------
# HEADER CLIENT
# ---------------------------------------------------------
client_info = CLIENTS[st.session_state.selected_client]
c_logo, c_title, c_user = st.columns([1, 4, 2])

with c_logo:
    st.image(client_info['logo'], width=80)
with c_title:
    st.title(f"Stock - {st.session_state.selected_client}")
with c_user:
    st.write(f"👤 **{st.session_state.user}** (`{st.session_state.role}`)")
    b1, b2 = st.columns(2)
    with b1:
        if st.button("Changer client"):
            st.session_state.selected_client = None
            st.rerun()
    with b2:
        if st.button("Déconnexion"):
            st.session_state.user = None
            st.session_state.selected_client = None
            st.rerun()

st.divider()

# ---------------------------------------------------------
# FONCTIONS BDD
# ---------------------------------------------------------
def fetch_mouvements(client):
    try:
        res = execute_query(conn.table("mouvements").select("*").eq("client", client).order("created_at", desc=True))
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def fetch_articles(client):
    try:
        res = execute_query(conn.table("articles").select("*").eq("client", client))
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def fetch_fournisseurs():
    try:
        res = execute_query(conn.table("fournisseurs").select("*"))
        return pd.DataFrame(res.data) if res.data else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def generate_n_bon(type_bon, client):
    prefix = "BE" if type_bon == "ENTREE" else "BS"
    df = fetch_mouvements(client)
    year = datetime.now().strftime("%Y")
    count = (len(df[df["type_mouvement"] == type_bon]) + 1) if not df.empty else 1
    return f"{prefix}-{year}-{client[:3].upper()}-{count:04d}"

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Données')
    return output.getvalue()

# ---------------------------------------------------------
# ONGLETS DE NAVIGATION
# ---------------------------------------------------------
tabs = [
    "📥 Bon d'Entrée (BE)", 
    "📤 Bon de Sortie (BS)", 
    "📊 État de Stock", 
    "📜 Historique & Imprimer", 
    "✏️ Modifier / Annuler Bon"
]
if st.session_state.role == "ADMIN":
    tabs.append("⚙️ Administration (Articles & Fournisseurs)")

active_tabs = st.tabs(tabs)

df_arts = fetch_articles(st.session_state.selected_client)
list_articles = df_arts["designation"].tolist() if not df_arts.empty else []

df_fourn = fetch_fournisseurs()
list_fournisseurs = df_fourn["nom"].tolist() if not df_fourn.empty else []

today_date = datetime.today().date()

# ---------------------------------------------------------
# 1. BON D'ENTRÉE (BE)
# ---------------------------------------------------------
with active_tabs[0]:
    st.subheader("📥 Créer un Bon d'Entrée (BE)")
    n_bon_be = generate_n_bon("ENTREE", st.session_state.selected_client)
    
    with st.form("form_be", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input("N° BON", value=n_bon_be, disabled=True)
            date_creation = st.date_input("Date Création", value=today_date, disabled=True)
        with c2:
            date_bl = st.date_input("Date BL (Max aujourd'hui)", value=today_date, max_value=today_date)
            n_bl = st.text_input("N° BL (Bon de Livraison)")
        with c3:
            fournisseur = st.selectbox("Fournisseur", options=["-- Sélectionner --"] + list_fournisseurs)
            article = st.selectbox("Article / Désignation", options=["-- Sélectionner --"] + list_articles)
        
        c4, c5 = st.columns(2)
        with c4:
            quantite = st.number_input("Quantité Entrée", min_value=1, step=1)
            img_file = st.file_uploader("Image du matériel (Optionnel)", type=["jpg", "jpeg", "png"])
        with c5:
            observation = st.text_area("Observation / Remarque", height=100)
            
        btn_be = st.form_submit_button("Valider le Bon d'Entrée", use_container_width=True)
        
        if btn_be:
            if article == "-- Sélectionner --" or fournisseur == "-- Sélectionner --" or not n_bl:
                st.error("Veuillez sélectionner un Article, un Fournisseur et saisir le N° BL.")
            else:
                img_b64 = base64.b64encode(img_file.read()).decode('utf-8') if img_file else None
                record = {
                    "client": st.session_state.selected_client,
                    "type_mouvement": "ENTREE",
                    "n_bon": n_bon_be,
                    "date_creation": str(date_creation),
                    "date_bl": str(date_bl),
                    "n_bl": n_bl,
                    "fournisseur": fournisseur,
                    "article": article,
                    "quantite": int(quantite),
                    "observation": observation,
                    "utilisateur": st.session_state.user,
                    "image_data": img_b64
                }
                conn.table("mouvements").insert(record).execute()
                st.success(f"Bon d'Entrée {n_bon_be} validé avec succès !")
                st.rerun()

# ---------------------------------------------------------
# 2. BON DE SORTIE (BS)
# ---------------------------------------------------------
with active_tabs[1]:
    st.subheader("📤 Créer un Bon de Sortie (BS)")
    n_bon_bs = generate_n_bon("SORTIE", st.session_state.selected_client)
    
    with st.form("form_bs", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input("N° BON", value=n_bon_bs, disabled=True)
            date_creation_bs = st.date_input("Date Création", value=today_date, disabled=True, key="dc_bs")
        with c2:
            date_bl_bs = st.date_input("Date BS/Demande (Max aujourd'hui)", value=today_date, max_value=today_date, key="dbl_bs")
            n_bl_bs = st.text_input("N° Order / Demande", key="nbl_bs")
        with c3:
            equipe_recup = st.text_input("Équipe Destinataire / Récupérateur", key="eq_bs")
            article_bs = st.selectbox("Article / Désignation", options=["-- Sélectionner --"] + list_articles, key="art_bs")
        
        c4, c5 = st.columns(2)
        with c4:
            quantite_bs = st.number_input("Quantité Sortie", min_value=1, step=1, key="q_bs")
            img_file_bs = st.file_uploader("Image du matériel (Optionnel)", type=["jpg", "jpeg", "png"], key="img_bs")
        with c5:
            observation_bs = st.text_area("Observation", height=100, key="obs_bs")
            
        btn_bs = st.form_submit_button("Valider le Bon de Sortie", use_container_width=True)
        
        if btn_bs:
            if article_bs == "-- Sélectionner --" or not equipe_recup or not n_bl_bs:
                st.error("Veuillez sélectionner l'Article, renseigner l'Équipe destinataire et le N° Order.")
            else:
                img_b64_bs = base64.b64encode(img_file_bs.read()).decode('utf-8') if img_file_bs else None
                record = {
                    "client": st.session_state.selected_client,
                    "type_mouvement": "SORTIE",
                    "n_bon": n_bon_bs,
                    "date_creation": str(date_creation_bs),
                    "date_bl": str(date_bl_bs),
                    "n_bl": n_bl_bs,
                    "equipe_recuperation": equipe_recup,
                    "article": article_bs,
                    "quantite": int(quantite_bs),
                    "observation": observation_bs,
                    "utilisateur": st.session_state.user,
                    "image_data": img_b64_bs
                }
                conn.table("mouvements").insert(record).execute()
                st.success(f"Bon de Sortie {n_bon_bs} validé avec succès !")
                st.rerun()

# ---------------------------------------------------------
# 3. ÉTAT DE STOCK
# ---------------------------------------------------------
with active_tabs[2]:
    st.subheader(f"📊 État du Stock Actuel - {st.session_state.selected_client}")
    df_mvt = fetch_mouvements(st.session_state.selected_client)
    
    if not df_mvt.empty:
        df_mvt['q_signe'] = df_mvt.apply(lambda r: r['quantite'] if r['type_mouvement'] == 'ENTREE' else -r['quantite'], axis=1)
        stock_summary = df_mvt.groupby('article').agg(
            Entrées=('quantite', lambda x: x[df_mvt.loc[x.index, 'type_mouvement'] == 'ENTREE'].sum()),
            Sorties=('quantite', lambda x: x[df_mvt.loc[x.index, 'type_mouvement'] == 'SORTIE'].sum()),
            Stock_Actuel=('q_signe', 'sum')
        ).reset_index()
        
        st.dataframe(stock_summary, use_container_width=True)
        st.download_button(
            "📥 Télécharger l'état de stock (Excel)",
            data=to_excel(stock_summary),
            file_name=f"Stock_{st.session_state.selected_client}_{today_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Aucun mouvement enregistré pour ce client.")

# ---------------------------------------------------------
# 4. HISTORIQUE & IMPRESSION
# ---------------------------------------------------------
with active_tabs[3]:
    st.subheader("📜 Historique des Bons (BE/BS)")
    df_all = fetch_mouvements(st.session_state.selected_client)
    
    if not df_all.empty:
        st.dataframe(df_all[["n_bon", "type_mouvement", "date_creation", "date_bl", "n_bl", "fournisseur", "equipe_recuperation", "article", "quantite", "utilisateur"]], use_container_width=True)
        
        st.divider()
        sel_bon = st.selectbox("Sélectionner un Bon à imprimer / consulter :", df_all["n_bon"].unique())
        
        if sel_bon:
            b_info = df_all[df_all["n_bon"] == sel_bon].iloc[0]
            st.markdown(f"### Détails du Bon : `{sel_bon}`")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.write(f"**Type :** {b_info['type_mouvement']}")
                st.write(f"**N° BL/Order :** {b_info['n_bl']}")
                st.write(f"**Date BL :** {b_info['date_bl']}")
                st.write(f"**Article :** {b_info['article']} | **Quantité :** {b_info['quantite']}")
                if b_info['type_mouvement'] == 'ENTREE':
                    st.write(f"**Fournisseur :** {b_info['fournisseur']}")
                else:
                    st.write(f"**Équipe Récupération :** {b_info['equipe_recuperation']}")
            
            with col_d2:
                st.write(f"**Créé par :** {b_info['utilisateur']} le {b_info['date_creation']}")
                st.write(f"**Observation :** {b_info['observation']}")
                if pd.notna(b_info.get('image_data')) and b_info['image_data']:
                    st.image(base64.b64decode(b_info['image_data']), caption="Photo du matériel", width=250)
            
            st.download_button(
                "📄 Télécharger le Bon (Excel)",
                data=to_excel(pd.DataFrame([b_info])),
                file_name=f"{sel_bon}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Historique vide.")

# ---------------------------------------------------------
# 5. MODIFIER / ANNULER UN BON
# ---------------------------------------------------------
with active_tabs[4]:
    st.subheader("✏️ Modifier ou Annuler un Bon")
    df_mod = fetch_mouvements(st.session_state.selected_client)
    
    if not df_mod.empty:
        bon_id_mod = st.selectbox("Sélectionnez le N° de Bon à modifier/supprimer :", df_mod["n_bon"].unique(), key="mod_sel")
        row_mod = df_mod[df_mod["n_bon"] == bon_id_mod].iloc[0]
        
        with st.form("form_edit_bon"):
            e_nbl = st.text_input("N° BL / Order :", value=str(row_mod["n_bl"]))
            e_qty = st.number_input("Quantité :", min_value=1, value=int(row_mod["quantite"]))
            e_obs = st.text_area("Observation :", value=str(row_mod["observation"] or ""))
            
            c_ed1, c_ed2 = st.columns(2)
            with c_ed1:
                btn_update = st.form_submit_button("Enregistrer les Modifications", use_container_width=True)
            with c_ed2:
                btn_delete = st.form_submit_button("🗑️ Supprimer ce Bon", use_container_width=True)
                
            if btn_update:
                conn.table("mouvements").update({
                    "n_bl": e_nbl,
                    "quantite": int(e_qty),
                    "observation": e_obs
                }).eq("id", int(row_mod["id"])).execute()
                st.success("Bon mis à jour avec succès !")
                st.rerun()
                
            if btn_delete:
                conn.table("mouvements").delete().eq("id", int(row_mod["id"])).execute()
                st.warning("Bon supprimé avec succès !")
                st.rerun()

# ---------------------------------------------------------
# 6. ADMINISTRATION (ADMIN ONLY)
# ---------------------------------------------------------
if st.session_state.role == "ADMIN" and len(active_tabs) > 5:
    with active_tabs[5]:
        st.subheader("⚙️ Administration des Référentiels")
        
        tab_adm1, tab_adm2 = st.tabs(["📦 Gestion des Articles", "🏭 Gestion des Fournisseurs"])
        
        with tab_adm1:
            st.write(f"**Articles enregistrés pour {st.session_state.selected_client} :**")
            if not df_arts.empty:
                st.dataframe(df_arts[["designation", "description"]], use_container_width=True)
                
                # Option de suppression d'article
                art_to_del = st.selectbox("Supprimer un article :", options=["-- Choisir --"] + df_arts["designation"].tolist(), key="del_art_sel")
                if st.button("🗑️ Supprimer l'article sélectionné") and art_to_del != "-- Choisir --":
                    conn.table("articles").delete().eq("client", st.session_state.selected_client).eq("designation", art_to_del).execute()
                    st.success("Article supprimé !")
                    st.rerun()
            else:
                st.write("Aucun article disponible pour le moment.")
            
            st.divider()
            st.markdown("#### Ajouter un nouvel article")
            new_art_des = st.text_input("Désignation de l'article")
            new_art_desc = st.text_input("Description / Code")
            if st.button("Ajouter l'article"):
                if new_art_des:
                    conn.table("articles").insert({
                        "client": st.session_state.selected_client,
                        "designation": new_art_des,
                        "description": new_art_desc
                    }).execute()
                    st.success("Article ajouté avec succès !")
                    st.rerun()
                else:
                    st.error("Veuillez renseigner au moins la désignation de l'article.")
        
        with tab_adm2:
            st.write("**Fournisseurs enregistrés :**")
            if not df_fourn.empty:
                st.dataframe(df_fourn[["nom", "contact"]], use_container_width=True)
                
                fourn_to_del = st.selectbox("Supprimer un fournisseur :", options=["-- Choisir --"] + df_fourn["nom"].tolist(), key="del_fourn_sel")
                if st.button("🗑️ Supprimer le fournisseur sélectionné") and fourn_to_del != "-- Choisir --":
                    conn.table("fournisseurs").delete().eq("nom", fourn_to_del).execute()
                    st.success("Fournisseur supprimé !")
                    st.rerun()
            else:
                st.write("Aucun fournisseur disponible.")
            
            st.divider()
            st.markdown("#### Ajouter un nouveau fournisseur")
            new_f_nom = st.text_input("Nom du fournisseur")
            new_f_contact = st.text_input("Contact / Email")
            if st.button("Ajouter le fournisseur"):
                if new_f_nom:
                    conn.table("fournisseurs").insert({
                        "nom": new_f_nom,
                        "contact": new_f_contact
                    }).execute()
                    st.success("Fournisseur ajouté avec succès !")
                    st.rerun()
                else:
                    st.error("Veuillez remplir le nom du fournisseur.")
