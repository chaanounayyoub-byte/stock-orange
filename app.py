import streamlit as st
import pandas as pd
import datetime
import json
import os
from fpdf import FPDF
from PIL import Image

# ==========================================
# CONFIGURATION DE LA PAGE
# ==========================================
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS", 
    page_icon="📦",
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Logos locaux
LOGOS = {
    "NOMATIS": "Logo Nomatis.jpg",
    "ORANGE": "Orange_logo.svg.webp",
    "INWI": "Logo INWI.jpg",
    "ZTE": "Logo ZTE.jpg"
}

def load_image(path):
    if os.path.exists(path):
        try:
            return Image.open(path)
        except Exception:
            return None
    return None

# ==========================================
# DESIGN ET THÈME AMÉLIORÉ
# ==========================================
def apply_advanced_theme():
    st.markdown("""
    <style>
        /* Fond global moderne */
        .stApp { background-color: #F8FAFC; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        
        /* Titres principaux */
        h1 { color: #0B4F6C !important; font-weight: 800; letter-spacing: -0.5px; }
        h2, h3 { color: #1E293B !important; font-weight: 700; }
        
        /* Style du formulaire de connexion */
        div[data-testid="stForm"] {
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 32px;
            background-color: #FFFFFF;
            box-shadow: 0px 10px 25px -5px rgba(0, 0, 0, 0.05);
        }
        
        /* Bouton dynamique de connexion */
        .btn-login-rouge button { 
            background-color: #EF4444 !important; 
            color: white !important; 
            font-weight: 700 !important; 
            border-radius: 8px !important;
            border: none !important;
            width: 100% !important;
            padding: 10px !important;
            transition: all 0.3s ease;
        }
        .btn-login-vert button { 
            background-color: #10B981 !important; 
            color: white !important; 
            font-weight: 700 !important; 
            border-radius: 8px !important;
            border: none !important;
            width: 100% !important;
            padding: 10px !important;
            transition: all 0.3s ease;
        }
        
        /* Cartes Opérateurs (Choix Client) */
        .client-card {
            background-color: #FFFFFF;
            border-radius: 12px;
            padding: 24px;
            text-align: center;
            border: 1px solid #E2E8F0;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            margin-bottom: 15px;
        }
        .client-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        }
        .card-orange { border-top: 5px solid #FF7900; }
        .card-inwi { border-top: 5px solid #A1006B; }
        .card-zte { border-top: 5px solid #005BAC; }
        
        /* Boutons Opérateurs */
        .btn-orange button { background-color: #FF7900 !important; color: white !important; font-weight: bold; border-radius: 6px; border: none; width: 100%; }
        .btn-inwi button { background-color: #A1006B !important; color: white !important; font-weight: bold; border-radius: 6px; border: none; width: 100%; }
        .btn-zte button { background-color: #005BAC !important; color: white !important; font-weight: bold; border-radius: 6px; border: none; width: 100%; }
        
        /* Style des métriques */
        div[data-testid="stMetricValue"] {
            font-size: 28px !important;
            font-weight: 700 !important;
            color: #0B4F6C !important;
        }
        
        /* Barre latérale (Sidebar) */
        section[data-testid="stSidebar"] {
            background-color: #1E293B;
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] p, section[data-testid="stSidebar"] label {
            color: #F8FAFC !important;
        }
    </style>
    """, unsafe_allow_html=True)

apply_advanced_theme()

# ==========================================
# BASE DE DONNÉES
# ==========================================
DB_FILE = "database.json"

def init_db():
    if not os.path.exists(DB_FILE):
        db = {
            "users": {
                "admin": {"password": "admin", "role": "admin", "last_login": ""}
            },
            "articles": [],
            "fournisseurs": ["NEC", "ZTE", "Intégral", "FO connect"],
            "equipes": ["Nabil Team", "Yassine Team", "Issa Team"],
            "transactions": []
        }
        save_db(db)
        return db
    else:
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {"users": {"admin": {"password": "admin", "role": "admin"}}, "articles": [], "fournisseurs": [], "equipes": [], "transactions": []}

def save_db(db):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=4)

db = init_db()

# ==========================================
# FONCTIONS MÉTIERS & GENERATION PDF
# ==========================================
def get_stock(client):
    stock = {}
    for art in db["articles"]:
        stock[art["designation"]] = {"ref": art["ref"], "qte": 0}
        
    for tr in db["transactions"]:
        if tr["client"] == client:
            for item in tr["articles"]:
                nom = item["designation"]
                qte = item["qte"]
                if nom not in stock:
                    stock[nom] = {"ref": item.get("ref", ""), "qte": 0}
                if tr["type"] in ["BE", "ADJ_PLUS"]:
                    stock[nom]["qte"] += qte
                elif tr["type"] in ["BS", "ADJ_MOINS"]:
                    stock[nom]["qte"] -= qte
    return stock

def generate_id(type_bon):
    today = datetime.datetime.now().strftime("%Y%m%d")
    count = sum(1 for t in db["transactions"] if t["type"] == type_bon and today in t["id"])
    return f"MW-{type_bon}-{today}-{(count + 1):02d}"

def generate_pdf(bon_data, client):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    
    logo_nomatis = LOGOS["NOMATIS"]
    if os.path.exists(logo_nomatis):
        try:
            pdf.image(logo_nomatis, x=10, y=10, w=35)
        except Exception:
            pdf.set_font("Arial", 'B', 14)
            pdf.text(10, 15, "NOMATIS")
    else:
        pdf.set_font("Arial", 'B', 14)
        pdf.text(10, 15, "NOMATIS")
        
    logo_client = LOGOS.get(client, "")
    if os.path.exists(logo_client):
        try:
            pdf.image(logo_client, x=160, y=10, w=35)
        except Exception:
            pdf.set_font("Arial", 'B', 12)
            pdf.text(160, 15, client)
    else:
        pdf.set_font("Arial", 'B', 12)
        pdf.text(160, 15, client)

    pdf.set_y(25)
    pdf.set_font("Arial", size=8)
    pdf.cell(100, 4, "NOMATIS", ln=1)
    pdf.cell(100, 4, "32 Rue Al Hatim", ln=1)
    pdf.cell(100, 4, "Les Orangers", ln=1)
    pdf.cell(100, 4, "10000", ln=1)
    pdf.ln(8)
    
    pdf.set_font("Arial", 'B', 14)
    titre = "Bon d'entree" if bon_data['type'] == 'BE' else "Bon de sortie"
    pdf.cell(0, 8, titre, ln=1, align='C')
    pdf.ln(4)
    
    pdf.set_font("Arial", 'B', 8)
    entetes_info = ["Bon de Livraison", "Date", "Fournisseur" if bon_data['type']=='BE' else "Equipe", "Lieu de livraison", "receptione par", "Stock"]
    col_w = [35, 25, 35, 35, 30, 30]
    
    for i, entete in enumerate(entetes_info):
        pdf.cell(col_w[i], 7, entete, border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", size=8)
    valeurs_info = [
        bon_data['id'],
        bon_data['date'],
        bon_data['fournisseur_equipe'],
        bon_data['destination'],
        bon_data['user'],
        client
    ]
    for i, val in enumerate(valeurs_info):
        pdf.cell(col_w[i], 7, str(val), border=1, align='C')
    pdf.ln(10)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(45, 7, "Référence", border=1, align='C')
    pdf.cell(115, 7, "Désignation", border=1, align='C')
    pdf.cell(30, 7, "Qté", border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", size=8)
    for art in bon_data['articles']:
        pdf.cell(45, 6, str(art.get('ref', '')), border=1)
        pdf.cell(115, 6, str(art['designation']), border=1)
        pdf.cell(30, 6, str(art['qte']), border=1, align='C')
        pdf.ln()

    return bytes(pdf.output())

# ==========================================
# 1. PAGE DE CONNEXION
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.role = None
    st.session_state.client = None

if not st.session_state.logged_in:
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        img_nom = load_image(LOGOS["NOMATIS"])
        if img_nom:
            st.image(img_nom, width=220)
        
        st.markdown("<h1 style='text-align: center;'>Gestion Stock MW NOMATIS</h1>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            
            btn_class = "btn-login-vert" if username and password else "btn-login-rouge"
            st.markdown(f'<div class="{btn_class}">', unsafe_allow_html=True)
            submitted = st.form_submit_button("SE CONNECTER")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if submitted:
                if username in db["users"] and db["users"][username]["password"] == password:
                    st.session_state.logged_in = True
                    st.session_state.user = username
                    st.session_state.role = db["users"][username]["role"]
                    db["users"][username]["last_login"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    save_db(db)
                    st.rerun()
                else:
                    st.error("Identifiants incorrects. Mot de passe par défaut : admin")
    st.stop()

# ==========================================
# 2. SÉLECTION DU CLIENT (ESPACES DÉDIÉS)
# ==========================================
if st.session_state.client is None:
    st.title(f"👋 Bienvenue, {st.session_state.user} !")
    st.markdown("### Veuillez sélectionner votre espace de gestion :")
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown('<div class="client-card card-orange">', unsafe_allow_html=True)
        img_orange = load_image(LOGOS["ORANGE"])
        if img_orange:
            st.image(img_orange, width=120)
        st.markdown("#### Stock ORANGE")
        st.markdown('<div class="btn-orange">', unsafe_allow_html=True)
        if st.button("Accéder au Stock", key="btn_or"):
            st.session_state.client = "ORANGE"
            st.rerun()
        st.markdown('</div></div>', unsafe_allow_html=True)
            
    with c2:
        st.markdown('<div class="client-card card-inwi">', unsafe_allow_html=True)
        img_inwi = load_image(LOGOS["INWI"])
        if img_inwi:
            st.image(img_inwi, width=120)
        st.markdown("#### Stock INWI")
        st.markdown('<div class="btn-inwi">', unsafe_allow_html=True)
        if st.button("Accéder au Stock", key="btn_in"):
            st.session_state.client = "INWI"
            st.rerun()
        st.markdown('</div></div>', unsafe_allow_html=True)
            
    with c3:
        st.markdown('<div class="client-card card-zte">', unsafe_allow_html=True)
        img_zte = load_image(LOGOS["ZTE"])
        if img_zte:
            st.image(img_zte, width=120)
        st.markdown("#### Stock ZTE")
        st.markdown('<div class="btn-zte">', unsafe_allow_html=True)
        if st.button("Accéder au Stock", key="btn_zt"):
            st.session_state.client = "ZTE"
            st.rerun()
        st.markdown('</div></div>', unsafe_allow_html=True)
        
    st.divider()
    with st.expander("👤 Mon Profil (Modifier nom / mot de passe)"):
        new_user = st.text_input("Nom d'utilisateur", value=st.session_state.user)
        new_pass = st.text_input("Nouveau mot de passe", type="password")
        if st.button("Enregistrer les modifications"):
            if new_user and new_user != st.session_state.user:
                db["users"][new_user] = db["users"].pop(st.session_state.user)
                st.session_state.user = new_user
            if new_pass:
                db["users"][st.session_state.user]["password"] = new_pass
            save_db(db)
            st.success("Profil mis à jour avec succès !")
    
    if st.button("Déconnexion"):
        st.session_state.clear()
        st.rerun()
    st.stop()

# ==========================================
# 3. APPLICATION PRINCIPALE
# ==========================================
client = st.session_state.client
role = st.session_state.role

# Sidebar
img_nom_side = load_image(LOGOS["NOMATIS"])
if img_nom_side:
    st.sidebar.image(img_nom_side, width=140)
st.sidebar.markdown(f"## Espace **{client}**")
st.sidebar.caption(f"Connecté : **{st.session_state.user}** (`{role}`)")
st.sidebar.divider()

if st.sidebar.button("🔄 Changer de Client", use_container_width=True):
    st.session_state.client = None
    st.rerun()
if st.sidebar.button("🚪 Déconnexion", use_container_width=True):
    st.session_state.clear()
    st.rerun()

st.sidebar.divider()

menus = ["Situation Stock", "Historique"]
if role in ["admin", "magasinier"]:
    menus = ["Bon d'Entrée (BE)", "Bon de Sortie (BS)"] + menus
if role == "admin":
    menus.append("Configuration")

choix_menu = st.sidebar.radio("Menu Navigation", menus)
liste_articles = [a["designation"] for a in db["articles"]]

# --- BON D'ENTRÉE (BE) ---
if choix_menu == "Bon d'Entrée (BE)":
    st.header("📥 Créer un Bon d'Entrée (BE)")
    
    col1, col2 = st.columns(2)
    with col1:
        date_be = st.date_input("Date du BE", max_value=datetime.date.today())
        fournisseur = st.selectbox("Fournisseur", db["fournisseurs"] + ["Autre..."])
        if fournisseur == "Autre...":
            fournisseur = st.text_input("Nom du nouveau Fournisseur")
    with col2:
        lieu = st.text_input("Lieu de livraison", "Dépôt Principal")
        remarque_bon = st.text_area("Remarque Générale", height=68)
        
    st.divider()
    st.subheader("Articles à intégrer")
    
    if "current_be_articles" not in st.session_state:
        st.session_state.current_be_articles = []

    with st.form("ajout_article_be", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            article_sel = st.selectbox("Sélectionner l'article", liste_articles if liste_articles else ["Aucun article"])
        with c2:
            qte = st.number_input("Quantité", min_value=1, value=1)
        with c3:
            remarque_art = st.text_input("Remarque article")
        
        if st.form_submit_button("➕ Ajouter la ligne"):
            if not liste_articles:
                st.error("Ajoutez d'abord des articles dans le menu Configuration !")
            else:
                ref = next((a["ref"] for a in db["articles"] if a["designation"] == article_sel), "")
                trouve = False
                for item in st.session_state.current_be_articles:
                    if item["designation"] == article_sel:
                        item["qte"] += qte
                        trouve = True
                        break
                if not trouve:
                    st.session_state.current_be_articles.append({"ref": ref, "designation": article_sel, "qte": qte, "remarque": remarque_art})
                st.rerun()

    if st.session_state.current_be_articles:
        df_be = pd.DataFrame(st.session_state.current_be_articles)
        st.dataframe(df_be, use_container_width=True)
        
        c_del, c_save = st.columns([1, 4])
        with c_del:
            if st.button("🗑️ Vider tout"):
                st.session_state.current_be_articles = []
                st.rerun()
        with c_save:
            if st.button("💾 Enregistrer et Générer le BE", type="primary", use_container_width=True):
                nouveau_be = {
                    "id": generate_id("BE"),
                    "type": "BE",
                    "date": date_be.strftime("%Y-%m-%d"),
                    "heure_saisie": datetime.datetime.now().strftime("%H:%M:%S"),
                    "client": client,
                    "user": st.session_state.user,
                    "fournisseur_equipe": fournisseur,
                    "destination": lieu,
                    "remarque": remarque_bon,
                    "articles": st.session_state.current_be_articles
                }
                db["transactions"].append(nouveau_be)
                if fournisseur not in db["fournisseurs"] and fournisseur:
                    db["fournisseurs"].append(fournisseur)
                save_db(db)
                
                st.success(f"✅ Bon d'Entrée {nouveau_be['id']} enregistré !")
                pdf_bytes = generate_pdf(nouveau_be, client)
                st.download_button("📄 Imprimer le BE (PDF)", data=pdf_bytes, file_name=f"{nouveau_be['id']}.pdf", mime='application/pdf')
                st.session_state.current_be_articles = []

# --- BON DE SORTIE (BS) ---
elif choix_menu == "Bon de Sortie (BS)":
    st.header("📤 Créer un Bon de Sortie (BS)")
    
    col1, col2 = st.columns(2)
    with col1:
        date_bs = st.date_input("Date du BS", max_value=datetime.date.today())
        equipe = st.selectbox("Équipe destinataire", db["equipes"])
    with col2:
        destination = st.text_input("Destination / Site")
        remarque_bon = st.text_area("Remarque Générale", height=68)
        
    st.divider()
    st.subheader("Articles à sortir")
    stock_actuel = get_stock(client)
    
    if "current_bs_articles" not in st.session_state:
        st.session_state.current_bs_articles = []

    with st.form("ajout_article_bs", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            article_sel = st.selectbox("Sélectionner l'article", liste_articles if liste_articles else ["Aucun article"])
        with c2:
            qte = st.number_input("Quantité", min_value=1, value=1)
        with c3:
            remarque_art = st.text_input("Remarque article")
        
        if st.form_submit_button("➕ Ajouter la ligne"):
            if not liste_articles:
                st.error("Aucun article configuré !")
            else:
                qte_deja = sum(item["qte"] for item in st.session_state.current_bs_articles if item["designation"] == article_sel)
                dispo = stock_actuel.get(article_sel, {}).get("qte", 0) - qte_deja
                
                if qte > dispo:
                    st.error(f"⚠️ Stock insuffisant ! Quantité disponible restante : {dispo}")
                else:
                    ref = next((a["ref"] for a in db["articles"] if a["designation"] == article_sel), "")
                    trouve = False
                    for item in st.session_state.current_bs_articles:
                        if item["designation"] == article_sel:
                            item["qte"] += qte
                            trouve = True
                            break
                    if not trouve:
                        st.session_state.current_bs_articles.append({"ref": ref, "designation": article_sel, "qte": qte, "remarque": remarque_art})
                    st.rerun()

    if st.session_state.current_bs_articles:
        df_bs = pd.DataFrame(st.session_state.current_bs_articles)
        st.dataframe(df_bs, use_container_width=True)
        
        c_del, c_save = st.columns([1, 4])
        with c_del:
            if st.button("🗑️ Vider tout"):
                st.session_state.current_bs_articles = []
                st.rerun()
        with c_save:
            if st.button("💾 Enregistrer et Générer le BS", type="primary", use_container_width=True):
                nouveau_bs = {
                    "id": generate_id("BS"),
                    "type": "BS",
                    "date": date_bs.strftime("%Y-%m-%d"),
                    "heure_saisie": datetime.datetime.now().strftime("%H:%M:%S"),
                    "client": client,
                    "user": st.session_state.user,
                    "fournisseur_equipe": equipe,
                    "destination": destination,
                    "remarque": remarque_bon,
                    "articles": st.session_state.current_bs_articles
                }
                db["transactions"].append(nouveau_bs)
                save_db(db)
                st.success(f"✅ Bon de Sortie {nouveau_bs['id']} enregistré !")
                
                pdf_bytes = generate_pdf(nouveau_bs, client)
                st.download_button("📄 Imprimer le BS (PDF)", data=pdf_bytes, file_name=f"{nouveau_bs['id']}.pdf", mime='application/pdf')
                st.session_state.current_bs_articles = []

# --- SITUATION STOCK ---
elif choix_menu == "Situation Stock":
    st.header(f"📊 Situation du Stock — {client}")
    stock = get_stock(client)
    
    if not stock:
        st.info("Aucun article répertorié dans ce stock.")
    else:
        df_stock = pd.DataFrame.from_dict(stock, orient='index').reset_index()
        df_stock.columns = ["Désignation", "Référence", "Quantité Disponible"]
        df_stock = df_stock[["Référence", "Désignation", "Quantité Disponible"]]
        
        total_arts = len(df_stock)
        total_qte = df_stock["Quantité Disponible"].sum()
        ruptures = len(df_stock[df_stock["Quantité Disponible"] <= 0])
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Nombre Références", total_arts)
        m2.metric("Total Unités en Stock", total_qte)
        m3.metric("Articles en Rupture", ruptures, delta_color="inverse")
        
        st.divider()
        st.dataframe(df_stock, use_container_width=True)

# --- HISTORIQUE ---
elif choix_menu == "Historique":
    st.header("🕒 Historique des Transactions")
    tab_be, tab_bs = st.tabs(["📥 Bons d'Entrée (BE)", "📤 Bons de Sortie (BS)"])
    
    def afficher_historique(type_bon):
        trans = [t for t in db["transactions"] if t["type"] == type_bon and t["client"] == client]
        if not trans:
            st.info(f"Aucun {type_bon} enregistré pour le moment.")
            return
            
        for t in reversed(trans):
            with st.expander(f"📄 {t['id']} — Date : {t['date']} à {t.get('heure_saisie', 'N/C')} (Par: {t['user']})"):
                st.write(f"**Tiers/Équipe:** {t['fournisseur_equipe']} | **Lieu:** {t['destination']}")
                if t.get('remarque'):
                    st.caption(f"Remarque: {t['remarque']}")
                st.table(pd.DataFrame(t['articles']))
                
                c1, c2 = st.columns(2)
                with c1:
                    pdf_bytes = generate_pdf(t, client)
                    st.download_button("🖨️ Télécharger le PDF", data=pdf_bytes, file_name=f"{t['id']}.pdf", mime='application/pdf', key=f"print_{t['id']}")
                with c2:
                    if role in ["admin", "magasinier"]:
                        if st.button("🗑️ Supprimer ce bon", key=f"del_{t['id']}"):
                            db["transactions"] = [x for x in db["transactions"] if x["id"] != t["id"]]
                            save_db(db)
                            st.success("Bon supprimé.")
                            st.rerun()

    with tab_be:
        afficher_historique("BE")
    with tab_bs:
        afficher_historique("BS")

# --- CONFIGURATION (ADMIN) ---
elif choix_menu == "Configuration":
    st.header("⚙️ Administration & Configuration")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Utilisateurs", "Articles", "Fournisseurs", "Équipes", "Ajustement Stock"])
    
    with tab1:
        st.subheader("Utilisateurs du Système")
        df_u = pd.DataFrame.from_dict(db["users"], orient='index').reset_index()
        df_u.columns = ["Nom", "Mot de passe", "Rôle", "Dernière Connexion"]
        st.dataframe(df_u, use_container_width=True)
        
        with st.form("add_user"):
            nom_u = st.text_input("Nom d'utilisateur")
            pass_u = st.text_input("Mot de passe")
            role_u = st.selectbox("Rôle", ["magasinier", "coordinateur", "coordinatrice", "admin"])
            if st.form_submit_button("Créer l'utilisateur"):
                if nom_u and pass_u:
                    db["users"][nom_u] = {"password": pass_u, "role": role_u, "last_login": ""}
                    save_db(db)
                    st.success("Utilisateur ajouté !")
                    st.rerun()

    with tab2:
        st.subheader("Référentiel des Articles")
        st.dataframe(pd.DataFrame(db["articles"]), use_container_width=True)
        with st.form("add_article"):
            ref = st.text_input("Référence Article")
            desig = st.text_input("Désignation Article")
            if st.form_submit_button("Ajouter l'Article"):
                if desig:
                    db["articles"].append({"ref": ref, "designation": desig})
                    save_db(db)
                    st.rerun()

    with tab3:
        st.subheader("Liste des Fournisseurs")
        st.write(db["fournisseurs"])
        with st.form("add_f"):
            f_nom = st.text_input("Nouveau Fournisseur")
            if st.form_submit_button("Ajouter"):
                if f_nom:
                    db["fournisseurs"].append(f_nom)
                    save_db(db)
                    st.rerun()

    with tab4:
        st.subheader("Équipes Projets")
        st.write(db["equipes"])
        with st.form("add_e"):
            e_nom = st.text_input("Nouvelle Équipe")
            if st.form_submit_button("Ajouter"):
                if e_nom:
                    db["equipes"].append(e_nom)
                    save_db(db)
                    st.rerun()

    with tab5:
        st.subheader(f"Ajustement Inventaire ({client})")
        with st.form("adjust_stock"):
            art_adj = st.selectbox("Article", liste_articles if liste_articles else ["Aucun"])
            type_adj = st.radio("Type d'ajustement", ["Ajouter (+)", "Retirer (-)"])
            qte_adj = st.number_input("Quantité", min_value=1)
            motif = st.text_input("Motif / Justification")
            if st.form_submit_button("Valider l'ajustement"):
                t_type = "ADJ_PLUS" if "+" in type_adj else "ADJ_MOINS"
                db["transactions"].append({
                    "id": f"MW-ADJ-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "type": t_type,
                    "date": datetime.datetime.now().strftime("%Y-%m-%d"),
                    "client": client,
                    "user": st.session_state.user,
                    "fournisseur_equipe": "Ajustement Manuel",
                    "destination": motif,
                    "remarque": "Ajustement Inventaire",
                    "articles": [{"designation": art_adj, "qte": qte_adj, "ref": ""}]
                })
                save_db(db)
                st.success("Ajustement appliqué !")
                st.rerun()
