import streamlit as st
import pandas as pd
import datetime
import json
import os
from fpdf import FPDF
import base64

# ==========================================
# CONFIGURATION DE LA PAGE & DESIGN
# ==========================================
st.set_page_config(page_title="Gestion Stock MW NOMATIS", layout="wide", initial_sidebar_state="expanded")

# CSS Personnalisé (Couleurs, Boutons)
def local_css():
    st.markdown("""
    <style>
        /* Couleurs principales: Bleu, Blanc, Vert */
        .stApp { background-color: #F8F9FA; }
        h1, h2, h3 { color: #0056b3; }
        
        /* Bouton Connexion dynamique */
        .btn-login-rouge button { background-color: #dc3545 !important; color: white !important; font-weight: bold; }
        .btn-login-vert button { background-color: #28a745 !important; color: white !important; font-weight: bold; }
        
        /* Boutons Clients */
        .btn-orange button { background-color: #FF7900 !important; color: white !important; font-weight: bold; }
        .btn-inwi button { background-color: #E30613 !important; color: white !important; font-weight: bold; }
        .btn-zte button { background-color: #005A9C !important; color: white !important; font-weight: bold; }
        
        .dataframe { font-size: 14px; }
    </style>
    """, unsafe_allow_html=True)

local_css()

# ==========================================
# BASE DE DONNÉES (Simulation via JSON local)
# ==========================================
DB_FILE = "database.json"

def init_db():
    if not os.path.exists(DB_FILE):
        db = {
            "users": {
                "admin": {"password": "admin", "role": "admin", "last_login": ""}
            },
            "articles": [], # {ref, designation}
            "fournisseurs": ["NEC", "ZTE", "Intégral", "FO connect"],
            "equipes": ["Nabil Team", "Yassine Team", "Issa Team"],
            "transactions": [] # {id, type (BE/BS/ADJ), date, client, user, fournisseur_equipe, destination, articles: [{ref, designation, qte, remarque}]}
        }
        save_db(db)
        return db
    else:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

def save_db(db):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=4)

db = init_db()

# ==========================================
# FONCTIONS MÉTIERS & STOCK
# ==========================================
def get_stock(client):
    stock = {}
    # Initialiser tous les articles à 0
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

# ==========================================
# GÉNÉRATION DU PDF (Le modèle strict)
# ==========================================
def generate_pdf(bon_data, client):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    
    # En-tête : Logos et Adresse
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(100, 8, "NOMATIS", ln=0)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(90, 8, f"Client : {client.upper()}", ln=1, align='R')
    
    pdf.set_font("Arial", size=9)
    pdf.cell(100, 5, "32 Rue Al Hatim", ln=1)
    pdf.cell(100, 5, "Les Orangers", ln=1)
    pdf.cell(100, 5, "10000", ln=1)
    pdf.ln(10)
    
    # Titre Central
    pdf.set_font("Arial", 'B', 16)
    titre = "Bon d'Entrée" if bon_data['type'] == 'BE' else "Bon de Sortie"
    pdf.cell(0, 10, titre, ln=1, align='C')
    pdf.ln(5)
    
    # Tableau Info Globales
    pdf.set_font("Arial", 'B', 8)
    entetes_info = ["N° Bon", "Date", "Fournisseur/Equipe", "Lieu/Dest", "Par", "Stock"]
    col_widths_info = [35, 25, 40, 40, 30, 20]
    
    for i, entete in enumerate(entetes_info):
        pdf.cell(col_widths_info[i], 8, entete, border=1, align='C')
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
        pdf.cell(col_widths_info[i], 8, str(val), border=1, align='C')
    pdf.ln(10)
    
    # Tableau Articles
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(40, 8, "Référence", border=1, align='C')
    pdf.cell(90, 8, "Désignation", border=1, align='C')
    pdf.cell(30, 8, "Qté", border=1, align='C')
    pdf.cell(30, 8, "Remarque", border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", size=9)
    for art in bon_data['articles']:
        pdf.cell(40, 8, str(art.get('ref', '')), border=1, align='C')
        pdf.cell(90, 8, str(art['designation']), border=1)
        pdf.cell(30, 8, str(art['qte']), border=1, align='C')
        pdf.cell(30, 8, str(art.get('remarque', '')), border=1)
        pdf.ln()

    # Footer Signatures
    pdf.ln(20)
    pdf.cell(95, 10, "Signature Magasinier :", ln=0)
    pdf.cell(95, 10, "Signature Réceptionnaire / Livreur :", ln=1)

    # Sortie
    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# 1. PAGE DE CONNEXION
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.role = None
    st.session_state.client = None

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.image("https://via.placeholder.com/150x50/FFFFFF/0056b3?text=NOMATIS", width=150) # Placeholder logo
        st.title("Gestion Stock MW NOMATIS")
        
        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            
            # CSS du bouton selon la saisie (Rouge -> Vert)
            btn_class = "btn-login-vert" if username and password else "btn-login-rouge"
            st.markdown(f'<div class="{btn_class}">', unsafe_allow_html=True)
            submitted = st.form_submit_button("SE CONNECTER")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if submitted:
                if username in db["users"] and db["users"][username]["password"] == password:
                    st.session_state.logged_in = True
                    st.session_state.user = username
                    st.session_state.role = db["users"][username]["role"]
                    # MAJ Date connexion
                    db["users"][username]["last_login"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    save_db(db)
                    st.success("Accès validé !")
                    st.rerun()
                else:
                    st.error("Identifiants incorrects")
    st.stop()

# ==========================================
# 2. SÉLECTION DU CLIENT
# ==========================================
if st.session_state.client is None:
    st.title(f"Bienvenue, {st.session_state.user} !")
    st.subheader("Sélectionnez l'espace client :")
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.image("https://via.placeholder.com/200x100/FF7900/FFFFFF?text=ORANGE", use_container_width=True)
        st.markdown('<div class="btn-orange">', unsafe_allow_html=True)
        if st.button("Accès au stock ORANGE", use_container_width=True):
            st.session_state.client = "ORANGE"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
            
    with c2:
        st.image("https://via.placeholder.com/200x100/E30613/FFFFFF?text=INWI", use_container_width=True)
        st.markdown('<div class="btn-inwi">', unsafe_allow_html=True)
        if st.button("Accès au stock INWI", use_container_width=True):
            st.session_state.client = "INWI"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
            
    with c3:
        st.image("https://via.placeholder.com/200x100/005A9C/FFFFFF?text=ZTE", use_container_width=True)
        st.markdown('<div class="btn-zte">', unsafe_allow_html=True)
        if st.button("Accès au stock ZTE", use_container_width=True):
            st.session_state.client = "ZTE"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
    st.divider()
    st.write("Espace personnel :")
    with st.expander("Modifier mes informations"):
        new_user = st.text_input("Nouveau nom", value=st.session_state.user)
        new_pass = st.text_input("Nouveau mot de passe", type="password")
        if st.button("Mettre à jour mon profil"):
            if new_user and new_user != st.session_state.user:
                db["users"][new_user] = db["users"].pop(st.session_state.user)
                st.session_state.user = new_user
            if new_pass:
                db["users"][st.session_state.user]["password"] = new_pass
            save_db(db)
            st.success("Profil mis à jour !")
    
    if st.button("Se déconnecter"):
        st.session_state.clear()
        st.rerun()
    st.stop()

# ==========================================
# 3. APPLICATION PRINCIPALE (5 Rubriques)
# ==========================================
client = st.session_state.client
role = st.session_state.role

st.sidebar.image("https://via.placeholder.com/150x50/FFFFFF/0056b3?text=NOMATIS")
st.sidebar.title(f"Stock {client}")
st.sidebar.write(f"Utilisateur : **{st.session_state.user}** ({role})")
if st.sidebar.button("Changer de Client"):
    st.session_state.client = None
    st.rerun()
if st.sidebar.button("Déconnexion"):
    st.session_state.clear()
    st.rerun()

# Menus basés sur les rôles
menus = ["Situation Stock", "Historique"]
if role in ["admin", "magasinier"]:
    menus = ["Bon d'Entrée (BE)", "Bon de Sortie (BS)"] + menus
if role == "admin":
    menus.append("Configuration")

choix_menu = st.sidebar.radio("Navigation", menus)

liste_articles = [a["designation"] for a in db["articles"]]

# --- RUBRIQUE : BON D'ENTRÉE (BE) ---
if choix_menu == "Bon d'Entrée (BE)":
    st.header("📥 Créer un Bon d'Entrée (BE)")
    
    col1, col2 = st.columns(2)
    with col1:
        date_be = st.date_input("Date du BE", max_value=datetime.date.today())
        fournisseur = st.selectbox("Fournisseur", db["fournisseurs"] + ["Autre..."])
        if fournisseur == "Autre...":
            fournisseur = st.text_input("Nouveau Fournisseur")
    with col2:
        lieu = st.text_input("Lieu de livraison", "Dépôt Principal")
        remarque_bon = st.text_area("Remarque Générale")
        
    st.subheader("Articles à réceptionner")
    if "current_be_articles" not in st.session_state:
        st.session_state.current_be_articles = []

    with st.form("ajout_article_be"):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            article_sel = st.selectbox("Article", liste_articles if liste_articles else ["Veuillez configurer les articles"])
        with c2:
            qte = st.number_input("Quantité", min_value=1, value=1)
        with c3:
            remarque_art = st.text_input("Remarque")
        
        if st.form_submit_button("Ajouter l'article"):
            if not liste_articles:
                st.error("Aucun article configuré !")
            else:
                ref = next(a["ref"] for a in db["articles"] if a["designation"] == article_sel)
                # Fusion automatique si l'article existe déjà dans le bon
                trouve = False
                for item in st.session_state.current_be_articles:
                    if item["designation"] == article_sel:
                        item["qte"] += qte
                        trouve = True
                        break
                if not trouve:
                    st.session_state.current_be_articles.append({"ref": ref, "designation": article_sel, "qte": qte, "remarque": remarque_art})
                st.success("Article ajouté.")
                st.rerun()
                
    if st.session_state.current_be_articles:
        st.table(pd.DataFrame(st.session_state.current_be_articles))
        if st.button("Vider la liste"):
            st.session_state.current_be_articles = []
            st.rerun()
            
        if st.button("Enregistrer le Bon d'Entrée", type="primary"):
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
            save_db(db)
            
            # Gestion du fournisseur si nouveau
            if fournisseur not in db["fournisseurs"] and fournisseur != "":
                db["fournisseurs"].append(fournisseur)
                save_db(db)
                
            st.success(f"Bon d'entrée {nouveau_be['id']} enregistré avec succès !")
            
            # Génération PDF
            pdf_bytes = generate_pdf(nouveau_be, client)
            st.download_button(label="📄 Télécharger le BE (PDF)", data=pdf_bytes, file_name=f"{nouveau_be['id']}.pdf", mime='application/pdf')
            
            # Reset
            st.session_state.current_be_articles = []

# --- RUBRIQUE : BON DE SORTIE (BS) ---
elif choix_menu == "Bon de Sortie (BS)":
    st.header("📤 Créer un Bon de Sortie (BS)")
    
    col1, col2 = st.columns(2)
    with col1:
        date_bs = st.date_input("Date du BS", max_value=datetime.date.today())
        equipe = st.selectbox("Équipe destinataire", db["equipes"])
    with col2:
        destination = st.text_input("Destination / Site")
        remarque_bon = st.text_area("Remarque Générale")
        
    st.subheader("Articles à sortir")
    stock_actuel = get_stock(client)
    
    if "current_bs_articles" not in st.session_state:
        st.session_state.current_bs_articles = []

    with st.form("ajout_article_bs"):
        c1, c2, c3 = st.columns([2, 1, 1])
        with c1:
            article_sel = st.selectbox("Article", liste_articles if liste_articles else ["Veuillez configurer les articles"])
        with c2:
            qte = st.number_input("Quantité", min_value=1, value=1)
        with c3:
            remarque_art = st.text_input("Remarque")
        
        if st.form_submit_button("Ajouter l'article"):
            if not liste_articles:
                st.error("Aucun article configuré !")
            else:
                # Vérification du stock dispo (Stock actuel - déjà mis dans le bon)
                qte_deja_au_bon = sum(item["qte"] for item in st.session_state.current_bs_articles if item["designation"] == article_sel)
                dispo = stock_actuel.get(article_sel, {}).get("qte", 0) - qte_deja_au_bon
                
                if qte > dispo:
                    st.error(f"Stock insuffisant ! Stock disponible restant : {dispo}")
                else:
                    ref = next(a["ref"] for a in db["articles"] if a["designation"] == article_sel)
                    trouve = False
                    for item in st.session_state.current_bs_articles:
                        if item["designation"] == article_sel:
                            item["qte"] += qte
                            trouve = True
                            break
                    if not trouve:
                        st.session_state.current_bs_articles.append({"ref": ref, "designation": article_sel, "qte": qte, "remarque": remarque_art})
                    st.success("Article ajouté.")
                    st.rerun()

    if st.session_state.current_bs_articles:
        st.table(pd.DataFrame(st.session_state.current_bs_articles))
        if st.button("Vider la liste"):
            st.session_state.current_bs_articles = []
            st.rerun()
            
        if st.button("Enregistrer le Bon de Sortie", type="primary"):
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
            st.success(f"Bon de sortie {nouveau_bs['id']} enregistré avec succès !")
            
            # Génération PDF
            pdf_bytes = generate_pdf(nouveau_bs, client)
            st.download_button(label="📄 Télécharger le BS (PDF)", data=pdf_bytes, file_name=f"{nouveau_bs['id']}.pdf", mime='application/pdf')
            
            st.session_state.current_bs_articles = []

# --- RUBRIQUE : SITUATION STOCK ---
elif choix_menu == "Situation Stock":
    st.header("📊 Situation du Stock Actuel")
    st.write(f"Stock en temps réel pour le client : **{client}**")
    
    stock = get_stock(client)
    if not stock:
        st.info("Le stock est vide.")
    else:
        df_stock = pd.DataFrame.from_dict(stock, orient='index').reset_index()
        df_stock.columns = ["Désignation", "Référence", "Quantité Disponible"]
        # Réorganiser les colonnes
        df_stock = df_stock[["Référence", "Désignation", "Quantité Disponible"]]
        st.dataframe(df_stock, use_container_width=True)

# --- RUBRIQUE : HISTORIQUE ---
elif choix_menu == "Historique":
    st.header("🕒 Historique des Mouvements")
    tab_be, tab_bs = st.tabs(["Bons d'Entrée (BE)", "Bons de Sortie (BS)"])
    
    def afficher_historique(type_bon):
        trans = [t for t in db["transactions"] if t["type"] == type_bon and t["client"] == client]
        if not trans:
            st.write(f"Aucun {type_bon} trouvé.")
            return
            
        for t in reversed(trans):
            with st.expander(f"{t['id']} | Date: {t['date']} | Par: {t['user']} | Fournisseur/Equipe: {t['fournisseur_equipe']}"):
                st.write(f"**Lieu/Destination:** {t['destination']} | **Remarque:** {t['remarque']}")
                st.table(pd.DataFrame(t['articles']))
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    pdf_bytes = generate_pdf(t, client)
                    st.download_button("🖨️ Imprimer PDF", data=pdf_bytes, file_name=f"{t['id']}.pdf", mime='application/pdf', key=f"print_{t['id']}")
                with col2:
                    if role in ["admin", "magasinier"]:
                        if st.button("🗑️ Supprimer", key=f"del_{t['id']}"):
                            db["transactions"] = [x for x in db["transactions"] if x["id"] != t["id"]]
                            save_db(db)
                            st.success("Bon supprimé ! Le stock a été ajusté.")
                            st.rerun()
                # La modification complète demanderait un formulaire complet pré-rempli (Simplifié ici pour des raisons d'espace, mais la suppression/recréation est la méthode la plus sûre comptablement)

    with tab_be:
        afficher_historique("BE")
    with tab_bs:
        afficher_historique("BS")

# --- RUBRIQUE : CONFIGURATION (Admin uniquement) ---
elif choix_menu == "Configuration":
    st.header("⚙️ Configuration Système")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Utilisateurs", "Articles", "Fournisseurs", "Équipes", "Ajustement Stock"])
    
    with tab1:
        st.subheader("Gestion des Utilisateurs")
        df_users = pd.DataFrame.from_dict(db["users"], orient='index').reset_index()
        df_users.columns = ["Nom", "Mot de passe", "Rôle", "Dernière Connexion"]
        st.dataframe(df_users)
        
        with st.form("add_user"):
            nom_u = st.text_input("Nom d'utilisateur")
            pass_u = st.text_input("Mot de passe")
            role_u = st.selectbox("Rôle", ["magasinier", "coordinateur", "coordinatrice", "admin"])
            if st.form_submit_button("Ajouter/Modifier Utilisateur"):
                db["users"][nom_u] = {"password": pass_u, "role": role_u, "last_login": ""}
                save_db(db)
                st.success("Utilisateur enregistré !")
                st.rerun()

    with tab2:
        st.subheader("Gestion des Articles (Référentiel global)")
        st.dataframe(pd.DataFrame(db["articles"]))
        with st.form("add_article"):
            ref = st.text_input("Référence")
            desig = st.text_input("Désignation (ex: Câble RJ45)")
            if st.form_submit_button("Ajouter l'article"):
                if desig:
                    db["articles"].append({"ref": ref, "designation": desig})
                    save_db(db)
                    st.success("Article ajouté !")
                    st.rerun()

    with tab3:
        st.subheader("Fournisseurs")
        st.write(db["fournisseurs"])
        with st.form("add_fournisseur"):
            f_nom = st.text_input("Nom fournisseur")
            if st.form_submit_button("Ajouter"):
                db["fournisseurs"].append(f_nom)
                save_db(db)
                st.rerun()

    with tab4:
        st.subheader("Équipes Projet")
        st.write(db["equipes"])
        with st.form("add_equipe"):
            e_nom = st.text_input("Nom de l'équipe")
            if st.form_submit_button("Ajouter"):
                db["equipes"].append(e_nom)
                save_db(db)
                st.rerun()

    with tab5:
        st.subheader(f"Ajustement Manuel du Stock ({client})")
        with st.form("adjust_stock"):
            art_adj = st.selectbox("Article", liste_articles if liste_articles else ["Vide"])
            type_adj = st.radio("Type d'ajustement", ["Ajouter au stock (+)", "Retirer du stock (-)"])
            qte_adj = st.number_input("Quantité à ajuster", min_value=1)
            motif = st.text_input("Motif de l'ajustement")
            if st.form_submit_button("Appliquer l'ajustement"):
                t_type = "ADJ_PLUS" if "Ajouter" in type_adj else "ADJ_MOINS"
                db["transactions"].append({
                    "id": f"MW-ADJ-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}",
                    "type": t_type,
                    "date": datetime.datetime.now().strftime("%Y-%m-%d"),
                    "client": client,
                    "user": st.session_state.user,
                    "fournisseur_equipe": "Manuel",
                    "destination": motif,
                    "remarque": "Ajustement Admin",
                    "articles": [{"designation": art_adj, "qte": qte_adj, "ref": ""}]
                })
                save_db(db)
                st.success("Stock ajusté avec succès !")
                st.rerun()
