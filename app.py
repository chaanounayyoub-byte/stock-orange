import streamlit as st
import pandas as pd
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(
    page_title="Gestion de Stock - Orange",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Fichiers de données CSV local
STOCK_FILE = "orange_stock.csv"
USERS_FILE = "utilisateurs.csv"

# Initialisation des fichiers par défaut si non existants
def init_data():
    if not os.path.exists(STOCK_FILE):
        stock_df = pd.DataFrame([
            {"DESIGNATION": "cable IF", "QUANTITE": 500, "REMARQUE": "RAS"},
            {"DESIGNATION": "clamps", "QUANTITE": 230, "REMARQUE": "a voir"},
            {"DESIGNATION": "KIT de Terre", "QUANTITE": 5, "REMARQUE": "OK"},
            {"DESIGNATION": "mise a la terre IF", "QUANTITE": 105, "REMARQUE": "RAS"},
            {"DESIGNATION": "Connecteur Droit", "QUANTITE": 5, "REMARQUE": "OK"},
            {"DESIGNATION": "connecteur codé", "QUANTITE": 5, "REMARQUE": "OK"},
            {"DESIGNATION": "support 0.6m/0.3m", "QUANTITE": 10, "REMARQUE": "RAS"},
            {"DESIGNATION": "Support 1.2m", "QUANTITE": 10, "REMARQUE": ""}
        ])
        stock_df.to_csv(STOCK_FILE, index=False)

    if not os.path.exists(USERS_FILE):
        users_df = pd.DataFrame([
            {"Login": "admin", "MotDePasse": "admin123", "Nom": "Ayyoub Chaanoun", "Profil": "ADMIN", "Actif": "OUI"},
            {"Login": "magasin1", "MotDePasse": "123456", "Nom": "Ahmed", "Profil": "MAGASINIER", "Actif": "OUI"},
            {"Login": "chef1", "MotDePasse": "123456", "Nom": "Mohamed", "Profil": "CONSULTATION", "Actif": "OUI"}
        ])
        users_df.to_csv(USERS_FILE, index=False)

init_data()

# Chargement et sauvegarde des données
def load_stock():
    return pd.read_csv(STOCK_FILE)

def save_stock(df):
    df.to_csv(STOCK_FILE, index=False)

def load_users():
    return pd.read_csv(USERS_FILE)

def save_users(df):
    df.to_csv(USERS_FILE, index=False)

# Session state pour l'authentification
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_info" not in st.session_state:
    st.session_state.user_info = None

# Styles CSS personnalisés (Thème Orange)
st.markdown("""
    <style>
    .main-header {
        font-size: 28px;
        font-weight: bold;
        color: #FF6600;
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #FF6600;
        color: white;
        border-radius: 6px;
        border: none;
    }
    .stButton>button:hover {
        background-color: #E65C00;
        color: white;
    }
    </style>
""", unsafe_allow_html=True)

# PAGE DE CONNEXION
if not st.session_state.authenticated:
    st.title("🔒 Connexion - Gestion de Stock")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("Veuillez vous identifier")
        login_input = st.text_input("Nom d'utilisateur (Login)")
        password_input = st.text_input("Mot de passe", type="password")
        
        if st.button("Se connecter", use_container_width=True):
            users_df = load_users()
            user = users_df[(users_df["Login"] == login_input) & (users_df["MotDePasse"] == password_input)]
            
            if not user.empty:
                if user.iloc[0]["Actif"] == "OUI":
                    st.session_state.authenticated = True
                    st.session_state.user_info = user.iloc[0].to_dict()
                    st.success(f"Bienvenue {st.session_state.user_info['Nom']} !")
                    st.rerun()
                else:
                    st.error("Ce compte est désactivé.")
            else:
                st.error("Identifiants incorrects.")

else:
    # APPLICATION PRINCIPALE
    user = st.session_state.user_info
    role = user["Profil"]
    
    # Barre latérale (Sidebar)
    st.sidebar.image("https://img.icons8.com/color/96/orange.png", width=60)
    st.sidebar.title("Navigation")
    st.sidebar.write(f"👤 **{user['Nom']}**")
    st.sidebar.caption(f"Rôle : `{role}`")
    
    menu_options = ["📋 Voir le Stock", "➕ Mouvement de Stock"]
    if role in ["ADMIN", "MAGASINIER"]:
        menu_options.append("📝 Ajouter / Modifier Article")
    if role == "ADMIN":
        menu_options.append("👥 Gestion Utilisateurs")
    
    choice = st.sidebar.radio("Aller vers :", menu_options)
    
    if st.sidebar.button("Déconnexion"):
        st.session_state.authenticated = False
        st.session_state.user_info = None
        st.rerun()

    stock_df = load_stock()

    # SECTION 1: VOIR LE STOCK
    if choice == "📋 Voir le Stock":
        st.markdown("<div class='main-header'>📦 État Actuel du Stock</div>", unsafe_allow_html=True)
        
        # Indicateurs clés (KPIs)
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Nombre de Références", len(stock_df))
        kpi2.metric("Quantité Totale en Stock", int(stock_df["QUANTITE"].sum()))
        low_stock_count = len(stock_df[stock_df["QUANTITE"] <= 10])
        kpi3.metric("Articles en Stock Faible (≤ 10)", low_stock_count, delta_color="inverse")
        
        st.markdown("---")
        
        # Recherche et Filtres
        search = st.text_input("🔍 Rechercher une désignation ou remarque :")
        if search:
            filtered_df = stock_df[
                stock_df["DESIGNATION"].str.contains(search, case=False, na=False) |
                stock_df["REMARQUE"].str.contains(search, case=False, na=False)
            ]
        else:
            filtered_df = stock_df

        st.dataframe(filtered_df, use_container_width=True, height=400)
        
        # Téléchargement CSV
        csv_data = stock_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Télécharger l'état de stock (CSV)", data=csv_data, file_name="stock_orange.csv", mime="text/csv")

    # SECTION 2: MOUVEMENT DE STOCK (ENTRÉE / SORTIE)
    elif choice == "➕ Mouvement de Stock":
        st.markdown("<div class='main-header'>🔄 Entrée / Sortie de Stock</div>", unsafe_allow_html=True)
        
        if role == "CONSULTATION":
            st.warning("🔒 Vous êtes en mode consultation. Vous ne pouvez pas modifier le stock.")
        else:
            article = st.selectbox("Sélectionner l'article :", stock_df["DESIGNATION"].unique())
            current_qty = stock_df.loc[stock_df["DESIGNATION"] == article, "QUANTITE"].values[0]
            st.info(f"Quantité actuelle pour **{article}** : `{current_qty}`")
            
            type_mvt = st.radio("Type de mouvement :", ["Entrée (+)", "Sortie (-)"])
            qty_change = st.number_input("Quantité :", min_value=1, value=1, step=1)
            remarque_mvt = st.text_input("Remarque / Justification :")
            
            if st.button("Valider le Mouvement"):
                if type_mvt == "Entrée (+)":
                    new_qty = current_qty + qty_change
                else:
                    new_qty = current_qty - qty_change
                    if new_qty < 0:
                        st.error("❌ Opération impossible : la quantité ne peut pas être négative !")
                        st.stop()
                
                stock_df.loc[stock_df["DESIGNATION"] == article, "QUANTITE"] = new_qty
                if remarque_mvt:
                    stock_df.loc[stock_df["DESIGNATION"] == article, "REMARQUE"] = remarque_mvt
                
                save_stock(stock_df)
                st.success(f"✅ Mouvement enregistré ! Nouvelle quantité pour **{article}** : {new_qty}")

    # SECTION 3: AJOUTER / MODIFIER ARTICLE
    elif choice == "📝 Ajouter / Modifier Article":
        st.markdown("<div class='main-header'>📝 Gestion des Références</div>", unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["➕ Ajouter un nouvel article", "✏️ Modifier / Supprimer"])
        
        with tab1:
            new_des = st.text_input("Désignation de l'article :")
            new_qty = st.number_input("Quantité initiale :", min_value=0, value=0)
            new_rem = st.text_input("Remarque :", value="RAS")
            
            if st.button("Ajouter à la base"):
                if new_des in stock_df["DESIGNATION"].values:
                    st.error("Cet article existe déjà !")
                elif new_des.strip() == "":
                    st.warning("Veuillez saisir une désignation valide.")
                else:
                    new_row = pd.DataFrame([{"DESIGNATION": new_des, "QUANTITE": new_qty, "REMARQUE": new_rem}])
                    stock_df = pd.concat([stock_df, new_row], ignore_index=True)
                    save_stock(stock_df)
                    st.success(f"Article **{new_des}** ajouté avec succès !")

        with tab2:
            article_to_edit = st.selectbox("Article à modifier :", stock_df["DESIGNATION"].unique())
            row_edit = stock_df[stock_df["DESIGNATION"] == article_to_edit].iloc[0]
            
            edit_qty = st.number_input("Modifier quantité :", value=int(row_edit["QUANTITE"]))
            edit_rem = st.text_input("Modifier remarque :", value=str(row_edit["REMARQUE"]))
            
            col_a, col_b = st.columns(2)
            with col_a:
                if st.button("Enregistrer les modifications"):
                    stock_df.loc[stock_df["DESIGNATION"] == article_to_edit, "QUANTITE"] = edit_qty
                    stock_df.loc[stock_df["DESIGNATION"] == article_to_edit, "REMARQUE"] = edit_rem
                    save_stock(stock_df)
                    st.success("Modifications enregistrées !")
            
            with col_b:
                if role == "ADMIN" and st.button("🗑️ Supprimer l'article"):
                    stock_df = stock_df[stock_df["DESIGNATION"] != article_to_edit]
                    save_stock(stock_df)
                    st.success("Article supprimé !")
                    st.rerun()

    # SECTION 4: GESTION UTILISATEURS (ADMIN ONLY)
    elif choice == "👥 Gestion Utilisateurs":
        st.markdown("<div class='main-header'>👥 Administration des Utilisateurs</div>", unsafe_allow_html=True)
        users_df = load_users()
        
        st.dataframe(users_df[["Login", "Nom", "Profil", "Actif"]], use_container_width=True)
        
        st.subheader("➕ Ajouter un utilisateur")
        u_login = st.text_input("Login :")
        u_pass = st.text_input("Mot de passe :")
        u_nom = st.text_input("Nom complet :")
        u_profil = st.selectbox("Rôle :", ["ADMIN", "MAGASINIER", "CONSULTATION"])
        
        if st.button("Créer l'utilisateur"):
            if u_login in users_df["Login"].values:
                st.error("Ce login existe déjà !")
            elif u_login and u_pass and u_nom:
                new_u = pd.DataFrame([{"Login": u_login, "MotDePasse": u_pass, "Nom": u_nom, "Profil": u_profil, "Actif": "OUI"}])
                users_df = pd.concat([users_df, new_u], ignore_index=True)
                save_users(users_df)
                st.success(f"Utilisateur {u_nom} créé avec succès !")
            else:
                st.warning("Veuillez remplir tous les champs.")
