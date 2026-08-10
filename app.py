from datetime import datetime, date
import os
import base64
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURATION ET THÈME VISUEL (Noir, Bleu, Blanc, Vert)
# ---------------------------------------------------------
st.set_page_config(
    page_title="gestion Stock MW NOMATIS",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Injection CSS Personnalisé
st.markdown(
    """
<style>
    /* Fond principal sombre et texte blanc */
    .stApp {
        background-color: #0F172A;
        color: #F8FAFC;
    }
    
    /* Titres et sous-titres */
    h1, h2, h3, h4, h5, h6, label {
        color: #F8FAFC !important;
    }

    /* Style des métriques (Vert) */
    [data-testid="stMetricValue"] {
        color: #10B981 !important;
    }
    
    /* Style global des boutons */
    .stButton>button {
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }

    /* Style du bouton SE CONNECTER (Rouge) */
    .btn-red button {
        background-color: #EF4444 !important;
        color: white !important;
        border: none !important;
        font-size: 16px !important;
        padding: 10px !important;
    }
    .btn-red button:hover {
        background-color: #DC2626 !important;
    }

    /* Style du bouton Accéder au Stock (Gras Bleu et Blanc) */
    .btn-stock button {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border: 2px solid #3B82F6 !important;
        font-size: 16px !important;
    }
    .btn-stock button:hover {
        background-color: #1D4ED8 !important;
        color: #FFFFFF !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# INITIALISATION DU SESSION STATE
# ---------------------------------------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_user" not in st.session_state:
    st.session_state.current_user = None
if "selected_client" not in st.session_state:
    st.session_state.selected_client = None

# Base de données utilisateurs
if "users_db" not in st.session_state:
    st.session_state.users_db = {
        "admin": {
            "password": "admin123",
            "name": "Administrateur",
            "role": "admin",
            "last_login": "Jamais",
        },
        "magasinier": {
            "password": "123",
            "name": "Magasinier Principal",
            "role": "magasinier",
            "last_login": "Jamais",
        },
    }

# Configuration Générale (Articles, Fournisseurs, Équipes)
if "config" not in st.session_state:
    st.session_state.config = {
        "articles": [
            "câble IF",
            "câble RJ45",
            "support 0.3 m",
            "support 0.6 m",
        ],
        "fournisseurs": ["NEC", "ZTE", "Intégral", "FO connect"],
        "equipes": ["Nabil Team", "Yassine Team", "Issam Team"],
    }

# Stocks et Registres de Bons
if "stock_db" not in st.session_state:
    st.session_state.stock_db = {
        "câble IF": 100,
        "câble RJ45": 250,
        "support 0.3 m": 40,
        "support 0.6 m": 25,
    }

if "be_list" not in st.session_state:
    st.session_state.be_list = []

if "bs_list" not in st.session_state:
    st.session_state.bs_list = []

# Paniers temporaires pour saisie multi-articles
if "temp_be_items" not in st.session_state:
    st.session_state.temp_be_items = []
if "temp_bs_items" not in st.session_state:
    st.session_state.temp_bs_items = []

CLIENTS_INFO = {
    "Orange": {"logo": "Orange_logo.svg.webp", "color": "#FF6600"},
    "Inwi": {"logo": "Logo INWI.jpg", "color": "#A1006B"},
    "ZTE": {"logo": "Logo ZTE.jpg", "color": "#005BAC"},
}


# ---------------------------------------------------------
# FONCTION DE GÉNÉRATION HTML POUR BE (MODELE EXACT IMAGE)
# ---------------------------------------------------------
def generate_be_html(be_data, client_logo_path):
    nomatis_logo_b64 = ""
    if os.path.exists("Logo Nomatis.jpg"):
        with open("Logo Nomatis.jpg", "rb") as f:
            nomatis_logo_b64 = base64.b64encode(f.read()).decode()

    client_logo_b64 = ""
    if os.path.exists(client_logo_path):
        with open(client_logo_path, "rb") as f:
            client_logo_b64 = base64.b64encode(f.read()).decode()

    items_rows = ""
    for item in be_data["items"]:
        items_rows += f"""
        <tr>
            <td style="border: 1px solid black; padding: 8px;">{item.get('Référence', '-')}</td>
            <td style="border: 1px solid black; padding: 8px; text-align: left;">{item.get('Article', '')}</td>
            <td style="border: 1px solid black; padding: 8px; text-align: center;">{item.get('Quantité', 0)}</td>
        </tr>
        """

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 30px; color: #000; background-color: #fff; }}
            .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; }}
            .address {{ font-size: 13px; line-height: 1.4; margin-top: 5px; }}
            .title {{ text-align: center; font-size: 24px; font-weight: bold; margin: 25px 0; }}
            .info-table, .items-table {{ width: 100%; border-collapse: collapse; text-align: center; }}
            .info-table th, .info-table td, .items-table th, .items-table td {{
                border: 2px solid black;
                padding: 8px;
                font-size: 13px;
            }}
            .info-table th, .items-table th {{ background-color: #f2f2f2; font-weight: bold; }}
            .items-table {{ margin-top: 20px; }}
        </style>
    </head>
    <body>

        <table class="header-table">
            <tr>
                <td style="width: 50%; vertical-align: top;">
                    {f'<img src="data:image/jpeg;base64,{nomatis_logo_b64}" height="55">' if nomatis_logo_b64 else '<h2>NOMATIS</h2>'}
                    <div class="address">
                        <strong>NOMATIS</strong><br>
                        32 Rue Al Hatim<br>
                        les Orangers<br>
                        10000
                    </div>
                </td>
                <td style="width: 50%; text-align: right; vertical-align: top;">
                    {f'<img src="data:image/jpeg;base64,{client_logo_b64}" height="55">' if client_logo_b64 else '<h2>LOGO CLIENT</h2>'}
                </td>
            </tr>
        </table>

        <div class="title">Bon d'entree</div>

        <table class="info-table">
            <thead>
                <tr>
                    <th>Bon de Livraison</th>
                    <th>Date</th>
                    <th>Fournisseur</th>
                    <th>Lieu de livraison</th>
                    <th>receptioné par</th>
                    <th>Stock</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>{be_data.get('id', '')}</td>
                    <td>{be_data.get('date_be', '')}</td>
                    <td>{be_data.get('fournisseur', '')}</td>
                    <td>{be_data.get('lieu_livraison', '')}</td>
                    <td>{be_data.get('receptionne_par', '')}</td>
                    <td>{be_data.get('client', '')}</td>
                </tr>
            </tbody>
        </table>

        <table class="items-table">
            <thead>
                <tr>
                    <th style="width: 25%;">Référence</th>
                    <th style="width: 60%;">Désignation</th>
                    <th style="width: 15%;">Qté</th>
                </tr>
            </thead>
            <tbody>
                {items_rows}
            </tbody>
        </table>

    </body>
    </html>
    """


# ---------------------------------------------------------
# ÉCRAN DE CONNEXION (LOGIN)
# ---------------------------------------------------------
if not st.session_state.logged_in:
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])

    with c2:
        with st.container(border=True):
            col_logo, col_title = st.columns([1, 3])
            with col_logo:
                if os.path.exists("Logo Nomatis.jpg"):
                    st.image("Logo Nomatis.jpg", width=85)
            with col_title:
                st.markdown(
                    "<h3 style='margin-top:10px;'>gestion Stock MW NOMATIS</h3>",
                    unsafe_allow_html=True,
                )

            st.write("---")
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")

            # Bouton de connexion ROUGE
            st.markdown('<div class="btn-red">', unsafe_allow_html=True)
            login_btn = st.button("SE CONNECTER", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

            if login_btn:
                if (
                    username in st.session_state.users_db
                    and st.session_state.users_db[username]["password"]
                    == password
                ):
                    st.session_state.logged_in = True
                    st.session_state.current_user = username
                    st.session_state.users_db[username]["last_login"] = (
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    )
                    st.success("Accès Validé !")
                    st.rerun()
                else:
                    st.error("Identifiants incorrects.")
    st.stop()


# ---------------------------------------------------------
# EN-TÊTE / NAVIGATION CLIENT
# ---------------------------------------------------------
head_col1, head_col2, head_col3 = st.columns([2, 3, 2])

with head_col1:
    if os.path.exists("Logo Nomatis.jpg"):
        st.image("Logo Nomatis.jpg", width=120)
    st.markdown(
        "<h3 style='margin:0;'>gestion Stock MW NOMATIS</h3>",
        unsafe_allow_html=True,
    )

with head_col3:
    user_info = st.session_state.users_db[st.session_state.current_user]
    st.write(f"👤 **{user_info['name']}** ({user_info['role'].upper()})")

    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        if st.session_state.selected_client and st.button("🔄 Changer Client"):
            st.session_state.selected_client = None
            st.rerun()
    with c_btn2:
        if st.button("🚪 Déconnexion"):
            st.session_state.logged_in = False
            st.session_state.selected_client = None
            st.rerun()

st.divider()


# ---------------------------------------------------------
# PAGE : SÉLECTION CLIENT & GESTION DES UTILISATEURS
# ---------------------------------------------------------
if not st.session_state.selected_client:
    st.subheader("Sélectionnez l'espace client :")

    cols = st.columns(3)
    for idx, (client_name, info) in enumerate(CLIENTS_INFO.items()):
        with cols[idx]:
            with st.container(border=True):
                # Images de mêmes dimensions
                if os.path.exists(info["logo"]):
                    st.image(info["logo"], use_container_width=True)
                else:
                    st.markdown(
                        f"<h1 style='text-align:center;'>{client_name[0]}</h1>",
                        unsafe_allow_html=True,
                    )

                st.markdown(
                    f"<h3 style='text-align:center; color:{info['color']};'>{client_name}</h3>",
                    unsafe_allow_html=True,
                )

                # Bouton Bleu et Blanc Gras
                st.markdown('<div class="btn-stock">', unsafe_allow_html=True)
                if st.button(
                    "Accéder au Stock",
                    key=f"acc_{client_name}",
                    use_container_width=True,
                ):
                    st.session_state.selected_client = client_name
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

    st.write("---")

    # GESTION DES COMPTES UTILISATEURS
    if user_info["role"] == "admin":
        st.subheader("🛠️ Admin : Gestion des Utilisateurs")
        tab_create, tab_edit = st.tabs(
            ["Créer un utilisateur", "Modifier / Voir les comptes"]
        )

        with tab_create:
            with st.form("form_create_user"):
                new_username = st.text_input("Identifiant")
                new_fullname = st.text_input("Nom complet")
                new_pass = st.text_input("Mot de passe", type="password")
                new_role = st.selectbox(
                    "Rôle",
                    ["admin", "magasinier", "coordinateur", "coordinatrice"],
                )
                if st.form_submit_button("Créer l'utilisateur"):
                    if new_username and new_pass:
                        st.session_state.users_db[new_username] = {
                            "password": new_pass,
                            "name": new_fullname,
                            "role": new_role,
                            "last_login": "Jamais",
                        }
                        st.success(
                            f"Utilisateur {new_username} créé avec succès !"
                        )
                        st.rerun()

        with tab_edit:
            user_data = []
            for u_id, u_info in st.session_state.users_db.items():
                user_data.append(
                    {
                        "Identifiant": u_id,
                        "Nom": u_info["name"],
                        "Rôle": u_info["role"],
                        "Dernière Connexion": u_info["last_login"],
                    }
                )
            st.dataframe(pd.DataFrame(user_data), use_container_width=True)

            selected_u = st.selectbox(
                "Choisir un compte à modifier",
                list(st.session_state.users_db.keys()),
            )
            u_to_mod = st.session_state.users_db[selected_u]

            with st.form("form_mod_user"):
                mod_name = st.text_input("Nom", value=u_to_mod["name"])
                mod_pass = st.text_input(
                    "Mot de passe", value=u_to_mod["password"]
                )
                mod_role = st.selectbox(
                    "Rôle",
                    ["admin", "magasinier", "coordinateur", "coordinatrice"],
                    index=[
                        "admin",
                        "magasinier",
                        "coordinateur",
                        "coordinatrice",
                    ].index(u_to_mod["role"]),
                )
                if st.form_submit_button("Mettre à jour le compte"):
                    st.session_state.users_db[selected_u]["name"] = mod_name
                    st.session_state.users_db[selected_u]["password"] = mod_pass
                    st.session_state.users_db[selected_u]["role"] = mod_role
                    st.success("Modifications enregistrées !")
                    st.rerun()

    else:
        st.subheader("⚙️ Modifier Mon Profil")
        with st.form("form_self_edit"):
            self_name = st.text_input("Nom complet", value=user_info["name"])
            self_pass = st.text_input(
                "Nouveau mot de passe", value=user_info["password"]
            )
            if st.form_submit_button("Mettre à jour"):
                st.session_state.users_db[st.session_state.current_user][
                    "name"
                ] = self_name
                st.session_state.users_db[st.session_state.current_user][
                    "password"
                ] = self_pass
                st.success("Profil mis à jour !")

    st.stop()


# ---------------------------------------------------------
# ESPACE CLIENT : LES 5 RUBRIQUES
# ---------------------------------------------------------
st.title(f"Espace Client : {st.session_state.selected_client}")

t_be, t_bs, t_stock, t_mods, t_config = st.tabs(
    [
        "📥 BE (Bon d'Entrée)",
        "📤 BS (Bon de Sortie)",
        "📊 Situation Stock",
        "✏️ Modification & Impression",
        "⚙️ Configuration",
    ]
)

# =========================================================
# 1. BON D'ENTRÉE (BE)
# =========================================================
with t_be:
    st.subheader("Créer un Bon d'Entrée (BE)")

    c_be1, c_be2, c_be3 = st.columns(3)
    with c_be1:
        saisie_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.text_input("Date/Heure Saisie (Auto)", value=saisie_dt, disabled=True)
    with c_be2:
        date_be = st.date_input(
            "Date du BE", max_value=date.today(), key="be_date"
        )
    with c_be3:
        num_be = st.text_input(
            "Bon de Livraison / N° BE",
            value=f"BE-{len(st.session_state.be_list)+1:04d}",
        )

    c_be4, c_be5, c_be6 = st.columns(3)
    with c_be4:
        fournis_opt = st.session_state.config["fournisseurs"] + [
            "Autre (Saisir)"
        ]
        sel_fournis = st.selectbox("Fournisseur", fournis_opt)
        fournisseur_final = (
            st.text_input("Nouveau Fournisseur")
            if sel_fournis == "Autre (Saisir)"
            else sel_fournis
        )
    with c_be5:
        lieu_livraison = st.text_input("Lieu de livraison", value="Magasin Principal")
    with c_be6:
        receptionne_par = st.text_input(
            "Réceptionné par", value=user_info["name"]
        )

    st.write("---")
    st.write("##### Articles du Bon d'Entrée")

    col_ref, col_art, col_qte, col_rem = st.columns([2, 3, 2, 3])
    with col_ref:
        ref_be = st.text_input("Référence", key="be_ref")
    with col_art:
        art_be = st.selectbox(
            "Désignation (Article)",
            st.session_state.config["articles"],
            key="be_art",
        )
    with col_qte:
        qte_be = st.number_input(
            "Quantité (> 0)", min_value=1, step=1, key="be_qte"
        )
    with col_rem:
        rem_be = st.text_input("Remarque", key="be_rem")

    if st.button("➕ Ajouter au bon"):
        st.session_state.temp_be_items.append(
            {
                "Référence": ref_be,
                "Article": art_be,
                "Quantité": qte_be,
                "Remarque": rem_be,
            }
        )

    if st.session_state.temp_be_items:
        st.write("##### Aperçu des articles :")
        st.dataframe(
            pd.DataFrame(st.session_state.temp_be_items),
            use_container_width=True,
        )

        if st.button("💾 Enregistrer le Bon d'Entrée"):
            new_be = {
                "id": num_be,
                "datetime_saisie": saisie_dt,
                "date_be": str(date_be),
                "fournisseur": fournisseur_final,
                "lieu_livraison": lieu_livraison,
                "receptionne_par": receptionne_par,
                "items": st.session_state.temp_be_items.copy(),
                "client": st.session_state.selected_client,
            }

            for item in st.session_state.temp_be_items:
                st.session_state.stock_db[item["Article"]] = (
                    st.session_state.stock_db.get(item["Article"], 0)
                    + item["Quantité"]
                )

            st.session_state.be_list.append(new_be)
            st.session_state.temp_be_items = []
            st.success(f"Bon d'Entrée {num_be} enregistré !")
            st.rerun()

# =========================================================
# 2. BON DE SORTIE (BS)
# =========================================================
with t_bs:
    st.subheader("Créer un Bon de Sortie (BS)")

    c_bs1, c_bs2, c_bs3 = st.columns(3)
    with c_bs1:
        st.text_input(
            "Date/Heure Saisie (Auto)",
            value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            disabled=True,
            key="bs_dt",
        )
    with c_bs2:
        date_bs = st.date_input(
            "Date du BS", max_value=date.today(), key="bs_date"
        )
    with c_bs3:
        num_bs = st.text_input(
            "N° de BS", value=f"BS-{len(st.session_state.bs_list)+1:04d}"
        )

    c_bs4, c_bs5 = st.columns(2)
    with c_bs4:
        equipe_bs = st.selectbox(
            "Équipe réceptrice", st.session_state.config["equipes"]
        )
    with c_bs5:
        dest_bs = st.text_input("Destination / Projet")

    st.write("---")
    st.write("##### Articles à sortir")

    col_art_s, col_qte_s, col_rem_s = st.columns([3, 2, 3])
    with col_art_s:
        art_bs = st.selectbox(
            "Article", st.session_state.config["articles"], key="bs_art"
        )
    with col_qte_s:
        stock_dispo = st.session_state.stock_db.get(art_bs, 0)
        st.caption(f"Stock disponible : **{stock_dispo}**")
        qte_bs = st.number_input(
            "Quantité", min_value=1, max_value=max(1, stock_dispo), key="bs_qte"
        )
    with col_rem_s:
        rem_bs = st.text_input("Remarque", key="bs_rem")

    if st.button("➕ Ajouter à la sortie"):
        if qte_bs > stock_dispo or stock_dispo <= 0:
            st.error("Quantité insuffisante en stock !")
        else:
            st.session_state.temp_bs_items.append(
                {
                    "Article": art_bs,
                    "Quantité": qte_bs,
                    "Remarque": rem_bs,
                }
            )

    if st.session_state.temp_bs_items:
        st.write("##### Aperçu des sorties :")
        st.dataframe(
            pd.DataFrame(st.session_state.temp_bs_items),
            use_container_width=True,
        )

        if st.button("💾 Enregistrer le Bon de Sortie"):
            if dest_bs:
                new_bs = {
                    "id": num_bs,
                    "datetime_saisie": datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                    "date_bs": str(date_bs),
                    "equipe": equipe_bs,
                    "destination": dest_bs,
                    "items": st.session_state.temp_bs_items.copy(),
                    "client": st.session_state.selected_client,
                }
                for item in st.session_state.temp_bs_items:
                    st.session_state.stock_db[item["Article"]] -= item[
                        "Quantité"
                    ]

                st.session_state.bs_list.append(new_bs)
                st.session_state.temp_bs_items = []
                st.success(f"Bon de Sortie {num_bs} enregistré !")
                st.rerun()

# =========================================================
# 3. SITUATION STOCK
# =========================================================
with t_stock:
    st.subheader(f"Situation du Stock - {st.session_state.selected_client}")

    records = []
    for art in st.session_state.config["articles"]:
        total_in = sum(
            item["Quantité"]
            for be in st.session_state.be_list
            if be["client"] == st.session_state.selected_client
            for item in be["items"]
            if item["Article"] == art
        )
        total_out = sum(
            item["Quantité"]
            for bs in st.session_state.bs_list
            if bs["client"] == st.session_state.selected_client
            for item in bs["items"]
            if item["Article"] == art
        )
        records.append(
            {
                "Article": art,
                "Total Entrées": total_in,
                "Total Sorties": total_out,
                "Stock Actuel": st.session_state.stock_db.get(art, 0),
            }
        )

    st.dataframe(pd.DataFrame(records), use_container_width=True)

# =========================================================
# 4. MODIFICATION ET IMPRESSION
# =========================================================
with t_mods:
    st.subheader("Consultation, Modification & Impression")

    type_bon = st.radio(
        "Type de bon :",
        ["Bons d'Entrée (BE)", "Bons de Sortie (BS)"],
        horizontal=True,
    )

    if type_bon == "Bons d'Entrée (BE)":
        client_bes = [
            b
            for b in st.session_state.be_list
            if b["client"] == st.session_state.selected_client
        ]
        if client_bes:
            sel_be_id = st.selectbox(
                "Choisir un BE :", [b["id"] for b in client_bes]
            )
            be_obj = next(b for b in client_bes if b["id"] == sel_be_id)

            st.write(
                f"**Date BE:** {be_obj['date_be']} | **Fournisseur:** {be_obj['fournisseur']} | **Réceptionné par:** {be_obj['receptionne_par']}"
            )
            st.dataframe(
                pd.DataFrame(be_obj["items"]), use_container_width=True
            )

            col_del, col_print = st.columns(2)
            with col_del:
                if st.button("❌ Supprimer ce BE"):
                    st.session_state.be_list = [
                        b
                        for b in st.session_state.be_list
                        if b["id"] != sel_be_id
                    ]
                    st.success("BE supprimé !")
                    st.rerun()

            with col_print:
                # Génération HTML conforme au modèle exact
                logo_client_path = CLIENTS_INFO[
                    st.session_state.selected_client
                ]["logo"]
                html_out = generate_be_html(be_obj, logo_client_path)

                st.download_button(
                    label="🖨️ Imprimer / Télécharger BE (Modèle Exact HTML/Word)",
                    data=html_out,
                    file_name=f"{be_obj['id']}.html",
                    mime="text/html",
                )
        else:
            st.info("Aucun Bon d'Entrée.")

    else:
        client_bss = [
            b
            for b in st.session_state.bs_list
            if b["client"] == st.session_state.selected_client
        ]
        if client_bss:
            sel_bs_id = st.selectbox(
                "Choisir un BS :", [b["id"] for b in client_bss]
            )
            bs_obj = next(b for b in client_bss if b["id"] == sel_bs_id)

            st.write(
                f"**Date BS:** {bs_obj['date_bs']} | **Équipe:** {bs_obj['equipe']} | **Destination:** {bs_obj['destination']}"
            )
            st.dataframe(
                pd.DataFrame(bs_obj["items"]), use_container_width=True
            )

            col_del2, col_print2 = st.columns(2)
            with col_del2:
                if st.button("❌ Supprimer ce BS"):
                    st.session_state.bs_list = [
                        b
                        for b in st.session_state.bs_list
                        if b["id"] != sel_bs_id
                    ]
                    st.success("BS supprimé !")
                    st.rerun()

            with col_print2:
                st.download_button(
                    label="🖨️ Imprimer / Télécharger BS",
                    data=f"BON DE SORTIE {bs_obj['id']}\nEquipe: {bs_obj['equipe']}\nDestination: {bs_obj['destination']}\nItems: {bs_obj['items']}",
                    file_name=f"{bs_obj['id']}.txt",
                )
        else:
            st.info("Aucun Bon de Sortie.")

# =========================================================
# 5. CONFIGURATION (ADMIN)
# =========================================================
with t_config:
    if user_info["role"] != "admin":
        st.warning("🔒 Section réservée à l'administrateur.")
    else:
        st.subheader("Configuration Système")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.write("##### Articles")
            n_art = st.text_input("Nouvel article")
            if st.button("Ajouter Article"):
                if (
                    n_art
                    and n_art not in st.session_state.config["articles"]
                ):
                    st.session_state.config["articles"].append(n_art)
                    st.session_state.stock_db[n_art] = 0
                    st.success("Ajouté !")
                    st.rerun()

        with c2:
            st.write("##### Fournisseurs")
            n_four = st.text_input("Nouveau fournisseur")
            if st.button("Ajouter Fournisseur"):
                if (
                    n_four
                    and n_four not in st.session_state.config["fournisseurs"]
                ):
                    st.session_state.config["fournisseurs"].append(n_four)
                    st.success("Ajouté !")
                    st.rerun()

        with c3:
            st.write("##### Équipes")
            n_eq = st.text_input("Nouvelle équipe")
            if st.button("Ajouter Équipe"):
                if n_eq and n_eq not in st.session_state.config["equipes"]:
                    st.session_state.config["equipes"].append(n_eq)
                    st.success("Ajoutée !")
                    st.rerun()

        st.divider()
        st.write("##### Ajustement direct du stock")
        a_mod = st.selectbox(
            "Article", st.session_state.config["articles"], key="cfg_art"
        )
        q_mod = st.number_input(
            "Nouvelle quantité",
            min_value=0,
            value=st.session_state.stock_db.get(a_mod, 0),
        )
        if st.button("Mettre à jour la quantité du stock"):
            st.session_state.stock_db[a_mod] = q_mod
            st.success("Stock mis à jour !")
