from datetime import datetime, date
import os
import base64
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURATION DE LA PAGE
# ---------------------------------------------------------
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
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

# Base de données utilisateurs par défaut
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
        "coord1": {
            "password": "123",
            "name": "Coordinateur Projet",
            "role": "coordinateur",
            "last_login": "Jamais",
        },
    }

# Configuration globale des référentiels (Gestion Admin)
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

# Stock par article (Initialisation)
if "stock_db" not in st.session_state:
    st.session_state.stock_db = {
        "câble IF": 100,
        "câble RJ45": 250,
        "support 0.3 m": 40,
        "support 0.6 m": 25,
    }

# Suivi des ajustements manuels par l'admin
if "manual_adjustments" not in st.session_state:
    st.session_state.manual_adjustments = {
        "câble IF": 0,
        "câble RJ45": 0,
        "support 0.3 m": 0,
        "support 0.6 m": 0,
    }

# Registres des Bons (BE et BS)
if "be_list" not in st.session_state:
    st.session_state.be_list = []
if "bs_list" not in st.session_state:
    st.session_state.bs_list = []

# Paniers temporaires de saisie multi-articles
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
# FONCTION DE GÉNÉRATION HTML POUR BONS (EXCEL FORMAT)
# ---------------------------------------------------------
def generate_be_html(be_data, client_logo_path, is_bs=False):
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
            <td style="border: 2px solid black; padding: 8px;">{item.get('Référence', '-')}</td>
            <td style="border: 2px solid black; padding: 8px; text-align: left;">{item.get('Article', '')}</td>
            <td style="border: 2px solid black; padding: 8px; text-align: center;">{item.get('Quantité', 0)}</td>
        </tr>
        """

    title_text = "Bon de sortie" if is_bs else "Bon d'entree"
    col3_header = "Équipe" if is_bs else "Fournisseur"
    col3_val = be_data.get("equipe", "") if is_bs else be_data.get("fournisseur", "")
    col4_header = "Destination" if is_bs else "Lieu de livraison"
    col4_val = be_data.get("destination", "") if is_bs else be_data.get("lieu_livraison", "")

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 30px; color: #000; background-color: #fff; }}
            .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; }}
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

        <div class="title">{title_text}</div>

        <table class="info-table">
            <thead>
                <tr>
                    <th>N° Bon</th>
                    <th>Date</th>
                    <th>{col3_header}</th>
                    <th>{col4_header}</th>
                    <th>Établi / Réceptionné par</th>
                    <th>Stock</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>{be_data.get('id', '')}</td>
                    <td>{be_data.get('date_be', be_data.get('date_bs', ''))}</td>
                    <td>{col3_val}</td>
                    <td>{col4_val}</td>
                    <td>{be_data.get('receptionne_par', st.session_state.users_db[st.session_state.current_user]['name'])}</td>
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


# =========================================================
# ÉCRAN DE CONNEXION (PAGE 1 : BLEU, BLANC, TOUCHES DE VERT)
# =========================================================
if not st.session_state.logged_in:
    st.markdown(
        """
    <style>
        .stApp {
            background-color: #0d1b2a;
            color: #ffffff;
        }
        h1, h2, h3, h4, label {
            color: #ffffff !important;
        }
        .login-card {
            background-color: #1b263b;
            padding: 30px;
            border-radius: 12px;
            border: 1px solid #415a77;
            box-shadow: 0 8px 16px rgba(0,0,0,0.4);
        }
        .btn-red button {
            background-color: #d90429 !important;
            color: #ffffff !important;
            border: none !important;
            font-size: 16px !important;
            font-weight: bold !important;
            border-radius: 6px !important;
            padding: 10px !important;
        }
    </style>
    """,
        unsafe_allow_html=True,
    )

    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])

    with c2:
        with st.container():
            st.markdown('<div class="login-card">', unsafe_allow_html=True)
            col_img, col_txt = st.columns([1, 3])
            with col_img:
                if os.path.exists("Logo Nomatis.jpg"):
                    st.image("Logo Nomatis.jpg", width=90)
                else:
                    st.markdown("## [LOGO]")
            with col_txt:
                st.markdown(
                    "<h2 style='margin-top:10px; color:#2ec4b6 !important;'>Gestion Stock MW NOMATIS</h2>",
                    unsafe_allow_html=True,
                )

            st.write("---")
            username_input = st.text_input("Nom d'utilisateur", key="user_in")
            password_input = st.text_input(
                "Mot de passe", type="password", key="pass_in"
            )

            st.markdown('<div class="btn-red">', unsafe_allow_html=True)
            login_clicked = st.button("SE CONNECTER", use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

            if login_clicked:
                if (
                    username_input in st.session_state.users_db
                    and st.session_state.users_db[username_input]["password"]
                    == password_input
                ):
                    st.session_state.logged_in = True
                    st.session_state.current_user = username_input
                    st.session_state.users_db[username_input][
                        "last_login"
                    ] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Bouton devient vert lors de la validation
                    st.markdown(
                        """
                    <style>
                        .btn-red button { background-color: #2ec4b6 !important; }
                    </style>
                    """,
                        unsafe_allow_html=True,
                    )
                    st.success("Accès validé ! Redirection...")
                    st.rerun()
                else:
                    st.error("Identifiants incorrects. Veuillez réessayer.")

            st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


# =========================================================
# APPLICATION PRINCIPALE APPRÈS CONNEXION
# =========================================================
current_user_info = st.session_state.users_db[st.session_state.current_user]
user_role = current_user_info["role"]

# ---------------------------------------------------------
# ÉCRAN DE SÉLECTION DU CLIENT (THÈME BLANC / ÉCRITURE NOIRE)
# ---------------------------------------------------------
if not st.session_state.selected_client:
    st.markdown(
        """
    <style>
        .stApp {
            background-color: #ffffff;
            color: #000000;
        }
        h1, h2, h3, h4, h5, h6, label, p, span, div {
            color: #000000 !important;
        }
        .client-card {
            border: 2px solid #e0e0e0;
            border-radius: 10px;
            padding: 15px;
            text-align: center;
            background-color: #f9f9f9;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }
        .btn-client-stock button {
            background-color: #0056b3 !important;
            color: #ffffff !important;
            font-weight: bold !important;
            border: none !important;
            border-radius: 5px !important;
            font-size: 15px !important;
            padding: 8px 16px !important;
        }
        .btn-client-stock button:hover {
            background-color: #003d82 !important;
            color: #ffffff !important;
        }
    </style>
    """,
        unsafe_allow_html=True,
    )

    c_head1, c_head2 = st.columns([3, 1])
    with c_head1:
        st.title("Gestion Stock MW NOMATIS")
        st.write(
            f"Bienvenue, **{current_user_info['name']}** ({user_role.upper()})"
        )
    with c_head2:
        if st.button("🚪 Déconnexion", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.selected_client = None
            st.rerun()

    st.divider()
    st.subheader("Sélectionnez l'espace client :")

    cols_clients = st.columns(3)
    for idx, (client_name, info) in enumerate(CLIENTS_INFO.items()):
        with cols_clients[idx]:
            st.markdown('<div class="client-card">', unsafe_allow_html=True)

            # Fixation de la taille d'image sans paramètre height invalide
            if os.path.exists(info["logo"]):
                st.image(info["logo"], width=120)
            else:
                st.markdown(
                    f"<h2 style='height:100px; line-height:100px;'>{client_name}</h2>",
                    unsafe_allow_html=True,
                )

            st.markdown(
                f"<h3 style='color:{info['color']} !important;'>{client_name}</h3>",
                unsafe_allow_html=True,
            )

            st.markdown(
                '<div class="btn-client-stock">', unsafe_allow_html=True
            )
            if st.button(
                "Accéder au stock",
                key=f"btn_acc_{client_name}",
                use_container_width=True,
            ):
                st.session_state.selected_client = client_name
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    # ---------------------------------------------------------
    # GESTION DES UTILISATEURS (PAGE CLIENT)
    # ---------------------------------------------------------
    if user_role == "admin":
        st.subheader("🛠️ Gestion des Utilisateurs (Administrateur)")
        tab_u_create, tab_u_manage = st.tabs(
            ["Créer un utilisateur", "Liste & Modifications"]
        )

        with tab_u_create:
            with st.form("form_create_user"):
                c_u1, c_u2 = st.columns(2)
                with c_u1:
                    new_u_id = st.text_input("Identifiant (login)")
                    new_u_name = st.text_input("Nom complet")
                with c_u2:
                    new_u_pass = st.text_input(
                        "Mot de passe", type="password"
                    )
                    new_u_role = st.selectbox(
                        "Rôle",
                        ["admin", "magasinier", "coordinateur", "coordinatrice"],
                    )

                if st.form_submit_button("➕ Enregistrer le nouvel utilisateur"):
                    if new_u_id and new_u_pass and new_u_name:
                        if new_u_id not in st.session_state.users_db:
                            st.session_state.users_db[new_u_id] = {
                                "password": new_u_pass,
                                "name": new_u_name,
                                "role": new_u_role,
                                "last_login": "Jamais",
                            }
                            st.success(
                                f"Compte '{new_u_id}' créé avec succès !"
                            )
                            st.rerun()
                        else:
                            st.error("Cet identifiant existe déjà.")
                    else:
                        st.error("Veuillez remplir tous les champs.")

        with tab_u_manage:
            u_table_data = []
            for u_key, u_val in st.session_state.users_db.items():
                u_table_data.append(
                    {
                        "Identifiant": u_key,
                        "Nom Complet": u_val["name"],
                        "Rôle": u_val["role"],
                        "Dernière Connexion": u_val.get(
                            "last_login", "Jamais"
                        ),
                    }
                )
            st.dataframe(pd.DataFrame(u_table_data), use_container_width=True)

            st.write("---")
            st.write("##### Modifier un compte existant")
            target_user = st.selectbox(
                "Sélectionner l'utilisateur à modifier",
                list(st.session_state.users_db.keys()),
            )
            selected_u_data = st.session_state.users_db[target_user]

            with st.form("form_edit_user"):
                edit_name = st.text_input(
                    "Nom complet", value=selected_u_data["name"]
                )
                edit_pass = st.text_input(
                    "Mot de passe", value=selected_u_data["password"]
                )
                edit_role = st.selectbox(
                    "Rôle",
                    ["admin", "magasinier", "coordinateur", "coordinatrice"],
                    index=[
                        "admin",
                        "magasinier",
                        "coordinateur",
                        "coordinatrice",
                    ].index(selected_u_data["role"]),
                )

                if st.form_submit_button("💾 Mettre à jour"):
                    st.session_state.users_db[target_user]["name"] = edit_name
                    st.session_state.users_db[target_user][
                        "password"
                    ] = edit_pass
                    st.session_state.users_db[target_user]["role"] = edit_role
                    st.success("Utilisateur mis à jour avec succès !")
                    st.rerun()

    else:
        st.subheader("⚙️ Mon Compte Utilisateur")
        with st.form("form_self_edit"):
            self_name = st.text_input(
                "Nom complet", value=current_user_info["name"]
            )
            self_pass = st.text_input(
                "Mot de passe", value=current_user_info["password"]
            )

            if st.form_submit_button("💾 Mettre à jour mes informations"):
                st.session_state.users_db[st.session_state.current_user][
                    "name"
                ] = self_name
                st.session_state.users_db[st.session_state.current_user][
                    "password"
                ] = self_pass
                st.success("Profil mis à jour !")
                st.rerun()

    st.stop()


# ---------------------------------------------------------
# DÉFINITION DU THÈME POUR LES 5 RUBRIQUES ESPACE CLIENT
# ---------------------------------------------------------
st.markdown(
    """
<style>
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
    }
    h1, h2, h3, h4, h5, h6, label, p, span {
        color: #f8fafc !important;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: bold;
    }
</style>
""",
    unsafe_allow_html=True,
)

c_top1, c_top2, c_top3 = st.columns([2, 3, 2])
with c_top1:
    if os.path.exists("Logo Nomatis.jpg"):
        st.image("Logo Nomatis.jpg", width=110)
with c_top2:
    st.title(f"Client : {st.session_state.selected_client}")
with c_top3:
    st.write(
        f"👤 **{current_user_info['name']}** ({user_role.upper()})"
    )
    c_sub1, c_sub2 = st.columns(2)
    with c_sub1:
        if st.button("🔄 Changer Client", use_container_width=True):
            st.session_state.selected_client = None
            st.rerun()
    with c_sub2:
        if st.button("🚪 Déconnexion", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.selected_client = None
            st.rerun()

st.divider()

# LES 5 RUBRIQUES
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
# RUBRIQUE 1 : BON D'ENTRÉE (BE)
# =========================================================
with t_be:
    st.subheader("Saisie d'un Bon d'Entrée (BE)")

    if user_role in ["coordinateur", "coordinatrice"]:
        st.warning(
            "🔒 Votre rôle ne vous permet pas de créer des Bons d'Entrée (Consultation et Impression uniquement)."
        )
    else:
        c_be1, c_be2, c_be3 = st.columns(3)
        with c_be1:
            dt_auto = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.text_input(
                "Date et Heure de Saisie (Auto)", value=dt_auto, disabled=True
            )
        with c_be2:
            date_be = st.date_input(
                "Date du BE", max_value=date.today(), key="be_date_input"
            )
        with c_be3:
            num_be = st.text_input(
                "Numéro de BE", value=f"BE-{len(st.session_state.be_list)+1:04d}"
            )

        c_be4, c_be5, c_be6 = st.columns(3)
        with c_be4:
            f_options = st.session_state.config["fournisseurs"] + [
                "Autre (Saisir)"
            ]
            sel_f = st.selectbox("Fournisseur", f_options)
            fournisseur_val = (
                st.text_input("Préciser Fournisseur")
                if sel_f == "Autre (Saisir)"
                else sel_f
            )
        with c_be5:
            lieu_liv = st.text_input(
                "Lieu de livraison", value="Magasin Principal"
            )
        with c_be6:
            st.text_input(
                "Réceptionné par (Auto)",
                value=current_user_info["name"],
                disabled=True,
            )

        st.write("---")
        st.write("##### Saisie des articles :")

        col_ref, col_art, col_qte, col_rem = st.columns([2, 3, 2, 3])
        with col_ref:
            ref_i = st.text_input("Référence", key="be_ref_in")
        with col_art:
            art_i = st.selectbox(
                "Article (Liste prédéfinie)",
                st.session_state.config["articles"],
                key="be_art_in",
            )
        with col_qte:
            qte_i = st.number_input(
                "Quantité (> 0)", min_value=1, step=1, key="be_qte_in"
            )
        with col_rem:
            rem_i = st.text_input("Remarque", key="be_rem_in")

        if st.button("➕ Ajouter l'article au bon"):
            if qte_i <= 0:
                st.error("La quantité doit être supérieure à 0.")
            else:
                found = False
                for item in st.session_state.temp_be_items:
                    if item["Article"] == art_i:
                        item["Quantité"] += qte_i
                        item["Remarque"] += f" | {rem_i}" if rem_i else ""
                        found = True
                        break
                if not found:
                    st.session_state.temp_be_items.append(
                        {
                            "Référence": ref_i,
                            "Article": art_i,
                            "Quantité": qte_i,
                            "Remarque": rem_i,
                        }
                    )
                st.success(f"Article {art_i} ajouté au bon.")

        if st.session_state.temp_be_items:
            st.write("##### Tableau des articles du Bon :")
            st.dataframe(
                pd.DataFrame(st.session_state.temp_be_items),
                use_container_width=True,
            )

            c_save_be, c_clr_be = st.columns([2, 1])
            with c_save_be:
                if st.button("💾 Enregistrer le Bon d'Entrée", use_container_width=True):
                    be_record = {
                        "id": num_be,
                        "datetime_saisie": dt_auto,
                        "date_be": str(date_be),
                        "fournisseur": fournisseur_val,
                        "lieu_livraison": lieu_liv,
                        "receptionne_par": current_user_info["name"],
                        "items": st.session_state.temp_be_items.copy(),
                        "client": st.session_state.selected_client,
                    }

                    for it in st.session_state.temp_be_items:
                        st.session_state.stock_db[it["Article"]] = (
                            st.session_state.stock_db.get(it["Article"], 0)
                            + it["Quantité"]
                        )

                    st.session_state.be_list.append(be_record)
                    st.session_state.temp_be_items = []
                    st.success(f"Bon d'Entrée {num_be} enregistré !")
                    st.rerun()

            with c_clr_be:
                if st.button("🗑️ Vider la liste", use_container_width=True):
                    st.session_state.temp_be_items = []
                    st.rerun()


# =========================================================
# RUBRIQUE 2 : BON DE SORTIE (BS)
# =========================================================
with t_bs:
    st.subheader("Saisie d'un Bon de Sortie (BS)")

    if user_role in ["coordinateur", "coordinatrice"]:
        st.warning(
            "🔒 Votre rôle ne vous permet pas de créer des Bons de Sortie (Consultation et Impression uniquement)."
        )
    else:
        c_bs1, c_bs2, c_bs3 = st.columns(3)
        with c_bs1:
            dt_bs_auto = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.text_input(
                "Date et Heure de Saisie (Auto)",
                value=dt_bs_auto,
                disabled=True,
                key="bs_dt_auto",
            )
        with c_bs2:
            date_bs = st.date_input(
                "Date du BS", max_value=date.today(), key="bs_date_input"
            )
        with c_bs3:
            num_bs = st.text_input(
                "Numéro de BS", value=f"BS-{len(st.session_state.bs_list)+1:04d}"
            )

        c_bs4, c_bs5 = st.columns(2)
        with c_bs4:
            eq_sel = st.selectbox(
                "Équipe réceptrice", st.session_state.config["equipes"]
            )
        with c_bs5:
            dest_bs = st.text_input("Destination / Projet")

        st.write("---")
        st.write("##### Saisie des articles à sortir :")

        col_bs_art, col_bs_qte, col_bs_rem = st.columns([3, 2, 3])
        with col_bs_art:
            art_bs_sel = st.selectbox(
                "Article", st.session_state.config["articles"], key="bs_art_sel"
            )
        with col_bs_qte:
            stk_dispo = st.session_state.stock_db.get(art_bs_sel, 0)
            st.caption(f"Stock disponible : **{stk_dispo}**")
            qte_bs_sel = st.number_input(
                "Quantité à sortir", min_value=1, step=1, key="bs_qte_sel"
            )
        with col_bs_rem:
            rem_bs_sel = st.text_input("Remarque", key="bs_rem_sel")

        if st.button("➕ Ajouter l'article à la sortie"):
            if qte_bs_sel <= 0:
                st.error("La quantité doit être supérieure à 0.")
            elif qte_bs_sel > stk_dispo:
                st.error(
                    f"Quantité insuffisante ! Stock disponible : {stk_dispo}"
                )
            else:
                found = False
                for item in st.session_state.temp_bs_items:
                    if item["Article"] == art_bs_sel:
                        if (item["Quantité"] + qte_bs_sel) > stk_dispo:
                            st.error(
                                "Le cumul dépasse le stock disponible !"
                            )
                            found = True
                            break
                        item["Quantité"] += qte_bs_sel
                        item["Remarque"] += (
                            f" | {rem_bs_sel}" if rem_bs_sel else ""
                        )
                        found = True
                        break
                if not found and qte_bs_sel <= stk_dispo:
                    st.session_state.temp_bs_items.append(
                        {
                            "Article": art_bs_sel,
                            "Quantité": qte_bs_sel,
                            "Remarque": rem_bs_sel,
                        }
                    )
                    st.success(f"Article {art_bs_sel} ajouté au bon.")

        if st.session_state.temp_bs_items:
            st.write("##### Tableau des articles du Bon de Sortie :")
            st.dataframe(
                pd.DataFrame(st.session_state.temp_bs_items),
                use_container_width=True,
            )

            c_save_bs, c_clr_bs = st.columns([2, 1])
            with c_save_bs:
                if st.button("💾 Enregistrer le Bon de Sortie", use_container_width=True):
                    if not dest_bs:
                        st.error("Veuillez renseigner la destination.")
                    else:
                        bs_record = {
                            "id": num_bs,
                            "datetime_saisie": dt_bs_auto,
                            "date_bs": str(date_bs),
                            "equipe": eq_sel,
                            "destination": dest_bs,
                            "items": st.session_state.temp_bs_items.copy(),
                            "client": st.session_state.selected_client,
                        }

                        for it in st.session_state.temp_bs_items:
                            st.session_state.stock_db[it["Article"]] -= it[
                                "Quantité"
                            ]

                        st.session_state.bs_list.append(bs_record)
                        st.session_state.temp_bs_items = []
                        st.success(f"Bon de Sortie {num_bs} enregistré !")
                        st.rerun()

            with c_clr_bs:
                if st.button("🗑️ Vider la liste BS", use_container_width=True):
                    st.session_state.temp_bs_items = []
                    st.rerun()


# =========================================================
# RUBRIQUE 3 : SITUATION DU STOCK
# =========================================================
with t_stock:
    st.subheader(
        f"📊 Situation du Stock à Jour — {st.session_state.selected_client}"
    )

    stock_summary = []
    for art in st.session_state.config["articles"]:
        tot_be = sum(
            it["Quantité"]
            for be in st.session_state.be_list
            if be["client"] == st.session_state.selected_client
            for it in be["items"]
            if it["Article"] == art
        )

        tot_bs = sum(
            it["Quantité"]
            for bs in st.session_state.bs_list
            if bs["client"] == st.session_state.selected_client
            for it in bs["items"]
            if it["Article"] == art
        )

        ajust_man = st.session_state.manual_adjustments.get(art, 0)
        stk_actuel = st.session_state.stock_db.get(art, 0)

        stock_summary.append(
            {
                "Article": art,
                "Total Entrées (BE)": tot_be,
                "Total Sorties (BS)": tot_bs,
                "Ajustements Manuels Admin": ajust_man,
                "Stock Actuel": stk_actuel,
            }
        )

    st.dataframe(pd.DataFrame(stock_summary), use_container_width=True)


# =========================================================
# RUBRIQUE 4 : MODIFICATION & IMPRESSION DES BONS
# =========================================================
with t_mods:
    st.subheader("✏️ Modification & Impression des Bons")

    bon_type = st.radio(
        "Sélectionnez le type de bon :",
        ["Bons d'Entrée (BE)", "Bons de Sortie (BS)"],
        horizontal=True,
    )

    if bon_type == "Bons d'Entrée (BE)":
        client_bes = [
            b
            for b in st.session_state.be_list
            if b["client"] == st.session_state.selected_client
        ]

        if not client_bes:
            st.info("Aucun Bon d'Entrée enregistré pour ce client.")
        else:
            sel_be_id = st.selectbox(
                "Choisir un Bon d'Entrée", [b["id"] for b in client_bes]
            )
            be_target = next(b for b in client_bes if b["id"] == sel_be_id)

            st.write("---")
            if user_role in ["admin", "magasinier"]:
                st.write("##### Modifier les informations du Bon :")
                with st.form("form_edit_be"):
                    e_fournis = st.text_input(
                        "Fournisseur", value=be_target["fournisseur"]
                    )
                    e_lieu = st.text_input(
                        "Lieu de livraison", value=be_target["lieu_livraison"]
                    )

                    if st.form_submit_button("💾 Enregistrer les modifications"):
                        be_target["fournisseur"] = e_fournis
                        be_target["lieu_livraison"] = e_lieu
                        st.success("Informations du BE mises à jour !")
                        st.rerun()

            st.write("##### Contenu du bon :")
            st.dataframe(
                pd.DataFrame(be_target["items"]), use_container_width=True
            )

            col_be_act1, col_be_act2 = st.columns(2)
            with col_be_act1:
                if user_role in ["admin", "magasinier"]:
                    if st.button("❌ Supprimer ce BE", use_container_width=True):
                        for it in be_target["items"]:
                            st.session_state.stock_db[it["Article"]] -= it[
                                "Quantité"
                            ]
                        st.session_state.be_list = [
                            b
                            for b in st.session_state.be_list
                            if b["id"] != sel_be_id
                        ]
                        st.success("BE supprimé et stock mis à jour !")
                        st.rerun()

            with col_be_act2:
                logo_path = CLIENTS_INFO[st.session_state.selected_client][
                    "logo"
                ]
                html_data = generate_be_html(be_target, logo_path, is_bs=False)
                st.download_button(
                    label="🖨️ Télécharger / Imprimer BE (Format HTML/Word)",
                    data=html_data,
                    file_name=f"{be_target['id']}.html",
                    mime="text/html",
                    use_container_width=True,
                )

    else:
        client_bss = [
            b
            for b in st.session_state.bs_list
            if b["client"] == st.session_state.selected_client
        ]

        if not client_bss:
            st.info("Aucun Bon de Sortie enregistré pour ce client.")
        else:
            sel_bs_id = st.selectbox(
                "Choisir un Bon de Sortie", [b["id"] for b in client_bss]
            )
            bs_target = next(b for b in client_bss if b["id"] == sel_bs_id)

            st.write("---")
            if user_role in ["admin", "magasinier"]:
                st.write("##### Modifier les informations du Bon :")
                with st.form("form_edit_bs"):
                    e_eq = st.selectbox(
                        "Équipe",
                        st.session_state.config["equipes"],
                        index=st.session_state.config["equipes"].index(
                            bs_target["equipe"]
                        )
                        if bs_target["equipe"]
                        in st.session_state.config["equipes"]
                        else 0,
                    )
                    e_dest = st.text_input(
                        "Destination", value=bs_target["destination"]
                    )

                    if st.form_submit_button("💾 Enregistrer les modifications"):
                        bs_target["equipe"] = e_eq
                        bs_target["destination"] = e_dest
                        st.success("Informations du BS mises à jour !")
                        st.rerun()

            st.write("##### Contenu du bon :")
            st.dataframe(
                pd.DataFrame(bs_target["items"]), use_container_width=True
            )

            col_bs_act1, col_bs_act2 = st.columns(2)
            with col_bs_act1:
                if user_role in ["admin", "magasinier"]:
                    if st.button("❌ Supprimer ce BS", use_container_width=True):
                        for it in bs_target["items"]:
                            st.session_state.stock_db[it["Article"]] += it[
                                "Quantité"
                            ]
                        st.session_state.bs_list = [
                            b
                            for b in st.session_state.bs_list
                            if b["id"] != sel_bs_id
                        ]
                        st.success("BS supprimé et stock réintégré !")
                        st.rerun()

            with col_bs_act2:
                logo_path = CLIENTS_INFO[st.session_state.selected_client][
                    "logo"
                ]
                html_bs_data = generate_be_html(bs_target, logo_path, is_bs=True)
                st.download_button(
                    label="🖨️ Télécharger / Imprimer BS (Format HTML/Word)",
                    data=html_bs_data,
                    file_name=f"{bs_target['id']}.html",
                    mime="text/html",
                    use_container_width=True,
                )


# =========================================================
# RUBRIQUE 5 : CONFIGURATION (RÉSERVÉE A L'ADMIN)
# =========================================================
with t_config:
    if user_role != "admin":
        st.error("🔒 Cette rubrique est strictly réservée à l'administrateur.")
    else:
        st.subheader("⚙️ Configuration Référentiels & Stock")

        col_cfg1, col_cfg2, col_cfg3 = st.columns(3)

        with col_cfg1:
            st.write("##### 📦 Articles")
            add_art = st.text_input("Nouvel article", key="cfg_add_art")
            if st.button("Ajouter Article"):
                if (
                    add_art
                    and add_art not in st.session_state.config["articles"]
                ):
                    st.session_state.config["articles"].append(add_art)
                    st.session_state.stock_db[add_art] = 0
                    st.success(f"Article '{add_art}' ajouté.")
                    st.rerun()

            st.write("---")
            del_art = st.selectbox(
                "Supprimer un article",
                st.session_state.config["articles"],
                key="cfg_del_art",
            )
            if st.button("Supprimer Article"):
                st.session_state.config["articles"].remove(del_art)
                st.session_state.stock_db.pop(del_art, None)
                st.success("Article supprimé.")
                st.rerun()

        with col_cfg2:
            st.write("##### 🏢 Fournisseurs")
            add_four = st.text_input("Nouveau fournisseur", key="cfg_add_four")
            if st.button("Ajouter Fournisseur"):
                if (
                    add_four
                    and add_four not in st.session_state.config["fournisseurs"]
                ):
                    st.session_state.config["fournisseurs"].append(add_four)
                    st.success(f"Fournisseur '{add_four}' ajouté.")
                    st.rerun()

            st.write("---")
            del_four = st.selectbox(
                "Supprimer Fournisseur",
                st.session_state.config["fournisseurs"],
                key="cfg_del_four",
            )
            if st.button("Supprimer Fournisseur"):
                st.session_state.config["fournisseurs"].remove(del_four)
                st.success("Fournisseur supprimé.")
                st.rerun()

        with col_cfg3:
            st.write("##### 👥 Équipes / Ressources")
            add_eq = st.text_input("Nouvelle équipe", key="cfg_add_eq")
            if st.button("Ajouter Équipe"):
                if add_eq and add_eq not in st.session_state.config["equipes"]:
                    st.session_state.config["equipes"].append(add_eq)
                    st.success(f"Équipe '{add_eq}' ajoutée.")
                    st.rerun()

            st.write("---")
            del_eq = st.selectbox(
                "Supprimer Équipe",
                st.session_state.config["equipes"],
                key="cfg_del_eq",
            )
            if st.button("Supprimer Équipe"):
                st.session_state.config["equipes"].remove(del_eq)
                st.success("Équipe supprimée.")
                st.rerun()

        st.divider()

        st.subheader("🛠️ Ajustement Manuel du Stock")
        c_adj1, c_adj2 = st.columns(2)
        with c_adj1:
            adj_art = st.selectbox(
                "Article à ajuster",
                st.session_state.config["articles"],
                key="adj_art_sel",
            )
        with c_adj2:
            current_stk = st.session_state.stock_db.get(adj_art, 0)
            new_stk_val = st.number_input(
                f"Nouveau stock pour {adj_art} (Actuel: {current_stk})",
                min_value=0,
                value=current_stk,
            )

        if st.button("💾 Valider l'Ajustement Manuel"):
            diff = new_stk_val - current_stk
            st.session_state.stock_db[adj_art] = new_stk_val
            st.session_state.manual_adjustments[adj_art] = (
                st.session_state.manual_adjustments.get(adj_art, 0) + diff
            )
            st.success(f"Stock de {adj_art} mis à jour à {new_stk_val} !")
            st.rerun()
