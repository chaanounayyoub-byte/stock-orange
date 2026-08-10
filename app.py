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
    initial_sidebar_state="collapsed",
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

if "stock_db" not in st.session_state:
    st.session_state.stock_db = {
        "câble IF": 100,
        "câble RJ45": 250,
        "support 0.3 m": 40,
        "support 0.6 m": 25,
    }

if "manual_adjustments" not in st.session_state:
    st.session_state.manual_adjustments = {
        "câble IF": 0,
        "câble RJ45": 0,
        "support 0.3 m": 0,
        "support 0.6 m": 0,
    }

if "be_list" not in st.session_state:
    st.session_state.be_list = []
if "bs_list" not in st.session_state:
    st.session_state.bs_list = []

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
# GÉNÉRATION DU MODÈLE D'IMPRESSION HTML (EXCEL TYPE)
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
            <td style="border: 1px solid #cbd5e1; padding: 10px;">{item.get('Référence', '-')}</td>
            <td style="border: 1px solid #cbd5e1; padding: 10px; text-align: left;">{item.get('Article', '')}</td>
            <td style="border: 1px solid #cbd5e1; padding: 10px; text-align: center; font-weight: bold;">{item.get('Quantité', 0)}</td>
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
            body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 40px; color: #1e293b; background-color: #fff; }}
            .header-table {{ width: 100%; border-collapse: collapse; margin-bottom: 25px; }}
            .address {{ font-size: 13px; color: #64748b; line-height: 1.5; margin-top: 8px; }}
            .title {{ text-align: center; font-size: 26px; font-weight: 800; text-transform: uppercase; color: #0f172a; margin: 30px 0; border-bottom: 2px solid #0f172a; padding-bottom: 10px; }}
            .info-table, .items-table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
            .info-table th, .info-table td, .items-table th, .items-table td {{
                border: 1px solid #94a3b8;
                padding: 10px;
                font-size: 13px;
            }}
            .info-table th, .items-table th {{ background-color: #f1f5f9; font-weight: bold; color: #0f172a; text-transform: uppercase; }}
        </style>
    </head>
    <body>

        <table class="header-table">
            <tr>
                <td style="width: 50%; vertical-align: top;">
                    {f'<img src="data:image/jpeg;base64,{nomatis_logo_b64}" height="60">' if nomatis_logo_b64 else '<h2 style="margin:0; color:#0284c7;">NOMATIS</h2>'}
                    <div class="address">
                        <strong>NOMATIS S.A.R.L</strong><br>
                        32 Rue Al Hatim, Les Orangers<br>
                        Rabat, Maroc
                    </div>
                </td>
                <td style="width: 50%; text-align: right; vertical-align: top;">
                    {f'<img src="data:image/jpeg;base64,{client_logo_b64}" height="60">' if client_logo_b64 else '<h2>CLIENT</h2>'}
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
                    <th>Agent</th>
                    <th>Espace Client</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><strong>{be_data.get('id', '')}</strong></td>
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
                    <th style="width: 60%;">Désignation Article</th>
                    <th style="width: 15%;">Quantité</th>
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
# PAGE 1 : CONNEXION (BLEU, BLANC, TOUCHES DE VERT, ROUGE)
# =========================================================
if not st.session_state.logged_in:
    st.markdown(
        """
    <style>
        .stApp {
            background: linear-gradient(135deg, #0b132b 0%, #1c2541 100%);
            color: #ffffff;
        }
        .login-wrapper {
            max-width: 450px;
            margin: 80px auto 0 auto;
            background: rgba(255, 255, 255, 0.05);
            backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 16px;
            padding: 40px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.4);
            text-align: center;
        }
        .login-title {
            font-size: 24px;
            font-weight: 700;
            color: #ffffff !important;
            margin-bottom: 5px;
        }
        .login-subtitle {
            font-size: 14px;
            color: #10b981 !important;
            font-weight: 600;
            margin-bottom: 25px;
        }
        /* Style des Inputs */
        div[data-baseweb="input"] {
            background-color: rgba(255, 255, 255, 0.08) !important;
            border-radius: 8px !important;
            border: 1px solid rgba(255, 255, 255, 0.2) !important;
            color: #ffffff !important;
        }
        div[data-baseweb="input"]:focus-within {
            border-color: #10b981 !important;
        }
        input {
            color: #ffffff !important;
        }
        label {
            color: #e2e8f0 !important;
            font-weight: 500;
        }
        /* Bouton rouge */
        .btn-connect button {
            background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
            color: #ffffff !important;
            border: none !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            border-radius: 8px !important;
            padding: 12px !important;
            transition: all 0.3s ease !important;
            margin-top: 15px;
        }
        .btn-connect button:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(239, 68, 68, 0.4) !important;
        }
        .btn-success button {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
            color: #ffffff !important;
            border: none !important;
            font-size: 16px !important;
            font-weight: 700 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
    </style>
    """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([1, 2, 1])

    with c2:
        st.markdown('<div class="login-wrapper">', unsafe_allow_html=True)

        if os.path.exists("Logo Nomatis.jpg"):
            st.image("Logo Nomatis.jpg", width=110)

        st.markdown(
            '<div class="login-title">Gestion Stock MW NOMATIS</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="login-subtitle">● Plateforme de Gestion Multi-Clients</div>',
            unsafe_allow_html=True,
        )

        username_input = st.text_input("Nom d'utilisateur", key="u_in")
        password_input = st.text_input(
            "Mot de passe", type="password", key="p_in"
        )

        st.markdown('<div class="btn-connect">', unsafe_allow_html=True)
        login_clicked = st.button(
            "SE CONNECTER", use_container_width=True, key="btn_log"
        )
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

                st.markdown(
                    """
                <style>
                    .btn-connect button {
                        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
                    }
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
# APPLICATION APRÈS CONNEXION
# =========================================================
current_user_info = st.session_state.users_db[st.session_state.current_user]
user_role = current_user_info["role"]

# ---------------------------------------------------------
# ÉCRAN DE SÉLECTION CLIENT (BLANC / NOIR ET BOUTON BLEU)
# ---------------------------------------------------------
if not st.session_state.selected_client:
    st.markdown(
        """
    <style>
        .stApp {
            background-color: #f8fafc;
            color: #0f172a;
        }
        h1, h2, h3, h4, h5, h6, label, p, span, div {
            color: #0f172a !important;
        }
        .header-bar {
            background-color: #ffffff;
            padding: 20px 30px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            margin-bottom: 30px;
        }
        .client-card-box {
            background: #ffffff;
            border: 2px solid #e2e8f0;
            border-radius: 16px;
            padding: 25px;
            text-align: center;
            transition: all 0.3s ease;
            box-shadow: 0 4px 12px rgba(0,0,0,0.02);
        }
        .client-card-box:hover {
            border-color: #2563eb;
            transform: translateY(-4px);
            box-shadow: 0 12px 24px rgba(37, 99, 235, 0.1);
        }
        .btn-acc-stock button {
            background-color: #1d4ed8 !important;
            color: #ffffff !important;
            font-weight: 700 !important;
            font-size: 15px !important;
            border: 2px solid #1d4ed8 !important;
            border-radius: 8px !important;
            padding: 10px 20px !important;
            transition: all 0.2s ease !important;
        }
        .btn-acc-stock button:hover {
            background-color: #ffffff !important;
            color: #1d4ed8 !important;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px;
            background-color: #f1f5f9;
            padding: 6px;
            border-radius: 10px;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 6px;
            font-weight: 600;
        }
    </style>
    """,
        unsafe_allow_html=True,
    )

    c_h1, c_h2 = st.columns([3, 1])
    with c_h1:
        st.title("Gestion Stock MW NOMATIS")
        st.markdown(
            f"Connecté en tant que : **{current_user_info['name']}** (`{user_role.upper()}`)"
        )
    with c_h2:
        if st.button("🚪 Déconnexion", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.selected_client = None
            st.rerun()

    st.divider()
    st.subheader("Sélectionnez l'espace client :")

    cols_clients = st.columns(3)
    for idx, (client_name, info) in enumerate(CLIENTS_INFO.items()):
        with cols_clients[idx]:
            st.markdown('<div class="client-card-box">', unsafe_allow_html=True)

            if os.path.exists(info["logo"]):
                st.image(info["logo"], width=130)
            else:
                st.markdown(
                    f"<h2 style='height:80px;'>{client_name}</h2>",
                    unsafe_allow_html=True,
                )

            st.markdown(
                f"<h3 style='color:{info['color']} !important; margin:15px 0;'>{client_name}</h3>",
                unsafe_allow_html=True,
            )

            st.markdown('<div class="btn-acc-stock">', unsafe_allow_html=True)
            if st.button(
                "Accéder au stock",
                key=f"btn_c_{client_name}",
                use_container_width=True,
            ):
                st.session_state.selected_client = client_name
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    # ---------------------------------------------------------
    # GESTION UTILISATEURS (PAGE D'ACCUEIL CLIENTS)
    # ---------------------------------------------------------
    if user_role == "admin":
        st.subheader("🛠️ Gestion des Utilisateurs (Administrateur)")
        t_user_create, t_user_list = st.tabs(
            ["Créer un utilisateur", "Liste & Modifications"]
        )

        with t_user_create:
            with st.form("form_create_user"):
                cu1, cu2 = st.columns(2)
                with cu1:
                    new_u_id = st.text_input("Identifiant (login)")
                    new_u_name = st.text_input("Nom complet")
                with cu2:
                    new_u_pass = st.text_input(
                        "Mot de passe", type="password"
                    )
                    new_u_role = st.selectbox(
                        "Rôle",
                        ["admin", "magasinier", "coordinateur", "coordinatrice"],
                    )

                if st.form_submit_button("➕ Enregistrer l'utilisateur"):
                    if new_u_id and new_u_pass and new_u_name:
                        if new_u_id not in st.session_state.users_db:
                            st.session_state.users_db[new_u_id] = {
                                "password": new_u_pass,
                                "name": new_u_name,
                                "role": new_u_role,
                                "last_login": "Jamais",
                            }
                            st.success(
                                f"Utilisateur '{new_u_id}' créé avec succès !"
                            )
                            st.rerun()
                        else:
                            st.error("Identifiant déjà existant.")
                    else:
                        st.error("Champs obligatoires manquants.")

        with t_user_list:
            u_data = []
            for k, v in st.session_state.users_db.items():
                u_data.append(
                    {
                        "Login": k,
                        "Nom Complet": v["name"],
                        "Rôle": v["role"],
                        "Dernière Connexion": v.get("last_login", "Jamais"),
                    }
                )
            st.dataframe(pd.DataFrame(u_data), use_container_width=True)

            st.write("---")
            st.write("##### Modifier un compte existant")
            target_user = st.selectbox(
                "Sélectionner l'utilisateur",
                list(st.session_state.users_db.keys()),
            )
            u_info = st.session_state.users_db[target_user]

            with st.form("form_edit_user"):
                e_name = st.text_input("Nom complet", value=u_info["name"])
                e_pass = st.text_input("Mot de passe", value=u_info["password"])
                e_role = st.selectbox(
                    "Rôle",
                    ["admin", "magasinier", "coordinateur", "coordinatrice"],
                    index=[
                        "admin",
                        "magasinier",
                        "coordinateur",
                        "coordinatrice",
                    ].index(u_info["role"]),
                )

                if st.form_submit_button("💾 Mettre à jour"):
                    st.session_state.users_db[target_user]["name"] = e_name
                    st.session_state.users_db[target_user]["password"] = e_pass
                    st.session_state.users_db[target_user]["role"] = e_role
                    st.success("Mise à jour effectuée !")
                    st.rerun()

    else:
        st.subheader("⚙️ Mon Compte Utilisateur")
        with st.form("form_self_edit"):
            s_name = st.text_input(
                "Nom complet", value=current_user_info["name"]
            )
            s_pass = st.text_input(
                "Mot de passe", value=current_user_info["password"]
            )

            if st.form_submit_button("💾 Sauvegarder les modifications"):
                st.session_state.users_db[st.session_state.current_user][
                    "name"
                ] = s_name
                st.session_state.users_db[st.session_state.current_user][
                    "password"
                ] = s_pass
                st.success("Informations mises à jour !")
                st.rerun()

    st.stop()


# ---------------------------------------------------------
# THÈME POUR LES 5 RUBRIQUES ESPACE CLIENT (SOMBRE PRO)
# ---------------------------------------------------------
st.markdown(
    """
<style>
    .stApp {
        background: #0f172a;
        color: #f8fafc;
    }
    h1, h2, h3, h4, h5, h6, label, p, span {
        color: #f8fafc !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #1e293b;
        padding: 8px;
        border-radius: 12px;
        border: 1px solid #334155;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #94a3b8 !important;
        font-weight: 600;
        padding: 10px 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
    }
    div[data-baseweb="input"], div[data-baseweb="select"] {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        color: #ffffff !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

# Header Espace Client
c_top1, c_top2, c_top3 = st.columns([2, 3, 2])
with c_top1:
    if os.path.exists("Logo Nomatis.jpg"):
        st.image("Logo Nomatis.jpg", width=110)
with c_top2:
    st.title(f"Espace Client : {st.session_state.selected_client}")
with c_top3:
    st.write(
        f"👤 **{current_user_info['name']}** (`{user_role.upper()}`)"
    )
    cb1, cb2 = st.columns(2)
    with cb1:
        if st.button("🔄 Changer Client", use_container_width=True):
            st.session_state.selected_client = None
            st.rerun()
    with cb2:
        if st.button("🚪 Déconnexion", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.selected_client = None
            st.rerun()

st.divider()

# LES 5 RUBRIQUES BONS ET STOCK
t_be, t_bs, t_stock, t_mods, t_config = st.tabs(
    [
        "📥 Bon d'Entrée (BE)",
        "📤 Bon de Sortie (BS)",
        "📊 Situation Stock",
        "✏️ Modifications & Impression",
        "⚙️ Configuration",
    ]
)

# =========================================================
# RUBRIQUE 1 : BON D'ENTRÉE (BE)
# =========================================================
with t_be:
    st.subheader("Création d'un Bon d'Entrée (BE)")

    if user_role in ["coordinateur", "coordinatrice"]:
        st.warning(
            "🔒 Rôle restreint : Consultation et impression des bons uniquement."
        )
    else:
        cbe1, cbe2, cbe3 = st.columns(3)
        with cbe1:
            dt_auto = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.text_input(
                "Date/Heure Saisie (Auto)", value=dt_auto, disabled=True
            )
        with cbe2:
            date_be = st.date_input(
                "Date du BE", max_value=date.today(), key="be_date_in"
            )
        with cbe3:
            num_be = st.text_input(
                "Numéro de BE", value=f"BE-{len(st.session_state.be_list)+1:04d}"
            )

        cbe4, cbe5, cbe6 = st.columns(3)
        with cbe4:
            f_options = st.session_state.config["fournisseurs"] + [
                "Autre (Saisir)"
            ]
            sel_f = st.selectbox("Fournisseur", f_options)
            fournisseur_val = (
                st.text_input("Saisir Fournisseur")
                if sel_f == "Autre (Saisir)"
                else sel_f
            )
        with cbe5:
            lieu_liv = st.text_input(
                "Lieu de livraison", value="Magasin Principal"
            )
        with cbe6:
            st.text_input(
                "Réceptionné par (Auto)",
                value=current_user_info["name"],
                disabled=True,
            )

        st.write("---")
        st.write("##### Articles du Bon :")

        cr, ca, cq, cm = st.columns([2, 3, 2, 3])
        with cr:
            ref_i = st.text_input("Référence", key="be_ref")
        with ca:
            art_i = st.selectbox(
                "Article", st.session_state.config["articles"], key="be_art"
            )
        with cq:
            qte_i = st.number_input(
                "Quantité (> 0)", min_value=1, step=1, key="be_qte"
            )
        with cm:
            rem_i = st.text_input("Remarque", key="be_rem")

        if st.button("➕ Ajouter l'article"):
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
            st.dataframe(
                pd.DataFrame(st.session_state.temp_be_items),
                use_container_width=True,
            )

            csave, cclr = st.columns([2, 1])
            with csave:
                if st.button("💾 Enregistrer le BE", use_container_width=True):
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
                    st.success(f"BE {num_be} enregistré !")
                    st.rerun()

            with cclr:
                if st.button("🗑️ Vider", use_container_width=True):
                    st.session_state.temp_be_items = []
                    st.rerun()


# =========================================================
# RUBRIQUE 2 : BON DE SORTIE (BS)
# =========================================================
with t_bs:
    st.subheader("Création d'un Bon de Sortie (BS)")

    if user_role in ["coordinateur", "coordinatrice"]:
        st.warning(
            "🔒 Rôle restreint : Consultation et impression des bons uniquement."
        )
    else:
        cbs1, cbs2, cbs3 = st.columns(3)
        with cbs1:
            dt_bs_auto = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.text_input(
                "Date/Heure Saisie (Auto)", value=dt_bs_auto, disabled=True
            )
        with cbs2:
            date_bs = st.date_input(
                "Date du BS", max_value=date.today(), key="bs_date_in"
            )
        with cbs3:
            num_bs = st.text_input(
                "Numéro de BS", value=f"BS-{len(st.session_state.bs_list)+1:04d}"
            )

        cbs4, cbs5 = st.columns(2)
        with cbs4:
            eq_sel = st.selectbox(
                "Équipe réceptrice", st.session_state.config["equipes"]
            )
        with cbs5:
            dest_bs = st.text_input("Destination / Projet")

        st.write("---")
        st.write("##### Articles du Bon :")

        csa, csq, csm = st.columns([3, 2, 3])
        with csa:
            art_bs_sel = st.selectbox(
                "Article", st.session_state.config["articles"], key="bs_art_sel"
            )
        with csq:
            stk_dispo = st.session_state.stock_db.get(art_bs_sel, 0)
            st.caption(f"Stock dispo : **{stk_dispo}**")
            qte_bs_sel = st.number_input(
                "Quantité à sortir", min_value=1, step=1, key="bs_qte_sel"
            )
        with csm:
            rem_bs_sel = st.text_input("Remarque", key="bs_rem_sel")

        if st.button("➕ Ajouter l'article à la sortie"):
            if qte_bs_sel <= 0:
                st.error("La quantité doit être supérieure à 0.")
            elif qte_bs_sel > stk_dispo:
                st.error(
                    f"Stock insuffisant ! Stock disponible : {stk_dispo}"
                )
            else:
                found = False
                for item in st.session_state.temp_bs_items:
                    if item["Article"] == art_bs_sel:
                        if (item["Quantité"] + qte_bs_sel) > stk_dispo:
                            st.error("Le cumul dépasse le stock !")
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
                    st.success(f"Article {art_bs_sel} ajouté.")

        if st.session_state.temp_bs_items:
            st.dataframe(
                pd.DataFrame(st.session_state.temp_bs_items),
                use_container_width=True,
            )

            csave_bs, cclr_bs = st.columns([2, 1])
            with csave_bs:
                if st.button("💾 Enregistrer le BS", use_container_width=True):
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
                        st.success(f"BS {num_bs} enregistré !")
                        st.rerun()

            with cclr_bs:
                if st.button("🗑️ Vider BS", use_container_width=True):
                    st.session_state.temp_bs_items = []
                    st.rerun()


# =========================================================
# RUBRIQUE 3 : SITUATION DU STOCK
# =========================================================
with t_stock:
    st.subheader(
        f"📊 Situation du Stock à Jour — Client {st.session_state.selected_client}"
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
# RUBRIQUE 4 : MODIFICATION & IMPRESSION BONS
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
            st.info("Aucun Bon d'Entrée disponible.")
        else:
            sel_be_id = st.selectbox(
                "Choisir un Bon d'Entrée", [b["id"] for b in client_bes]
            )
            be_target = next(b for b in client_bes if b["id"] == sel_be_id)

            st.write("---")
            if user_role in ["admin", "magasinier"]:
                with st.form("form_edit_be"):
                    e_fournis = st.text_input(
                        "Fournisseur", value=be_target["fournisseur"]
                    )
                    e_lieu = st.text_input(
                        "Lieu de livraison", value=be_target["lieu_livraison"]
                    )

                    if st.form_submit_button("💾 Sauvegarder modifications"):
                        be_target["fournisseur"] = e_fournis
                        be_target["lieu_livraison"] = e_lieu
                        st.success("Modifications enregistrées !")
                        st.rerun()

            st.dataframe(
                pd.DataFrame(be_target["items"]), use_container_width=True
            )

            cb_act1, cb_act2 = st.columns(2)
            with cb_act1:
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

            with cb_act2:
                logo_path = CLIENTS_INFO[st.session_state.selected_client][
                    "logo"
                ]
                html_data = generate_be_html(be_target, logo_path, is_bs=False)
                st.download_button(
                    label="🖨️ Imprimer / Télécharger BE (HTML/Word)",
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
            st.info("Aucun Bon de Sortie disponible.")
        else:
            sel_bs_id = st.selectbox(
                "Choisir un Bon de Sortie", [b["id"] for b in client_bss]
            )
            bs_target = next(b for b in client_bss if b["id"] == sel_bs_id)

            st.write("---")
            if user_role in ["admin", "magasinier"]:
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

                    if st.form_submit_button("💾 Sauvegarder modifications"):
                        bs_target["equipe"] = e_eq
                        bs_target["destination"] = e_dest
                        st.success("Modifications enregistrées !")
                        st.rerun()

            st.dataframe(
                pd.DataFrame(bs_target["items"]), use_container_width=True
            )

            cbs_act1, cbs_act2 = st.columns(2)
            with cbs_act1:
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
                        st.success("BS supprimé et stock restitué !")
                        st.rerun()

            with cbs_act2:
                logo_path = CLIENTS_INFO[st.session_state.selected_client][
                    "logo"
                ]
                html_bs_data = generate_be_html(bs_target, logo_path, is_bs=True)
                st.download_button(
                    label="🖨️ Imprimer / Télécharger BS (HTML/Word)",
                    data=html_bs_data,
                    file_name=f"{bs_target['id']}.html",
                    mime="text/html",
                    use_container_width=True,
                )


# =========================================================
# RUBRIQUE 5 : CONFIGURATION ADMIN
# =========================================================
with t_config:
    if user_role != "admin":
        st.error("🔒 Section réservée à l'administrateur.")
    else:
        st.subheader("⚙️ Configuration Référentiels & Stock")

        cfg1, cfg2, cfg3 = st.columns(3)

        with cfg1:
            st.write("##### 📦 Articles")
            add_art = st.text_input("Nouvel article", key="cfg_art_add")
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
                key="cfg_art_del",
            )
            if st.button("Supprimer Article"):
                st.session_state.config["articles"].remove(del_art)
                st.session_state.stock_db.pop(del_art, None)
                st.success("Article supprimé.")
                st.rerun()

        with cfg2:
            st.write("##### 🏢 Fournisseurs")
            add_four = st.text_input("Nouveau fournisseur", key="cfg_four_add")
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
                key="cfg_four_del",
            )
            if st.button("Supprimer Fournisseur"):
                st.session_state.config["fournisseurs"].remove(del_four)
                st.success("Fournisseur supprimé.")
                st.rerun()

        with cfg3:
            st.write("##### 👥 Équipes / Ressources")
            add_eq = st.text_input("Nouvelle équipe", key="cfg_eq_add")
            if st.button("Ajouter Équipe"):
                if add_eq and add_eq not in st.session_state.config["equipes"]:
                    st.session_state.config["equipes"].append(add_eq)
                    st.success(f"Équipe '{add_eq}' ajoutée.")
                    st.rerun()

            st.write("---")
            del_eq = st.selectbox(
                "Supprimer Équipe",
                st.session_state.config["equipes"],
                key="cfg_eq_del",
            )
            if st.button("Supprimer Équipe"):
                st.session_state.config["equipes"].remove(del_eq)
                st.success("Équipe supprimée.")
                st.rerun()

        st.divider()

        st.subheader("🛠️ Ajustement Manuel du Stock")
        cadj1, cadj2 = st.columns(2)
        with cadj1:
            adj_art = st.selectbox(
                "Article à ajuster",
                st.session_state.config["articles"],
                key="adj_art",
            )
        with cadj2:
            curr_stk = st.session_state.stock_db.get(adj_art, 0)
            new_stk_val = st.number_input(
                f"Nouveau stock pour {adj_art} (Actuel: {curr_stk})",
                min_value=0,
                value=curr_stk,
            )

        if st.button("💾 Appliquer l'Ajustement Manuel"):
            diff = new_stk_val - curr_stk
            st.session_state.stock_db[adj_art] = new_stk_val
            st.session_state.manual_adjustments[adj_art] = (
                st.session_state.manual_adjustments.get(adj_art, 0) + diff
            )
            st.success(f"Stock mis à jour pour {adj_art} ({new_stk_val}).")
            st.rerun()
