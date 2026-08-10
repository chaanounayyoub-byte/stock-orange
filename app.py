import os
import pandas as pd
import streamlit as st

# Config de la page
st.set_page_config(
    page_title="Gestion de Stock multi-clients", page_icon="📦", layout="wide"
)

# ---------------------------------------------------------
# DICTIONNAIRE CLIENTS & LOGOS LOCAUX
# ---------------------------------------------------------
CLIENTS = {
    "Orange": {
        "color": "#FF6600",
        "logo": "Orange_logo.svg.webp",
    },
    "Inwi": {
        "color": "#A1006B",
        "logo": "Logo INWI.jpg",
    },
    "ZTE": {
        "color": "#005BAC",
        "logo": "Logo ZTE.jpg",
    },
}

# Initialisation du Session State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user" not in st.session_state:
    st.session_state.user = None
if "selected_client" not in st.session_state:
    st.session_state.selected_client = None


# ---------------------------------------------------------
# ÉCRAN DE CONNEXION (LOGIN)
# ---------------------------------------------------------
if not st.session_state.logged_in:
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.container(border=True):
            # En-tête avec logo Nomatis si présent
            if os.path.exists("Logo Nomatis.jpg"):
                st.image("Logo Nomatis.jpg", use_container_width=True)
            else:
                st.title("📦 Application Stock")

            st.subheader("Connexion")
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")

            if st.button("Se connecter", use_container_width=True):
                if username == "admin" and password == "admin123":
                    st.session_state.logged_in = True
                    st.session_state.user = username
                    st.success("Connexion réussie !")
                    st.rerun()
                else:
                    st.error("Identifiants incorrects")
    st.stop()


# ---------------------------------------------------------
# BARRE D'EN-TÊTE / PROFIL & DÉCONNEXION
# ---------------------------------------------------------
c_head1, c_head2, c_head3 = st.columns([2, 4, 2])
with c_head1:
    if os.path.exists("Logo Nomatis.jpg"):
        st.image("Logo Nomatis.jpg", width=150)
    else:
        st.markdown("### 📦 Nomatis Stock")

with c_head3:
    st.write(f"👤 Connecté en tant que : **{st.session_state.user}**")
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.session_state.selected_client:
            if st.button("🔄 Changer Client"):
                st.session_state.selected_client = None
                st.rerun()
    with col_b2:
        if st.button("🚪 Déconnexion"):
            st.session_state.logged_in = False
            st.session_state.selected_client = None
            st.rerun()

st.divider()


# ---------------------------------------------------------
# ÉCRAN DE SÉLECTION DU CLIENT
# ---------------------------------------------------------
if not st.session_state.selected_client:
    st.title(f"👋 Bienvenue, {st.session_state.user}")
    st.subheader("Sélectionnez l'espace Client :")
    st.write("")

    cols = st.columns(3)
    for idx, (client_name, info) in enumerate(CLIENTS.items()):
        with cols[idx]:
            with st.container(border=True):
                # Affichage sécurisé de l'image locale uploadée
                if os.path.exists(info["logo"]):
                    st.image(info["logo"], use_container_width=True)
                else:
                    st.warning(f"Image {info['logo']} introuvable")

                st.markdown(
                    f"<h2 style='text-align: center; color: {info['color']};"
                    f" margin-top: 10px;'>{client_name}</h2>",
                    unsafe_allow_html=True,
                )
                st.write("")
                if st.button(
                    f"Accéder au Stock {client_name}",
                    key=f"btn_{client_name}",
                    use_container_width=True,
                ):
                    st.session_state.selected_client = client_name
                    st.rerun()
    st.stop()


# ---------------------------------------------------------
# ESPACE CLIENT SÉLECTIONNÉ
# ---------------------------------------------------------
current_client = st.session_state.selected_client
client_info = CLIENTS[current_client]

st.title(f"Espace de Gestion : {current_client}")

# Exemple de navigation dans l'espace client
tabs = st.tabs(
    [
        "📊 Tableau de bord",
        "📥 Entrées Stock",
        "📤 Sorties Stock",
        "📋 Inventaire",
    ]
)

with tabs[0]:
    st.subheader(f"Statistiques générales - {current_client}")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Articles", "1,240", "12")
    col2.metric("Entrées ce mois", "350", "5%")
    col3.metric("Sorties ce mois", "180", "-2%")

with tabs[1]:
    st.subheader("Ajouter du Stock")
    # Votre logique d'ajout de stock...

with tabs[2]:
    st.subheader("Retirer du Stock")
    # Votre logique de sortie de stock...

with tabs[3]:
    st.subheader("Inventaire Actuel")
    # Exemple de tableau d'inventaire
    sample_data = pd.DataFrame(
        {
            "Code Article": ["ART-001", "ART-002", "ART-003"],
            "Désignation": ["Câble Fibre 10m", "Routeur Wi-Fi", "Connecteur RJ45"],
            "Quantité": [150, 42, 1200],
            "Statut": ["En stock", "Faible stock", "En stock"],
        }
    )
    st.dataframe(sample_data, use_container_width=True)
