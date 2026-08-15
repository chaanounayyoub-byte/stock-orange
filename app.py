import io
import re
import sqlite3
from datetime import date, datetime

import docx
import pandas as pd
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

import streamlit as st

# =========================================================
# CONFIGURATION DE LA PAGE
# =========================================================
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS",
    page_icon="📦",
    layout="wide",
)

# =========================================================
# STYLE CSS PERSONNALISÉ
# =========================================================
st.markdown(
    """
<style>
    .login-header {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 15px;
        margin-bottom: 5px;
    }
    .login-title {
        color: #1a365d;
        font-size: 2.2rem;
        font-weight: bold;
        margin: 0;
    }
    .login-subtitle {
        color: #10b981;
        font-size: 1.1rem;
        text-align: center;
        margin-bottom: 25px;
        font-weight: 500;
    }
    .header-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 20px;
    }
    .client-orange { color: #FF6600 !important; font-weight: bold; }
    .client-inwi { color: #8A2BE2 !important; font-weight: bold; }
    .client-zte { color: #0052CC !important; font-weight: bold; }
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# FONCTION DE GÉNÉRATION DU PDF (BON D'ENTRÉE)
# =========================================================
def generate_pdf_bon_entree(bon_data, items_data, client_name):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30
    )
    elements = []
    styles = getSampleStyleSheet()

    # Dynamic Client Color
    color_map = {
        "ORANGE": colors.HexColor("#FF6600"),
        "INWI": colors.HexColor("#8A2BE2"),
        "ZTE": colors.HexColor("#0052CC"),
    }
    client_color = color_map.get(client_name.upper(), colors.navy)

    # 1. En-tête (Logos + Adresse)
    company_info = "<b>NOMATIS</b><br/>32 Rue Al Hatim<br/>les Orangers<br/>10000"
    header_data = [
        [
            Paragraph("<b>// NOMATIS</b>", styles["Heading2"]),
            "",
            Paragraph(f"<b>LOGO {client_name.upper()}</b>", styles["Heading2"]),
        ],
        [Paragraph(company_info, styles["Normal"]), "", ""],
    ]
    header_table = Table(header_data, colWidths=[200, 150, 180])
    header_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (2, 0), (2, 0), "RIGHT"),
            ]
        )
    )
    elements.append(header_table)
    elements.append(Table([[""]], colWidths=[530], rowHeights=[20]))

    # 2. Titre du Bon
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        alignment=1,
        textColor=colors.black,
        fontSize=18,
    )
    elements.append(Paragraph("<b>Bon d'entrée</b>", title_style))
    elements.append(Table([[""]], colWidths=[530], rowHeights=[15]))

    # 3. Métadonnées du Bon
    meta_data = [
        ["N° Bon", "Date", "Fournisseur", "Lieu de livraison", "Réceptionné par", "Stock"],
        [
            bon_data.get("id", ""),
            bon_data.get("date_bon", ""),
            bon_data.get("fournisseur", ""),
            bon_data.get("lieu_livraison", ""),
            bon_data.get("user", ""),
            bon_data.get("stock", ""),
        ],
    ]
    meta_table = Table(meta_data, colWidths=[80, 80, 100, 100, 90, 80])
    meta_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    elements.append(meta_table)
    elements.append(Table([[""]], colWidths=[530], rowHeights=[15]))

    # 4. Tableau des Articles
    items_table_data = [["Référence", "Désignation", "Qté"]]
    for item in items_data:
        items_table_data.append([item["reference"], item["designation"], item["qte"]])

    items_table = Table(items_table_data, colWidths=[130, 320, 80])
    items_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("ALIGN", (0, 0), (0, -1), "LEFT"),
                ("ALIGN", (1, 0), (1, -1), "LEFT"),
                ("ALIGN", (2, 0), (2, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    elements.append(items_table)
    elements.append(Table([[""]], colWidths=[530], rowHeights=[40]))

    # 5. Pied de page (Signature & Date de saisie)
    footer_data = [
        ["Signature / Cachet Magasinier", ""],
        ["", ""],
        [f"Saisie le : {datetime.now().strftime('%Y-%m-%d %H:%M')}", ""],
    ]
    footer_table = Table(footer_data, colWidths=[300, 230], rowHeights=[20, 40, 20])
    footer_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    elements.append(footer_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer


# =========================================================
# ÉCRAN DE CONNEXION
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(
            """
            <div class="login-header">
                <span style="font-size: 2rem;">🔹</span>
                <h1 class="login-title">Gestion Stock MW NOMATIS</h1>
            </div>
            <div class="login-subtitle">
                espace de gestion et suivi de stock MW
            </div>
        """,
            unsafe_allow_html=True,
        )

        with st.form("login_form"):
            st.subheader("Connexion Sécurisée")
            username = st.text_input("Identifiant / Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("SE CONNECTER", use_container_width=True)

            if submit:
                if username and password:  # Remplacer par votre logique d'authentification
                    st.session_state.authenticated = True
                    st.session_state.user = username
                    st.session_state.role = "ADMIN"
                    st.session_state.client = "ORANGE"  # Valeur par défaut
                    st.rerun()
                else:
                    st.error("Identifiants invalides.")
    st.stop()

# =========================================================
# HEADER APPLICATIF (POST-CONNEXION)
# =========================================================
selected_client = st.sidebar.selectbox(
    "Espace Client", ["ORANGE", "INWI", "ZTE"], key="client_select"
)

client_css_class = f"client-{selected_client.lower()}"

st.markdown(
    f"""
    <div class="header-card">
        <h2 style="margin:0;">
            Gestion Stock MW NOMATIS — 
            <span class="{client_css_class}">ESPACE {selected_client}</span>
        </h2>
        <p style="margin:5px 0 0 0; color:#10b981; font-weight:500;">
            Utilisateur : {st.session_state.get('user', 'Utilisateur')} ({st.session_state.get('role', 'USER')})
        </p>
    </div>
""",
    unsafe_allow_html=True,
)

# Exemple de bouton de téléchargement de Bon de Test
st.subheader("📄 Test de Génération de Bon")
sample_bon = {
    "id": "BE-2026-001",
    "date_bon": "2026-08-15",
    "fournisseur": "Fournisseur A",
    "lieu_livraison": "Magasin Principal",
    "user": st.session_state.get("user", "Admin"),
    "stock": "Stock MW",
}

sample_items = [
    {"reference": "REF-001", "designation": "Antenne MW 0.6m", "qte": 4},
    {"reference": "REF-002", "designation": "ODU Radio Unit", "qte": 2},
]

pdf_buffer = generate_pdf_bon_entree(sample_bon, sample_items, selected_client)

st.download_button(
    label=f"📥 Télécharger Bon d'Entrée ({selected_client})",
    data=pdf_buffer,
    file_name=f"Bon_Entree_{selected_client}_{sample_bon['id']}.pdf",
    mime="application/pdf",
)
