import io
import sqlite3
from datetime import date, datetime

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

# =========================================================
# CONFIGURATION DE LA PAGE & STYLES CSS
# =========================================================
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS",
    page_icon="📦",
    layout="wide",
)

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
# BASE DE DONNÉES (MODÈLE DÉMO)
# =========================================================
conn = sqlite3.connect("stock.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute(
    """
CREATE TABLE IF NOT EXISTS bons (
    id TEXT PRIMARY KEY,
    type_bon TEXT,
    date_bon TEXT,
    fournisseur TEXT,
    lieu_livraison TEXT,
    equipe TEXT,
    destination TEXT,
    user TEXT,
    stock TEXT
)
"""
)
conn.commit()


def execute(query, params=()):
    cursor.execute(query, params)
    conn.commit()


def active_names(table_type):
    return ["Fournisseur A", "Fournisseur B", "Equipe 1", "Equipe 2"]


# =========================================================
# GÉNÉRATION PDF (CONFORME AUX CAPTURES)
# =========================================================
def generate_pdf_bon_entree(bon_data, items_data, client_name):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20,
    )
    elements = []
    styles = getSampleStyleSheet()

    color_map = {
        "ORANGE": "#FF6600",
        "INWI": "#8A2BE2",
        "ZTE": "#0052CC",
    }
    client_color_hex = color_map.get(client_name.upper(), "#000000")

    client_header_style = ParagraphStyle(
        "ClientHeader",
        parent=styles["Heading2"],
        alignment=2,
        textColor=colors.HexColor(client_color_hex),
        fontSize=14,
    )

    company_info = "<b>NOMATIS</b><br/>32 Rue Al Hatim<br/>les Orangers<br/>10000"
    header_data = [
        [
            Paragraph("<b>// NOMATIS</b>", styles["Heading1"]),
            "",
            Paragraph(
                f"<b>logo {client_name.lower()}</b>", client_header_style
            ),
        ],
        [Paragraph(company_info, styles["Normal"]), "", ""],
    ]
    header_table = Table(header_data, colWidths=[200, 150, 200])
    header_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (2, 0), (2, 0), "RIGHT"),
            ]
        )
    )
    elements.append(header_table)
    elements.append(Table([[""]], colWidths=[550], rowHeights=[15]))

    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        alignment=1,
        textColor=colors.black,
        fontSize=16,
    )
    elements.append(Paragraph("<b>Bon d'entrée</b>", title_style))
    elements.append(Table([[""]], colWidths=[550], rowHeights=[10]))

    meta_data = [
        [
            "N° Bon",
            "Date",
            "Founisseur",
            "Lieu de livraison",
            "receptioné par",
            "Stock",
        ],
        [
            str(bon_data.get("id", "")),
            str(bon_data.get("date_bon", "")),
            str(bon_data.get("fournisseur", "")),
            str(bon_data.get("lieu_livraison", "")),
            str(bon_data.get("user", "")),
            str(bon_data.get("stock", "")),
        ],
    ]
    meta_table = Table(meta_data, colWidths=[90, 90, 100, 110, 90, 70])
    meta_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    elements.append(meta_table)
    elements.append(Table([[""]], colWidths=[550], rowHeights=[10]))

    items_table_data = [["Référence", "Désignation", "Qté"]]
    for item in items_data:
        items_table_data.append(
            [
                str(item.get("reference", "")),
                str(item.get("designation", "")),
                str(item.get("qte", "")),
            ]
        )

    while len(items_table_data) < 5:
        items_table_data.append(["", "", ""])

    items_table = Table(items_table_data, colWidths=[120, 350, 80])
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
    elements.append(Table([[""]], colWidths=[550], rowHeights=[30]))

    footer_data = [
        ["Signature / Cachet Magasinier", ""],
        ["", ""],
        [f"Saisie le : {datetime.now().strftime('%Y-%m-%d %H:%M')}", ""],
    ]
    footer_table = Table(
        footer_data, colWidths=[300, 250], rowHeights=[15, 40, 15]
    )
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
# CONNEXION ET AUTHENTIFICATION
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    _, col2, _ = st.columns([1, 2, 1])
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
            submit = st.form_submit_button(
                "SE CONNECTER", use_container_width=True
            )

            if submit:
                if username and password:
                    st.session_state.authenticated = True
                    st.session_state.user = username
                    st.session_state.role = "ADMIN"
                    st.session_state.client = "ORANGE"
                    st.rerun()
                else:
                    st.error("Identifiants invalides.")
    st.stop()

# =========================================================
# BARRE LATÉRALE ET EN-TÊTE DYNAMIQUE
# =========================================================
selected_client = st.sidebar.selectbox(
    "Espace Client", ["ORANGE", "INWI", "ZTE"], key="client_select"
)
st.session_state.client = selected_client
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

# =========================================================
# SECTION HISTORIQUE & FORMULAIRE D'ÉDITION CORRIGÉ
# =========================================================
st.subheader("📜 Historique des Bons")

# Création d'une variable bon_detail fictive pour la démonstration
selected_bon_id = "BE-2026-001"
target_type = "BE"
bon_detail = {
    "id": selected_bon_id,
    "date_bon": "2026-08-15",
    "fournisseur": "Fournisseur A",
    "lieu_livraison": "Magasin Principal",
    "user": st.session_state.user,
    "stock": "Stock MW",
}

with st.expander("✏️ Modifier les informations du Bon"):
    if "bon_detail" in locals() and bon_detail:
        with st.form("form_edit_bon"):
            raw_date = bon_detail.get("date_bon", "")
            try:
                parsed_date = datetime.strptime(
                    str(raw_date), "%Y-%m-%d"
                ).date()
            except (ValueError, TypeError):
                parsed_date = date.today()

            today = date.today()
            safe_value = min(parsed_date, today)

            mod_date = st.date_input(
                "Nouvelle Date",
                value=safe_value,
                max_value=today,
                key=f"edit_date_{selected_bon_id}",
            )

            if target_type == "BE":
                fournisseurs_list = active_names("fournisseurs")
                current_fourn = bon_detail.get("fournisseur", "")
                fourn_index = (
                    fournisseurs_list.index(current_fourn)
                    if current_fourn in fournisseurs_list
                    else 0
                )

                mod_fourn = st.selectbox(
                    "Fournisseur",
                    fournisseurs_list,
                    index=fourn_index,
                    key=f"edit_fourn_{selected_bon_id}",
                )
                mod_lieu = st.text_input(
                    "Lieu Livraison",
                    value=bon_detail.get("lieu_livraison") or "",
                    key=f"edit_lieu_{selected_bon_id}",
                )
            else:
                equipes_list = active_names("equipes")
                current_eq = bon_detail.get("equipe", "")
                eq_index = (
                    equipes_list.index(current_eq)
                    if current_eq in equipes_list
                    else 0
                )

                mod_eq = st.selectbox(
                    "Équipe",
                    equipes_list,
                    index=eq_index,
                    key=f"edit_eq_{selected_bon_id}",
                )
                mod_dest = st.text_input(
                    "Destination",
                    value=bon_detail.get("destination") or "",
                    key=f"edit_dest_{selected_bon_id}",
                )

            if st.form_submit_button("Valider les modifications"):
                st.session_state[f"confirm_edit_{selected_bon_id}"] = True

        if st.session_state.get(f"confirm_edit_{selected_bon_id}", False):
            st.warning(
                "Confirmez-vous la modification des informations de ce bon ?"
            )
            col_yes, col_no = st.columns(2)
            with col_yes:
                if st.button(
                    "✅ Oui, Confirmer", key=f"btn_confirm_{selected_bon_id}"
                ):
                    if target_type == "BE":
                        execute(
                            "UPDATE bons SET date_bon=?, fournisseur=?, lieu_livraison=? WHERE id=?",
                            (
                                str(mod_date),
                                mod_fourn,
                                mod_lieu,
                                selected_bon_id,
                            ),
                        )
                    else:
                        execute(
                            "UPDATE bons SET date_bon=?, equipe=?, destination=? WHERE id=?",
                            (
                                str(mod_date),
                                mod_eq,
                                mod_dest,
                                selected_bon_id,
                            ),
                        )
                    st.session_state[f"confirm_edit_{selected_bon_id}"] = False
                    st.success("Bon mis à jour !")
                    st.rerun()

            with col_no:
                if st.button("❌ Annuler", key=f"btn_cancel_{selected_bon_id}"):
                    st.session_state[f"confirm_edit_{selected_bon_id}"] = False
                    st.rerun()
    else:
        st.info("Veuillez sélectionner un bon valide dans la liste.")

# Téléchargement PDF
items_sample = [
    {"reference": "REF-MW-01", "designation": "Antenne MW 0.6m", "qte": "2"},
    {"reference": "REF-MW-02", "designation": "ODU Radio Unit", "qte": "1"},
]

pdf_file = generate_pdf_bon_entree(
    bon_detail, items_sample, st.session_state.client
)
st.download_button(
    label="📥 Imprimer / Télécharger le Bon (PDF)",
    data=pdf_file,
    file_name=f"Bon_Entree_{selected_bon_id}.pdf",
    mime="application/pdf",
    key=f"dl_pdf_{selected_bon_id}",
)
