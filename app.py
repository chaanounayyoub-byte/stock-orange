import os
import sqlite3
from io import BytesIO

import docx
from docx import Document
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image as RLImage,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import streamlit as st

# =========================================================
# 1. CONFIGURATION & COULEURS PAR CLIENT
# =========================================================
CLIENTS = {
    "orange": {"color": "#FF6600", "logo": "logo_orange.png"},
    "inwi": {"color": "#A1006B", "logo": "logo_inwi.png"},
    "zte": {"color": "#005BAC", "logo": "logo_zte.png"},
}

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS", layout="wide"
)

# Initialisation des variables de session (Ajustez selon votre logique)
if "client" not in st.session_state:
    st.session_state["client"] = "orange"
if "user" not in st.session_state:
    st.session_state["user"] = {"fullname": "Ayyoub chaanoun"}
if "role" not in st.session_state:
    st.session_state["role"] = "admin"

CLIENT = st.session_state["client"]
ROLE = st.session_state["role"]
CURRENT_USER = st.session_state["user"]


# =========================================================
# 2. HELPER BASE DE DONNÉES (EXEMPLE SQLITE)
# =========================================================
def query(sql, params=(), one=False):
    """Remplacez / adaptez avec votre propre fonction de connexion DB."""
    try:
        conn = sqlite3.connect("stock.db")
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(sql, params)
        rv = cur.fetchall()
        conn.close()
        return (rv[0] if rv else None) if one else rv
    except Exception:
        # Fallback pour test de génération sans DB réelle
        if one:
            return {
                "id": params[0] if params else 1,
                "number": "BE-2024-001",
                "type": "BE",
                "date_bon": "2024-01-15",
                "fournisseur": "Fournisseur A",
                "lieu_livraison": "Magasin Principal",
                "receptionne_par": "Ayyoub",
                "equipe": "Équipe N1",
                "destination": "Site RABAT-01",
                "created_by": "Ayyoub",
                "client": CLIENT,
                "datetime_saisie": "2024-01-15 10:30",
            }
        return [
            {"reference": "REF-001", "article": "Câble FO 24B", "quantity": 10},
            {
                "reference": "REF-002",
                "article": "SFP+ 10G LR",
                "quantity": 5,
            },
        ]


# =========================================================
# 3. GÉNÉRATION PDF (REPORTLAB - CONFORME MODÈLES BE/BS)
# =========================================================
def generate_pdf(bon_id):
    bon = query("SELECT * FROM bons WHERE id=?", (bon_id,), one=True)
    items = query(
        "SELECT bi.*, a.name AS article FROM bon_items bi JOIN articles"
        " a ON a.id=bi.article_id WHERE bi.bon_id=?",
        (bon_id,),
    )
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )
    styles = getSampleStyleSheet()
    story = []

    # En-tête : Logos
    logo_nomatis_path = "logo_nomatis.png"
    logo_client_info = CLIENTS.get(str(bon["client"]).lower(), {})
    logo_client_path = logo_client_info.get("logo")

    col_nomatis = Paragraph("<b>NOMATIS</b>", styles["Normal"])
    if os.path.exists(logo_nomatis_path):
        try:
            col_nomatis = RLImage(
                logo_nomatis_path, width=45 * mm, height=18 * mm
            )
        except Exception:
            pass

    col_client = Paragraph(
        f"<b>{str(bon['client']).upper()}</b>", styles["Normal"]
    )
    if logo_client_path and os.path.exists(logo_client_path):
        try:
            col_client = RLImage(
                logo_client_path, width=45 * mm, height=18 * mm
            )
        except Exception:
            pass

    header_table = Table(
        [[col_nomatis, "", col_client]], colWidths=[60 * mm, 70 * mm, 60 * mm]
    )
    header_table.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (2, 0), (2, 0), "RIGHT"),
        ])
    )
    story.append(header_table)
    story.append(Spacer(1, 5))

    # Adresse de la société
    adresse_text = (
        "<b>NOMATIS</b><br/>32 Rue Al Hatim<br/>les Orangers<br/>10000"
    )
    story.append(Paragraph(adresse_text, styles["Normal"]))
    story.append(Spacer(1, 15))

    # Titre du bon
    title_text = "Bon d'entrée" if bon["type"] == "BE" else "Bon de sortie"
    story.append(
        Paragraph(
            f"<font size=18><b><center>{title_text}</center></b></font>",
            styles["Normal"],
        )
    )
    story.append(Spacer(1, 10))

    # Tableau d'informations (6 colonnes adaptées)
    if bon["type"] == "BE":
        info_headers = [
            "N° Bon",
            "Date",
            "Fournisseur",
            "Lieu de livraison",
            "receptioné par",
            "Stock",
        ]
        info_values = [
            bon["number"],
            bon["date_bon"],
            bon["fournisseur"] or "",
            bon["lieu_livraison"] or "",
            bon["receptionne_par"] or "",
            bon["client"],
        ]
    else:
        info_headers = [
            "N° Bon",
            "Date",
            "Équipe/ST",
            "Destination / Site",
            "validé  par",
            "Stock",
        ]
        info_values = [
            bon["number"],
            bon["date_bon"],
            bon["equipe"] or "",
            bon["destination"] or "",
            bon["created_by"],
            bon["client"],
        ]

    info_data = [info_headers, info_values]
    info_table = Table(
        info_data,
        colWidths=[31 * mm, 25 * mm, 35 * mm, 35 * mm, 34 * mm, 30 * mm],
    )
    info_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("PADDING", (0, 0), (-1, -1), 4),
        ])
    )
    story.append(info_table)
    story.append(Spacer(1, 15))

    # Tableau des articles
    items_data = [["Référence", "Désignation", "Qté"]]
    for item in items:
        items_data.append([
            item["reference"] or "-",
            item["article"],
            str(item["quantity"]),
        ])

    while len(items_data) < 5:
        items_data.append(["", "", ""])

    items_table = Table(items_data, colWidths=[45 * mm, 115 * mm, 30 * mm])
    items_table.setStyle(
        TableStyle([
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 0), (1, -1), "LEFT"),
            ("ALIGN", (2, 0), (2, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("PADDING", (0, 0), (-1, -1), 5),
        ])
    )
    story.append(items_table)
    story.append(Spacer(1, 20))

    # Signature et date de saisie
    story.append(
        Paragraph("<b>Signature / Cachet Magasinier</b>", styles["Normal"])
    )
    story.append(Spacer(1, 25))
    story.append(
        Paragraph(f"Saisie le : {bon['datetime_saisie']}", styles["Normal"])
    )

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# 4. GÉNÉRATION DOCX (WORD - CONFORME MODÈLES BE/BS)
# =========================================================
def generate_docx(bon_id):
    bon = query("SELECT * FROM bons WHERE id=?", (bon_id,), one=True)
    items = query(
        "SELECT bi.*, a.name AS article FROM bon_items bi JOIN articles"
        " a ON a.id=bi.article_id WHERE bi.bon_id=?",
        (bon_id,),
    )
    doc = Document()

    # Adresse
    p_addr = doc.add_paragraph()
    p_addr.add_run("NOMATIS\n32 Rue Al Hatim\nles Orangers\n10000").bold = True

    # Titre
    p_title = doc.add_paragraph()
    p_title.alignment = 1
    title_run = p_title.add_run(
        "Bon d'entrée" if bon["type"] == "BE" else "Bon de sortie"
    )
    title_run.bold = True
    title_run.font.size = docx.shared.Pt(16)

    # Tableau d'informations
    t_info = doc.add_table(rows=2, cols=6)
    t_info.style = "Table Grid"

    if bon["type"] == "BE":
        headers = [
            "N° Bon",
            "Date",
            "Fournisseur",
            "Lieu de livraison",
            "receptioné par",
            "Stock",
        ]
        values = [
            bon["number"],
            bon["date_bon"],
            bon["fournisseur"],
            bon["lieu_livraison"],
            bon["receptionne_par"],
            bon["client"],
        ]
    else:
        headers = [
            "N° Bon",
            "Date",
            "Équipe/ST",
            "Destination / Site",
            "validé  par",
            "Stock",
        ]
        values = [
            bon["number"],
            bon["date_bon"],
            bon["equipe"],
            bon["destination"],
            bon["created_by"],
            bon["client"],
        ]

    for i in range(6):
        t_info.rows[0].cells[i].text = headers[i]
        t_info.rows[1].cells[i].text = str(values[i] or "")

    doc.add_paragraph()

    # Tableau des articles
    t_items = doc.add_table(rows=1, cols=3)
    t_items.style = "Table Grid"
    hdr = t_items.rows[0].cells
    hdr[0].text = "Référence"
    hdr[1].text = "Désignation"
    hdr[2].text = "Qté"

    for item in items:
        row = t_items.add_row().cells
        row[0].text = item["reference"] or ""
        row[1].text = item["article"]
        row[2].text = str(item["quantity"])

    doc.add_paragraph()
    doc.add_paragraph("Signature / Cachet Magasinier")
    doc.add_paragraph(f"Saisie le : {bon['datetime_saisie']}")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# 5. EN-TÊTE PRINCIPAL DE L'APPLICATION (COULEUR DYNAMIQUE)
# =========================================================
client_color = CLIENTS.get(CLIENT.lower(), {}).get("color", "#2563EB")

h1, h2 = st.columns([3, 1])
with h1:
    st.markdown(
        f"""
        <div class="main-header">
            <div class="main-title" style="font-size: 24px; font-weight: bold; color: #1E293B;">
                Gestion Stock MW NOMATIS — <span style="color: {client_color};">ESPACE {CLIENT.upper()}</span>
            </div>
            <div class="subtitle" style="color: #10B981; font-weight: 600; margin-top: 4px;">
                Utilisateur : {CURRENT_USER['fullname']} ({ROLE.upper()})
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# Exemple de sélecteur client pour la démonstration dynamique en direct
selected_client = st.selectbox(
    "Changer de client (Test dynamique)", ["orange", "inwi", "zte"]
)
if selected_client != st.session_state["client"]:
    st.session_state["client"] = selected_client
    st.rerun()
