import io
from datetime import date, datetime
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# =========================================================
# 1. FONCTION DE GÉNÉRATION DU PDF (MODÈLE EXACT DE VOS CAPTURES)
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

    # Dynamic Client Color
    color_map = {
        "ORANGE": "#FF6600",
        "INWI": "#8A2BE2",
        "ZTE": "#0052CC",
    }
    client_color_hex = color_map.get(client_name.upper(), "#000000")

    # Style Header Client
    client_header_style = ParagraphStyle(
        "ClientHeader",
        parent=styles["Heading2"],
        alignment=2,  # Align Droite
        textColor=colors.HexColor(client_color_hex),
        fontSize=14,
    )

    # Header Data: Logo Nomatis / Client Logo
    company_info = "<b>NOMATIS</b><br/>32 Rue Al Hatim<br/>les Orangers<br/>10000"
    header_data = [
        [
            Paragraph("<b>// NOMATIS</b>", styles["Heading1"]),
            "",
            Paragraph(f"<b>logo {client_name.lower()}</b>", client_header_style),
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

    # Titre du Bon
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        alignment=1,
        textColor=colors.black,
        fontSize=16,
    )
    elements.append(Paragraph("<b>Bon d'entrée</b>", title_style))
    elements.append(Table([[""]], colWidths=[550], rowHeights=[10]))

    # Méta-données du Bon (Tableau de 2 lignes)
    meta_data = [
        ["N° Bon", "Date", "Fournisseur", "Lieu de livraison", "receptioné par", "Stock"],
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

    # Tableau des Articles
    items_table_data = [["Référence", "Désignation", "Qté"]]
    for item in items_data:
        items_table_data.append(
            [str(item.get("reference", "")), str(item.get("designation", "")), str(item.get("qte", ""))]
        )

    # Compléter avec des lignes vides si nécessaire pour garder la forme du modèle
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

    # Pied de Page (Signature et Saisie le)
    footer_data = [
        ["Signature / Cachet Magasinier", ""],
        ["", ""],
        [f"Saisie le : {datetime.now().strftime('%Y-%m-%d %H:%M')}", ""],
    ]
    footer_table = Table(footer_data, colWidths=[300, 250], rowHeights=[15, 40, 15])
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
# 2. EXTRAIT CORRIGÉ : DANS L'ONGLET HISTORIQUE
# =========================================================

with st.expander("✏️ Modifier les informations du Bon"):
    with st.form("form_edit_bon"):
        # Parsing sécurisé de la date
        try:
            parsed_date = datetime.strptime(
                bon_detail["date_bon"], "%Y-%m-%d"
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

# Bouton de téléchargement PDF qui utilise la fonction mise à jour
pdf_file = generate_pdf_bon_entree(
    bon_detail, items_list, st.session_state.get("client", "ORANGE")
)
st.download_button(
    label="📥 Imprimer / Télécharger le Bon (PDF)",
    data=pdf_file,
    file_name=f"Bon_Entree_{selected_bon_id}.pdf",
    mime="application/pdf",
    key=f"dl_pdf_{selected_bon_id}",
)
