import streamlit as st
import pandas as pd
import datetime
import json
import os
from io import BytesIO

from fpdf import FPDF
from PIL import Image

try:
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


# ============================================================
# CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_FILE = "database.json"

LOGOS = {
    "NOMATIS": "Logo Nomatis.jpg",
    "ORANGE": "Orange_logo.svg.webp",
    "INWI": "Logo INWI.jpg",
    "ZTE": "Logo ZTE.jpg",
}

CLIENTS = ["ORANGE", "INWI", "ZTE"]
ROLES = ["admin", "magasinier", "coordinateur", "coordinatrice"]

CREATE_EDIT_ROLES = {"admin", "magasinier"}
PRINT_ROLES = {"admin", "magasinier", "coordinateur", "coordinatrice"}


# ============================================================
# DESIGN
# ============================================================
def apply_theme():
    st.markdown(
        """
        <style>
        .stApp { background-color: #F4F7FA; }
        h1, h2, h3 { color: #0B4F6C !important; font-weight: 700 !important; }

        [data-testid="stSidebar"] {
            background-color: #FFFFFF;
            border-right: 1px solid #DCE3EA;
        }

        div[data-testid="stForm"] {
            border: 1px solid #DCE1E8;
            border-radius: 12px;
            padding: 24px;
            background: #FFFFFF;
            box-shadow: 0 3px 12px rgba(0,0,0,.04);
        }

        .login-card {
            background: #FFFFFF;
            border: 1px solid #DCE1E8;
            border-radius: 14px;
            padding: 30px;
            box-shadow: 0 6px 20px rgba(0,0,0,.06);
        }

        .client-card {
            background: #FFFFFF;
            border: 1px solid #DCE1E8;
            border-radius: 14px;
            padding: 22px;
            min-height: 300px;
            box-shadow: 0 4px 14px rgba(0,0,0,.05);
            text-align: center;
        }

        .btn-login-red button {
            background: #D9534F !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            border: none !important;
            width: 100% !important;
            border-radius: 7px !important;
        }

        .btn-login-green button {
            background: #28A745 !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            border: none !important;
            width: 100% !important;
            border-radius: 7px !important;
        }

        .btn-orange button {
            background: #FF7900 !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            border: none !important;
            width: 100% !important;
        }

        .btn-inwi button {
            background: #A1006B !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            border: none !important;
            width: 100% !important;
        }

        .btn-zte button {
            background: #005BAC !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            border: none !important;
            width: 100% !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


apply_theme()


# ============================================================
# BASE DE DONNEES
# ============================================================
def default_db():
    return {
        "users": {
            "admin": {
                "password": "admin",
                "role": "admin",
                "last_login": "",
            }
        },
        "articles": [],
        "fournisseurs": ["NEC", "ZTE", "Intégral", "FO connect"],
        "equipes": ["Nabil Team", "Yassine Team", "Issa Team"],
        "transactions": [],
    }


def init_db():
    if not os.path.exists(DB_FILE):
        db = default_db()
        save_db(db)
        return db

    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            db = json.load(f)
    except Exception:
        db = default_db()

    defaults = default_db()
    for key, value in defaults.items():
        if key not in db:
            db[key] = value

    if "admin" not in db["users"]:
        db["users"]["admin"] = defaults["users"]["admin"]

    return db


def save_db(db):
    temp_file = DB_FILE + ".tmp"
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=4, ensure_ascii=False)
    os.replace(temp_file, DB_FILE)


db = init_db()


# ============================================================
# OUTILS
# ============================================================
def load_image(path):
    if not os.path.exists(path):
        return None
    try:
        return Image.open(path)
    except Exception:
        return None


def now():
    return datetime.datetime.now()


def today():
    return datetime.date.today()


def safe_int(value, default=0):
    try:
        return int(value)
    except Exception:
        return default


def format_qte(value):
    try:
        n = float(value)
        return str(int(n)) if n.is_integer() else f"{n:g}"
    except Exception:
        return str(value)


def article_designations():
    return [
        a.get("designation", "")
        for a in db["articles"]
        if a.get("designation", "").strip()
    ]


def article_ref(designation):
    for article in db["articles"]:
        if article.get("designation") == designation:
            return article.get("ref", "")
    return ""


def next_article_ref():
    """Génère automatiquement MW-001, MW-002... sans réutiliser un index supprimé."""
    max_num = 0
    for article in db.get("articles", []):
        ref = str(article.get("ref", "")).strip().upper()
        if ref.startswith("MW-"):
            try:
                max_num = max(max_num, int(ref.split("-", 1)[1]))
            except Exception:
                pass
    for transaction in db.get("transactions", []):
        for item in transaction.get("articles", []):
            ref = str(item.get("ref", "")).strip().upper()
            if ref.startswith("MW-"):
                try:
                    max_num = max(max_num, int(ref.split("-", 1)[1]))
                except Exception:
                    pass
    return f"MW-{max_num + 1:03d}"


ARTICLE_CATEGORIES = {
    "Sans catégorie": {},
    "Support": {
        "dimensions": ["0.3 m", "0.6 m", "1.2 m", "1.8 m"]
    },
    "Jarretière": {
        "types": ["LC/LC", "LC/SC-APC", "LC/SC-UPC", "LC/FC", "FC/FC"],
        "longueurs": ["3 m", "5 m", "10 m", "30 m", "60 m"],
    },
}


def build_article_designation(category, dimension="", fibre_type="", longueur="", base_name=""):
    if category == "Support":
        return f"Support {dimension}".strip()
    if category == "Jarretière":
        return f"Jarretière {fibre_type} {longueur}".strip()
    return base_name.strip()


def generate_bon_id(type_bon):
    """
    Format obligatoire :
      BE-MW-YYYYMMDD-01
      BS-MW-YYYYMMDD-01

    Le compteur est séparé pour les BE et les BS
    et recommence à 01 chaque nouveau jour.
    """
    date_str = now().strftime("%Y%m%d")
    prefix = "BE-MW-" if type_bon == "BE" else "BS-MW-"

    existing = [
        t for t in db["transactions"]
        if t.get("type") == type_bon
        and str(t.get("id", "")).startswith(prefix + date_str + "-")
    ]

    max_number = 0
    for transaction in existing:
        try:
            number = int(str(transaction["id"]).rsplit("-", 1)[1])
            max_number = max(max_number, number)
        except Exception:
            pass

    return f"{prefix}{date_str}-{max_number + 1:02d}"


def get_stock(client, exclude_transaction_id=None):
    """
    Calcule le stock à partir de tous les mouvements du client.
    Les modifications/suppressions de bons sont donc automatiquement
    répercutées dans la situation stock.
    """
    stock = {
        a.get("designation"): {
            "ref": a.get("ref", ""),
            "qte": 0,
        }
        for a in db["articles"]
        if a.get("designation")
    }

    for transaction in db["transactions"]:
        if transaction.get("client") != client:
            continue
        if exclude_transaction_id and transaction.get("id") == exclude_transaction_id:
            continue

        for item in transaction.get("articles", []):
            designation = item.get("designation", "")
            qte = safe_int(item.get("qte", 0))

            if not designation:
                continue

            if designation not in stock:
                stock[designation] = {
                    "ref": item.get("ref", ""),
                    "qte": 0,
                }

            if transaction.get("type") in ("BE", "ADJ_PLUS"):
                stock[designation]["qte"] += qte
            elif transaction.get("type") in ("BS", "ADJ_MOINS"):
                stock[designation]["qte"] -= qte

    return stock


def stock_dataframe(client):
    stock = get_stock(client)
    rows = []

    for designation, data in stock.items():
        rows.append({
            "Référence": data.get("ref", ""),
            "Désignation": designation,
            "Quantité Disponible": data.get("qte", 0),
        })

    return pd.DataFrame(
        rows,
        columns=["Référence", "Désignation", "Quantité Disponible"],
    )


def validate_articles(articles):
    allowed = set(article_designations())

    if not articles:
        return False, "Le bon doit contenir au moins un article."

    for item in articles:
        designation = item.get("designation", "")
        qte = safe_int(item.get("qte", 0))

        if designation not in allowed:
            return False, (
                f"L'article « {designation} » n'existe pas "
                "dans le référentiel administrateur."
            )

        if qte <= 0:
            return False, (
                f"La quantité de « {designation} » "
                "doit être supérieure à 0."
            )

    return True, ""


def validate_bs_stock(client, articles, old_transaction_id=None):
    stock = get_stock(
        client,
        exclude_transaction_id=old_transaction_id,
    )

    requested = {}

    for item in articles:
        designation = item.get("designation", "")
        requested[designation] = (
            requested.get(designation, 0)
            + safe_int(item.get("qte", 0))
        )

    for designation, requested_qte in requested.items():
        available = safe_int(
            stock.get(designation, {}).get("qte", 0)
        )

        if requested_qte > available:
            return False, (
                f"Stock insuffisant pour « {designation} ». "
                f"Disponible : {available} | Demandé : {requested_qte}."
            )

    return True, ""


def get_transaction(transaction_id):
    for transaction in db["transactions"]:
        if transaction.get("id") == transaction_id:
            return transaction
    return None


def get_transactions(client, type_bon=None):
    result = [
        t for t in db["transactions"]
        if t.get("client") == client
    ]

    if type_bon:
        result = [
            t for t in result
            if t.get("type") == type_bon
        ]

    return result


def can_edit():
    return st.session_state.role in {"admin", "magasinier"}


def can_print():
    return st.session_state.role in {
        "admin",
        "magasinier",
        "coordinateur",
        "coordinatrice",
    }


def confirmation(label, key):
    return st.checkbox(label, key=key)


def all_article_remarks_present(articles):
    return all(str(item.get("remarque", "")).strip() for item in articles)


# ============================================================
# PDF
# ============================================================
def pdf_safe(text):
    text = "" if text is None else str(text)

    replacements = {
        "’": "'",
        "–": "-",
        "—": "-",
        "œ": "oe",
        "Œ": "OE",
        "é": "e",
        "è": "e",
        "ê": "e",
        "ë": "e",
        "à": "a",
        "â": "a",
        "ä": "a",
        "î": "i",
        "ï": "i",
        "ô": "o",
        "ö": "o",
        "ù": "u",
        "û": "u",
        "ü": "u",
        "ç": "c",
    }

    for source, destination in replacements.items():
        text = text.replace(source, destination)

    return text


def generate_pdf(bon_data, client):
    pdf = FPDF(
        orientation="P",
        unit="mm",
        format="A4",
    )
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()

    # Logo NOMATIS
    if os.path.exists(LOGOS["NOMATIS"]):
        try:
            pdf.image(
                LOGOS["NOMATIS"],
                x=12,
                y=10,
                w=42,
            )
        except Exception:
            pdf.set_font("Arial", "B", 15)
            pdf.text(12, 20, "NOMATIS")
    else:
        pdf.set_font("Arial", "B", 15)
        pdf.text(12, 20, "NOMATIS")

    # Logo client
    logo_client = LOGOS.get(client, "")
    if os.path.exists(logo_client):
        try:
            pdf.image(
                logo_client,
                x=162,
                y=10,
                w=35,
            )
        except Exception:
            pdf.set_font("Arial", "B", 12)
            pdf.text(162, 20, pdf_safe(client))
    else:
        pdf.set_font("Arial", "B", 12)
        pdf.text(162, 20, pdf_safe(client))

    # Coordonnées
    pdf.set_xy(12, 40)
    pdf.set_font("Arial", "", 8)
    pdf.cell(100, 4, "NOMATIS", ln=1)
    pdf.cell(100, 4, "32 Rue Al Hatim", ln=1)
    pdf.cell(100, 4, "Les Orangers", ln=1)
    pdf.cell(100, 4, "10000", ln=1)

    pdf.ln(15)

    is_be = bon_data.get("type") == "BE"
    title = "Bon d'entree" if is_be else "Bon de sortie"

    pdf.set_font("Arial", "B", 16)
    pdf.cell(
        0,
        9,
        title,
        ln=1,
        align="C",
    )
    pdf.ln(4)

    # Tableau d'informations
    headers = [
        "Bon de Livraison",
        "Date",
        "Fournisseur" if is_be else "Equipe",
        "Lieu de livraison" if is_be else "Destination",
        "receptionne par",
        "Stock",
    ]

    values = [
        bon_data.get("id", ""),
        bon_data.get("date", ""),
        bon_data.get("fournisseur_equipe", ""),
        bon_data.get("destination", ""),
        bon_data.get("user", ""),
        client,
    ]

    widths = [34, 25, 36, 35, 31, 29]

    pdf.set_font("Arial", "B", 8)
    for i, header in enumerate(headers):
        pdf.cell(
            widths[i],
            7,
            pdf_safe(header),
            border=1,
            align="C",
        )
    pdf.ln()

    pdf.set_font("Arial", "", 8)
    for i, value in enumerate(values):
        pdf.cell(
            widths[i],
            8,
            pdf_safe(value),
            border=1,
            align="C",
        )
    pdf.ln(10)

    # Tableau articles
    pdf.set_font("Arial", "B", 8)
    pdf.cell(40, 7, "Reference", border=1, align="C")
    pdf.cell(120, 7, "Designation", border=1, align="C")
    pdf.cell(30, 7, "Qte", border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial", "", 8)

    article_count = len(bon_data.get("articles", []))

    for item in bon_data.get("articles", []):
        pdf.cell(
            40,
            7,
            pdf_safe(item.get("ref", "")),
            border=1,
        )
        pdf.cell(
            120,
            7,
            pdf_safe(item.get("designation", "")),
            border=1,
        )
        pdf.cell(
            30,
            7,
            format_qte(item.get("qte", 0)),
            border=1,
            align="C",
        )
        pdf.ln()

        if item.get("remarque"):
            pdf.set_font("Arial", "I", 7)
            pdf.cell(40, 5, "", border="LR")
            pdf.cell(
                120,
                5,
                "Remarque : " + pdf_safe(item.get("remarque", "")),
                border="LR",
            )
            pdf.cell(30, 5, "", border="LR")
            pdf.ln()
            pdf.set_font("Arial", "", 8)

    for _ in range(max(0, 8 - article_count)):
        pdf.cell(40, 7, "", border=1)
        pdf.cell(120, 7, "", border=1)
        pdf.cell(30, 7, "", border=1)
        pdf.ln()

    if bon_data.get("remarque"):
        pdf.ln(4)
        pdf.set_font("Arial", "B", 8)
        pdf.cell(25, 6, "Remarque :", ln=0)
        pdf.set_font("Arial", "", 8)
        pdf.multi_cell(
            0,
            6,
            pdf_safe(bon_data.get("remarque", "")),
        )

    pdf.ln(8)
    pdf.set_font("Arial", "", 7)
    pdf.cell(
        0,
        5,
        "Document genere par Gestion Stock MW NOMATIS",
        align="C",
    )

    return bytes(pdf.output(dest="S"))


# ============================================================
# WORD
# ============================================================
try:
    from docx import Document
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


def set_docx_cell(cell, text, bold=False, size=8, align=None):
    cell.text = ""
    paragraph = cell.paragraphs[0]

    if align is not None:
        paragraph.alignment = align

    run = paragraph.add_run(pdf_safe(text))
    run.bold = bold
    run.font.size = Pt(size)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def generate_docx(bon_data, client):
    if not DOCX_AVAILABLE:
        return None

    document = Document()
    section = document.sections[0]

    section.top_margin = Cm(1.2)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    header_table = document.add_table(
        rows=1,
        cols=3,
    )
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cell1, cell2, cell3 = header_table.rows[0].cells

    if os.path.exists(LOGOS["NOMATIS"]):
        try:
            paragraph = cell1.paragraphs[0]
            paragraph.add_run().add_picture(
                LOGOS["NOMATIS"],
                width=Cm(4.0),
            )
        except Exception:
            set_docx_cell(
                cell1,
                "NOMATIS",
                bold=True,
                size=14,
            )
    else:
        set_docx_cell(
            cell1,
            "NOMATIS",
            bold=True,
            size=14,
        )

    set_docx_cell(cell2, "")

    if os.path.exists(LOGOS.get(client, "")):
        try:
            paragraph = cell3.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            paragraph.add_run().add_picture(
                LOGOS[client],
                width=Cm(3.2),
            )
        except Exception:
            set_docx_cell(
                cell3,
                client,
                bold=True,
                size=12,
            )
    else:
        set_docx_cell(
            cell3,
            client,
            bold=True,
            size=12,
        )

    paragraph = document.add_paragraph()
    run = paragraph.add_run(
        "NOMATIS\n"
        "32 Rue Al Hatim\n"
        "Les Orangers\n"
        "10000"
    )
    run.font.size = Pt(8)

    title = (
        "Bon d'entree"
        if bon_data.get("type") == "BE"
        else "Bon de sortie"
    )

    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run(title)
    run.bold = True
    run.font.size = Pt(16)

    is_be = bon_data.get("type") == "BE"

    headers = [
        "Bon de Livraison",
        "Date",
        "Fournisseur" if is_be else "Equipe",
        "Lieu de livraison" if is_be else "Destination",
        "receptionne par",
        "Stock",
    ]

    values = [
        bon_data.get("id", ""),
        bon_data.get("date", ""),
        bon_data.get("fournisseur_equipe", ""),
        bon_data.get("destination", ""),
        bon_data.get("user", ""),
        client,
    ]

    info_table = document.add_table(
        rows=2,
        cols=6,
    )
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_table.style = "Table Grid"

    for i in range(6):
        set_docx_cell(
            info_table.cell(0, i),
            headers[i],
            bold=True,
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )
        set_docx_cell(
            info_table.cell(1, i),
            values[i],
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    document.add_paragraph()

    article_table = document.add_table(
        rows=1,
        cols=3,
    )
    article_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    article_table.style = "Table Grid"

    for i, header in enumerate(
        ["Reference", "Designation", "Qte"]
    ):
        set_docx_cell(
            article_table.cell(0, i),
            header,
            bold=True,
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

    for item in bon_data.get("articles", []):
        cells = article_table.add_row().cells

        set_docx_cell(
            cells[0],
            item.get("ref", ""),
            size=8,
        )
        set_docx_cell(
            cells[1],
            item.get("designation", ""),
            size=8,
        )
        set_docx_cell(
            cells[2],
            format_qte(item.get("qte", 0)),
            size=8,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        )

        if item.get("remarque"):
            cells = article_table.add_row().cells
            set_docx_cell(cells[0], "", size=7)
            set_docx_cell(
                cells[1],
                "Remarque : " + item.get("remarque", ""),
                size=7,
            )
            set_docx_cell(cells[2], "", size=7)

    for _ in range(
        max(0, 8 - len(bon_data.get("articles", [])))
    ):
        cells = article_table.add_row().cells
        for cell in cells:
            set_docx_cell(cell, "", size=8)

    if bon_data.get("remarque"):
        paragraph = document.add_paragraph()
        run = paragraph.add_run(
            "Remarque : " + str(bon_data.get("remarque", ""))
        )
        run.font.size = Pt(8)

    buffer = BytesIO()
    document.save(buffer)

    return buffer.getvalue()


# ============================================================
# SESSION
# ============================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user = None
    st.session_state.role = None
    st.session_state.client = None

if "current_be_articles" not in st.session_state:
    st.session_state.current_be_articles = []

if "current_bs_articles" not in st.session_state:
    st.session_state.current_bs_articles = []

if "editing_transaction_id" not in st.session_state:
    st.session_state.editing_transaction_id = None


# ============================================================
# CONNEXION
# ============================================================
if not st.session_state.logged_in:
    _, center, _ = st.columns([1, 1.4, 1])

    with center:
        st.markdown(
            '<div class="login-card">',
            unsafe_allow_html=True,
        )

        logo = load_image(LOGOS["NOMATIS"])
        if logo:
            st.image(logo, width=190)

        st.markdown(
            "<h1 style='text-align:center;'>"
            "Gestion Stock MW NOMATIS"
            "</h1>",
            unsafe_allow_html=True,
        )

        st.markdown(
            "<p style='text-align:center;color:#6C757D;'>"
            "Gestion des stocks multi-clients"
            "</p>",
            unsafe_allow_html=True,
        )

        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input(
                "Mot de passe",
                type="password",
            )

            button_class = (
                "btn-login-green"
                if username.strip() and password
                else "btn-login-red"
            )

            st.markdown(
                f'<div class="{button_class}">',
                unsafe_allow_html=True,
            )

            submitted = st.form_submit_button(
                "SE CONNECTER",
                use_container_width=True,
            )

            st.markdown("</div>", unsafe_allow_html=True)

            if submitted:
                if (
                    username in db["users"]
                    and db["users"][username].get("password") == password
                ):
                    st.session_state.logged_in = True
                    st.session_state.user = username
                    st.session_state.role = db["users"][username].get(
                        "role",
                        "magasinier",
                    )

                    db["users"][username]["last_login"] = (
                        now().strftime("%Y-%m-%d %H:%M:%S")
                    )

                    save_db(db)
                    st.rerun()
                else:
                    st.error("Identifiants incorrects.")

        st.markdown("</div>", unsafe_allow_html=True)

    st.stop()


# ============================================================
# SELECTION CLIENT
# ============================================================
if st.session_state.client is None:
    st.title(
        f"Bienvenue, {st.session_state.user} !"
    )
    st.subheader("Sélectionnez l'espace client")

    columns = st.columns(3)

    client_classes = {
        "ORANGE": "btn-orange",
        "INWI": "btn-inwi",
        "ZTE": "btn-zte",
    }

    for column, client_name in zip(columns, CLIENTS):
        with column:
            st.markdown(
                '<div class="client-card">',
                unsafe_allow_html=True,
            )

            image = load_image(LOGOS[client_name])
            if image:
                st.image(image, width=180)

            st.markdown(
                f"<h3 style='text-align:center;'>"
                f"{client_name}"
                f"</h3>",
                unsafe_allow_html=True,
            )

            st.markdown(
                f'<div class="{client_classes[client_name]}">',
                unsafe_allow_html=True,
            )

            if st.button(
                f"ACCÈS AU STOCK {client_name}",
                key=f"client_{client_name}",
                use_container_width=True,
            ):
                st.session_state.client = client_name
                st.session_state.current_be_articles = []
                st.session_state.current_bs_articles = []
                st.session_state.editing_transaction_id = None
                st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    st.subheader("Mon compte")

    with st.expander(
        "Modifier mon nom / mot de passe"
    ):
        current_user = st.session_state.user

        new_user = st.text_input(
            "Nom d'utilisateur",
            value=current_user,
        )

        new_password = st.text_input(
            "Nouveau mot de passe",
            type="password",
        )

        confirm_account = confirmation("Je confirme la modification de mon compte.", "confirm_my_account")
        if st.button("Mettre à jour mon compte"):
            if not confirm_account:
                st.warning("Veuillez confirmer la modification avant de continuer.")
                st.stop()
            new_user = new_user.strip()

            if not new_user:
                st.error(
                    "Le nom d'utilisateur ne peut pas être vide."
                )
            elif (
                new_user != current_user
                and new_user in db["users"]
            ):
                st.error(
                    "Ce nom d'utilisateur existe déjà."
                )
            else:
                user_data = db["users"].pop(current_user)

                if new_password:
                    user_data["password"] = new_password

                db["users"][new_user] = user_data
                st.session_state.user = new_user

                save_db(db)

                st.success(
                    "Compte mis à jour."
                )
                st.rerun()

    if st.button("Se déconnecter"):
        st.session_state.clear()
        st.rerun()

    st.stop()


# ============================================================
# APPLICATION PRINCIPALE
# ============================================================
client = st.session_state.client
role = st.session_state.role

sidebar_logo = load_image(
    LOGOS["NOMATIS"]
)

if sidebar_logo:
    st.sidebar.image(
        sidebar_logo,
        width=145,
    )

st.sidebar.title(
    f"Stock {client}"
)

st.sidebar.caption(
    f"Utilisateur : {st.session_state.user} | "
    f"Rôle : {role}"
)

if st.sidebar.button(
    "Changer de client",
    use_container_width=True,
):
    st.session_state.client = None
    st.session_state.current_be_articles = []
    st.session_state.current_bs_articles = []
    st.session_state.editing_transaction_id = None
    st.rerun()

if st.sidebar.button(
    "Déconnexion",
    use_container_width=True,
):
    st.session_state.clear()
    st.rerun()


# Navigation selon rôle
menus = [
    "Situation Stock",
    "Historique",
]

if role in {"admin", "magasinier"}:
    menus = [
        "Bon d'Entrée (BE)",
        "Bon de Sortie (BS)",
    ] + menus

if role == "admin":
    menus.append("Configuration")

# Si une modification a été demandée depuis l'historique,
# afficher directement le bon correspondant.
editing_id = st.session_state.editing_transaction_id
editing_transaction = (
    get_transaction(editing_id)
    if editing_id
    else None
)

if editing_transaction:
    if editing_transaction.get("type") == "BE":
        default_menu = "Bon d'Entrée (BE)"
    else:
        default_menu = "Bon de Sortie (BS)"
else:
    default_menu = menus[0]

if (
    "selected_menu" not in st.session_state
    or st.session_state.selected_menu not in menus
    or editing_transaction
):
    st.session_state.selected_menu = default_menu

choix_menu = st.sidebar.radio(
    "Navigation",
    menus,
    key="selected_menu",
)


# ============================================================
# BON D'ENTREE
# ============================================================
if choix_menu == "Bon d'Entrée (BE)":
    st.header("📥 Bon d'Entrée (BE)")

    editing_id = st.session_state.editing_transaction_id
    editing = (
        get_transaction(editing_id)
        if editing_id
        else None
    )

    if editing and editing.get("type") != "BE":
        editing = None
        st.session_state.editing_transaction_id = None

    if editing:
        st.info(
            f"Modification du bon : {editing['id']}"
        )

        default_date = datetime.date.fromisoformat(
            editing["date"]
        )
        default_supplier = editing.get(
            "fournisseur_equipe",
            "",
        )
        default_destination = editing.get(
            "destination",
            "",
        )
        default_remarque = editing.get(
            "remarque",
            "",
        )
    else:
        default_date = today()

        default_supplier = (
            db["fournisseurs"][0]
            if db["fournisseurs"]
            else ""
        )

        default_destination = "Dépôt Principal"
        default_remarque = ""

    # Liste de travail persistante
    if editing:
        work_key = f"be_working_{editing_id}"

        if work_key not in st.session_state:
            st.session_state[work_key] = [
                dict(item)
                for item in editing.get(
                    "articles",
                    [],
                )
            ]

        working_articles = (
            st.session_state[work_key]
        )
    else:
        working_articles = (
            st.session_state.current_be_articles
        )

    col1, col2 = st.columns(2)

    with col1:
        date_be = st.date_input(
            "Date du BE",
            value=default_date,
            max_value=today(),
            key=f"be_date_{editing_id or 'new'}",
        )

        supplier_options = (
            list(db["fournisseurs"])
            + ["Autre..."]
        )

        supplier_index = (
            supplier_options.index(
                default_supplier
            )
            if default_supplier in supplier_options
            else 0
        )

        selected_supplier = st.selectbox(
            "Fournisseur",
            supplier_options,
            index=supplier_index,
            key=f"be_supplier_{editing_id or 'new'}",
        )

        if selected_supplier == "Autre...":
            fournisseur = st.text_input(
                "Nom du nouveau fournisseur",
                value=(
                    ""
                    if default_supplier in supplier_options
                    else default_supplier
                ),
                key=f"be_new_supplier_{editing_id or 'new'}",
            ).strip()
        else:
            fournisseur = selected_supplier

    with col2:
        lieu = st.text_input(
            "Lieu de livraison",
            value=default_destination,
            key=f"be_lieu_{editing_id or 'new'}",
        )

        remarque_bon = st.text_area(
            "Remarque générale",
            value=default_remarque,
            key=f"be_remarque_{editing_id or 'new'}",
        )

    if editing:
        st.markdown(
            f"**Numéro BE :** `{editing['id']}`  \n"
            f"**Date et heure de saisie :** "
            f"`{editing.get('date','')} "
            f"{editing.get('heure_saisie','')}`  \n"
            f"**Réceptionné par :** "
            f"`{editing.get('user','')}`"
        )
    else:
        st.caption(
            "Numéro automatique : "
            "BE-MW-YYYYMMDD-01"
        )

    st.subheader("Articles")

    articles_list = article_designations()

    if not articles_list:
        st.warning(
            "Aucun article n'est configuré. "
            "L'administrateur doit d'abord ajouter les articles."
        )
    else:
        with st.form(
            f"be_add_article_{editing_id or 'new'}"
        ):
            c1, c2, c3 = st.columns(
                [2, 1, 2]
            )

            with c1:
                article_sel = st.selectbox(
                    "Article",
                    articles_list,
                    key=f"be_article_{editing_id or 'new'}",
                )

            with c2:
                qte = st.number_input(
                    "Quantité",
                    min_value=1,
                    value=1,
                    step=1,
                    key=f"be_qte_{editing_id or 'new'}",
                )

            with c3:
                remarque_article = st.text_input(
                    "Remarque",
                    key=f"be_art_rem_{editing_id or 'new'}",
                )

            add_article = st.form_submit_button(
                "Ajouter l'article"
            )

            if add_article:
                if not remarque_article.strip():
                    st.error("La remarque de l'article est obligatoire.")
                elif qte <= 0:
                    st.error(
                        "La quantité doit être supérieure à 0."
                    )
                else:
                    found = False
                    ref = article_ref(
                        article_sel
                    )

                    for item in working_articles:
                        if item.get("designation") == article_sel:
                            item["qte"] = (
                                safe_int(
                                    item.get(
                                        "qte",
                                        0,
                                    )
                                )
                                + int(qte)
                            )

                            if remarque_article.strip():
                                old_remark = item.get(
                                    "remarque",
                                    "",
                                )
                                item["remarque"] = (
                                    f"{old_remark} | "
                                    f"{remarque_article.strip()}"
                                    if old_remark
                                    else remarque_article.strip()
                                )

                            found = True
                            break

                    if not found:
                        working_articles.append({
                            "ref": ref,
                            "designation": article_sel,
                            "qte": int(qte),
                            "remarque": remarque_article.strip(),
                        })

                    st.rerun()

    if working_articles:
        st.markdown("### Articles du bon")

        rows = [
            {
                "Référence": item.get(
                    "ref",
                    "",
                ),
                "Désignation": item.get(
                    "designation",
                    "",
                ),
                "Quantité": item.get(
                    "qte",
                    0,
                ),
                "Remarque": item.get(
                    "remarque",
                    "",
                ),
            }
            for item in working_articles
        ]

        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown(
            "#### Modifier / supprimer une ligne"
        )

        selected_line = st.selectbox(
            "Article",
            [
                item.get("designation")
                for item in working_articles
            ],
            key=f"be_line_select_{editing_id or 'new'}",
        )

        line_index = next(
            (
                i
                for i, item
                in enumerate(working_articles)
                if item.get("designation")
                == selected_line
            ),
            None,
        )

        if line_index is not None:
            c1, c2, c3 = st.columns(
                [1, 1, 1]
            )

            with c1:
                line_qty = st.number_input(
                    "Nouvelle quantité",
                    min_value=1,
                    value=safe_int(
                        working_articles[
                            line_index
                        ].get(
                            "qte",
                            1,
                        )
                    ),
                    step=1,
                    key=f"be_line_qty_{editing_id or 'new'}",
                )

            with c2:
                line_remark = st.text_input(
                    "Nouvelle remarque",
                    value=working_articles[
                        line_index
                    ].get(
                        "remarque",
                        "",
                    ),
                    key=f"be_line_remark_{editing_id or 'new'}",
                )

            with c3:
                confirm_be_line = confirmation("Confirmer la modification de la ligne.", f"confirm_be_line_{editing_id or 'new'}")
                if st.button(
                    "Appliquer",
                    key=f"be_apply_{editing_id or 'new'}",
                    use_container_width=True,
                ):
                    if not confirm_be_line:
                        st.warning("Veuillez confirmer la modification de la ligne.")
                        st.stop()
                    if not line_remark.strip():
                        st.error("La remarque de l'article est obligatoire.")
                        st.stop()
                    working_articles[
                        line_index
                    ]["qte"] = int(line_qty)

                    working_articles[
                        line_index
                    ]["remarque"] = line_remark

                    st.rerun()

                if st.button(
                    "Supprimer",
                    key=f"be_remove_{editing_id or 'new'}",
                    use_container_width=True,
                ):
                    working_articles.pop(
                        line_index
                    )
                    st.rerun()

        if st.button(
            "Vider la liste",
            key=f"be_clear_{editing_id or 'new'}",
        ):
            working_articles.clear()
            st.rerun()

        st.divider()

        if editing:
            confirm_be_edit = confirmation("Je confirme l'enregistrement des modifications du BE.", f"confirm_be_edit_{editing_id}")
            c1, c2 = st.columns(2)

            with c1:
                if st.button(
                    "💾 Enregistrer les modifications",
                    type="primary",
                    use_container_width=True,
                ):
                    if not confirm_be_edit:
                        st.warning("Veuillez confirmer l'enregistrement des modifications.")
                        st.stop()
                    if date_be > today():
                        st.error(
                            "La date du BE ne peut pas "
                            "dépasser aujourd'hui."
                        )
                    elif not fournisseur:
                        st.error(
                            "Le fournisseur est obligatoire."
                        )
                    else:
                        valid, message = (
                            validate_articles(
                                working_articles
                            )
                        )

                        if not valid:
                            st.error(message)
                        elif not lieu.strip():
                            st.error("Le lieu de livraison est obligatoire.")
                        elif not remarque_bon.strip():
                            st.error("La remarque générale est obligatoire.")
                        elif not all_article_remarks_present(working_articles):
                            st.error("La remarque de chaque article est obligatoire.")
                        else:
                            editing["date"] = (
                                date_be.strftime(
                                    "%Y-%m-%d"
                                )
                            )
                            editing[
                                "fournisseur_equipe"
                            ] = fournisseur
                            editing[
                                "destination"
                            ] = lieu
                            editing[
                                "remarque"
                            ] = remarque_bon
                            editing[
                                "articles"
                            ] = [
                                dict(item)
                                for item
                                in working_articles
                            ]

                            if (
                                fournisseur
                                not in db["fournisseurs"]
                            ):
                                db["fournisseurs"].append(
                                    fournisseur
                                )

                            save_db(db)

                            st.session_state.pop(
                                work_key,
                                None,
                            )
                            st.session_state.editing_transaction_id = None

                            st.success(
                                f"BE {editing['id']} "
                                "modifié avec succès."
                            )
                            st.rerun()

            with c2:
                if st.button(
                    "Annuler la modification",
                    use_container_width=True,
                ):
                    st.session_state.pop(
                        work_key,
                        None,
                    )
                    st.session_state.editing_transaction_id = None
                    st.rerun()

        else:
            confirm_be_new = confirmation("Je confirme l'enregistrement définitif du Bon d'Entrée.", "confirm_be_new")
            if st.button(
                "💾 Enregistrer le Bon d'Entrée",
                type="primary",
                use_container_width=True,
            ):
                if date_be > today():
                    st.error(
                        "La date du BE ne peut pas "
                        "dépasser aujourd'hui."
                    )
                elif not fournisseur:
                    st.error(
                        "Le fournisseur est obligatoire."
                    )
                else:
                    valid, message = (
                        validate_articles(
                            working_articles
                        )
                    )

                    if not valid:
                        st.error(message)
                    elif not lieu.strip():
                        st.error("Le lieu de livraison est obligatoire.")
                    elif not remarque_bon.strip():
                        st.error("La remarque générale est obligatoire.")
                    elif not all_article_remarks_present(working_articles):
                        st.error("La remarque de chaque article est obligatoire.")
                    else:
                        new_be = {
                            "id": generate_bon_id("BE"),
                            "type": "BE",
                            "date": date_be.strftime(
                                "%Y-%m-%d"
                            ),
                            "heure_saisie": now().strftime(
                                "%H:%M:%S"
                            ),
                            "client": client,
                            "user": st.session_state.user,
                            "fournisseur_equipe": fournisseur,
                            "destination": lieu,
                            "remarque": remarque_bon,
                            "articles": [
                                dict(item)
                                for item
                                in working_articles
                            ],
                        }

                        db["transactions"].append(
                            new_be
                        )

                        if (
                            fournisseur
                            not in db["fournisseurs"]
                        ):
                            db["fournisseurs"].append(
                                fournisseur
                            )

                        save_db(db)
                        st.session_state.current_be_articles = []

                        st.success(
                            f"BE {new_be['id']} "
                            "enregistré avec succès."
                        )

                        pdf_data = generate_pdf(
                            new_be,
                            client,
                        )

                        st.download_button(
                            "📄 Télécharger / imprimer le BE en PDF",
                            data=pdf_data,
                            file_name=f"{new_be['id']}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                        )

                        if DOCX_AVAILABLE:
                            docx_data = generate_docx(
                                new_be,
                                client,
                            )

                            st.download_button(
                                "📝 Télécharger le BE en Word",
                                data=docx_data,
                                file_name=f"{new_be['id']}.docx",
                                mime=(
                                    "application/vnd.openxmlformats-"
                                    "officedocument.wordprocessingml.document"
                                ),
                                use_container_width=True,
                            )


# ============================================================
# BON DE SORTIE
# ============================================================
elif choix_menu == "Bon de Sortie (BS)":
    st.header("📤 Bon de Sortie (BS)")

    editing_id = st.session_state.editing_transaction_id
    editing = (
        get_transaction(editing_id)
        if editing_id
        else None
    )

    if editing and editing.get("type") != "BS":
        editing = None
        st.session_state.editing_transaction_id = None

    if editing:
        st.info(
            f"Modification du bon : {editing['id']}"
        )

        default_date = datetime.date.fromisoformat(
            editing["date"]
        )
        default_team = editing.get(
            "fournisseur_equipe",
            "",
        )
        default_destination = editing.get(
            "destination",
            "",
        )
        default_remarque = editing.get(
            "remarque",
            "",
        )
    else:
        default_date = today()

        default_team = (
            db["equipes"][0]
            if db["equipes"]
            else ""
        )

        default_destination = ""
        default_remarque = ""

    if editing:
        work_key = f"bs_working_{editing_id}"

        if work_key not in st.session_state:
            st.session_state[work_key] = [
                dict(item)
                for item in editing.get(
                    "articles",
                    [],
                )
            ]

        working_articles = (
            st.session_state[work_key]
        )
    else:
        working_articles = (
            st.session_state.current_bs_articles
        )

    col1, col2 = st.columns(2)

    with col1:
        date_bs = st.date_input(
            "Date du BS",
            value=default_date,
            max_value=today(),
            key=f"bs_date_{editing_id or 'new'}",
        )

        team_options = list(db["equipes"])

        if not team_options:
            team_options = [
                "Aucune équipe configurée"
            ]

        team_index = (
            team_options.index(
                default_team
            )
            if default_team in team_options
            else 0
        )

        equipe = st.selectbox(
            "Équipe destinataire",
            team_options,
            index=team_index,
            key=f"bs_team_{editing_id or 'new'}",
        )

    with col2:
        destination = st.text_input(
            "Destination / Site",
            value=default_destination,
            key=f"bs_destination_{editing_id or 'new'}",
        )

        remarque_bon = st.text_area(
            "Remarque générale",
            value=default_remarque,
            key=f"bs_remarque_{editing_id or 'new'}",
        )

    if editing:
        st.markdown(
            f"**Numéro BS :** `{editing['id']}`  \n"
            f"**Date et heure de saisie :** "
            f"`{editing.get('date','')} "
            f"{editing.get('heure_saisie','')}`  \n"
            f"**Équipe :** "
            f"`{editing.get('fournisseur_equipe','')}`"
        )
    else:
        st.caption(
            "Numéro automatique : "
            "BS-MW-YYYYMMDD-01"
        )

    st.subheader("Articles")

    articles_list = article_designations()

    if not articles_list:
        st.warning(
            "Aucun article n'est configuré. "
            "L'administrateur doit d'abord ajouter les articles."
        )
    else:
        stock_form = get_stock(
            client,
            exclude_transaction_id=(
                editing_id
                if editing
                else None
            ),
        )

        with st.form(
            f"bs_add_article_{editing_id or 'new'}"
        ):
            c1, c2, c3 = st.columns(
                [2, 1, 2]
            )

            with c1:
                article_sel = st.selectbox(
                    "Article",
                    articles_list,
                    key=f"bs_article_{editing_id or 'new'}",
                )

            with c2:
                qte = st.number_input(
                    "Quantité",
                    min_value=1,
                    value=1,
                    step=1,
                    key=f"bs_qte_{editing_id or 'new'}",
                )

            with c3:
                remarque_article = st.text_input(
                    "Remarque",
                    key=f"bs_art_rem_{editing_id or 'new'}",
                )

            available = safe_int(
                stock_form.get(
                    article_sel,
                    {},
                ).get(
                    "qte",
                    0,
                )
            )

            st.caption(
                f"Stock disponible : {available}"
            )

            add_article = st.form_submit_button(
                "Ajouter l'article"
            )

            if add_article:
                if not remarque_article.strip():
                    st.error("La remarque de l'article est obligatoire.")
                    st.stop()
                already = sum(
                    safe_int(
                        item.get(
                            "qte",
                            0,
                        )
                    )
                    for item
                    in working_articles
                    if item.get(
                        "designation"
                    ) == article_sel
                )

                remaining = (
                    available
                    - already
                )

                if qte <= 0:
                    st.error(
                        "La quantité doit être supérieure à 0."
                    )
                elif qte > remaining:
                    st.error(
                        "Stock insuffisant. "
                        f"Disponible restant : {remaining}."
                    )
                else:
                    found = False
                    ref = article_ref(
                        article_sel
                    )

                    for item in working_articles:
                        if item.get(
                            "designation"
                        ) == article_sel:
                            item["qte"] = (
                                safe_int(
                                    item.get(
                                        "qte",
                                        0,
                                    )
                                )
                                + int(qte)
                            )

                            if remarque_article.strip():
                                old_remark = item.get(
                                    "remarque",
                                    "",
                                )

                                item["remarque"] = (
                                    f"{old_remark} | "
                                    f"{remarque_article.strip()}"
                                    if old_remark
                                    else remarque_article.strip()
                                )

                            found = True
                            break

                    if not found:
                        working_articles.append({
                            "ref": ref,
                            "designation": article_sel,
                            "qte": int(qte),
                            "remarque": remarque_article.strip(),
                        })

                    st.rerun()

    if working_articles:
        st.markdown("### Articles du bon")

        rows = [
            {
                "Référence": item.get(
                    "ref",
                    "",
                ),
                "Désignation": item.get(
                    "designation",
                    "",
                ),
                "Quantité": item.get(
                    "qte",
                    0,
                ),
                "Remarque": item.get(
                    "remarque",
                    "",
                ),
            }
            for item in working_articles
        ]

        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown(
            "#### Modifier / supprimer une ligne"
        )

        selected_line = st.selectbox(
            "Article",
            [
                item.get("designation")
                for item in working_articles
            ],
            key=f"bs_line_select_{editing_id or 'new'}",
        )

        line_index = next(
            (
                i
                for i, item
                in enumerate(working_articles)
                if item.get("designation")
                == selected_line
            ),
            None,
        )

        if line_index is not None:
            c1, c2, c3 = st.columns(
                [1, 1, 1]
            )

            with c1:
                line_qty = st.number_input(
                    "Nouvelle quantité",
                    min_value=1,
                    value=safe_int(
                        working_articles[
                            line_index
                        ].get(
                            "qte",
                            1,
                        )
                    ),
                    step=1,
                    key=f"bs_line_qty_{editing_id or 'new'}",
                )

            with c2:
                line_remark = st.text_input(
                    "Nouvelle remarque",
                    value=working_articles[
                        line_index
                    ].get(
                        "remarque",
                        "",
                    ),
                    key=f"bs_line_remark_{editing_id or 'new'}",
                )

            with c3:
                confirm_bs_line = confirmation("Confirmer la modification de la ligne.", f"confirm_bs_line_{editing_id or 'new'}")
                if st.button(
                    "Appliquer",
                    key=f"bs_apply_{editing_id or 'new'}",
                    use_container_width=True,
                ):
                    if not confirm_bs_line:
                        st.warning("Veuillez confirmer la modification de la ligne.")
                        st.stop()
                    if not line_remark.strip():
                        st.error("La remarque de l'article est obligatoire.")
                        st.stop()
                    working_articles[
                        line_index
                    ]["qte"] = int(line_qty)

                    working_articles[
                        line_index
                    ]["remarque"] = line_remark

                    st.rerun()

                if st.button(
                    "Supprimer",
                    key=f"bs_remove_{editing_id or 'new'}",
                    use_container_width=True,
                ):
                    working_articles.pop(
                        line_index
                    )
                    st.rerun()

        if st.button(
            "Vider la liste",
            key=f"bs_clear_{editing_id or 'new'}",
        ):
            working_articles.clear()
            st.rerun()

        st.divider()

        if editing:
            c1, c2 = st.columns(2)

            with c1:
                if st.button(
                    "💾 Enregistrer les modifications",
                    type="primary",
                    use_container_width=True,
                ):
                    if date_bs > today():
                        st.error(
                            "La date du BS ne peut pas "
                            "dépasser aujourd'hui."
                        )
                    elif (
                        not equipe
                        or equipe
                        == "Aucune équipe configurée"
                    ):
                        st.error(
                            "L'équipe est obligatoire."
                        )
                    else:
                        valid, message = (
                            validate_articles(
                                working_articles
                            )
                        )

                        if not valid:
                            st.error(message)
                        else:
                            valid_stock, stock_message = (
                                validate_bs_stock(
                                    client,
                                    working_articles,
                                    old_transaction_id=editing["id"],
                                )
                            )

                            if not valid_stock:
                                st.error(
                                    stock_message
                                )
                            else:
                                editing["date"] = (
                                    date_bs.strftime(
                                        "%Y-%m-%d"
                                    )
                                )
                                editing[
                                    "fournisseur_equipe"
                                ] = equipe
                                editing[
                                    "destination"
                                ] = destination
                                editing[
                                    "remarque"
                                ] = remarque_bon
                                editing[
                                    "articles"
                                ] = [
                                    dict(item)
                                    for item
                                    in working_articles
                                ]

                                save_db(db)

                                st.session_state.pop(
                                    work_key,
                                    None,
                                )
                                st.session_state.editing_transaction_id = None

                                st.success(
                                    f"BS {editing['id']} "
                                    "modifié avec succès."
                                )
                                st.rerun()

            with c2:
                if st.button(
                    "Annuler la modification",
                    use_container_width=True,
                ):
                    st.session_state.pop(
                        work_key,
                        None,
                    )
                    st.session_state.editing_transaction_id = None
                    st.rerun()

        else:
            confirm_bs_new = confirmation("Je confirme l'enregistrement définitif du Bon de Sortie.", "confirm_bs_new")
            if st.button(
                "💾 Enregistrer le Bon de Sortie",
                type="primary",
                use_container_width=True,
            ):
                if date_bs > today():
                    st.error(
                        "La date du BS ne peut pas "
                        "dépasser aujourd'hui."
                    )
                elif (
                    not equipe
                    or equipe
                    == "Aucune équipe configurée"
                ):
                    st.error(
                        "L'équipe est obligatoire."
                    )
                else:
                    valid, message = (
                        validate_articles(
                            working_articles
                        )
                    )

                    if not valid:
                        st.error(message)
                    elif not destination.strip():
                        st.error("La destination est obligatoire.")
                    elif not remarque_bon.strip():
                        st.error("La remarque générale est obligatoire.")
                    elif not all_article_remarks_present(working_articles):
                        st.error("La remarque de chaque article est obligatoire.")
                    else:
                        valid_stock, stock_message = (
                            validate_bs_stock(
                                client,
                                working_articles,
                            )
                        )

                        if not valid_stock:
                            st.error(
                                stock_message
                            )
                        else:
                            new_bs = {
                                "id": generate_bon_id("BS"),
                                "type": "BS",
                                "date": date_bs.strftime(
                                    "%Y-%m-%d"
                                ),
                                "heure_saisie": now().strftime(
                                    "%H:%M:%S"
                                ),
                                "client": client,
                                "user": st.session_state.user,
                                "fournisseur_equipe": equipe,
                                "destination": destination,
                                "remarque": remarque_bon,
                                "articles": [
                                    dict(item)
                                    for item
                                    in working_articles
                                ],
                            }

                            db["transactions"].append(
                                new_bs
                            )

                            save_db(db)

                            st.session_state.current_bs_articles = []

                            st.success(
                                f"BS {new_bs['id']} "
                                "enregistré avec succès."
                            )

                            pdf_data = generate_pdf(
                                new_bs,
                                client,
                            )

                            st.download_button(
                                "📄 Télécharger / imprimer le BS en PDF",
                                data=pdf_data,
                                file_name=f"{new_bs['id']}.pdf",
                                mime="application/pdf",
                                use_container_width=True,
                            )

                            if DOCX_AVAILABLE:
                                docx_data = generate_docx(
                                    new_bs,
                                    client,
                                )

                                st.download_button(
                                    "📝 Télécharger le BS en Word",
                                    data=docx_data,
                                    file_name=f"{new_bs['id']}.docx",
                                    mime=(
                                        "application/vnd.openxmlformats-"
                                        "officedocument.wordprocessingml.document"
                                    ),
                                    use_container_width=True,
                                )


# ============================================================
# SITUATION STOCK
# ============================================================
elif choix_menu == "Situation Stock":
    st.header(
        f"📊 Situation Stock — {client}"
    )

    df_stock = stock_dataframe(client)

    if df_stock.empty:
        st.info(
            "Aucun article configuré."
        )
    else:
        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Articles",
            len(df_stock),
        )

        c2.metric(
            "Quantité totale",
            int(
                df_stock[
                    "Quantité Disponible"
                ].sum()
            ),
        )

        c3.metric(
            "Client",
            client,
        )

        st.dataframe(
            df_stock,
            use_container_width=True,
            hide_index=True,
        )

        csv_data = df_stock.to_csv(
            index=False
        ).encode(
            "utf-8-sig"
        )

        st.download_button(
            "📥 Exporter la situation stock",
            data=csv_data,
            file_name=(
                f"Stock_{client}_"
                f"{now().strftime('%Y%m%d_%H%M%S')}.csv"
            ),
            mime="text/csv",
        )

        if st.button(
            "🖨️ Générer la situation stock PDF"
        ):
            pdf = FPDF(
                orientation="P",
                unit="mm",
                format="A4",
            )
            pdf.add_page()

            if os.path.exists(
                LOGOS["NOMATIS"]
            ):
                try:
                    pdf.image(
                        LOGOS["NOMATIS"],
                        x=12,
                        y=10,
                        w=38,
                    )
                except Exception:
                    pass

            pdf.set_xy(12, 34)
            pdf.set_font(
                "Arial",
                "B",
                14,
            )

            pdf.cell(
                0,
                8,
                pdf_safe(
                    f"Situation Stock - {client}"
                ),
                ln=1,
                align="C",
            )

            pdf.set_font(
                "Arial",
                "",
                8,
            )

            pdf.cell(
                0,
                5,
                pdf_safe(
                    "Date / heure : "
                    + now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                ),
                ln=1,
                align="C",
            )

            pdf.ln(5)

            pdf.set_font(
                "Arial",
                "B",
                8,
            )

            pdf.cell(
                40,
                7,
                "Reference",
                border=1,
                align="C",
            )
            pdf.cell(
                100,
                7,
                "Designation",
                border=1,
                align="C",
            )
            pdf.cell(
                40,
                7,
                "Quantite",
                border=1,
                align="C",
            )
            pdf.ln()

            pdf.set_font(
                "Arial",
                "",
                8,
            )

            for _, row in df_stock.iterrows():
                pdf.cell(
                    40,
                    7,
                    pdf_safe(
                        row["Référence"]
                    ),
                    border=1,
                )

                pdf.cell(
                    100,
                    7,
                    pdf_safe(
                        row["Désignation"]
                    ),
                    border=1,
                )

                pdf.cell(
                    40,
                    7,
                    format_qte(
                        row[
                            "Quantité Disponible"
                        ]
                    ),
                    border=1,
                    align="C",
                )

                pdf.ln()

            stock_pdf = bytes(
                pdf.output(dest="S")
            )

            st.download_button(
                "📄 Télécharger la situation stock PDF",
                data=stock_pdf,
                file_name=(
                    f"Situation_Stock_{client}_"
                    f"{now().strftime('%Y%m%d_%H%M%S')}.pdf"
                ),
                mime="application/pdf",
            )


# ============================================================
# HISTORIQUE
# ============================================================
elif choix_menu == "Historique":
    st.header("🕒 Historique")

    tab_be, tab_bs, tab_mov = st.tabs(
        [
            "Bons d'Entrée (BE)",
            "Bons de Sortie (BS)",
            "Mouvements de stock",
        ]
    )

    def history_section(type_bon):
        transactions = get_transactions(
            client,
            type_bon,
        )

        if not transactions:
            st.info(
                f"Aucun {type_bon} enregistré pour {client}."
            )
            return

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            filter_id = st.text_input(
                "N° Bon",
                key=f"history_id_{type_bon}",
            ).strip().lower()

        with c2:
            filter_article = st.text_input(
                "Article",
                key=f"history_article_{type_bon}",
            ).strip().lower()

        with c3:
            third_choices = sorted({
                t.get(
                    "fournisseur_equipe",
                    "",
                )
                for t in transactions
                if t.get(
                    "fournisseur_equipe"
                )
            })

            if type_bon == "BE":
                third_label = "Fournisseur"
                all_label = "Tous"
            else:
                third_label = "Équipe"
                all_label = "Toutes"

            filter_third = st.selectbox(
                third_label,
                [all_label] + third_choices,
                key=f"history_third_{type_bon}",
            )

        with c4:
            user_choices = sorted({
                t.get("user", "")
                for t in transactions
                if t.get("user")
            })

            filter_user = st.selectbox(
                "Utilisateur",
                ["Tous"] + user_choices,
                key=f"history_user_{type_bon}",
            )

        filtered = []

        for transaction in transactions:
            if (
                filter_id
                and filter_id
                not in str(
                    transaction.get(
                        "id",
                        "",
                    )
                ).lower()
            ):
                continue

            if filter_article:
                article_found = any(
                    filter_article
                    in str(
                        item.get(
                            "designation",
                            "",
                        )
                    ).lower()
                    for item
                    in transaction.get(
                        "articles",
                        [],
                    )
                )

                if not article_found:
                    continue

            third = transaction.get(
                "fournisseur_equipe",
                "",
            )

            if (
                filter_third != all_label
                and third != filter_third
            ):
                continue

            if (
                filter_user != "Tous"
                and transaction.get(
                    "user"
                ) != filter_user
            ):
                continue

            filtered.append(
                transaction
            )

        st.caption(
            f"{len(filtered)} bon(s) trouvé(s)."
        )

        for transaction in reversed(
            filtered
        ):
            title = (
                f"{transaction.get('id')} | "
                f"Date : {transaction.get('date')} | "
                f"Utilisateur : {transaction.get('user')}"
            )

            with st.expander(title):
                c1, c2, c3 = st.columns(3)

                c1.write(
                    f"**Client :** "
                    f"{transaction.get('client')}"
                )

                c2.write(
                    f"**Fournisseur / Équipe :** "
                    f"{transaction.get('fournisseur_equipe')}"
                )

                c3.write(
                    f"**Lieu / Destination :** "
                    f"{transaction.get('destination')}"
                )

                st.write(
                    f"**Date :** "
                    f"{transaction.get('date')} | "
                    f"**Heure de saisie :** "
                    f"{transaction.get('heure_saisie', '')}"
                )

                article_rows = [
                    {
                        "Référence": item.get(
                            "ref",
                            "",
                        ),
                        "Désignation": item.get(
                            "designation",
                            "",
                        ),
                        "Quantité": item.get(
                            "qte",
                            0,
                        ),
                        "Remarque": item.get(
                            "remarque",
                            "",
                        ),
                    }
                    for item
                    in transaction.get(
                        "articles",
                        [],
                    )
                ]

                st.dataframe(
                    pd.DataFrame(article_rows),
                    use_container_width=True,
                    hide_index=True,
                )

                if transaction.get(
                    "remarque"
                ):
                    st.write(
                        "**Remarque :** "
                        + str(
                            transaction.get(
                                "remarque"
                            )
                        )
                    )

                c1, c2, c3 = st.columns(3)

                with c1:
                    if can_print():
                        pdf_data = generate_pdf(
                            transaction,
                            client,
                        )

                        st.download_button(
                            "🖨️ Imprimer PDF",
                            data=pdf_data,
                            file_name=(
                                f"{transaction['id']}.pdf"
                            ),
                            mime="application/pdf",
                            key=(
                                f"pdf_{transaction['id']}"
                            ),
                            use_container_width=True,
                        )

                with c2:
                    if (
                        can_print()
                        and DOCX_AVAILABLE
                    ):
                        docx_data = generate_docx(
                            transaction,
                            client,
                        )

                        st.download_button(
                            "📝 Imprimer Word",
                            data=docx_data,
                            file_name=(
                                f"{transaction['id']}.docx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.wordprocessingml.document"
                            ),
                            key=(
                                f"docx_{transaction['id']}"
                            ),
                            use_container_width=True,
                        )

                with c3:
                    if can_edit():
                        confirm_edit_bon = confirmation("Confirmer la modification de ce bon.", f"confirm_edit_bon_{transaction['id']}")
                        if st.button(
                            "✏️ Modifier",
                            key=(
                                f"edit_{transaction['id']}"
                            ),
                            use_container_width=True,
                        ):
                            if not confirm_edit_bon:
                                st.warning("Veuillez confirmer la modification.")
                                st.stop()
                            st.session_state.editing_transaction_id = (
                                transaction["id"]
                            )

                            st.session_state.current_be_articles = []
                            st.session_state.current_bs_articles = []

                            if transaction.get(
                                "type"
                            ) == "BE":
                                st.session_state.selected_menu = (
                                    "Bon d'Entrée (BE)"
                                )
                            else:
                                st.session_state.selected_menu = (
                                    "Bon de Sortie (BS)"
                                )

                            st.rerun()

                if can_edit():
                    confirm_delete_bon = confirmation("Confirmer la suppression définitive de ce bon.", f"confirm_delete_bon_{transaction['id']}")
                    if st.button(
                        "🗑️ Supprimer ce bon",
                        key=(
                            f"delete_{transaction['id']}"
                        ),
                        use_container_width=True,
                    ):
                        if not confirm_delete_bon:
                            st.warning("Veuillez confirmer la suppression.")
                            st.stop()
                        db["transactions"] = [
                            item
                            for item
                            in db["transactions"]
                            if item.get("id")
                            != transaction.get("id")
                        ]

                        save_db(db)

                        st.success(
                            f"{transaction['id']} supprimé. "
                            "La situation du stock a été recalculée."
                        )

                        st.rerun()

    with tab_be:
        history_section("BE")

    with tab_bs:
        history_section("BS")

    with tab_mov:
        transactions = get_transactions(
            client
        )

        if not transactions:
            st.info(
                "Aucun mouvement pour ce client."
            )
        else:
            movement_rows = []

            for transaction in transactions:
                for item in transaction.get(
                    "articles",
                    [],
                ):
                    qte = safe_int(
                        item.get(
                            "qte",
                            0,
                        )
                    )

                    movement_rows.append({
                        "Date": transaction.get(
                            "date",
                            "",
                        ),
                        "Heure": transaction.get(
                            "heure_saisie",
                            "",
                        ),
                        "Type": transaction.get(
                            "type",
                            "",
                        ),
                        "N° Bon": transaction.get(
                            "id",
                            "",
                        ),
                        "Référence": item.get(
                            "ref",
                            "",
                        ),
                        "Désignation": item.get(
                            "designation",
                            "",
                        ),
                        "Entrée": (
                            qte
                            if transaction.get(
                                "type"
                            )
                            in (
                                "BE",
                                "ADJ_PLUS",
                            )
                            else 0
                        ),
                        "Sortie": (
                            qte
                            if transaction.get(
                                "type"
                            )
                            in (
                                "BS",
                                "ADJ_MOINS",
                            )
                            else 0
                        ),
                        "Fournisseur / Équipe": transaction.get(
                            "fournisseur_equipe",
                            "",
                        ),
                        "Utilisateur": transaction.get(
                            "user",
                            "",
                        ),
                        "Destination / Motif": transaction.get(
                            "destination",
                            "",
                        ),
                    })

            movement_df = pd.DataFrame(
                movement_rows
            )

            c1, c2, c3 = st.columns(3)

            with c1:
                movement_article = st.text_input(
                    "Article",
                    key="movement_article",
                ).strip().lower()

            with c2:
                movement_third = st.text_input(
                    "Fournisseur / Équipe",
                    key="movement_third",
                ).strip().lower()

            with c3:
                movement_type = st.selectbox(
                    "Type",
                    [
                        "Tous",
                        "BE",
                        "BS",
                        "ADJ_PLUS",
                        "ADJ_MOINS",
                    ],
                    key="movement_type",
                )

            if movement_article:
                movement_df = movement_df[
                    movement_df[
                        "Désignation"
                    ]
                    .astype(str)
                    .str.lower()
                    .str.contains(
                        movement_article,
                        na=False,
                    )
                ]

            if movement_third:
                movement_df = movement_df[
                    movement_df[
                        "Fournisseur / Équipe"
                    ]
                    .astype(str)
                    .str.lower()
                    .str.contains(
                        movement_third,
                        na=False,
                    )
                ]

            if movement_type != "Tous":
                movement_df = movement_df[
                    movement_df["Type"]
                    == movement_type
                ]

            st.dataframe(
                movement_df.sort_values(
                    [
                        "Date",
                        "Heure",
                    ],
                    ascending=False,
                ),
                use_container_width=True,
                hide_index=True,
            )


# ============================================================
# CONFIGURATION ADMIN
# ============================================================
elif choix_menu == "Configuration":
    if role != "admin":
        st.error(
            "Accès réservé à l'administrateur."
        )
        st.stop()

    st.header(
        "⚙️ Configuration Admin"
    )

    tab_users, tab_articles, tab_suppliers, tab_teams, tab_adjust = st.tabs(
        [
            "Utilisateurs",
            "Articles",
            "Fournisseurs",
            "Équipes",
            "Ajustement Stock",
        ]
    )

    # --------------------------------------------------------
    # UTILISATEURS
    # --------------------------------------------------------
    with tab_users:
        st.subheader(
            "Gestion des utilisateurs"
        )

        user_rows = []

        for username, info in db["users"].items():
            user_rows.append({
                "Nom": username,
                "Rôle": info.get(
                    "role",
                    "",
                ),
                "Dernière connexion": info.get(
                    "last_login",
                    "",
                ),
            })

        st.dataframe(
            pd.DataFrame(user_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown(
            "### Créer un utilisateur"
        )

        with st.form("create_user"):
            c1, c2 = st.columns(2)

            with c1:
                username = st.text_input(
                    "Nom d'utilisateur"
                )
                password = st.text_input(
                    "Mot de passe",
                    type="password",
                )

            with c2:
                user_role = st.selectbox(
                    "Rôle",
                    [
                        "admin",
                        "magasinier",
                        "coordinateur",
                        "coordinatrice",
                    ],
                )

            st.checkbox("Je confirme la création de cet utilisateur.", key="confirm_create_user")
            create = st.form_submit_button(
                "Créer l'utilisateur",
                type="primary",
            )

            if create:
                if not st.session_state.get("confirm_create_user", False):
                    st.warning("Veuillez confirmer la création de l'utilisateur.")
                    st.stop()
                username = username.strip()

                if not username or not password:
                    st.error(
                        "Nom et mot de passe obligatoires."
                    )
                elif username in db["users"]:
                    st.error(
                        "Cet utilisateur existe déjà."
                    )
                else:
                    db["users"][username] = {
                        "password": password,
                        "role": user_role,
                        "last_login": "",
                    }

                    save_db(db)

                    st.success(
                        "Utilisateur créé."
                    )
                    st.rerun()

        st.markdown(
            "### Modifier un utilisateur"
        )

        usernames = list(
            db["users"].keys()
        )

        selected_user = st.selectbox(
            "Utilisateur",
            usernames,
            key="selected_user_admin",
        )

        current_info = db["users"][
            selected_user
        ]

        with st.form("edit_user"):
            new_name = st.text_input(
                "Nouveau nom",
                value=selected_user,
            )

            new_password = st.text_input(
                "Nouveau mot de passe "
                "(laisser vide pour conserver)",
                type="password",
            )

            current_role = current_info.get(
                "role",
                "magasinier",
            )

            new_role = st.selectbox(
                "Nouveau rôle",
                [
                    "admin",
                    "magasinier",
                    "coordinateur",
                    "coordinatrice",
                ],
                index=(
                    [
                        "admin",
                        "magasinier",
                        "coordinateur",
                        "coordinatrice",
                    ].index(
                        current_role
                    )
                    if current_role
                    in [
                        "admin",
                        "magasinier",
                        "coordinateur",
                        "coordinatrice",
                    ]
                    else 1
                ),
            )

            st.checkbox("Je confirme la modification de cet utilisateur.", key="confirm_edit_user")
            save_user = st.form_submit_button(
                "Enregistrer les modifications"
            )

            if save_user:
                if not st.session_state.get("confirm_edit_user", False):
                    st.warning("Veuillez confirmer la modification.")
                    st.stop()
                new_name = new_name.strip()

                if not new_name:
                    st.error(
                        "Le nom ne peut pas être vide."
                    )
                elif (
                    new_name != selected_user
                    and new_name in db["users"]
                ):
                    st.error(
                        "Ce nom existe déjà."
                    )
                elif (
                    selected_user == "admin"
                    and new_role != "admin"
                ):
                    st.error(
                        "Le compte admin principal "
                        "doit rester admin."
                    )
                else:
                    info = db["users"].pop(
                        selected_user
                    )

                    info["role"] = new_role

                    if new_password:
                        info["password"] = new_password

                    db["users"][new_name] = info

                    save_db(db)

                    if (
                        st.session_state.user
                        == selected_user
                    ):
                        st.session_state.user = new_name
                        st.session_state.role = new_role

                    st.success(
                        "Utilisateur modifié."
                    )
                    st.rerun()

        if selected_user != "admin":
            confirm_delete_user = confirmation("Je confirme la suppression définitive de cet utilisateur.", "confirm_delete_user")
            if st.button(
                "🗑️ Supprimer l'utilisateur"
            ):
                if not confirm_delete_user:
                    st.warning("Veuillez confirmer la suppression.")
                    st.stop()
                db["users"].pop(
                    selected_user,
                    None,
                )

                save_db(db)

                st.success(
                    "Utilisateur supprimé."
                )
                st.rerun()

    # --------------------------------------------------------
    # ARTICLES
    # --------------------------------------------------------
    with tab_articles:
        st.subheader("Référentiel Articles")

        article_rows = []
        for item in db["articles"]:
            article_rows.append({
                "Index": item.get("ref", ""),
                "Catégorie": item.get("categorie", "Sans catégorie"),
                "Type": item.get("type", ""),
                "Caractéristique": item.get("caracteristique", ""),
                "Désignation": item.get("designation", ""),
            })

        if article_rows:
            st.dataframe(pd.DataFrame(article_rows), use_container_width=True, hide_index=True)
        else:
            st.info("Aucun article configuré.")

        st.markdown("### Ajouter un article")
        st.caption("L'index est généré automatiquement au format MW-001, MW-002, MW-003...")

        with st.form("add_article"):
            category = st.selectbox("Catégorie", list(ARTICLE_CATEGORIES.keys()))

            if category == "Sans catégorie":
                base_name = st.text_input("Désignation", placeholder="Ex. Câble IF ou Clamp")
                dimension = ""
                fibre_type = ""
                longueur = ""
            elif category == "Support":
                dimension = st.selectbox("Dimension du support", ARTICLE_CATEGORIES[category]["dimensions"])
                base_name = ""
                fibre_type = ""
                longueur = ""
            else:
                fibre_type = st.selectbox("Type de jarretière", ARTICLE_CATEGORIES[category]["types"])
                longueur = st.selectbox("Longueur", ARTICLE_CATEGORIES[category]["longueurs"])
                dimension = ""
                base_name = ""

            initial_qty = st.number_input(f"Quantité initiale pour {client}", min_value=0, value=0, step=1)
            st.checkbox("Je confirme l'ajout de cet article.", key="confirm_add_article")
            add = st.form_submit_button("Ajouter l'article", type="primary")

            if add:
                if not st.session_state.get("confirm_add_article", False):
                    st.warning("Veuillez confirmer l'ajout de l'article.")
                    st.stop()
                designation = build_article_designation(category, dimension, fibre_type, longueur, base_name)
                duplicate = any(a.get("designation", "").strip().lower() == designation.lower() for a in db["articles"])

                if not designation:
                    st.error("La désignation est obligatoire.")
                elif duplicate:
                    st.error("Cet article existe déjà.")
                else:
                    new_ref = next_article_ref()
                    db["articles"].append({
                        "ref": new_ref,
                        "designation": designation,
                        "categorie": category,
                        "type": fibre_type,
                        "caracteristique": dimension or longueur,
                    })

                    if initial_qty > 0:
                        db["transactions"].append({
                            "id": f"ADJ-INIT-{client}-{now().strftime('%Y%m%d%H%M%S')}",
                            "type": "ADJ_PLUS",
                            "date": now().strftime("%Y-%m-%d"),
                            "heure_saisie": now().strftime("%H:%M:%S"),
                            "client": client,
                            "user": st.session_state.user,
                            "fournisseur_equipe": "Ajustement",
                            "destination": "Stock initial",
                            "remarque": "Quantité initiale",
                            "articles": [{"designation": designation, "qte": int(initial_qty), "ref": new_ref, "remarque": "Stock initial"}],
                        })

                    save_db(db)
                    st.success(f"Article {new_ref} ajouté avec succès.")
                    st.rerun()

        if db["articles"]:
            st.markdown("### Modifier un article")
            article_labels = [f"{a.get('ref','')} — {a.get('designation','')}" for a in db["articles"]]
            selected_label = st.selectbox("Article", article_labels, key="selected_article_admin")
            article_index = article_labels.index(selected_label)
            current = db["articles"][article_index]

            with st.form("edit_article"):
                st.text_input("Index", value=current.get("ref", ""), disabled=True)
                new_designation = st.text_input("Désignation", value=current.get("designation", ""))
                st.caption(f"Catégorie : {current.get('categorie','Sans catégorie')} | Caractéristique : {current.get('caracteristique','')}")
                st.checkbox("Je confirme la modification de cet article.", key="confirm_edit_article")
                save_article = st.form_submit_button("Enregistrer les modifications", type="primary")

                if save_article:
                    if not st.session_state.get("confirm_edit_article", False):
                        st.warning("Veuillez confirmer la modification de l'article.")
                        st.stop()
                    new_designation = new_designation.strip()
                    duplicate = any(i != article_index and a.get("designation", "").lower() == new_designation.lower() for i, a in enumerate(db["articles"]))
                    if not new_designation:
                        st.error("La désignation est obligatoire.")
                    elif duplicate:
                        st.error("Cette désignation existe déjà.")
                    else:
                        old_designation = current.get("designation", "")
                        current["designation"] = new_designation
                        for transaction in db["transactions"]:
                            for item in transaction.get("articles", []):
                                if item.get("designation") == old_designation:
                                    item["designation"] = new_designation
                                    item["ref"] = current.get("ref", "")
                        save_db(db)
                        st.success("Article modifié.")
                        st.rerun()

            confirm_delete_article = confirmation("Je confirme la suppression définitive de cet article.", "confirm_delete_article")
            if st.button("🗑️ Supprimer l'article"):
                if not confirm_delete_article:
                    st.warning("Veuillez confirmer la suppression de l'article.")
                    st.stop()
                used_in_history = any(any(item.get("designation") == current.get("designation") for item in transaction.get("articles", [])) for transaction in db["transactions"])
                if used_in_history:
                    st.error("Impossible de supprimer cet article : il possède déjà des mouvements historiques.")
                else:
                    db["articles"].pop(article_index)
                    save_db(db)
                    st.success("Article supprimé.")
                    st.rerun()

    # --------------------------------------------------------
    # FOURNISSEURS
    # --------------------------------------------------------
    with tab_suppliers:
        st.subheader(
            "Fournisseurs"
        )

        st.write(
            ", ".join(
                db["fournisseurs"]
            )
            if db["fournisseurs"]
            else "Aucun fournisseur."
        )

        with st.form("add_supplier"):
            supplier = st.text_input(
                "Nouveau fournisseur"
            )

            st.checkbox("Je confirme l'ajout de ce fournisseur.", key="confirm_add_supplier")
            add_supplier = st.form_submit_button(
                "Ajouter"
            )

            if add_supplier:
                if not st.session_state.get("confirm_add_supplier", False):
                    st.warning("Veuillez confirmer l'ajout du fournisseur.")
                    st.stop()
                supplier = supplier.strip()

                if not supplier:
                    st.error(
                        "Le fournisseur ne peut pas être vide."
                    )
                elif (
                    supplier
                    in db["fournisseurs"]
                ):
                    st.error(
                        "Ce fournisseur existe déjà."
                    )
                else:
                    db["fournisseurs"].append(
                        supplier
                    )

                    save_db(db)

                    st.success(
                        "Fournisseur ajouté."
                    )
                    st.rerun()

        if db["fournisseurs"]:
            supplier_delete = st.selectbox(
                "Fournisseur à supprimer",
                db["fournisseurs"],
                key="supplier_delete",
            )

            confirm_delete_supplier = confirmation("Je confirme la suppression de ce fournisseur.", "confirm_delete_supplier")
            if st.button(
                "🗑️ Supprimer le fournisseur"
            ):
                if not confirm_delete_supplier:
                    st.warning("Veuillez confirmer la suppression.")
                    st.stop()
                db["fournisseurs"].remove(
                    supplier_delete
                )

                save_db(db)

                st.success(
                    "Fournisseur supprimé du référentiel."
                )
                st.rerun()

    # --------------------------------------------------------
    # EQUIPES
    # --------------------------------------------------------
    with tab_teams:
        st.subheader(
            "Équipes Projet"
        )

        st.write(
            ", ".join(
                db["equipes"]
            )
            if db["equipes"]
            else "Aucune équipe."
        )

        with st.form("add_team"):
            team = st.text_input(
                "Nouvelle équipe"
            )

            st.checkbox("Je confirme l'ajout de cette équipe.", key="confirm_add_team")
            add_team = st.form_submit_button(
                "Ajouter"
            )

            if add_team:
                if not st.session_state.get("confirm_add_team", False):
                    st.warning("Veuillez confirmer l'ajout de l'équipe.")
                    st.stop()
                team = team.strip()

                if not team:
                    st.error(
                        "Le nom de l'équipe ne peut pas être vide."
                    )
                elif team in db["equipes"]:
                    st.error(
                        "Cette équipe existe déjà."
                    )
                else:
                    db["equipes"].append(
                        team
                    )

                    save_db(db)

                    st.success(
                        "Équipe ajoutée."
                    )
                    st.rerun()

        if db["equipes"]:
            team_delete = st.selectbox(
                "Équipe à supprimer",
                db["equipes"],
                key="team_delete",
            )

            confirm_delete_team = confirmation("Je confirme la suppression de cette équipe.", "confirm_delete_team")
            if st.button(
                "🗑️ Supprimer l'équipe"
            ):
                if not confirm_delete_team:
                    st.warning("Veuillez confirmer la suppression.")
                    st.stop()
                db["equipes"].remove(
                    team_delete
                )

                save_db(db)

                st.success(
                    "Équipe supprimée du référentiel."
                )
                st.rerun()

    # --------------------------------------------------------
    # AJUSTEMENT STOCK
    # --------------------------------------------------------
    with tab_adjust:
        st.subheader(
            f"Ajustement Manuel — {client}"
        )

        st.warning(
            "Fonction réservée à l'administrateur. "
            "Chaque ajustement est enregistré comme un mouvement."
        )

        articles = article_designations()

        if not articles:
            st.info(
                "Ajoutez d'abord des articles."
            )
        else:
            with st.form("adjust_stock"):
                adjustment_article = st.selectbox(
                    "Article",
                    articles,
                )

                direction = st.radio(
                    "Sens",
                    [
                        "Ajouter au stock (+)",
                        "Retirer du stock (-)",
                    ],
                    horizontal=True,
                )

                adjustment_qty = st.number_input(
                    "Quantité",
                    min_value=1,
                    value=1,
                    step=1,
                )

                reason = st.text_input(
                    "Motif"
                )
                st.checkbox("Je confirme l'ajustement manuel du stock.", key="confirm_adjustment")

                apply_adjustment = st.form_submit_button(
                    "Appliquer l'ajustement",
                    type="primary",
                )

                if apply_adjustment:
                    if not st.session_state.get("confirm_adjustment", False):
                        st.warning("Veuillez confirmer l'ajustement avant de continuer.")
                        st.stop()
                    if not reason.strip():
                        st.error(
                            "Le motif est obligatoire."
                        )
                    else:
                        adjustment_type = (
                            "ADJ_PLUS"
                            if "+"
                            in direction
                            else "ADJ_MOINS"
                        )

                        if (
                            adjustment_type
                            == "ADJ_MOINS"
                        ):
                            available = safe_int(
                                get_stock(
                                    client
                                )
                                .get(
                                    adjustment_article,
                                    {},
                                )
                                .get(
                                    "qte",
                                    0,
                                )
                            )

                            if (
                                adjustment_qty
                                > available
                            ):
                                st.error(
                                    f"Impossible de retirer "
                                    f"{adjustment_qty}. "
                                    f"Stock disponible : "
                                    f"{available}."
                                )
                                st.stop()

                        db["transactions"].append({
                            "id": (
                                f"ADJ-{client}-"
                                f"{now().strftime('%Y%m%d%H%M%S')}"
                            ),
                            "type": adjustment_type,
                            "date": now().strftime(
                                "%Y-%m-%d"
                            ),
                            "heure_saisie": now().strftime(
                                "%H:%M:%S"
                            ),
                            "client": client,
                            "user": st.session_state.user,
                            "fournisseur_equipe": "Ajustement",
                            "destination": reason.strip(),
                            "remarque": "Ajustement manuel",
                            "articles": [{
                                "designation": adjustment_article,
                                "qte": int(adjustment_qty),
                                "ref": article_ref(
                                    adjustment_article
                                ),
                                "remarque": reason.strip(),
                            }],
                        })

                        save_db(db)

                        st.success(
                            "Ajustement enregistré et "
                            "appliqué à la situation du stock."
                        )

                        st.rerun()
