import base64
from datetime import date, datetime
from io import BytesIO
import json
import os
import sqlite3
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import pandas as pd

from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

# =========================================================
# CONFIGURATION DE LA PAGE STREAMLIT
# =========================================================
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_TITLE = "Gestion Stock MW NOMATIS"
DB_FILE = "stock_mw.db"

CLIENTS = {
    "Orange": {"logo": "Orange_logo.svg.webp", "color": "#FF6600"},
    "Inwi": {"logo": "Logo INWI.jpg", "color": "#A1006B"},
    "ZTE": {"logo": "Logo ZTE.jpg", "color": "#005BAC"},
}

ROLES = ["admin", "magasinier", "coordinateur", "coordinatrice"]
DEFAULT_ARTICLES = [
    "Câble IF",
    "Câble RJ45",
    "Support 0.3 m",
    "Support 0.6 m",
    "Support 1.2 m",
    "ODU 18GHz",
    "Antenne 0.6m",
]
DEFAULT_FOURNISSEURS = ["NEC", "ZTE", "Intégral", "FO Connect"]
DEFAULT_EQUIPES = ["Nabil Team", "Yassine Team", "Issam Team"]
DEFAULT_RESOURCES = ["Nabil", "Yassine", "Issam"]

# =========================================================
# SYSTEME DE STYLES ET INJECTION CSS DYNAMIQUE
# =========================================================
# Définition dynamique des couleurs selon la validation d'accès
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

btn_access_color = "#10B981" if st.session_state.logged_in else "#EF4444"
btn_access_hover = "#059669" if st.session_state.logged_in else "#DC2626"

css_code = f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        background-color: #F0F4F8 !important;
        color: #0F172A !important;
    }}

    .stApp {{
        background-color: #F0F4F8;
    }}

    /* En-tête principal de la page */
    .main-header {{
        background: #FFFFFF;
        border-top: 5px solid #2563EB;
        border-bottom: 2px solid #10B981;
        border-radius: 8px;
        padding: 16px 24px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }}

    .main-title {{
        color: #1E3A8A !important;
        font-weight: 800;
        font-size: 24px;
        margin: 0;
    }}

    .subtitle {{
        color: #10B981 !important;
        font-size: 14px;
        font-weight: 600;
        margin-top: 4px;
    }}

    /* Cartes d'affichage et conteneurs */
    .glass-card {{
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }}

    /* Stylisation du Bouton de Connexion / Validation */
    div[data-testid="stFormSubmitButton"] > button {{
        background-color: {btn_access_color} !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        border: none !important;
        padding: 10px 20px !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
        font-size: 16px !important;
    }}

    div[data-testid="stFormSubmitButton"] > button:hover {{
        background-color: {btn_access_hover} !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
    }}

    /* Boutons standards Streamlit */
    .stButton > button {{
        border-radius: 6px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
    }}

    /* Champs de formulaires et Entrées de texte */
    .stTextInput input, .stSelectbox div[data-baseweb="select"], .stNumberInput input, .stDateInput input {{
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #94A3B8 !important;
        border-radius: 6px !important;
    }}

    .stTextInput input:focus, .stNumberInput input:focus {{
        border-color: #2563EB !important;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.15) !important;
    }}

    /* Style des Onglets (Tabs) */
    div[data-baseweb="tab-list"] {{
        gap: 6px;
        background-color: #E2E8F0;
        padding: 6px;
        border-radius: 8px;
    }}

    button[data-baseweb="tab"] {{
        border-radius: 6px !important;
        font-weight: 600 !important;
        color: #475569 !important;
    }}

    button[aria-selected="true"] {{
        background-color: #2563EB !important;
        color: #FFFFFF !important;
    }}

    /* Style des Métriques */
    [data-testid="stMetricValue"] {{
        color: #2563EB !important;
        font-weight: 800 !important;
    }}
</style>
"""
st.markdown(css_code, unsafe_allow_html=True)


# =========================================================
# GESTION DE LA BASE DE DONNÉES SQLITE
# =========================================================
def get_conn():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            fullname TEXT NOT NULL,
            role TEXT NOT NULL,
            last_login TEXT
        );

        CREATE TABLE IF NOT EXISTS articles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS fournisseurs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS equipes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS resources (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS stock (
            client TEXT NOT NULL,
            article_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY (client, article_id)
        );

        CREATE TABLE IF NOT EXISTS bons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            number TEXT NOT NULL,
            client TEXT NOT NULL,
            date_bon TEXT NOT NULL,
            datetime_saisie TEXT NOT NULL,
            fournisseur TEXT,
            lieu_livraison TEXT,
            receptionne_par TEXT,
            equipe TEXT,
            resource TEXT,
            destination TEXT,
            created_by TEXT NOT NULL,
            UNIQUE(type, number, client)
        );

        CREATE TABLE IF NOT EXISTS bon_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bon_id INTEGER NOT NULL,
            article_id INTEGER NOT NULL,
            reference TEXT,
            quantity INTEGER NOT NULL,
            remarque TEXT,
            FOREIGN KEY(bon_id) REFERENCES bons(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client TEXT NOT NULL,
            article_id INTEGER NOT NULL,
            movement_type TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            reference_bon TEXT,
            username TEXT NOT NULL,
            created_at TEXT NOT NULL,
            comment TEXT,
            fournisseur TEXT,
            equipe TEXT
        );
        """
    )

    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users VALUES (?,?,?,?,?)",
            ("admin", "admin123", "Administrateur Système", "admin", "Jamais"),
        )
        cur.execute(
            "INSERT INTO users VALUES (?,?,?,?,?)",
            ("magasinier", "123", "Magasinier Principal", "magasinier", "Jamais"),
        )
        cur.execute(
            "INSERT INTO users VALUES (?,?,?,?,?)",
            (
                "coordinateur",
                "123",
                "Coordinateur Projets",
                "coordinateur",
                "Jamais",
            ),
        )

    for name in DEFAULT_ARTICLES:
        cur.execute(
            "INSERT OR IGNORE INTO articles(name,active) VALUES(?,1)", (name,)
        )
    for name in DEFAULT_FOURNISSEURS:
        cur.execute(
            "INSERT OR IGNORE INTO fournisseurs(name,active) VALUES(?,1)",
            (name,),
        )
    for name in DEFAULT_EQUIPES:
        cur.execute(
            "INSERT OR IGNORE INTO equipes(name,active) VALUES(?,1)", (name,)
        )
    for name in DEFAULT_RESOURCES:
        cur.execute(
            "INSERT OR IGNORE INTO resources(name,active) VALUES(?,1)", (name,)
        )

    conn.commit()
    conn.close()


init_db()


# =========================================================
# FONCTIONS REQUÊTES ET SERVICES
# =========================================================
def query(sql, params=(), one=False):
    conn = get_conn()
    cur = conn.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return rows[0] if one and rows else (None if one else rows)


def execute(sql, params=()):
    conn = get_conn()
    cur = conn.execute(sql, params)
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    return last_id


def active_names(table):
    rows = query(f"SELECT name FROM {table} WHERE active=1 ORDER BY name")
    return [r["name"] for r in rows]


def article_id_by_name(name):
    row = query("SELECT id FROM articles WHERE name=?", (name,), one=True)
    return row["id"] if row else None


def current_stock(client, article_id):
    row = query(
        "SELECT quantity FROM stock WHERE client=? AND article_id=?",
        (client, article_id),
        one=True,
    )
    return int(row["quantity"]) if row else 0


def set_stock(client, article_id, quantity):
    execute(
        """
        INSERT INTO stock(client,article_id,quantity) VALUES(?,?,?)
        ON CONFLICT(client,article_id) DO UPDATE SET quantity=excluded.quantity
        """,
        (client, article_id, int(quantity)),
    )


def add_movement(
    client,
    article_id,
    m_type,
    qty,
    ref_bon,
    user,
    comment="",
    fournisseur="",
    equipe="",
):
    execute(
        """
        INSERT INTO movements(client,article_id,movement_type,quantity,reference_bon,username,created_at,comment,fournisseur,equipe)
        VALUES(?,?,?,?,?,?,?,?,?,?)
        """,
        (
            client,
            article_id,
            m_type,
            int(qty),
            ref_bon,
            user,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            comment,
            fournisseur,
            equipe,
        ),
    )


def generate_bon_number(b_type, client):
    today_str = date.today().strftime("%Y%m%d")
    prefix = f"{b_type} MW-{today_str}-"
    rows = query(
        "SELECT number FROM bons WHERE type=? AND client=? AND number LIKE ?",
        (b_type, client, f"{prefix}%"),
    )
    max_seq = 0
    for r in rows:
        try:
            seq = int(r["number"].split("-")[-1])
            if seq > max_seq:
                max_seq = seq
        except ValueError:
            pass
    return f"{prefix}{max_seq + 1:02d}"


def user_info(username):
    return query("SELECT * FROM users WHERE username=?", (username,), one=True)


def normalized_logo(path, size=(300, 120)):
    if not os.path.exists(path):
        return None
    try:
        img = Image.open(path).convert("RGB")
        canvas = Image.new("RGB", size, "#FFFFFF")
        contained = ImageOps.contain(img, size)
        x = (size[0] - contained.width) // 2
        y = (size[1] - contained.height) // 2
        canvas.paste(contained, (x, y))
        return canvas
    except Exception:
        return None


# =========================================================
# GÉNÉRATION DE DOCUMENTS (PDF & DOCX)
# =========================================================
def generate_pdf(bon_id):
    bon = query("SELECT * FROM bons WHERE id=?", (bon_id,), one=True)
    items = query(
        "SELECT bi.*, a.name AS article FROM bon_items bi JOIN articles a ON a.id=bi.article_id WHERE bi.bon_id=?",
        (bon_id,),
    )
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )
    styles = getSampleStyleSheet()
    story = []

    logo_info = CLIENTS.get(bon["client"], {})
    logo_path = logo_info.get("logo")

    header_data = []
    nomatis_logo_path = "nomatis_logo.png"

    cell_nomatis = (
        Paragraph(
            "<b><font size=16 color='#1E3A8A'>NOMATIS</font></b><br/><font size=9 color='#10B981'>MW Stock Engine</font>",
            styles["Normal"],
        )
        if not os.path.exists(nomatis_logo_path)
        else RLImage(nomatis_logo_path, width=40 * mm, height=15 * mm)
    )

    cell_client = (
        Paragraph(
            f"<b><font size=14 color='#0F172A'>{bon['client']}</font></b>",
            styles["Normal"],
        )
        if not (logo_path and os.path.exists(logo_path))
        else RLImage(logo_path, width=40 * mm, height=15 * mm)
    )

    header_data.append([cell_nomatis, cell_client])
    header_table = Table(header_data, colWidths=[90 * mm, 90 * mm])
    header_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ]
        )
    )
    story.append(header_table)
    story.append(Spacer(1, 10))

    title = "BON D'ENTRÉE" if bon["type"] == "BE" else "BON DE SORTIE"
    story.append(
        Paragraph(
            f"<font size=16 color='#2563EB'><b>{title} : {bon['number']}</b></font>",
            styles["Title"],
        )
    )
    story.append(Spacer(1, 10))

    if bon["type"] == "BE":
        info = [
            ["N° Bon", bon["number"], "Date Bon", bon["date_bon"]],
            [
                "Saisie le",
                bon["datetime_saisie"],
                "Fournisseur",
                bon["fournisseur"] or "-",
            ],
            [
                "Lieu Livraison",
                bon["lieu_livraison"] or "-",
                "Réceptionné par",
                bon["receptionne_par"] or "-",
            ],
            ["Client / Projet", bon["client"], "Créé par", bon["created_by"]],
        ]
    else:
        info = [
            ["N° Bon", bon["number"], "Date Bon", bon["date_bon"]],
            [
                "Saisie le",
                bon["datetime_saisie"],
                "Équipe Destination",
                bon["equipe"] or "-",
            ],
            [
                "Ressource / Tech",
                bon["resource"] or "-",
                "Destination / Site",
                bon["destination"] or "-",
            ],
            ["Client / Projet", bon["client"], "Créé par", bon["created_by"]],
        ]

    info_table = Table(info, colWidths=[35 * mm, 55 * mm, 35 * mm, 55 * mm])
    info_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                (
                    "BACKGROUND",
                    (0, 0),
                    (0, -1),
                    colors.HexColor("#F1F5F9"),
                ),
                (
                    "BACKGROUND",
                    (2, 0),
                    (2, -1),
                    colors.HexColor("#F1F5F9"),
                ),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("PADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(info_table)
    story.append(Spacer(1, 15))

    data = [["Référence", "Désignation Article", "Quantité", "Remarque"]]
    for item in items:
        data.append([
            item["reference"] or "-",
            item["article"],
            str(item["quantity"]),
            item["remarque"] or "-",
        ])

    items_table = Table(data, colWidths=[35 * mm, 70 * mm, 25 * mm, 50 * mm])
    items_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2563EB")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(items_table)
    story.append(Spacer(1, 20))

    sig_data = [
        ["Signature / Cachet Magasinier", "Signature / Cachet Destinataire"]
    ]
    sig_table = Table(sig_data, colWidths=[90 * mm, 90 * mm])
    sig_table.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
            ]
        )
    )
    story.append(sig_table)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def generate_docx(bon_id):
    bon = query("SELECT * FROM bons WHERE id=?", (bon_id,), one=True)
    items = query(
        "SELECT bi.*, a.name AS article FROM bon_items bi JOIN articles a ON a.id=bi.article_id WHERE bi.bon_id=?",
        (bon_id,),
    )
    doc = Document()

    p_title = doc.add_paragraph()
    r_title = p_title.add_run(
        f"BON DE {'ENTRÉE' if bon['type']=='BE' else 'SORTIE'} — {bon['number']}"
    )
    r_title.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(37, 99, 235)

    p_info = doc.add_paragraph()
    p_info.add_run(f"Client / Projet : {bon['client']}\n")
    p_info.add_run(f"Date Bon : {bon['date_bon']}\n")
    p_info.add_run(f"Saisie le : {bon['datetime_saisie']}\n")
    if bon["type"] == "BE":
        p_info.add_run(f"Fournisseur : {bon['fournisseur'] or '-'}\n")
        p_info.add_run(f"Lieu Livraison : {bon['lieu_livraison'] or '-'}\n")
        p_info.add_run(f"Réceptionné par : {bon['receptionne_par'] or '-'}\n")
    else:
        p_info.add_run(f"Équipe Destination : {bon['equipe'] or '-'}\n")
        p_info.add_run(f"Ressource / Tech : {bon['resource'] or '-'}\n")
        p_info.add_run(f"Destination / Site : {bon['destination'] or '-'}\n")
    p_info.add_run(f"Créé par : {bon['created_by']}")

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Référence"
    hdr_cells[1].text = "Désignation Article"
    hdr_cells[2].text = "Quantité"
    hdr_cells[3].text = "Remarque"

    for item in items:
        row_cells = table.add_row().cells
        row_cells[0].text = item["reference"] or "-"
        row_cells[1].text = item["article"]
        row_cells[2].text = str(item["quantity"])
        row_cells[3].text = item["remarque"] or "-"

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# INITIALISATION DES ÉTATS DE SESSION (SESSION STATE)
# =========================================================
if "current_user" not in st.session_state:
    st.session_state.current_user = None
if "selected_client" not in st.session_state:
    st.session_state.selected_client = None
if "temp_be_items" not in st.session_state:
    st.session_state.temp_be_items = []
if "temp_bs_items" not in st.session_state:
    st.session_state.temp_bs_items = []

# =========================================================
# ÉCRAN DE CONNEXION (LOGIN)
# =========================================================
if not st.session_state.logged_in:
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.8, 1])
    with col:
        logo_path = "nomatis_logo.png"
        if os.path.exists(logo_path):
            st.image(logo_path, width=200)

        st.markdown(
            """
            <div class="glass-card" style="text-align: center;">
                <h1 style="color: #1E3A8A; margin-bottom: 0px; font-weight: 800;">Gestion Stock MW NOMATIS</h1>
                <p style="color: #10B981; font-size: 15px; font-weight:600;">Espace de Gestion et Suivi de Stock Telecom</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("login_form"):
            st.markdown("##### Connexion Sécurisée")
            username = st.text_input("Identifiant / Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("SE CONNECTER")

            if submit:
                user = user_info(username)
                if user and user["password"] == password:
                    execute(
                        "UPDATE users SET last_login=? WHERE username=?",
                        (
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            username,
                        ),
                    )
                    st.session_state.logged_in = True
                    st.session_state.current_user = username
                    st.rerun()
                else:
                    st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

CURRENT_USER = user_info(st.session_state.current_user)
ROLE = CURRENT_USER["role"]


# =========================================================
# SÉLECTION DU CLIENT ET ESPACE DE TRAVAIL
# =========================================================
if not st.session_state.selected_client:
    st.markdown(
        f"""
        <div class="main-header">
            <div class="main-title">{APP_TITLE}</div>
            <div class="subtitle">Sélectionnez le compte client pour accéder au stock dédié</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Affichage du profil utilisateur avec possibilité de modifier son propre compte
    with st.expander(
        f"👤 Profil connecté : {CURRENT_USER['fullname']} ({ROLE.upper()}) — Modifier mes identifiants"
    ):
        with st.form("user_self_edit"):
            new_fullname = st.text_input(
                "Nom complet", value=CURRENT_USER["fullname"]
            )
            new_pwd = st.text_input(
                "Nouveau mot de passe",
                value=CURRENT_USER["password"],
                type="password",
            )
            if st.form_submit_button("Mettre à jour mon compte"):
                execute(
                    "UPDATE users SET fullname=?, password=? WHERE username=?",
                    (new_fullname, new_pwd, CURRENT_USER["username"]),
                )
                st.success("Compte mis à jour avec succès !")
                st.rerun()

    cols = st.columns(3)
    for idx, (client, info) in enumerate(CLIENTS.items()):
        with cols[idx]:
            st.markdown(
                f'<div class="glass-card" style="border-top: 4px solid {info["color"]}; text-align: center;">',
                unsafe_allow_html=True,
            )
            logo = normalized_logo(info["logo"])
            if logo:
                st.image(logo, use_container_width=True)
            else:
                st.markdown(
                    f"<h2 style='color:{info['color']}'>{client}</h2>",
                    unsafe_allow_html=True,
                )

            st.markdown(
                f"<h3 style='margin-top: 10px;'>{client}</h3>",
                unsafe_allow_html=True,
            )

            # Bouton d'accès personnalisé en gras avec la couleur du client
            btn_style = f"""
            <style>
            div[data-testid="stButton"] > button[key="select_{client}"] {{
                background-color: {info["color"]} !important;
                color: #FFFFFF !important;
                font-weight: 800 !important;
            }}
            </style>
            """
            st.markdown(btn_style, unsafe_allow_html=True)

            if st.button(
                f"Accès au Stock {client}",
                key=f"select_{client}",
                use_container_width=True,
            ):
                st.session_state.selected_client = client
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

CLIENT = st.session_state.selected_client


# =========================================================
# APPLICATION PRINCIPALE ET BARRE D'EN-TÊTE
# =========================================================
h1, h2 = st.columns([3, 1])
with h1:
    st.markdown(
        f"""
        <div class="main-header">
            <div class="main-title">{APP_TITLE} — ESPACE {CLIENT.upper()}</div>
            <div class="subtitle">Utilisateur : {CURRENT_USER['fullname']} ({ROLE.upper()})</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with h2:
    if st.button("🔄 Changer Client", use_container_width=True):
        st.session_state.selected_client = None
        st.rerun()
    if st.button("🚪 Déconnexion", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.selected_client = None
        st.rerun()


# Definition des 5 rubriques requises
tabs_labels = [
    "📥 Bon d'Entrée (BE)",
    "📤 Bon de Sortie (BS)",
    "📊 Situation Stock",
    "📜 Historique",
    "⚙️ Configuration",
]
tabs = st.tabs(tabs_labels)


# =========================================================
# RUBRIQUE 1 : BON D'ENTRÉE (BE)
# =========================================================
with tabs[0]:
    if ROLE in ["admin", "magasinier"]:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Création Bon d'Entrée (BE)")

        # Génération automatique des informations
        auto_date_saisie = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        auto_num_be = generate_bon_number("BE", CLIENT)

        st.info(
            f"🕒 **Date & Heure de saisie (Automatique) :** {auto_date_saisie} | 🔢 **N° BE Généré :** `{auto_num_be}`"
        )

        c1, c2 = st.columns(2)
        date_be = c1.date_input(
            "Date du Bon d'Entrée*",
            value=date.today(),
            max_value=date.today(),
            key="be_date_input",
        )

        f_list = active_names("fournisseurs")
        fourn_option = c2.selectbox(
            "Fournisseur*", f_list + ["+ Ajouter un autre fournisseur"]
        )
        if fourn_option == "+ Ajouter un autre fournisseur":
            fournisseur_final = st.text_input("Nom du nouveau fournisseur*")
        else:
            fournisseur_final = fourn_option

        lieu_livraison = st.text_input(
            "Lieu de Livraison*", value="Magasin Principal NOMATIS"
        )
        st.text_input(
            "Réceptionné par (Automatique)",
            value=CURRENT_USER["fullname"],
            disabled=True,
        )

        st.markdown("---")
        st.markdown("##### Saisie des articles pour le BE")

        art_list = active_names("articles")
        r1, r2, r3, r4 = st.columns([3, 1.5, 2, 3])
        sel_article = r1.selectbox("Article (Sélection obligatoire)*", art_list)
        ref_article = r2.text_input("Référence (Optionnel)", key="be_ref")
        qty_article = r3.number_input(
            "Quantité*", min_value=1, value=1, step=1, key="be_qty"
        )
        rem_article = r4.text_input("Remarque", key="be_rem")

        if st.button("➕ Ajouter l'article au tableau", key="add_be_line"):
            if not sel_article:
                st.error("Veuillez sélectionner un article valide.")
            elif qty_article <= 0:
                st.error("La quantité doit être strictement supérieure à 0.")
            else:
                # Regroupement et somme automatique des quantités si même article
                found = False
                for item in st.session_state.temp_be_items:
                    if item["article"] == sel_article:
                        item["quantity"] += qty_article
                        if rem_article:
                            item["remarque"] += f" | {rem_article}"
                        found = True
                        break
                if not found:
                    st.session_state.temp_be_items.append({
                        "article": sel_article,
                        "reference": ref_article,
                        "quantity": qty_article,
                        "remarque": rem_article,
                    })
                st.success(f"Article {sel_article} ajouté.")
                st.rerun()

        # Affichage du petit tableau au-dessous
        if st.session_state.temp_be_items:
            st.markdown("###### Articles ajoutés au Bon d'Entrée :")
            df_be_temp = pd.DataFrame(st.session_state.temp_be_items)
            st.dataframe(df_be_temp, use_container_width=True)

            b1, b2 = st.columns(2)
            if b1.button(
                "💾 Valider et Enregistrer le BE",
                key="btn_save_be",
                use_container_width=True,
            ):
                if not fournisseur_final.strip() or not lieu_livraison.strip():
                    st.error("Tous les champs obligatoires doivent être remplis !")
                else:
                    st.session_state.confirm_be_save = True

            if st.session_state.get("confirm_be_save", False):
                st.warning(
                    "⚠️ Confirmez-vous l'enregistrement définitif de ce Bon d'Entrée ?"
                )
                cb1, cb2 = st.columns(2)
                if cb1.button("✅ Oui, Confirmer l'enregistrement (BE)"):
                    # Si c'est un nouveau fournisseur ajouté à la volée, on l'enregistre dans la base
                    if (
                        fourn_option == "+ Ajouter un autre fournisseur"
                        and fournisseur_final.strip()
                    ):
                        execute(
                            "INSERT OR IGNORE INTO fournisseurs(name,active) VALUES(?,1)",
                            (fournisseur_final.strip(),),
                        )

                    bon_id = execute(
                        """
                        INSERT INTO bons (type,number,client,date_bon,datetime_saisie,fournisseur,lieu_livraison,receptionne_par,created_by)
                        VALUES (?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            "BE",
                            auto_num_be,
                            CLIENT,
                            str(date_be),
                            auto_date_saisie,
                            fournisseur_final,
                            lieu_livraison,
                            CURRENT_USER["fullname"],
                            CURRENT_USER["username"],
                        ),
                    )

                    for item in st.session_state.temp_be_items:
                        art_id = article_id_by_name(item["article"])
                        execute(
                            "INSERT INTO bon_items (bon_id,article_id,reference,quantity,remarque) VALUES (?,?,?,?,?)",
                            (
                                bon_id,
                                art_id,
                                item["reference"],
                                item["quantity"],
                                item["remarque"],
                            ),
                        )
                        set_stock(
                            CLIENT,
                            art_id,
                            current_stock(CLIENT, art_id) + item["quantity"],
                        )
                        add_movement(
                            CLIENT,
                            art_id,
                            "BE",
                            item["quantity"],
                            auto_num_be,
                            CURRENT_USER["username"],
                            item["remarque"],
                            fournisseur=fournisseur_final,
                        )

                    st.session_state.temp_be_items = []
                    st.session_state.confirm_be_save = False
                    st.session_state.last_created_bon_id = bon_id
                    st.success("Bon d'Entrée enregistré avec succès !")
                    st.rerun()

                if cb2.button("❌ Annuler"):
                    st.session_state.confirm_be_save = False
                    st.rerun()

            if b2.button(
                "🗑️ Vider le tableau BE", use_container_width=True
            ):
                st.session_state.temp_be_items = []
                st.rerun()

        if "last_created_bon_id" in st.session_state:
            last_id = st.session_state.last_created_bon_id
            st.markdown("---")
            st.markdown("##### Impression du Bon d'Entrée Généré")
            p1, p2 = st.columns(2)
            pdf_bytes = generate_pdf(last_id)
            docx_bytes = generate_docx(last_id)
            p1.download_button(
                "📄 Imprimer / Télécharger en PDF",
                pdf_bytes,
                file_name=f"BE_{last_id}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            p2.download_button(
                "📝 Télécharger au format Word (.docx)",
                docx_bytes,
                file_name=f"BE_{last_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.warning(
            "Seuls l'Admin et le Magasinier ont le droit de créer des Bons d'Entrée."
        )


# =========================================================
# RUBRIQUE 2 : BON DE SORTIE (BS)
# =========================================================
with tabs[1]:
    if ROLE in ["admin", "magasinier"]:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Création Bon de Sortie (BS)")

        auto_date_saisie_bs = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        auto_num_bs = generate_bon_number("BS", CLIENT)

        st.info(
            f"🕒 **Date & Heure de saisie (Automatique) :** {auto_date_saisie_bs} | 🔢 **N° BS Généré :** `{auto_num_bs}`"
        )

        c1, c2, c3 = st.columns(3)
        date_bs = c1.date_input(
            "Date du Bon de Sortie*",
            value=date.today(),
            max_value=date.today(),
            key="bs_date_input",
        )
        equipe_bs = c2.selectbox(
            "Équipe de Projet Destinataire*", active_names("equipes")
        )
        resource_bs = c3.selectbox(
            "Ressource / Technicien*", active_names("resources")
        )
        destination_bs = st.text_input(
            "Destination / Site Telecom*", value="Site Telecom"
        )

        st.markdown("---")
        st.markdown("##### Sélection et sortie de matériel")

        art_list = active_names("articles")
        r1, r2, r3, r4 = st.columns([3, 1.5, 2, 3])
        sel_article_bs = r1.selectbox(
            "Article*", art_list, key="bs_art_select"
        )
        ref_article_bs = r2.text_input("Référence", key="bs_ref")
        qty_article_bs = r3.number_input(
            "Quantité*", min_value=1, value=1, step=1, key="bs_qty"
        )
        rem_article_bs = r4.text_input("Remarque", key="bs_rem")

        if st.button("➕ Ajouter au Bon de Sortie", key="add_bs_line"):
            art_id = article_id_by_name(sel_article_bs)
            stk_dispo = current_stock(CLIENT, art_id)

            # Calcul de la quantité déjà réservée dans le tableau temporaire
            qty_already_in_temp = sum(
                item["quantity"]
                for item in st.session_state.temp_bs_items
                if item["article"] == sel_article_bs
            )
            if (qty_article_bs + qty_already_in_temp) > stk_dispo:
                st.error(
                    f"Stock insuffisant ! Disponible en stock : {stk_dispo} (Déjà dans le bon : {qty_already_in_temp})"
                )
            else:
                found = False
                for item in st.session_state.temp_bs_items:
                    if item["article"] == sel_article_bs:
                        item["quantity"] += qty_article_bs
                        if rem_article_bs:
                            item["remarque"] += f" | {rem_article_bs}"
                        found = True
                        break
                if not found:
                    st.session_state.temp_bs_items.append({
                        "article": sel_article_bs,
                        "reference": ref_article_bs,
                        "quantity": qty_article_bs,
                        "remarque": rem_article_bs,
                    })
                st.success(f"Article {sel_article_bs} ajouté au BS.")
                st.rerun()

        # Affichage du tableau de prévisualisation
        if st.session_state.temp_bs_items:
            st.markdown("###### Articles ajoutés au Bon de Sortie :")
            st.dataframe(
                pd.DataFrame(st.session_state.temp_bs_items),
                use_container_width=True,
            )

            b1, b2 = st.columns(2)
            if b1.button(
                "💾 Valider et Enregistrer le BS",
                key="btn_save_bs",
                use_container_width=True,
            ):
                if not destination_bs.strip():
                    st.error("Tous les champs obligatoires doivent être remplis !")
                else:
                    st.session_state.confirm_bs_save = True

            if st.session_state.get("confirm_bs_save", False):
                st.warning(
                    "⚠️ Confirmez-vous l'enregistrement définitif de ce Bon de Sortie ?"
                )
                cb1, cb2 = st.columns(2)
                if cb1.button("✅ Oui, Confirmer l'enregistrement (BS)"):
                    bon_id = execute(
                        """
                        INSERT INTO bons (type,number,client,date_bon,datetime_saisie,equipe,resource,destination,created_by)
                        VALUES (?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            "BS",
                            auto_num_bs,
                            CLIENT,
                            str(date_bs),
                            auto_date_saisie_bs,
                            equipe_bs,
                            resource_bs,
                            destination_bs,
                            CURRENT_USER["username"],
                        ),
                    )

                    for item in st.session_state.temp_bs_items:
                        art_id = article_id_by_name(item["article"])
                        execute(
                            "INSERT INTO bon_items (bon_id,article_id,reference,quantity,remarque) VALUES (?,?,?,?,?)",
                            (
                                bon_id,
                                art_id,
                                item["reference"],
                                item["quantity"],
                                item["remarque"],
                            ),
                        )
                        set_stock(
                            CLIENT,
                            art_id,
                            current_stock(CLIENT, art_id) - item["quantity"],
                        )
                        add_movement(
                            CLIENT,
                            art_id,
                            "BS",
                            item["quantity"],
                            auto_num_bs,
                            CURRENT_USER["username"],
                            item["remarque"],
                            equipe=equipe_bs,
                        )

                    st.session_state.temp_bs_items = []
                    st.session_state.confirm_bs_save = False
                    st.session_state.last_created_bs_id = bon_id
                    st.success("Bon de Sortie enregistré avec succès !")
                    st.rerun()

                if cb2.button("❌ Annuler (BS)"):
                    st.session_state.confirm_bs_save = False
                    st.rerun()

            if b2.button(
                "🗑️ Vider le tableau BS", use_container_width=True
            ):
                st.session_state.temp_bs_items = []
                st.rerun()

        if "last_created_bs_id" in st.session_state:
            last_bs_id = st.session_state.last_created_bs_id
            st.markdown("---")
            st.markdown("##### Impression du Bon de Sortie Généré")
            p1, p2 = st.columns(2)
            pdf_bytes = generate_pdf(last_bs_id)
            docx_bytes = generate_docx(last_bs_id)
            p1.download_button(
                "📄 Imprimer / Télécharger en PDF",
                pdf_bytes,
                file_name=f"BS_{last_bs_id}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            p2.download_button(
                "📝 Télécharger au format Word (.docx)",
                docx_bytes,
                file_name=f"BS_{last_bs_id}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.warning(
            "Seuls l'Admin et le Magasinier ont le droit de créer des Bons de Sortie."
        )


# =========================================================
# RUBRIQUE 3 : SITUATION STOCK
# =========================================================
with tabs[2]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.subheader(f"Situation du Stock en Temps Réel — Client {CLIENT}")

    rows = query(
        """
        SELECT 
            a.id AS article_id,
            a.name AS Article,
            COALESCE((SELECT SUM(quantity) FROM movements WHERE client=? AND article_id=a.id AND movement_type IN ('BE', 'AJUSTEMENT_PLUS')), 0) AS Total_Entrees,
            COALESCE((SELECT SUM(quantity) FROM movements WHERE client=? AND article_id=a.id AND movement_type IN ('BS', 'AJUSTEMENT_MOINS')), 0) AS Total_Sorties,
            COALESCE(s.quantity, 0) AS Stock_Actuel
        FROM articles a
        LEFT JOIN stock s ON s.article_id = a.id AND s.client = ?
        WHERE a.active = 1
        ORDER BY a.name
        """,
        (CLIENT, CLIENT, CLIENT),
    )

    df_stock = pd.DataFrame([dict(r) for r in rows])

    if not df_stock.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Nombre d'Articles", len(df_stock))
        c2.metric("Total Entrées Historiques", int(df_stock["Total_Entrees"].sum()))
        c3.metric("Stock Physique Actuel", int(df_stock["Stock_Actuel"].sum()))

        st.dataframe(
            df_stock[
                [
                    "Article",
                    "Total_Entrees",
                    "Total_Sorties",
                    "Stock_Actuel",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

        # Possibilité d'imprimer l'état de stock pour TOUS les utilisateurs
        st.markdown("---")
        csv_data = df_stock.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Exporter / Imprimer l'État du Stock (CSV)",
            data=csv_data,
            file_name=f"Situation_Stock_{CLIENT}_{date.today()}.csv",
            mime="text/csv",
        )
    else:
        st.info("Aucun article configuré dans le référentiel.")
    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# RUBRIQUE 4 : HISTORIQUE
# =========================================================
with tabs[3]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.subheader("Historique des Bons, Modification & Consultation des Mouvements")

    h_tab1, h_tab2 = st.tabs([
        "📄 Gestion des Bons (BE & BS)",
        "📜 Historique Détaillé des Mouvements",
    ])

    with h_tab1:
        type_filter = st.radio(
            "Sélectionner le type de Bon :",
            ["Bons d'Entrée (BE)", "Bons de Sortie (BS)"],
            horizontal=True,
        )
        target_type = "BE" if "Entrée" in type_filter else "BS"

        bons = query(
            "SELECT * FROM bons WHERE client=? AND type=? ORDER BY id DESC",
            (CLIENT, target_type),
        )
        if bons:
            opts = [
                f"{b['number']} | Date: {b['date_bon']} | Créé par: {b['created_by']} | ID:{b['id']}"
                for b in bons
            ]
            selected_bon_str = st.selectbox(
                f"Sélectionner un Bon {target_type}", opts
            )
            selected_bon_id = int(selected_bon_str.split("ID:")[1])

            bon_detail = query(
                "SELECT * FROM bons WHERE id=?", (selected_bon_id,), one=True
            )
            items_detail = query(
                "SELECT bi.*, a.name AS article FROM bon_items bi JOIN articles a ON a.id=bi.article_id WHERE bi.bon_id=?",
                (selected_bon_id,),
            )

            st.markdown("###### Articles contenus dans ce bon :")
            st.dataframe(
                pd.DataFrame([dict(i) for i in items_detail])[
                    ["reference", "article", "quantity", "remarque"]
                ],
                use_container_width=True,
            )

            # Option d'impression disponible pour tous
            col_print1, col_print2 = st.columns(2)
            pdf_b = generate_pdf(selected_bon_id)
            docx_b = generate_docx(selected_bon_id)
            col_print1.download_button(
                "📄 Imprimer / Télécharger en PDF",
                pdf_b,
                file_name=f"{bon_detail['number']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            col_print2.download_button(
                "📝 Télécharger en Word (.docx)",
                docx_b,
                file_name=f"{bon_detail['number']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

            # Options de modification / suppression réservées au Magasinier et Admin
            if ROLE in ["admin", "magasinier"]:
                st.markdown("---")
                st.markdown("##### Options d'Édition du Bon")

                with st.expander("✏️ Modifier les informations du Bon"):
                    with st.form("form_edit_bon"):
                        mod_date = st.date_input(
                            "Nouvelle Date",
                            value=datetime.strptime(
                                bon_detail["date_bon"], "%Y-%m-%d"
                            ).date(),
                            max_value=date.today(),
                        )
                        if target_type == "BE":
                            mod_fourn = st.selectbox(
                                "Fournisseur",
                                active_names("fournisseurs"),
                                index=0,
                            )
                            mod_lieu = st.text_input(
                                "Lieu Livraison",
                                value=bon_detail["lieu_livraison"] or "",
                            )
                        else:
                            mod_eq = st.selectbox(
                                "Équipe", active_names("equipes"), index=0
                            )
                            mod_dest = st.text_input(
                                "Destination",
                                value=bon_detail["destination"] or "",
                            )

                        if st.form_submit_button("Valider les modifications"):
                            st.session_state.confirm_edit_bon = True

                    if st.session_state.get("confirm_edit_bon", False):
                        st.warning(
                            "Confirmez-vous la modification des informations de ce bon ?"
                        )
                        if st.button("✅ Oui, Confirmer Modification"):
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
                            st.session_state.confirm_edit_bon = False
                            st.success("Bon mis à jour !")
                            st.rerun()

                if st.button(
                    "🚨 Supprimer ce bon (Restaure le stock)",
                    use_container_width=True,
                ):
                    st.session_state.confirm_delete_bon = True

                if st.session_state.get("confirm_delete_bon", False):
                    st.error(
                        "⚠️ ATTENTION : La suppression de ce bon va annuler tous les mouvements de stock associés !"
                    )
                    c_del1, c_del2 = st.columns(2)
                    if c_del1.button("✅ Oui, Supprimer Définitivement"):
                        for it in items_detail:
                            curr_s = current_stock(CLIENT, it["article_id"])
                            if target_type == "BE":
                                set_stock(
                                    CLIENT,
                                    it["article_id"],
                                    curr_s - it["quantity"],
                                )
                            else:
                                set_stock(
                                    CLIENT,
                                    it["article_id"],
                                    curr_s + it["quantity"],
                                )

                        execute(
                            "DELETE FROM bon_items WHERE bon_id=?",
                            (selected_bon_id,),
                        )
                        execute(
                            "DELETE FROM bons WHERE id=?", (selected_bon_id,)
                        )
                        execute(
                            "DELETE FROM movements WHERE reference_bon=?",
                            (bon_detail["number"],),
                        )

                        st.session_state.confirm_delete_bon = False
                        st.success("Bon supprimé et stock réajusté.")
                        st.rerun()
                    if c_del2.button("Annuler"):
                        st.session_state.confirm_delete_bon = False
                        st.rerun()
        else:
            st.info(f"Aucun Bon de type {target_type} enregistré.")

    with h_tab2:
        st.markdown("##### Recherche et Filtres d'Historique")
        col_f1, col_f2, col_f3 = st.columns(3)
        f_art = col_f1.selectbox(
            "Filtrer par Article", ["Tous"] + active_names("articles")
        )
        f_fourn = col_f2.selectbox(
            "Filtrer par Fournisseur", ["Tous"] + active_names("fournisseurs")
        )
        f_eq = col_f3.selectbox(
            "Filtrer par Équipe", ["Tous"] + active_names("equipes")
        )

        query_sql = """
            SELECT m.created_at AS Date_Mouvement, m.movement_type AS Type, m.reference_bon AS N_Bon,
                   a.name AS Article, m.quantity AS Quantite, m.fournisseur AS Fournisseur,
                   m.equipe AS Equipe, m.username AS Operateur, m.comment AS Remarque
            FROM movements m
            JOIN articles a ON a.id = m.article_id
            WHERE m.client = ?
        """
        params = [CLIENT]

        if f_art != "Tous":
            query_sql += " AND a.name = ?"
            params.append(f_art)
        if f_fourn != "Tous":
            query_sql += " AND m.fournisseur = ?"
            params.append(f_fourn)
        if f_eq != "Tous":
            query_sql += " AND m.equipe = ?"
            params.append(f_eq)

        query_sql += " ORDER BY m.id DESC"

        movs = query(query_sql, tuple(params))
        if movs:
            st.dataframe(
                pd.DataFrame([dict(m) for m in movs]), use_container_width=True
            )
        else:
            st.info("Aucun mouvement ne correspond aux filtres.")

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# RUBRIQUE 5 : CONFIGURATION (ADMIN SEULEMENT ET AJUSTEMENTS)
# =========================================================
with tabs[4]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)

    if ROLE == "admin":
        st.subheader("⚙️ Configuration Système & Administration")

        cfg_tab1, cfg_tab2, cfg_tab3, cfg_tab4 = st.tabs([
            "📦 Articles & Stocks",
            "🏢 Fournisseurs & Équipes",
            "👥 Gestion Utilisateurs",
            "🛠️ Ajustement Manuel Stock",
        ])

        with cfg_tab1:
            st.markdown("##### Référentiel des Articles")
            c1, c2 = st.columns([3, 1])
            new_art_name = c1.text_input(
                "Nom de l'article (ex: Câble IF, Support 0.3 m)"
            )
            init_stock_qty = c2.number_input(
                "Stock Initial", min_value=0, value=0
            )

            if st.button("➕ Ajouter l'Article"):
                if new_art_name.strip():
                    st.session_state.confirm_add_art = True

            if st.session_state.get("confirm_add_art", False):
                st.warning(
                    f"Confirmez-vous l'ajout de l'article '{new_art_name}' ?"
                )
                if st.button("✅ Confirmer Ajout Article"):
                    art_id = execute(
                        "INSERT OR IGNORE INTO articles(name,active) VALUES(?,1)",
                        (new_art_name.strip(),),
                    )
                    if init_stock_qty > 0 and art_id:
                        set_stock(CLIENT, art_id, init_stock_qty)
                        add_movement(
                            CLIENT,
                            art_id,
                            "AJUSTEMENT_PLUS",
                            init_stock_qty,
                            "INITIAL",
                            CURRENT_USER["username"],
                            "Stock Initial",
                        )
                    st.session_state.confirm_add_art = False
                    st.success("Article ajouté avec succès !")
                    st.rerun()

            st.markdown("---")
            st.markdown("##### Liste des Articles Existants")
            arts = query("SELECT * FROM articles WHERE active=1")
            for a in arts:
                col_a1, col_a2 = st.columns([4, 1])
                col_a1.write(f"• **{a['name']}**")
                if col_a2.button("Desactiver", key=f"del_art_{a['id']}"):
                    execute(
                        "UPDATE articles SET active=0 WHERE id=?", (a["id"],)
                    )
                    st.success("Article désactivé.")
                    st.rerun()

        with cfg_tab2:
            st.markdown("##### Fournisseurs et Équipes")
            cf1, cf2 = st.columns(2)

            with cf1:
                st.markdown("###### Fournisseurs")
                new_f = st.text_input("Nouveau Fournisseur")
                if st.button("Ajouter Fournisseur"):
                    if new_f.strip():
                        execute(
                            "INSERT OR IGNORE INTO fournisseurs(name,active) VALUES(?,1)",
                            (new_f.strip(),),
                        )
                        st.success("Fournisseur ajouté !")
                        st.rerun()

            with cf2:
                st.markdown("###### Équipes Projet")
                new_eq_input = st.text_input("Nouvelle Équipe")
                if st.button("Ajouter Équipe"):
                    if new_eq_input.strip():
                        execute(
                            "INSERT OR IGNORE INTO equipes(name,active) VALUES(?,1)",
                            (new_eq_input.strip(),),
                        )
                        st.success("Équipe ajoutée !")
                        st.rerun()

        with cfg_tab3:
            st.markdown("##### Gestion des Utilisateurs")
            st.markdown("###### Créer un Utilisateur")
            with st.form("form_create_user"):
                u_username = st.text_input("Nom d'utilisateur*")
                u_fullname = st.text_input("Nom Complet*")
                u_password = st.text_input("Mot de passe*", type="password")
                u_role = st.selectbox("Rôle*", ROLES)
                if st.form_submit_button("Créer l'utilisateur"):
                    st.session_state.confirm_create_u = True

            if st.session_state.get("confirm_create_u", False):
                st.warning(
                    f"Confirmez-vous la création de l'utilisateur {u_username} ?"
                )
                if st.button("✅ Confirmer Création User"):
                    execute(
                        "INSERT INTO users VALUES (?,?,?,?,?)",
                        (
                            u_username,
                            u_password,
                            u_fullname,
                            u_role,
                            "Jamais",
                        ),
                    )
                    st.session_state.confirm_create_u = False
                    st.success("Utilisateur créé !")
                    st.rerun()

            st.markdown("---")
            st.markdown("###### Liste des Utilisateurs")
            all_users = query("SELECT * FROM users")
            df_u = pd.DataFrame([dict(u) for u in all_users])[
                ["username", "fullname", "role", "last_login"]
            ]
            st.dataframe(df_u, use_container_width=True)

        with cfg_tab4:
            st.markdown("##### Ajustement Manuel du Stock")
            sel_art_adj = st.selectbox(
                "Article à ajuster",
                active_names("articles"),
                key="adj_art_select",
            )
            art_id_adj = article_id_by_name(sel_art_adj)
            curr_qty_adj = current_stock(CLIENT, art_id_adj)

            st.write(f"Stock Actuel : **{curr_qty_adj}**")
            new_qty_adj = st.number_input(
                "Nouveau Stock Désiré", min_value=0, value=curr_qty_adj
            )
            adj_comment = st.text_input(
                "Motif de l'ajustement*", value="Inventaire physique"
            )

            if st.button("💾 Appliquer l'Ajustement"):
                st.session_state.confirm_adj = True

            if st.session_state.get("confirm_adj", False):
                st.warning("Confirmez-vous la modification manuelle du stock ?")
                if st.button("✅ Oui, Confirmer Ajustement"):
                    diff = new_qty_adj - curr_qty_adj
                    set_stock(CLIENT, art_id_adj, new_qty_adj)
                    m_type = "AJUSTEMENT_PLUS" if diff >= 0 else "AJUSTEMENT_MOINS"
                    add_movement(
                        CLIENT,
                        art_id_adj,
                        m_type,
                        abs(diff),
                        "MANUEL",
                        CURRENT_USER["username"],
                        adj_comment,
                    )
                    st.session_state.confirm_adj = False
                    st.success("Stock ajusté avec succès !")
                    st.rerun()
    else:
        st.warning(
            "🔒 La rubrique Configuration est strictement réservée à l'Administrateur."
        )

    st.markdown("</div>", unsafe_allow_html=True)
