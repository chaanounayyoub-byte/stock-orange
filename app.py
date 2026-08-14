from datetime import datetime, date
from io import BytesIO
import os
import sqlite3
import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
)
from docx import Document

# =========================================================
# CONFIGURATION & THÈME LIGHT SAAS MINIMALISTE (DESIGN DE RÉFÉRENCE)
# =========================================================
st.set_page_config(
    page_title="NOMATIS — MW Stock Engine",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_TITLE = "NOMATIS — MW Stock Engine"
DB_FILE = "stock_mw.db"

CLIENTS = {
    "Orange": {"logo": "Orange_logo.svg.webp", "color": "#FF6600"},
    "Inwi": {"logo": "Logo INWI.jpg", "color": "#A1006B"},
    "ZTE": {"logo": "Logo ZTE.jpg", "color": "#005BAC"},
}

ROLES = ["admin", "magasinier", "coordinateur", "coordinatrice"]
DEFAULT_ARTICLES = ["Câble IF", "Câble RJ45", "Support 0.3 m", "Support 0.6 m", "ODU 18GHz", "Antenne 0.6m"]
DEFAULT_FOURNISSEURS = ["NEC", "ZTE", "Intégral", "FO Connect"]
DEFAULT_EQUIPES = ["Nabil Team", "Yassine Team", "Issam Team"]
DEFAULT_RESOURCES = ["Nabil", "Yassine", "Issam"]

PERMISSIONS = {
    "admin": {"be", "bs", "stock", "edit", "config", "import"},
    "magasinier": {"be", "bs", "stock", "edit", "import"},
    "coordinateur": {"stock", "print"},
    "coordinatrice": {"stock", "print"},
}

# Injection CSS — Style SaaS Lumineux & Épuré
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background-color: #F8FAFC !important;
            color: #0F172A !important;
        }

        .stApp {
            background-color: #F8FAFC;
        }

        /* En-tête principal */
        .main-header {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 20px 28px;
            margin-bottom: 24px;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        }

        .main-title {
            color: #0F172A !important;
            font-weight: 700;
            font-size: 24px;
            margin: 0;
        }

        .subtitle {
            color: #64748B !important;
            font-size: 14px;
            margin-top: 4px;
        }

        /* Cartes Blanches Épurées */
        .glass-card {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05), 0 1px 2px 0 rgba(0, 0, 0, 0.03);
        }

        /* Boutons (Bleu Indigo Moderne) */
        .stButton > button, div[data-testid="stFormSubmitButton"] > button {
            background-color: #2563EB !important;
            color: #FFFFFF !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            border: none !important;
            padding: 10px 20px !important;
            transition: all 0.2s ease !important;
            min-height: 42px !important;
        }

        .stButton > button:hover, div[data-testid="stFormSubmitButton"] > button:hover {
            background-color: #1D4ED8 !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2) !important;
        }

        /* Champs de Saisie (Inputs) */
        .stTextInput input, .stSelectbox div[data-baseweb="select"], .stNumberInput input, .stDateInput input {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 8px !important;
        }

        .stTextInput input:focus, .stNumberInput input:focus {
            border-color: #2563EB !important;
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1) !important;
        }

        /* Onglets / Tabs */
        div[data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #F1F5F9;
            padding: 6px;
            border-radius: 10px;
            border: 1px solid #E2E8F0;
        }

        button[data-baseweb="tab"] {
            border-radius: 6px !important;
            font-weight: 500 !important;
            color: #64748B !important;
        }

        button[aria-selected="true"] {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
        }

        /* Métriques */
        [data-testid="stMetricValue"] {
            color: #2563EB !important;
            font-weight: 700 !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# BASE DE DONNÉES (SQLite)
# =========================================================
def get_conn():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
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
            comment TEXT
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

    for name in DEFAULT_ARTICLES:
        cur.execute("INSERT OR IGNORE INTO articles(name,active) VALUES(?,1)", (name,))
    for name in DEFAULT_FOURNISSEURS:
        cur.execute("INSERT OR IGNORE INTO fournisseurs(name,active) VALUES(?,1)", (name,))
    for name in DEFAULT_EQUIPES:
        cur.execute("INSERT OR IGNORE INTO equipes(name,active) VALUES(?,1)", (name,))
    for name in DEFAULT_RESOURCES:
        cur.execute("INSERT OR IGNORE INTO resources(name,active) VALUES(?,1)", (name,))

    conn.commit()
    conn.close()


init_db()


# =========================================================
# UTILITAIRES & FONCTIONS DE BASE
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
    row = query("SELECT quantity FROM stock WHERE client=? AND article_id=?", (client, article_id), one=True)
    return int(row["quantity"]) if row else 0


def set_stock(client, article_id, quantity):
    execute(
        """
        INSERT INTO stock(client,article_id,quantity) VALUES(?,?,?)
        ON CONFLICT(client,article_id) DO UPDATE SET quantity=excluded.quantity
        """,
        (client, article_id, int(quantity)),
    )


def add_movement(client, article_id, m_type, qty, ref_bon, user, comment=""):
    execute(
        """
        INSERT INTO movements(client,article_id,movement_type,quantity,reference_bon,username,created_at,comment)
        VALUES(?,?,?,?,?,?,?,?)
        """,
        (client, article_id, m_type, int(qty), ref_bon, user, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), comment),
    )


def user_info(username):
    return query("SELECT * FROM users WHERE username=?", (username,), one=True)


def can(role, permission):
    return permission in PERMISSIONS.get(role, set())


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
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15 * mm, leftMargin=15 * mm, topMargin=15 * mm, bottomMargin=15 * mm)
    styles = getSampleStyleSheet()
    story = []

    logo_info = CLIENTS.get(bon["client"], {})
    logo_path = logo_info.get("logo")
    
    header_data = []
    if logo_path and os.path.exists(logo_path):
        try:
            img_rl = RLImage(logo_path, width=40*mm, height=18*mm)
            header_data.append([img_rl, Paragraph(f"<font size=16 color='#0F172A'><b>NOMATIS — MW STOCK</b></font>", styles["Normal"])])
        except Exception:
            header_data.append([Paragraph(f"<font size=16 color='#0F172A'><b>NOMATIS</b></font>", styles["Normal"])])
    else:
        header_data.append([Paragraph(f"<font size=16 color='#0F172A'><b>NOMATIS</b></font>", styles["Normal"])])

    header_table = Table(header_data, colWidths=[50*mm, 130*mm])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(header_table)
    story.append(Spacer(1, 10))

    title = "BON D'ENTRÉE" if bon["type"] == "BE" else "BON DE SORTIE"
    story.append(Paragraph(f"<font size=18 color='#2563EB'><b>{title} — {bon['number']}</b></font>", styles["Title"]))
    story.append(Spacer(1, 10))

    if bon["type"] == "BE":
        info = [
            ["N° Bon", bon["number"], "Date Bon", bon["date_bon"]],
            ["Fournisseur", bon["fournisseur"] or "", "Lieu Livraison", bon["lieu_livraison"] or ""],
            ["Réceptionné par", bon["receptionne_par"] or "", "Client / Projet", bon["client"]],
        ]
    else:
        info = [
            ["N° Bon", bon["number"], "Date Bon", bon["date_bon"]],
            ["Équipe Destination", bon["equipe"] or "", "Ressource", bon["resource"] or ""],
            ["Destination / Site", bon["destination"] or "", "Client / Projet", bon["client"]],
        ]

    info_table = Table(info, colWidths=[35 * mm, 55 * mm, 35 * mm, 55 * mm])
    info_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F1F5F9")),
                ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#F1F5F9")),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(info_table)
    story.append(Spacer(1, 15))

    data = [["Référence", "Désignation Article", "Quantité", "Remarque"]]
    for item in items:
        data.append([item["reference"] or "-", item["article"], str(item["quantity"]), item["remarque"] or "-"])

    items_table = Table(data, colWidths=[35 * mm, 70 * mm, 25 * mm, 50 * mm])
    items_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2563EB")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(items_table)
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
    doc.add_heading(f"BON DE {'ENTRÉE' if bon['type']=='BE' else 'SORTIE'} - {bon['number']}", level=1)
    p = doc.add_paragraph()
    p.add_run(f"Client / Projet : {bon['client']}\nDate Bon : {bon['date_bon']}\nCréé par : {bon['created_by']}")

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Référence"
    hdr_cells[1].text = "Article"
    hdr_cells[2].text = "Quantité"
    hdr_cells[3].text = "Remarque"

    for item in items:
        row_cells = table.add_row().cells
        row_cells[0].text = item["reference"] or ""
        row_cells[1].text = item["article"]
        row_cells[2].text = str(item["quantity"])
        row_cells[3].text = item["remarque"] or ""

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# ÉTAT SESSION & AUTHENTIFICATION
# =========================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "current_user" not in st.session_state:
    st.session_state.current_user = None
if "selected_client" not in st.session_state:
    st.session_state.selected_client = None
if "temp_be_items" not in st.session_state:
    st.session_state.temp_be_items = []
if "temp_bs_items" not in st.session_state:
    st.session_state.temp_bs_items = []


def login_screen():
    st.markdown("<br><br>", unsafe_allow_html=True)
    _, col, _ = st.columns([1, 1.8, 1])
    with col:
        st.markdown(
            """
            <div class="glass-card" style="text-align: center;">
                <h1 style="color: #2563EB; margin-bottom: 0px; font-weight: 800;">NOMATIS</h1>
                <p style="color: #64748B; font-size: 14px;">Plateforme de Gestion de Stock Microwave</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            st.markdown("##### Connexion Espace Sécurisé")
            username = st.text_input("Identifiant")
            password = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("SE CONNECTER", use_container_width=True)

            if submit:
                user = user_info(username)
                if user and user["password"] == password:
                    execute("UPDATE users SET last_login=? WHERE username=?", (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), username))
                    st.session_state.logged_in = True
                    st.session_state.current_user = username
                    st.rerun()
                else:
                    st.error("Identifiants invalides.")


if not st.session_state.logged_in:
    login_screen()
    st.stop()

CURRENT_USER = user_info(st.session_state.current_user)
ROLE = CURRENT_USER["role"]


# =========================================================
# SELECTION DU CLIENT / PROJET
# =========================================================
if not st.session_state.selected_client:
    st.markdown(
        """
        <div class="main-header">
            <div class="main-title">Espaces Clients</div>
            <div class="subtitle">Sélectionnez le compte client pour accéder au stock dédié</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    cols = st.columns(3)
    for idx, (client, info) in enumerate(CLIENTS.items()):
        with cols[idx]:
            st.markdown(f'<div class="glass-card" style="border-top: 4px solid {info["color"]};">', unsafe_allow_html=True)
            logo = normalized_logo(info["logo"])
            if logo:
                st.image(logo, use_container_width=True)
            else:
                st.markdown(f"<h2 style='text-align:center; color:{info['color']}'>{client}</h2>", unsafe_allow_html=True)

            st.markdown(f"<h3 style='text-align:center; margin-top: 10px;'>{client}</h3>", unsafe_allow_html=True)
            if st.button(f"Ouvrir Espace {client}", key=f"select_{client}", use_container_width=True):
                st.session_state.selected_client = client
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

CLIENT = st.session_state.selected_client


# =========================================================
# APPLICATION PRINCIPALE APRES LOGIN & SELECTION CLIENT
# =========================================================
h1, h2 = st.columns([3, 1])
with h1:
    st.markdown(
        f"""
        <div class="main-header">
            <div class="main-title">NOMATIS — {CLIENT}</div>
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

tabs_labels = ["📥 Bon d'Entrée (BE)", "📤 Bon de Sortie (BS)", "📊 État du Stock", "✏️ Modification / Impression"]
if can(ROLE, "import"):
    tabs_labels.append("📁 Import / Export Excel")
if can(ROLE, "config"):
    tabs_labels.append("⚙️ Configuration")

tabs = st.tabs(tabs_labels)

# ---------------------------------------------------------
# 📥 TAB 1: BON D'ENTRÉE (BE)
# ---------------------------------------------------------
with tabs[0]:
    if can(ROLE, "be"):
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Nouveau Bon d'Entrée (BE)")
        c1, c2, c3 = st.columns(3)
        date_be = c1.date_input("Date Bon", value=date.today(), key="be_d")
        num_be = c2.text_input("N° BE", value=f"BE-{date.today().strftime('%Y%m%d')}-01")
        fournisseur = c3.selectbox("Fournisseur", active_names("fournisseurs"))
        lieu = st.text_input("Lieu de Livraison", value="Magasin Principal NOMATIS")

        st.markdown("---")
        st.markdown("##### Ajouter des articles au BE")
        r1, r2, r3, r4 = st.columns([2, 3, 1.5, 3])
        ref = r1.text_input("Référence", key="be_ref")
        art = r2.selectbox("Article", active_names("articles"), key="be_art")
        qty = r3.number_input("Qté", min_value=1, value=1, key="be_qty")
        rem = r4.text_input("Remarque", key="be_rem")

        if st.button("➕ Ajouter la ligne", key="add_be_line"):
            st.session_state.temp_be_items.append({"ref": ref, "art": art, "qty": qty, "rem": rem})

        if st.session_state.temp_be_items:
            st.markdown("###### Articles à valider :")
            st.dataframe(pd.DataFrame(st.session_state.temp_be_items), use_container_width=True)
            
            b1, b2 = st.columns(2)
            if b1.button("💾 Valider et Enregistrer le BE", use_container_width=True):
                bon_id = execute(
                    "INSERT INTO bons (type,number,client,date_bon,datetime_saisie,fournisseur,lieu_livraison,receptionne_par,created_by) VALUES (?,?,?,?,?,?,?,?,?)",
                    ("BE", num_be, CLIENT, str(date_be), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), fournisseur, lieu, CURRENT_USER["fullname"], CURRENT_USER["username"]),
                )
                for item in st.session_state.temp_be_items:
                    art_id = article_id_by_name(item["art"])
                    execute("INSERT INTO bon_items (bon_id,article_id,reference,quantity,remarque) VALUES (?,?,?,?,?)", (bon_id, art_id, item["ref"], item["qty"], item["rem"]))
                    set_stock(CLIENT, art_id, current_stock(CLIENT, art_id) + item["qty"])
                    add_movement(CLIENT, art_id, "BE", item["qty"], num_be, CURRENT_USER["username"], item["rem"])
                st.session_state.temp_be_items = []
                st.success("Bon d'Entrée enregistré avec succès !")
                st.rerun()
            if b2.button("🗑️ Vider le tableau", use_container_width=True):
                st.session_state.temp_be_items = []
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.warning("Vous n'avez pas la permission de créer des bons d'entrée.")

# ---------------------------------------------------------
# 📤 TAB 2: BON DE SORTIE (BS)
# ---------------------------------------------------------
with tabs[1]:
    if can(ROLE, "bs"):
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Nouveau Bon de Sortie (BS)")
        c1, c2, c3 = st.columns(3)
        date_bs = c1.date_input("Date Bon", value=date.today(), key="bs_d")
        num_bs = c2.text_input("N° BS", value=f"BS-{date.today().strftime('%Y%m%d')}-01")
        equipe = c3.selectbox("Équipe Destination", active_names("equipes"))

        c4, c5 = st.columns(2)
        resource = c4.selectbox("Ressource / Technicien", active_names("resources"))
        destination = c5.text_input("Destination / Site", value="Site Telecom")

        st.markdown("---")
        st.markdown("##### Sélection matériel")
        r1, r2, r3, r4 = st.columns([2, 3, 1.5, 3])
        ref = r1.text_input("Référence", key="bs_ref")
        art = r2.selectbox("Article", active_names("articles"), key="bs_art")
        qty = r3.number_input("Qté", min_value=1, value=1, key="bs_qty")
        rem = r4.text_input("Remarque", key="bs_rem")

        if st.button("➕ Ajouter au BS", key="add_bs_line"):
            art_id = article_id_by_name(art)
            stk_dispo = current_stock(CLIENT, art_id)
            if qty > stk_dispo:
                st.error(f"Stock insuffisant ! Disponible : {stk_dispo}")
            else:
                st.session_state.temp_bs_items.append({"ref": ref, "art": art, "qty": qty, "rem": rem})

        if st.session_state.temp_bs_items:
            st.markdown("###### Articles à sortir :")
            st.dataframe(pd.DataFrame(st.session_state.temp_bs_items), use_container_width=True)

            b1, b2 = st.columns(2)
            if b1.button("💾 Valider et Enregistrer le BS", use_container_width=True):
                bon_id = execute(
                    "INSERT INTO bons (type,number,client,date_bon,datetime_saisie,equipe,resource,destination,created_by) VALUES (?,?,?,?,?,?,?,?,?)",
                    ("BS", num_bs, CLIENT, str(date_bs), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), equipe, resource, destination, CURRENT_USER["username"]),
                )
                for item in st.session_state.temp_bs_items:
                    art_id = article_id_by_name(item["art"])
                    execute("INSERT INTO bon_items (bon_id,article_id,reference,quantity,remarque) VALUES (?,?,?,?,?)", (bon_id, art_id, item["ref"], item["qty"], item["rem"]))
                    set_stock(CLIENT, art_id, current_stock(CLIENT, art_id) - item["qty"])
                    add_movement(CLIENT, art_id, "BS", item["qty"], num_bs, CURRENT_USER["username"], item["rem"])
                st.session_state.temp_bs_items = []
                st.success("Bon de Sortie enregistré avec succès !")
                st.rerun()
            if b2.button("🗑️ Vider la liste", use_container_width=True):
                st.session_state.temp_bs_items = []
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.warning("Vous n'avez pas la permission de créer des bons de sortie.")

# ---------------------------------------------------------
# 📊 TAB 3: ÉTAT DU STOCK & HISTORIQUE
# ---------------------------------------------------------
with tabs[2]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.subheader(f"État du Stock en Temps Réel — {CLIENT}")

    rows = query(
        """
        SELECT a.name AS Article, COALESCE(s.quantity, 0) AS Stock
        FROM articles a
        LEFT JOIN stock s ON s.article_id = a.id AND s.client = ?
        WHERE a.active = 1 ORDER BY a.name
        """,
        (CLIENT,),
    )
    df_stock = pd.DataFrame([dict(r) for r in rows])

    if not df_stock.empty:
        c1, c2 = st.columns(2)
        c1.metric("Articles Référencés", len(df_stock))
        c2.metric("Total Unités en Stock", int(df_stock["Stock"].sum()))
        st.dataframe(df_stock, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("Historique des Mouvements")
    movs = query(
        """
        SELECT m.created_at AS Date, m.movement_type AS Type, m.reference_bon AS Bon,
               a.name AS Article, m.quantity AS Quantité, m.username AS Opérateur, m.comment AS Remarque
        FROM movements m JOIN articles a ON a.id = m.article_id
        WHERE m.client = ? ORDER BY m.id DESC
        """,
        (CLIENT,),
    )
    if movs:
        st.dataframe(pd.DataFrame([dict(m) for m in movs]), use_container_width=True, hide_index=True)
    else:
        st.info("Aucun mouvement enregistré.")
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# ✏️ TAB 4: MODIFICATION & IMPRESSION DES BONS
# ---------------------------------------------------------
with tabs[3]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.subheader("Consultation, Impression & Edition des Bons")
    bons = query("SELECT * FROM bons WHERE client=? ORDER BY id DESC", (CLIENT,))
    if bons:
        opts = [f"{b['type']} - {b['number']} ({b['date_bon']}) - ID:{b['id']}" for b in bons]
        sel = st.selectbox("Sélectionner un Bon", opts)
        selected_id = int(sel.split("ID:")[1])

        c1, c2 = st.columns(2)
        with c1:
            pdf_b = generate_pdf(selected_id)
            st.download_button("📄 Télécharger en PDF", pdf_b, file_name=f"Bon_{selected_id}.pdf", mime="application/pdf", use_container_width=True)
        with c2:
            docx_b = generate_docx(selected_id)
            st.download_button("📝 Télécharger en Word (.docx)", docx_b, file_name=f"Bon_{selected_id}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

        if can(ROLE, "edit"):
            st.markdown("---")
            st.markdown("##### Actions de Modification / Suppression")
            b_data = query("SELECT * FROM bons WHERE id=?", (selected_id,), one=True)
            
            with st.expander("✏️ Modifier les détails du bon"):
                with st.form("edit_bon_form"):
                    new_date = st.date_input("Date Bon", value=datetime.strptime(b_data["date_bon"], "%Y-%m-%d").date())
                    if b_data["type"] == "BE":
                        new_fourn = st.selectbox("Fournisseur", active_names("fournisseurs"), index=0)
                        new_lieu = st.text_input("Lieu Livraison", value=b_data["lieu_livraison"] or "")
                        submit_edit = st.form_submit_button("Mettre à jour")
                        if submit_edit:
                            execute("UPDATE bons SET date_bon=?, fournisseur=?, lieu_livraison=? WHERE id=?", (str(new_date), new_fourn, new_lieu, selected_id))
                            st.success("Bon mis à jour !")
                            st.rerun()
                    else:
                        new_eq = st.selectbox("Équipe", active_names("equipes"), index=0)
                        new_dest = st.text_input("Destination", value=b_data["destination"] or "")
                        submit_edit = st.form_submit_button("Mettre à jour")
                        if submit_edit:
                            execute("UPDATE bons SET date_bon=?, equipe=?, destination=? WHERE id=?", (str(new_date), new_eq, new_dest, selected_id))
                            st.success("Bon mis à jour !")
                            st.rerun()

            if st.button("🚨 Supprimer ce bon (Restaure le stock)", use_container_width=True):
                # Restauration du stock avant suppression
                items_to_revert = query("SELECT article_id, quantity FROM bon_items WHERE bon_id=?", (selected_id,))
                for it in items_to_revert:
                    curr = current_stock(CLIENT, it["article_id"])
                    if b_data["type"] == "BE":
                        set_stock(CLIENT, it["article_id"], curr - it["quantity"])
                    else:
                        set_stock(CLIENT, it["article_id"], curr + it["quantity"])
                
                execute("DELETE FROM bon_items WHERE bon_id=?", (selected_id,))
                execute("DELETE FROM bons WHERE id=?", (selected_id,))
                st.success("Bon supprimé et stock réajusté !")
                st.rerun()
    else:
        st.info("Aucun bon enregistré pour le moment.")
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 📁 TAB 5: IMPORT / EXPORT EXCEL (SI AUTORISÉ)
# ---------------------------------------------------------
tab_idx = 4
if can(ROLE, "import"):
    with tabs[tab_idx]:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Importation & Exportation Excel")

        c_exp, c_imp = st.columns(2)

        with c_exp:
            st.markdown("##### 📤 Export de Données")
            # Export Stock
            rows_stk = query("SELECT a.name AS Article, COALESCE(s.quantity,0) AS Quantité FROM articles a LEFT JOIN stock s ON s.article_id=a.id AND s.client=?", (CLIENT,))
            df_stk_exp = pd.DataFrame([dict(r) for r in rows_stk])
            
            output_stk = BytesIO()
            with pd.ExcelWriter(output_stk, engine='openpyxl') as writer:
                df_stk_exp.to_excel(writer, index=False, sheet_name='Stock')
            
            st.download_button(
                "📊 Télécharger l'état du stock (.xlsx)",
                output_stk.getvalue(),
                file_name=f"Stock_{CLIENT}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.markdown("<br>", unsafe_allow_html=True)
            # Modèle d'importation Excel
            df_model = pd.DataFrame({
                "Référence": ["REF-001", "REF-002"],
                "Article": [DEFAULT_ARTICLES[0], DEFAULT_ARTICLES[1]],
                "Quantité": [10, 5],
                "Remarque": ["RAS", "Urgent"]
            })
            output_mod = BytesIO()
            with pd.ExcelWriter(output_mod, engine='openpyxl') as writer:
                df_model.to_excel(writer, index=False, sheet_name='Modele_Import')

            st.download_button(
                "📑 Télécharger un modèle d'import Excel",
                output_mod.getvalue(),
                file_name="Modele_Import_Bon.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with c_imp:
            st.markdown("##### 📥 Import massif de Bon via Excel")
            imp_type = st.radio("Type de bon à créer", ["BE (Entrée)", "BS (Sortie)"], horizontal=True)
            file_up = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])

            if file_up:
                try:
                    df_up = pd.read_excel(file_up)
                    st.write("Aperçu des données :", df_up.head())

                    if st.button("🚀 Valider l'importation de ce fichier"):
                        b_type = "BE" if "BE" in imp_type else "BS"
                        prefix = "BE-IMP" if b_type == "BE" else "BS-IMP"
                        num_imp = f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S')}"

                        bon_id = execute(
                            "INSERT INTO bons (type,number,client,date_bon,datetime_saisie,created_by) VALUES (?,?,?,?,?,?)",
                            (b_type, num_imp, CLIENT, str(date.today()), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), CURRENT_USER["username"])
                        )

                        for _, row in df_up.iterrows():
                            art_name = str(row["Article"]).strip()
                            qty_val = int(row["Quantité"])
                            ref_val = str(row.get("Référence", ""))
                            rem_val = str(row.get("Remarque", ""))

                            # Insertion article s'il n'existe pas
                            execute("INSERT OR IGNORE INTO articles (name, active) VALUES (?, 1)", (art_name,))
                            art_id = article_id_by_name(art_name)

                            execute("INSERT INTO bon_items (bon_id,article_id,reference,quantity,remarque) VALUES (?,?,?,?,?)", (bon_id, art_id, ref_val, qty_val, rem_val))

                            curr = current_stock(CLIENT, art_id)
                            new_qty = curr + qty_val if b_type == "BE" else max(0, curr - qty_val)
                            set_stock(CLIENT, art_id, new_qty)
                            add_movement(CLIENT, art_id, b_type, qty_val, num_imp, CURRENT_USER["username"], "Import Excel")

                        st.success("Importation effectuée avec succès !")
                        st.rerun()
                except Exception as e:
                    st.error(f"Erreur lors de la lecture du fichier : {e}")

        st.markdown("</div>", unsafe_allow_html=True)
    tab_idx += 1

# ---------------------------------------------------------
# ⚙️ TAB 6: CONFIGURATION (ADMIN)
# ---------------------------------------------------------
if can(ROLE, "config"):
    with tabs[tab_idx]:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Configuration Système")

        sub_tab1, sub_tab2 = st.tabs(["📋 Référentiels Stock", "👥 Gestion des Utilisateurs"])

        with sub_tab1:
            cat = st.radio("Entité à gérer", ["Articles", "Fournisseurs", "Équipes", "Ressources"], horizontal=True)
            t_map = {"Articles": "articles", "Fournisseurs": "fournisseurs", "Équipes": "equipes", "Ressources": "resources"}

            c1, c2 = st.columns([1, 2])
            with c1:
                new_val = st.text_input(f"Ajouter dans {cat}")
                if st.button("Enregistrer Entité"):
                    if new_val:
                        execute(f"INSERT OR IGNORE INTO {t_map[cat]} (name, active) VALUES (?, 1)", (new_val.strip(),))
                        st.success("Élément ajouté !")
                        st.rerun()
            with c2:
                items = query(f"SELECT id, name, active FROM {t_map[cat]} ORDER BY name")
                if items:
                    st.dataframe(pd.DataFrame([dict(i) for i in items]), use_container_width=True, hide_index=True)

        with sub_tab2:
            st.markdown("##### Utilisateurs enregistrés")
            u_col1, u_col2 = st.columns([1, 2])
            with u_col1:
                with st.form("add_user_form"):
                    st.markdown("###### Nouvel Utilisateur")
                    u_username = st.text_input("Identifiant")
                    u_pass = st.text_input("Mot de passe", type="password")
                    u_full = st.text_input("Nom Complet")
                    u_role = st.selectbox("Rôle", ROLES)
                    if st.form_submit_button("Créer Compte"):
                        if u_username and u_pass and u_full:
                            execute("INSERT OR REPLACE INTO users VALUES (?,?,?,?,?)", (u_username, u_pass, u_full, u_role, "Jamais"))
                            st.success("Utilisateur créé / mis à jour !")
                            st.rerun()
            with u_col2:
                users = query("SELECT username, fullname, role, last_login FROM users")
                st.dataframe(pd.DataFrame([dict(u) for u in users]), use_container_width=True, hide_index=True)

        st.markdown("</div>", unsafe_allow_html=True)
