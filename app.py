from datetime import datetime, date
from io import BytesIO
import os
import sqlite3
import base64
import html

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
)
from docx import Document
from docx.shared import Cm


# =========================================================
# CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="Gestion Stock MW NOMATIS",
    page_icon="📦",
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

DEFAULT_ARTICLES = ["câble IF", "câble RJ45", "support 0.3 m", "support 0.6 m"]
DEFAULT_FOURNISSEURS = ["NEC", "ZTE", "Intégral", "FO connect"]
DEFAULT_EQUIPES = ["Nabil Team", "Yassine Team", "Issa Team"]
DEFAULT_RESOURCES = ["Nabil", "Yassine", "Issam"]

# Permissions demandées
PERMISSIONS = {
    "admin": {"be", "bs", "stock", "edit", "config"},
    "magasinier": {"be", "bs", "stock", "edit"},
    "coordinateur": {"stock", "print"},
    "coordinatrice": {"stock", "print"},
}


# =========================================================
# STYLE BLEU / BLANC / UN PEU DE VERT
# =========================================================
st.markdown(
    """
    <style>
        .stApp {
            background: #F8FAFC;
            color: #0F172A;
        }

        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }

        h1, h2, h3, h4, h5, h6, label, p, span, div {
            color: #0F172A;
        }

        .main-title {
            color: #0B4EA2 !important;
            font-weight: 800;
            font-size: 30px;
            margin-top: 5px;
            margin-bottom: 3px;
        }

        .subtitle {
            color: #475569 !important;
            font-size: 14px;
        }

        .login-card {
            background: white;
            border: 1px solid #D7E2F0;
            border-radius: 16px;
            padding: 30px;
            box-shadow: 0 8px 25px rgba(15, 23, 42, 0.08);
        }

        .section-card {
            background: white;
            border: 1px solid #D7E2F0;
            border-radius: 12px;
            padding: 18px;
            margin-bottom: 12px;
        }

        .stButton > button {
            border-radius: 8px;
            font-weight: 700;
            min-height: 42px;
        }

        .login-btn > button {
            background: #DC2626 !important;
            color: white !important;
            border: 0 !important;
        }

        .login-btn > button:hover {
            background: #B91C1C !important;
        }

        .valid-btn > button {
            background: #16A34A !important;
            color: white !important;
            border: 0 !important;
        }

        .stock-btn > button {
            background: #2563EB !important;
            color: white !important;
            border: 2px solid #1D4ED8 !important;
            font-weight: 800 !important;
        }

        .stock-btn > button:hover {
            background: #1D4ED8 !important;
        }

        .metric-box {
            background: white;
            border: 1px solid #D7E2F0;
            border-radius: 10px;
            padding: 12px;
            text-align: center;
        }

        .small-note {
            color: #64748B;
            font-size: 12px;
        }

        div[data-baseweb="tab-list"] {
            gap: 8px;
        }

        button[data-baseweb="tab"] {
            font-weight: 700;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# DATABASE
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

    # Compte admin par défaut
    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users(username,password,fullname,role,last_login) VALUES(?,?,?,?,?)",
            ("admin", "admin123", "Administrateur", "admin", "Jamais"),
        )
        cur.execute(
            "INSERT INTO users(username,password,fullname,role,last_login) VALUES(?,?,?,?,?)",
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
# HELPERS
# =========================================================
def query(sql, params=(), one=False):
    conn = get_conn()
    cur = conn.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    if one:
        return rows[0] if rows else None
    return rows


def execute(sql, params=()):
    conn = get_conn()
    cur = conn.execute(sql, params)
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    return last_id


def execute_many(sql, data):
    conn = get_conn()
    conn.executemany(sql, data)
    conn.commit()
    conn.close()


def active_names(table):
    allowed = {"articles", "fournisseurs", "equipes", "resources"}
    if table not in allowed:
        raise ValueError("Table non autorisée")
    rows = query(f"SELECT name FROM {table} WHERE active=1 ORDER BY name")
    return [r["name"] for r in rows]


def article_rows():
    return query("SELECT id,name FROM articles WHERE active=1 ORDER BY name")


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
        INSERT INTO stock(client,article_id,quantity)
        VALUES(?,?,?)
        ON CONFLICT(client,article_id)
        DO UPDATE SET quantity=excluded.quantity
        """,
        (client, article_id, int(quantity)),
    )


def add_movement(client, article_id, movement_type, quantity, reference_bon, username, comment=""):
    execute(
        """
        INSERT INTO movements
        (client,article_id,movement_type,quantity,reference_bon,username,created_at,comment)
        VALUES(?,?,?,?,?,?,?,?)
        """,
        (
            client,
            article_id,
            movement_type,
            int(quantity),
            reference_bon,
            username,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            comment,
        ),
    )


def user_info(username):
    return query("SELECT * FROM users WHERE username=?", (username,), one=True)


def can(role, permission):
    return permission in PERMISSIONS.get(role, set())


def format_date(dt):
    return dt.strftime("%Y-%m-%d")


def make_safe_filename(text):
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in text)


def logo_bytes(path):
    if not os.path.exists(path):
        return None
    with open(path, "rb") as f:
        return f.read()


def normalized_logo(path, size=(280, 150)):
    """Logo client affiché dans un cadre identique sans déformation."""
    if not os.path.exists(path):
        return None
    try:
        img = Image.open(path).convert("RGB")
        canvas = Image.new("RGB", size, "white")
        contained = ImageOps.contain(img, size)
        x = (size[0] - contained.width) // 2
        y = (size[1] - contained.height) // 2
        canvas.paste(contained, (x, y))
        return canvas
    except Exception:
        return None


# =========================================================
# DOCUMENTS BE / BS
# =========================================================
def get_bon_items(bon_id):
    return query(
        """
        SELECT bi.*, a.name AS article
        FROM bon_items bi
        JOIN articles a ON a.id=bi.article_id
        WHERE bi.bon_id=?
        ORDER BY bi.id
        """,
        (bon_id,),
    )


def get_bon(bon_id):
    return query("SELECT * FROM bons WHERE id=?", (bon_id,), one=True)


def build_document_data(bon_id):
    bon = get_bon(bon_id)
    items = get_bon_items(bon_id)
    return bon, items


def generate_pdf(bon_id):
    bon, items = build_document_data(bon_id)
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    story = []

    # Logos
    logo_list = []
    nomatis = logo_bytes("Logo Nomatis.jpg")
    client_logo = logo_bytes(CLIENTS.get(bon["client"], {}).get("logo", ""))

    if nomatis:
        p = "nomatis_tmp.jpg"
        with open(p, "wb") as f:
            f.write(nomatis)
        logo_list.append(RLImage(p, width=35 * mm, height=18 * mm))
    else:
        logo_list.append(Paragraph("<b>NOMATIS</b>", styles["Normal"]))

    if client_logo:
        p2 = "client_tmp.jpg"
        with open(p2, "wb") as f:
            f.write(client_logo)
        logo_list.append(RLImage(p2, width=35 * mm, height=18 * mm))
    else:
        logo_list.append(Paragraph(bon["client"], styles["Normal"]))

    header = Table([logo_list], colWidths=[90 * mm, 90 * mm])
    header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(header)
    story.append(Spacer(1, 8))

    title = "BON D'ENTRÉE" if bon["type"] == "BE" else "BON DE SORTIE"
    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Spacer(1, 6))

    if bon["type"] == "BE":
        info = [
            ["Bon", bon["number"], "Date", bon["date_bon"]],
            ["Fournisseur", bon["fournisseur"] or "", "Lieu", bon["lieu_livraison"] or ""],
            ["Réceptionné par", bon["receptionne_par"] or "", "Client", bon["client"]],
        ]
    else:
        info = [
            ["Bon", bon["number"], "Date", bon["date_bon"]],
            ["Équipe", bon["equipe"] or "", "Ressource", bon["resource"] or ""],
            ["Destination", bon["destination"] or "", "Client", bon["client"]],
        ]

    info_table = Table(info, colWidths=[32 * mm, 58 * mm, 32 * mm, 58 * mm])
    info_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.8, colors.black),
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#EAF2FF")),
                ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#EAF2FF")),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(info_table)
    story.append(Spacer(1, 10))

    data = [["Référence", "Désignation", "Qté", "Remarque"]]
    for item in items:
        data.append(
            [
                item["reference"] or "",
                item["article"],
                str(item["quantity"]),
                item["remarque"] or "",
            ]
        )

    items_table = Table(data, colWidths=[35 * mm, 65 * mm, 20 * mm, 60 * mm], repeatRows=1)
    items_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.8, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2FF")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (2, 1), (2, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(items_table)
    story.append(Spacer(1, 12))
    story.append(
        Paragraph(
            f"Saisie le {bon['datetime_saisie']} par {bon['created_by']}",
            styles["Normal"],
        )
    )

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def generate_docx(bon_id):
    bon, items = build_document_data(bon_id)
    doc = Document()

    # Logo NOMATIS
    if os.path.exists("Logo Nomatis.jpg"):
        doc.add_picture("Logo Nomatis.jpg", width=Cm(3.0))

    title = "BON D'ENTRÉE" if bon["type"] == "BE" else "BON DE SORTIE"
    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run(title)
    run.bold = True
    run.font.size = None

    if bon["type"] == "BE":
        rows = [
            ["Bon", bon["number"], "Date", bon["date_bon"]],
            ["Fournisseur", bon["fournisseur"] or "", "Lieu", bon["lieu_livraison"] or ""],
            ["Réceptionné par", bon["receptionne_par"] or "", "Client", bon["client"]],
        ]
    else:
        rows = [
            ["Bon", bon["number"], "Date", bon["date_bon"]],
            ["Équipe", bon["equipe"] or "", "Ressource", bon["resource"] or ""],
            ["Destination", bon["destination"] or "", "Client", bon["client"]],
        ]

    table = doc.add_table(rows=len(rows), cols=4)
    table.style = "Table Grid"
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            table.cell(r, c).text = str(value)

    doc.add_paragraph()
    items_table = doc.add_table(rows=1, cols=4)
    items_table.style = "Table Grid"
    headers = ["Référence", "Désignation", "Qté", "Remarque"]
    for c, h in enumerate(headers):
        items_table.cell(0, c).text = h

    for item in items:
        cells = items_table.add_row().cells
        cells[0].text = item["reference"] or ""
        cells[1].text = item["article"]
        cells[2].text = str(item["quantity"])
        cells[3].text = item["remarque"] or ""

    doc.add_paragraph(
        f"Saisie le {bon['datetime_saisie']} par {bon['created_by']}"
    )

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================================================
# AUTHENTIFICATION
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
    left, center, right = st.columns([1, 2, 1])

    with center:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)

        c1, c2 = st.columns([1, 3])
        with c1:
            if os.path.exists("Logo Nomatis.jpg"):
                st.image("Logo Nomatis.jpg", width=80)

        with c2:
            st.markdown(
                '<div class="main-title">Gestion Stock MW NOMATIS</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                '<div class="subtitle">Gestion des entrées, sorties et stock MW</div>',
                unsafe_allow_html=True,
            )

        st.divider()

        username = st.text_input("Nom d'utilisateur", key="login_user")
        password = st.text_input("Mot de passe", type="password", key="login_password")

        st.markdown('<div class="login-btn">', unsafe_allow_html=True)
        connect = st.button("SE CONNECTER", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if connect:
            user = user_info(username)
            if user and user["password"] == password:
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                execute("UPDATE users SET last_login=? WHERE username=?", (now, username))
                st.session_state.logged_in = True
                st.session_state.current_user = username
                st.success("Accès validé ✓")
                st.rerun()
            else:
                st.error("Nom d'utilisateur ou mot de passe incorrect.")

        st.markdown("</div>", unsafe_allow_html=True)


if not st.session_state.logged_in:
    login_screen()
    st.stop()


CURRENT_USER = user_info(st.session_state.current_user)
ROLE = CURRENT_USER["role"]


# =========================================================
# SÉLECTION CLIENT
# =========================================================
def client_selection():
    st.markdown(
        '<div class="main-title">Gestion Stock MW NOMATIS</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="subtitle">Sélectionnez l’espace client</div>',
        unsafe_allow_html=True,
    )
    st.divider()

    cols = st.columns(3)

    for idx, (client, info) in enumerate(CLIENTS.items()):
        with cols[idx]:
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            logo = normalized_logo(info["logo"], (300, 150))
            if logo:
                st.image(logo, use_container_width=True)
            else:
                st.markdown(
                    f"<h2 style='text-align:center'>{client}</h2>",
                    unsafe_allow_html=True,
                )

            st.markdown(
                f"<h3 style='text-align:center;color:{info['color']}'>{client}</h3>",
                unsafe_allow_html=True,
            )

            st.markdown('<div class="stock-btn">', unsafe_allow_html=True)
            if st.button(
                "ACCÉDER AU STOCK",
                key=f"client_{client}",
                use_container_width=True,
            ):
                st.session_state.selected_client = client
                st.session_state.temp_be_items = []
                st.session_state.temp_bs_items = []
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    st.divider()

    # Gestion des utilisateurs ADMIN
    if ROLE == "admin":
        admin_users_management()
    else:
        profile_management()


def profile_management():
    st.subheader("👤 Mon compte")
    with st.form("profile_form"):
        new_name = st.text_input("Nom complet", value=CURRENT_USER["fullname"])
        new_password = st.text_input("Nouveau mot de passe", value=CURRENT_USER["password"], type="password")

        if st.form_submit_button("Enregistrer les modifications"):
            if not new_name.strip() or not new_password:
                st.error("Le nom et le mot de passe sont obligatoires.")
            else:
                execute(
                    "UPDATE users SET fullname=?, password=? WHERE username=?",
                    (new_name.strip(), new_password, st.session_state.current_user),
                )
                st.success("Votre compte a été mis à jour.")
                st.rerun()


def admin_users_management():
    st.subheader("🛠️ Administration des utilisateurs")
    tab_create, tab_edit = st.tabs(["Créer un utilisateur", "Modifier un utilisateur"])

    with tab_create:
        with st.form("create_user_form"):
            username = st.text_input("Identifiant")
            fullname = st.text_input("Nom complet")
            password = st.text_input("Mot de passe", type="password")
            role = st.selectbox("Rôle", ROLES)

            if st.form_submit_button("Créer l'utilisateur"):
                if not username.strip() or not fullname.strip() or not password:
                    st.error("Tous les champs sont obligatoires.")
                elif user_info(username.strip()):
                    st.error("Cet identifiant existe déjà.")
                else:
                    execute(
                        "INSERT INTO users(username,password,fullname,role,last_login) VALUES(?,?,?,?,?)",
                        (username.strip(), password, fullname.strip(), role, "Jamais"),
                    )
                    st.success("Utilisateur créé avec succès.")
                    st.rerun()

    with tab_edit:
        users = query(
            "SELECT username,fullname,role,last_login FROM users ORDER BY username"
        )
        if users:
            st.dataframe(
                pd.DataFrame([dict(u) for u in users]),
                use_container_width=True,
                hide_index=True,
            )

            usernames = [u["username"] for u in users]
            selected = st.selectbox("Utilisateur à modifier", usernames)
            selected_user = user_info(selected)

            with st.form("edit_user_form"):
                fullname = st.text_input("Nom", value=selected_user["fullname"])
                password = st.text_input(
                    "Mot de passe",
                    value=selected_user["password"],
                    type="password",
                )
                role = st.selectbox(
                    "Rôle",
                    ROLES,
                    index=ROLES.index(selected_user["role"]),
                )

                if st.form_submit_button("Mettre à jour"):
                    execute(
                        "UPDATE users SET fullname=?,password=?,role=? WHERE username=?",
                        (fullname.strip(), password, role, selected),
                    )
                    st.success("Utilisateur mis à jour.")
                    st.rerun()


if not st.session_state.selected_client:
    client_selection()
    st.stop()


CLIENT = st.session_state.selected_client

# =========================================================
# HEADER APPLICATION
# =========================================================
head1, head2, head3 = st.columns([2, 4, 2])

with head1:
    if os.path.exists("Logo Nomatis.jpg"):
        st.image("Logo Nomatis.jpg", width=90)

with head2:
    st.markdown(
        '<div class="main-title">Gestion Stock MW NOMATIS</div>',
        unsafe_allow_html=True,
    )
    st.caption(f"Espace client : {CLIENT}")

with head3:
    st.write(f"👤 {CURRENT_USER['fullname']}")
    st.caption(f"Rôle : {ROLE}")

    if st.button("🚪 Déconnexion", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.current_user = None
        st.session_state.selected_client = None
        st.session_state.temp_be_items = []
        st.session_state.temp_bs_items = []
        st.rerun()

st.divider()

# Compte utilisateur accessible à tout moment
with st.expander("👤 Modifier mon compte"):
    with st.form("account_edit_anywhere"):
        account_name = st.text_input("Nom complet", value=CURRENT_USER["fullname"])
        account_password = st.text_input(
            "Mot de passe", value=CURRENT_USER["password"], type="password"
        )
        if st.form_submit_button("Enregistrer mon compte"):
            if not account_name.strip() or not account_password:
                st.error("Le nom et le mot de passe sont obligatoires.")
            else:
                execute(
                    "UPDATE users SET fullname=?, password=? WHERE username=?",
                    (account_name.strip(), account_password, st.session_state.current_user),
                )
                st.success("Compte mis à jour avec succès.")
                st.rerun()


# =========================================================
# 5 RUBRIQUES
# =========================================================
tab_names = [
    "📥 BE",
    "📤 BS",
    "📊 Situation Stock",
    "✏️ Modification & Impression",
    "⚙️ Configuration",
]
t_be, t_bs, t_stock, t_mods, t_config = st.tabs(tab_names)


# =========================================================
# BE
# =========================================================
with t_be:
    if not can(ROLE, "be"):
        st.info("Votre rôle ne permet pas de créer ou modifier un Bon d'Entrée.")
    else:
        st.subheader("Bon d'Entrée")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input(
                "Date / heure de saisie",
                value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                disabled=True,
            )
        with c2:
            date_be = st.date_input(
                "Date du BE *",
                value=date.today(),
                max_value=date.today(),
                key="be_date",
            )
        with c3:
            next_be = query(
                "SELECT COUNT(*) AS n FROM bons WHERE type='BE' AND client=?",
                (CLIENT,),
                one=True,
            )["n"] + 1
            num_be = st.text_input("N° BE *", value=f"BE-{date.today().strftime('%Y%m%d')}-{next_be:04d}")

        c4, c5, c6 = st.columns(3)
        with c4:
            suppliers = active_names("fournisseurs")
            supplier_options = suppliers + ["Autre (saisie manuelle)"]
            supplier_choice = st.selectbox("Fournisseur *", supplier_options)
            if supplier_choice == "Autre (saisie manuelle)":
                fournisseur = st.text_input("Nouveau fournisseur *")
            else:
                fournisseur = supplier_choice
        with c5:
            lieu = st.text_input("Lieu de livraison *", value="Magasin Principal")
        with c6:
            reception = st.text_input(
                "Réceptionné par",
                value=CURRENT_USER["fullname"],
                disabled=True,
            )

        st.markdown("##### Articles du Bon d'Entrée")
        refs, arts, qtys, rems = st.columns([2, 4, 2, 3])

        with refs:
            ref = st.text_input("Référence")
        with arts:
            article_names = active_names("articles")
            if article_names:
                article = st.selectbox("Article *", article_names, key="be_article")
            else:
                article = None
                st.warning("Aucun article configuré.")
        with qtys:
            qty = st.number_input("Quantité *", min_value=1, step=1, value=1, key="be_qty")
        with rems:
            remarque = st.text_input("Remarque", key="be_rem")

        if st.button("➕ Ajouter l'article au BE", key="add_be"):
            if not article:
                st.error("Sélectionnez un article.")
            elif qty <= 0:
                st.error("La quantité doit être supérieure à 0.")
            else:
                st.session_state.temp_be_items.append(
                    {
                        "Référence": ref.strip(),
                        "Article": article,
                        "Quantité": int(qty),
                        "Remarque": remarque.strip(),
                    }
                )
                st.success("Article ajouté au BE.")

        if st.session_state.temp_be_items:
            st.markdown("##### Aperçu du BE")
            st.dataframe(
                pd.DataFrame(st.session_state.temp_be_items),
                use_container_width=True,
                hide_index=True,
            )

            if st.button("🗑️ Vider les articles du BE", key="clear_be"):
                st.session_state.temp_be_items = []
                st.rerun()

            if st.button("💾 Enregistrer le Bon d'Entrée", key="save_be"):
                if not num_be.strip() or not fournisseur.strip() or not lieu.strip():
                    st.error("N° BE, fournisseur et lieu de livraison sont obligatoires.")
                elif not st.session_state.temp_be_items:
                    st.error("Ajoutez au moins un article.")
                elif query(
                    "SELECT id FROM bons WHERE type='BE' AND number=? AND client=?",
                    (num_be.strip(), CLIENT),
                    one=True,
                ):
                    st.error("Ce numéro de BE existe déjà pour ce client.")
                else:
                    bon_id = execute(
                        """
                        INSERT INTO bons
                        (type,number,client,date_bon,datetime_saisie,fournisseur,lieu_livraison,
                         receptionne_par,created_by)
                        VALUES(?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            "BE",
                            num_be.strip(),
                            CLIENT,
                            str(date_be),
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            fournisseur.strip(),
                            lieu.strip(),
                            CURRENT_USER["fullname"],
                            st.session_state.current_user,
                        ),
                    )

                    for item in st.session_state.temp_be_items:
                        aid = article_id_by_name(item["Article"])
                        execute(
                            """
                            INSERT INTO bon_items
                            (bon_id,article_id,reference,quantity,remarque)
                            VALUES(?,?,?,?,?)
                            """,
                            (
                                bon_id,
                                aid,
                                item["Référence"],
                                item["Quantité"],
                                item["Remarque"],
                            ),
                        )
                        old = current_stock(CLIENT, aid)
                        set_stock(CLIENT, aid, old + item["Quantité"])
                        add_movement(
                            CLIENT,
                            aid,
                            "BE",
                            item["Quantité"],
                            num_be.strip(),
                            st.session_state.current_user,
                            item["Remarque"],
                        )

                    st.session_state.temp_be_items = []
                    st.session_state["last_saved_bon"] = bon_id
                    st.success(f"Bon d'Entrée {num_be} enregistré avec succès.")
                    st.rerun()


# =========================================================
# BS
# =========================================================
with t_bs:
    if not can(ROLE, "bs"):
        st.info("Votre rôle ne permet pas de créer ou modifier un Bon de Sortie.")
    else:
        st.subheader("Bon de Sortie")

        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input(
                "Date / heure de saisie",
                value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                disabled=True,
                key="bs_datetime",
            )
        with c2:
            date_bs = st.date_input(
                "Date du BS *",
                value=date.today(),
                max_value=date.today(),
                key="bs_date",
            )
        with c3:
            next_bs = query(
                "SELECT COUNT(*) AS n FROM bons WHERE type='BS' AND client=?",
                (CLIENT,),
                one=True,
            )["n"] + 1
            num_bs = st.text_input(
                "N° BS *",
                value=f"BS-{date.today().strftime('%Y%m%d')}-{next_bs:04d}",
            )

        c4, c5, c6 = st.columns(3)
        with c4:
            teams = active_names("equipes")
            equipe = st.selectbox("Équipe réceptrice *", teams) if teams else ""
        with c5:
            resources = active_names("resources")
            resource = st.selectbox("Ressource projet", [""] + resources)
        with c6:
            destination = st.text_input("Destination / Projet *")

        st.markdown("##### Articles à sortir")
        arts, qtys, rems = st.columns([4, 2, 4])

        with arts:
            article_names = active_names("articles")
            article = st.selectbox("Article *", article_names, key="bs_article") if article_names else None

        with qtys:
            available = current_stock(CLIENT, article_id_by_name(article)) if article else 0
            st.caption(f"Stock disponible : **{available}**")
            qty = st.number_input(
                "Quantité *",
                min_value=1,
                max_value=max(1, available),
                value=1,
                step=1,
                key="bs_qty",
            )

        with rems:
            remarque = st.text_input("Remarque", key="bs_rem")

        if st.button("➕ Ajouter l'article au BS", key="add_bs"):
            if not article:
                st.error("Sélectionnez un article.")
            elif available <= 0:
                st.error("Cet article n'est pas disponible en stock.")
            elif qty <= 0:
                st.error("La quantité doit être supérieure à 0.")
            elif qty > available:
                st.error("La quantité dépasse le stock disponible.")
            else:
                # Empêche de dépasser le stock avec plusieurs lignes du même article
                already = sum(
                    x["Quantité"]
                    for x in st.session_state.temp_bs_items
                    if x["Article"] == article
                )
                if already + qty > available:
                    st.error("La quantité cumulée de cet article dépasse le stock disponible.")
                else:
                    st.session_state.temp_bs_items.append(
                        {
                            "Article": article,
                            "Quantité": int(qty),
                            "Remarque": remarque.strip(),
                        }
                    )
                    st.success("Article ajouté au BS.")

        if st.session_state.temp_bs_items:
            st.markdown("##### Aperçu du BS")
            st.dataframe(
                pd.DataFrame(st.session_state.temp_bs_items),
                use_container_width=True,
                hide_index=True,
            )

            if st.button("🗑️ Vider les articles du BS", key="clear_bs"):
                st.session_state.temp_bs_items = []
                st.rerun()

            if st.button("💾 Enregistrer le Bon de Sortie", key="save_bs"):
                if not num_bs.strip() or not destination.strip() or not equipe:
                    st.error("N° BS, équipe et destination sont obligatoires.")
                elif not st.session_state.temp_bs_items:
                    st.error("Ajoutez au moins un article.")
                elif query(
                    "SELECT id FROM bons WHERE type='BS' AND number=? AND client=?",
                    (num_bs.strip(), CLIENT),
                    one=True,
                ):
                    st.error("Ce numéro de BS existe déjà pour ce client.")
                else:
                    # Vérification finale avant modification du stock
                    errors = []
                    for item in st.session_state.temp_bs_items:
                        aid = article_id_by_name(item["Article"])
                        if item["Quantité"] <= 0:
                            errors.append(f"{item['Article']} : quantité invalide.")
                        elif item["Quantité"] > current_stock(CLIENT, aid):
                            errors.append(f"{item['Article']} : stock insuffisant.")

                    if errors:
                        for e in errors:
                            st.error(e)
                    else:
                        bon_id = execute(
                            """
                            INSERT INTO bons
                            (type,number,client,date_bon,datetime_saisie,equipe,resource,
                             destination,created_by)
                            VALUES(?,?,?,?,?,?,?,?,?)
                            """,
                            (
                                "BS",
                                num_bs.strip(),
                                CLIENT,
                                str(date_bs),
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                equipe,
                                resource,
                                destination.strip(),
                                st.session_state.current_user,
                            ),
                        )

                        for item in st.session_state.temp_bs_items:
                            aid = article_id_by_name(item["Article"])
                            execute(
                                """
                                INSERT INTO bon_items
                                (bon_id,article_id,reference,quantity,remarque)
                                VALUES(?,?,?,?,?)
                                """,
                                (
                                    bon_id,
                                    aid,
                                    "",
                                    item["Quantité"],
                                    item["Remarque"],
                                ),
                            )
                            old = current_stock(CLIENT, aid)
                            set_stock(CLIENT, aid, old - item["Quantité"])
                            add_movement(
                                CLIENT,
                                aid,
                                "BS",
                                item["Quantité"],
                                num_bs.strip(),
                                st.session_state.current_user,
                                item["Remarque"],
                            )

                        st.session_state.temp_bs_items = []
                        st.session_state["last_saved_bon"] = bon_id
                        st.success(f"Bon de Sortie {num_bs} enregistré avec succès.")
                        st.rerun()


# =========================================================
# SITUATION STOCK
# =========================================================
with t_stock:
    st.subheader(f"Situation du Stock — {CLIENT}")

    rows = []
    for a in article_rows():
        aid = a["id"]

        be = query(
            """
            SELECT COALESCE(SUM(quantity),0) AS q
            FROM movements
            WHERE client=? AND article_id=? AND movement_type='BE'
            """,
            (CLIENT, aid),
            one=True,
        )["q"]

        bs = query(
            """
            SELECT COALESCE(SUM(quantity),0) AS q
            FROM movements
            WHERE client=? AND article_id=? AND movement_type='BS'
            """,
            (CLIENT, aid),
            one=True,
        )["q"]

        manual = query(
            """
            SELECT COALESCE(SUM(quantity),0) AS q
            FROM movements
            WHERE client=? AND article_id=? AND movement_type='MANUEL'
            """,
            (CLIENT, aid),
            one=True,
        )["q"]

        rows.append(
            {
                "Article": a["name"],
                "Entrées BE": int(be),
                "Sorties BS": int(bs),
                "Ajustements manuels": int(manual),
                "Stock actuel": current_stock(CLIENT, aid),
            }
        )

    if rows:
        df_stock = pd.DataFrame(rows)
        st.dataframe(df_stock, use_container_width=True, hide_index=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Articles", len(rows))
        with c2:
            st.metric("Total unités en stock", int(df_stock["Stock actuel"].sum()))
        with c3:
            st.metric("Sorties BS", int(df_stock["Sorties BS"].sum()))
    else:
        st.info("Aucun article configuré.")


# =========================================================
# MODIFICATION & IMPRESSION
# =========================================================
with t_mods:
    st.subheader("Modification et impression des Bons")

    if ROLE in ("coordinateur", "coordinatrice"):
        st.info("Votre rôle permet la consultation et l'impression des BE/BS, mais pas leur modification.")

    type_choice = st.radio(
        "Choisir le type de bon",
        ["BE", "BS"],
        horizontal=True,
        key="mod_type",
    )

    bons = query(
        "SELECT * FROM bons WHERE client=? AND type=? ORDER BY id DESC",
        (CLIENT, type_choice),
    )

    if not bons:
        st.info(f"Aucun {type_choice} enregistré pour {CLIENT}.")
    else:
        labels = [f"{b['number']} — {b['date_bon']}" for b in bons]
        selected_label = st.selectbox("Choisir un bon", labels)
        selected_id = bons[labels.index(selected_label)]["id"]
        bon = get_bon(selected_id)
        items = get_bon_items(selected_id)

        st.write(
            f"**N° :** {bon['number']}  |  **Date :** {bon['date_bon']}  |  "
            f"**Saisi par :** {bon['created_by']}"
        )

        if bon["type"] == "BE":
            st.write(
                f"**Fournisseur :** {bon['fournisseur']} | "
                f"**Lieu :** {bon['lieu_livraison']} | "
                f"**Réceptionné par :** {bon['receptionne_par']}"
            )
        else:
            st.write(
                f"**Équipe :** {bon['equipe']} | "
                f"**Ressource :** {bon['resource'] or '-'} | "
                f"**Destination :** {bon['destination']}"
            )

        st.dataframe(
            pd.DataFrame([dict(i) for i in items]),
            use_container_width=True,
            hide_index=True,
        )

        # Impression accessible à tous les rôles autorisés
        pdf_data = generate_pdf(selected_id)
        docx_data = generate_docx(selected_id)

        p1, p2 = st.columns(2)
        with p1:
            st.download_button(
                "📄 Télécharger en PDF",
                data=pdf_data,
                file_name=f"{make_safe_filename(bon['number'])}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with p2:
            st.download_button(
                "📝 Télécharger en Word",
                data=docx_data,
                file_name=f"{make_safe_filename(bon['number'])}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        if can(ROLE, "edit"):
            st.divider()
            st.markdown("##### Modifier le bon")

            with st.form(f"edit_bon_{selected_id}"):
                if bon["type"] == "BE":
                    new_date = st.date_input(
                        "Date du BE",
                        value=datetime.strptime(bon["date_bon"], "%Y-%m-%d").date(),
                        max_value=date.today(),
                    )
                    suppliers = active_names("fournisseurs")
                    supplier_options = suppliers + ["Autre (saisie manuelle)"]
                    current_supplier = bon["fournisseur"] or ""
                    if current_supplier in suppliers:
                        sup_index = supplier_options.index(current_supplier)
                        supplier_choice = st.selectbox(
                            "Fournisseur",
                            supplier_options,
                            index=sup_index,
                        )
                        new_supplier = supplier_choice
                    else:
                        supplier_choice = st.selectbox(
                            "Fournisseur",
                            supplier_options,
                            index=len(supplier_options) - 1,
                        )
                        new_supplier = (
                            st.text_input("Nouveau fournisseur", value=current_supplier)
                            if supplier_choice == "Autre (saisie manuelle)"
                            else supplier_choice
                        )
                    new_place = st.text_input(
                        "Lieu de livraison",
                        value=bon["lieu_livraison"] or "",
                    )
                else:
                    new_date = st.date_input(
                        "Date du BS",
                        value=datetime.strptime(bon["date_bon"], "%Y-%m-%d").date(),
                        max_value=date.today(),
                    )
                    teams = active_names("equipes")
                    resources = active_names("resources")
                    team_current = bon["equipe"] or ""
                    resource_current = bon["resource"] or ""
                    new_team = st.selectbox(
                        "Équipe",
                        teams,
                        index=teams.index(team_current) if team_current in teams else 0,
                    )
                    resource_options = [""] + resources
                    new_resource = st.selectbox(
                        "Ressource",
                        resource_options,
                        index=resource_options.index(resource_current)
                        if resource_current in resource_options else 0,
                    )
                    new_destination = st.text_input(
                        "Destination / Projet",
                        value=bon["destination"] or "",
                    )

                st.markdown("**Articles du bon**")
                edited_items = []
                article_names = active_names("articles")

                for idx, item in enumerate(items):
                    c1, c2, c3, c4 = st.columns([3, 3, 1.5, 3])
                    with c1:
                        old_article = item["article"]
                        new_article = st.selectbox(
                            "Article",
                            article_names,
                            index=article_names.index(old_article)
                            if old_article in article_names else 0,
                            key=f"edit_art_{selected_id}_{idx}",
                        )
                    with c2:
                        new_ref = st.text_input(
                            "Référence",
                            value=item["reference"] or "",
                            key=f"edit_ref_{selected_id}_{idx}",
                        )
                    with c3:
                        new_qty = st.number_input(
                            "Qté",
                            min_value=1,
                            value=max(1, int(item["quantity"])),
                            step=1,
                            key=f"edit_qty_{selected_id}_{idx}",
                        )
                    with c4:
                        new_rem = st.text_input(
                            "Remarque",
                            value=item["remarque"] or "",
                            key=f"edit_rem_{selected_id}_{idx}",
                        )

                    edited_items.append(
                        {
                            "id": item["id"],
                            "article": new_article,
                            "reference": new_ref,
                            "quantity": int(new_qty),
                            "remarque": new_rem,
                        }
                    )

                if st.form_submit_button("💾 Enregistrer les modifications"):
                    # Vérification du futur stock : on annule d'abord l'ancien mouvement,
                    # puis on applique le nouveau.
                    old_items = items

                    # Stock impact des anciens articles
                    old_sign = 1 if bon["type"] == "BE" else -1
                    for old in old_items:
                        old_aid = article_id_by_name(old["article"])
                        old_stock = current_stock(CLIENT, old_aid)
                        set_stock(CLIENT, old_aid, old_stock - old_sign * int(old["quantity"]))

                    # Contrôle BS avant application
                    if bon["type"] == "BS":
                        required = {}
                        for e in edited_items:
                            aid = article_id_by_name(e["article"])
                            required[aid] = required.get(aid, 0) + e["quantity"]
                        bad = []
                        for aid, req in required.items():
                            if current_stock(CLIENT, aid) < req:
                                arow = query("SELECT name FROM articles WHERE id=?", (aid,), one=True)
                                bad.append(f"{arow['name']} : stock insuffisant ({current_stock(CLIENT, aid)} disponible).")
                        if bad:
                            for msg in bad:
                                st.error(msg)
                            st.stop()

                    # Supprimer anciennes lignes et mouvements associés
                    execute("DELETE FROM bon_items WHERE bon_id=?", (selected_id,))
                    execute(
                        "DELETE FROM movements WHERE reference_bon=? AND client=?",
                        (bon["number"], CLIENT),
                    )

                    # Mise à jour entête
                    if bon["type"] == "BE":
                        execute(
                            """
                            UPDATE bons
                            SET date_bon=?, fournisseur=?, lieu_livraison=?
                            WHERE id=?
                            """,
                            (str(new_date), new_supplier.strip(), new_place.strip(), selected_id),
                        )
                    else:
                        execute(
                            """
                            UPDATE bons
                            SET date_bon=?, equipe=?, resource=?, destination=?
                            WHERE id=?
                            """,
                            (str(new_date), new_team, new_resource, new_destination.strip(), selected_id),
                        )

                    sign = 1 if bon["type"] == "BE" else -1

                    for e in edited_items:
                        aid = article_id_by_name(e["article"])
                        execute(
                            """
                            INSERT INTO bon_items
                            (bon_id,article_id,reference,quantity,remarque)
                            VALUES(?,?,?,?,?)
                            """,
                            (
                                selected_id,
                                aid,
                                e["reference"].strip(),
                                e["quantity"],
                                e["remarque"].strip(),
                            ),
                        )

                        old_stock = current_stock(CLIENT, aid)
                        set_stock(CLIENT, aid, old_stock + sign * e["quantity"])
                        add_movement(
                            CLIENT,
                            aid,
                            bon["type"],
                            e["quantity"],
                            bon["number"],
                            st.session_state.current_user,
                            e["remarque"],
                        )

                    st.success("Bon modifié avec succès et stock recalculé.")
                    st.rerun()

            if st.button("❌ Supprimer ce bon", key=f"delete_{selected_id}"):
                # Annuler l'impact stock avant suppression
                sign = 1 if bon["type"] == "BE" else -1
                for item in items:
                    aid = article_id_by_name(item["article"])
                    old_stock = current_stock(CLIENT, aid)
                    set_stock(CLIENT, aid, old_stock - sign * int(item["quantity"]))

                execute("DELETE FROM bon_items WHERE bon_id=?", (selected_id,))
                execute("DELETE FROM movements WHERE reference_bon=? AND client=?", (bon["number"], CLIENT))
                execute("DELETE FROM bons WHERE id=?", (selected_id,))
                st.success("Bon supprimé et stock recalculé.")
                st.rerun()


# =========================================================
# CONFIGURATION ADMIN
# =========================================================
with t_config:
    if ROLE != "admin":
        st.warning("🔒 Cette rubrique est réservée à l'administrateur.")
    else:
        st.subheader("Configuration système")

        c1, c2 = st.columns(2)

        with c1:
            st.markdown("##### 📦 Articles")
            article_new = st.text_input("Nouvel article", key="new_article")
            initial_qty = st.number_input(
                "Quantité initiale (0 si aucune)",
                min_value=0,
                value=0,
                step=1,
                key="initial_article_qty",
            )

            if st.button("➕ Ajouter l'article", key="add_article"):
                if not article_new.strip():
                    st.error("Le nom de l'article est obligatoire.")
                elif query("SELECT id FROM articles WHERE name=?", (article_new.strip(),), one=True):
                    st.error("Cet article existe déjà.")
                else:
                    aid = execute(
                        "INSERT INTO articles(name,active) VALUES(?,1)",
                        (article_new.strip(),),
                    )
                    for client in CLIENTS:
                        set_stock(client, aid, int(initial_qty))
                    st.success("Article ajouté.")
                    st.rerun()

            articles_all = query("SELECT id,name,active FROM articles ORDER BY name")
            if articles_all:
                selected_article = st.selectbox(
                    "Article à gérer",
                    [a["name"] for a in articles_all],
                    key="cfg_article_select",
                )
                selected_article_row = next(a for a in articles_all if a["name"] == selected_article)

                col_a1, col_a2 = st.columns(2)
                with col_a1:
                    renamed = st.text_input(
                        "Modifier le nom",
                        value=selected_article_row["name"],
                        key="rename_article",
                    )
                    if st.button("✏️ Modifier l'article", key="rename_article_btn"):
                        if not renamed.strip():
                            st.error("Le nom ne peut pas être vide.")
                        else:
                            try:
                                execute(
                                    "UPDATE articles SET name=? WHERE id=?",
                                    (renamed.strip(), selected_article_row["id"]),
                                )
                                st.success("Article modifié.")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.error("Ce nom existe déjà.")

                with col_a2:
                    if st.button("🗑️ Supprimer l'article", key="delete_article"):
                        used = query(
                            "SELECT COUNT(*) AS n FROM bon_items WHERE article_id=?",
                            (selected_article_row["id"],),
                            one=True,
                        )["n"]
                        stock_used = sum(
                            current_stock(c, selected_article_row["id"]) for c in CLIENTS
                        )
                        if used or stock_used:
                            st.error(
                                "Impossible de supprimer cet article : il possède un historique ou un stock. "
                                "Vous pouvez le désactiver."
                            )
                        else:
                            execute(
                                "UPDATE articles SET active=0 WHERE id=?",
                                (selected_article_row["id"],),
                            )
                            st.success("Article supprimé de la liste active.")
                            st.rerun()

        with c2:
            st.markdown("##### 🏢 Fournisseurs")
            new_supplier = st.text_input("Nouveau fournisseur", key="new_supplier")
            if st.button("➕ Ajouter le fournisseur", key="add_supplier"):
                if not new_supplier.strip():
                    st.error("Nom obligatoire.")
                elif query("SELECT id FROM fournisseurs WHERE name=?", (new_supplier.strip(),), one=True):
                    st.error("Ce fournisseur existe déjà.")
                else:
                    execute(
                        "INSERT INTO fournisseurs(name,active) VALUES(?,1)",
                        (new_supplier.strip(),),
                    )
                    st.success("Fournisseur ajouté.")
                    st.rerun()

            st.write("Fournisseurs actifs :", ", ".join(active_names("fournisseurs")))

            st.markdown("##### 👷 Équipes")
            new_team = st.text_input("Nouvelle équipe", key="new_team")
            if st.button("➕ Ajouter l'équipe", key="add_team"):
                if not new_team.strip():
                    st.error("Nom obligatoire.")
                elif query("SELECT id FROM equipes WHERE name=?", (new_team.strip(),), one=True):
                    st.error("Cette équipe existe déjà.")
                else:
                    execute(
                        "INSERT INTO equipes(name,active) VALUES(?,1)",
                        (new_team.strip(),),
                    )
                    st.success("Équipe ajoutée.")
                    st.rerun()

            st.write("Équipes actives :", ", ".join(active_names("equipes")))

            st.markdown("##### 👤 Ressources projets")
            new_resource = st.text_input("Nouvelle ressource", key="new_resource")
            if st.button("➕ Ajouter la ressource", key="add_resource"):
                if not new_resource.strip():
                    st.error("Nom obligatoire.")
                elif query("SELECT id FROM resources WHERE name=?", (new_resource.strip(),), one=True):
                    st.error("Cette ressource existe déjà.")
                else:
                    execute(
                        "INSERT INTO resources(name,active) VALUES(?,1)",
                        (new_resource.strip(),),
                    )
                    st.success("Ressource ajoutée.")
                    st.rerun()

            st.write("Ressources actives :", ", ".join(active_names("resources")))

        st.divider()

        st.markdown("##### 🔧 Ajustement manuel du stock")
        st.caption(
            "Toute modification manuelle est enregistrée comme mouvement MANUEL et apparaît dans la situation du stock."
        )

        c1, c2, c3 = st.columns(3)
        with c1:
            clients_options = list(CLIENTS.keys())
            manual_client = st.selectbox("Client", clients_options, key="manual_client")
        with c2:
            article_options = active_names("articles")
            manual_article = st.selectbox("Article", article_options, key="manual_article")
        with c3:
            aid = article_id_by_name(manual_article)
            old_qty = current_stock(manual_client, aid)
            new_qty = st.number_input(
                "Nouvelle quantité",
                min_value=0,
                value=old_qty,
                step=1,
                key="manual_qty",
            )

        manual_comment = st.text_input("Motif de l'ajustement", key="manual_comment")

        if st.button("💾 Appliquer l'ajustement manuel", key="apply_manual"):
            if new_qty < 0:
                st.error("La quantité ne peut pas être négative.")
            elif new_qty == old_qty:
                st.info("Aucun changement.")
            else:
                delta = int(new_qty - old_qty)
                set_stock(manual_client, aid, int(new_qty))
                add_movement(
                    manual_client,
                    aid,
                    "MANUEL",
                    delta,
                    "AJUSTEMENT_ADMIN",
                    st.session_state.current_user,
                    manual_comment.strip(),
                )
                st.success(
                    f"Stock de {manual_article} pour {manual_client} mis à jour : "
                    f"{old_qty} → {new_qty}."
                )
                st.rerun()
