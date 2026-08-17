import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
from io import BytesIO

# Imports pour génération PDF et Word
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------------------------------------------------
# CONFIGURATION STREAMLIT & STYLE CSS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Gestion de Stock Télécom",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stApp { max-width: 100%; }
    .glass-card {
        background: white;
        padding: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 20px;
        border: 1px solid #e9ecef;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
    }
    .stMetric {
        background: #f1f3f5;
        padding: 12px;
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# BASE DE DONNÉES & FONCTIONS HELPER
# ---------------------------------------------------------
DB_FILE = "telecom_stock.db"

def get_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    c = conn.cursor()
    
    # Tables
    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        username TEXT PRIMARY KEY,
        password TEXT,
        fullname TEXT,
        role TEXT,
        last_login TEXT
    )""")
    
    c.execute("""
    CREATE TABLE IF NOT EXISTS articles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        active INTEGER DEFAULT 1
    )""")
    
    c.execute("""
    CREATE TABLE IF NOT EXISTS fournisseurs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        active INTEGER DEFAULT 1
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS equipes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        active INTEGER DEFAULT 1
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS resources (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        active INTEGER DEFAULT 1
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS stock (
        client TEXT,
        article_id INTEGER,
        quantity INTEGER,
        PRIMARY KEY (client, article_id)
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS bons (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        type TEXT, -- BE ou BS
        number TEXT UNIQUE,
        client TEXT,
        date_bon TEXT,
        datetime_saisie TEXT,
        fournisseur TEXT,
        lieu_livraison TEXT,
        equipe TEXT,
        resource TEXT,
        destination TEXT,
        created_by TEXT
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS bon_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bon_id INTEGER,
        article_id INTEGER,
        reference TEXT,
        quantity INTEGER,
        remarque TEXT
    )""")

    c.execute("""
    CREATE TABLE IF NOT EXISTS movements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT,
        client TEXT,
        article_id INTEGER,
        m_type TEXT, -- ENTREE ou SORTIE
        qty INTEGER,
        ref_bon TEXT,
        user TEXT,
        comment TEXT,
        fournisseur TEXT,
        equipe TEXT
    )""")

    # Admin par défaut (Changer en production)
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('admin', 'admin123', 'Administrateur Principal', 'admin', 'Jamais')")

    # Données de démonstration
    arts = ["Câble Fibre Optique 24FO (mètre)", "Jarretière Optique SC/APC-LC/UPC", "Splicer Fusion Fitel", "Module SFP+ 10G", "Odf 24 Ports"]
    for a in arts:
        c.execute("INSERT OR IGNORE INTO articles (name, active) VALUES (?, 1)", (a,))

    fourns = ["Huawei Technologies", "Nokia Telecom", "Sagemcom", "FiberHome"]
    for f in fourns:
        c.execute("INSERT OR IGNORE INTO fournisseurs (name, active) VALUES (?, 1)", (f,))

    eqs = ["Équipe Build Fibre A", "Équipe Maintenance B", "Équipe Radio 5G"]
    for eq in eqs:
        c.execute("INSERT OR IGNORE INTO equipes (name, active) VALUES (?, 1)", (eq,))

    res = ["Karim Benani", "Youssef Amrani", "Rachid El Amrani"]
    for r in res:
        c.execute("INSERT OR IGNORE INTO resources (name, active) VALUES (?, 1)", (r,))

    conn.commit()
    conn.close()

init_db()

def query(sql, params=()):
    conn = get_connection()
    c = conn.cursor()
    c.execute(sql, params)
    rows = c.fetchall()
    conn.close()
    return rows

def execute(sql, params=()):
    conn = get_connection()
    c = conn.cursor()
    c.execute(sql, params)
    last_id = c.lastrowid
    conn.commit()
    conn.close()
    return last_id

def active_names(table):
    rows = query(f"SELECT name FROM {table} WHERE active=1 ORDER BY name")
    return [r["name"] for r in rows]

def article_id_by_name(name):
    rows = query("SELECT id FROM articles WHERE name=?", (name,))
    return rows[0]["id"] if rows else None

def current_stock(client, article_id):
    rows = query("SELECT quantity FROM stock WHERE client=? AND article_id=?", (client, article_id))
    return rows[0]["quantity"] if rows else 0

def set_stock(client, article_id, new_qty):
    execute("""
        INSERT INTO stock (client, article_id, quantity) VALUES (?, ?, ?)
        ON CONFLICT(client, article_id) DO UPDATE SET quantity=excluded.quantity
    """, (client, article_id, new_qty))

def add_movement(client, article_id, m_type, qty, ref_bon, user, comment="", fournisseur="", equipe=""):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    execute("""
        INSERT INTO movements (timestamp, client, article_id, m_type, qty, ref_bon, user, comment, fournisseur, equipe)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (ts, client, article_id, m_type, qty, ref_bon, user, comment, fournisseur, equipe))

def generate_bon_number(b_type, client):
    prefix = f"{b_type}-{client[:3].upper()}-{datetime.now().strftime('%Y%m')}"
    rows = query("SELECT number FROM bons WHERE number LIKE ? ORDER BY id DESC LIMIT 1", (f"{prefix}%",))
    if rows:
        last_num = rows[0]["number"]
        try:
            seq = int(last_num.split("-")[-1]) + 1
        except:
            seq = 1
    else:
        seq = 1
    return f"{prefix}-{seq:04d}"

# ---------------------------------------------------------
# GÉNÉRATION DE DOCUMENTS (PDF & WORD)
# ---------------------------------------------------------
def generate_pdf(bon_id):
    b_rows = query("SELECT * FROM bons WHERE id=?", (bon_id,))
    if not b_rows:
        return None
    b = b_rows[0]
    
    items = query("""
        SELECT i.*, a.name as article_name 
        FROM bon_items i 
        JOIN articles a ON i.article_id = a.id 
        WHERE i.bon_id=?
    """, (bon_id,))

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    styles = getSampleStyleSheet()

    # Style
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1E3A8A'), alignment=1)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#4B5563'), alignment=1)

    story.append(Paragraph(f"<b>BON DE {'RÉCEPTION / ENTRÉE' if b['type']=='BE' else 'SORTIE / LIVRAISON'}</b>", title_style))
    story.append(Paragraph(f"Client / Opérateur : <b>{b['client']}</b> | Référence : <b>{b['number']}</b>", subtitle_style))
    story.append(Spacer(1, 15))

    # General Info Table
    if b['type'] == 'BE':
        info_data = [
            [f"N° Bon: {b['number']}", f"Date d'entrée: {b['date_bon']}"],
            [f"Fournisseur: {b['fournisseur']}", f"Lieu de livraison: {b['lieu_livraison']}"],
            [f"Saisi le: {b['datetime_saisie']}", f"Saisi par: {b['created_by']}"]
        ]
    else:
        info_data = [
            [f"N° Bon: {b['number']}", f"Date de sortie: {b['date_bon']}"],
            [f"Équipe: {b['equipe']}", f"Technicien / Ressource: {b['resource']}"],
            [f"Destination / Site: {b['destination']}", f"Saisi par: {b['created_by']}"]
        ]

    t_info = Table(info_data, colWidths=[260, 260])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F3F4F6')),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#1F2937')),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('PADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 15))

    # Items Table
    table_data = [["N°", "Désignation Article", "Référence", "Quantité", "Remarques"]]
    for idx, item in enumerate(items, 1):
        table_data.append([
            str(idx),
            item['article_name'],
            item['reference'] or "-",
            str(item['quantity']),
            item['remarque'] or "-"
        ])

    t_items = Table(table_data, colWidths=[30, 200, 100, 60, 130])
    t_items.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('ALIGN', (3,1), (3,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('PADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_items)
    story.append(Spacer(1, 30))

    # Signatures
    sig_data = [["Visa Magasinier / Gestionnaire", "Visa Bénéficiaire / Transporteur"]]
    t_sig = Table(sig_data, colWidths=[260, 260])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 40),
    ]))
    story.append(t_sig)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

def generate_docx(bon_id):
    b_rows = query("SELECT * FROM bons WHERE id=?", (bon_id,))
    if not b_rows:
        return None
    b = b_rows[0]
    
    items = query("""
        SELECT i.*, a.name as article_name 
        FROM bon_items i 
        JOIN articles a ON i.article_id = a.id 
        WHERE i.bon_id=?
    """, (bon_id,))

    doc = docx.Document()

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(f"BON DE {'RÉCEPTION (BE)' if b['type']=='BE' else 'SORTIE (BS)'}")
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = RGBColor(30, 58, 138)

    sub = doc.add_paragraph(f"Client: {b['client']} | N°: {b['number']} | Date: {b['date_bon']}")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # Table Articles
    t = doc.add_table(rows=1, cols=4)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    hdr[0].text = "Article"
    hdr[1].text = "Référence"
    hdr[2].text = "Quantité"
    hdr[3].text = "Remarques"

    for item in items:
        row = t.add_row().cells
        row[0].text = str(item['article_name'])
        row[1].text = str(item['reference'] or "")
        row[2].text = str(item['quantity'])
        row[3].text = str(item['remarque'] or "")

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------------------------------------
# AUTHENTIFICATION & SESSIONS
# ---------------------------------------------------------
if "user" not in st.session_state:
    st.session_state.user = None
if "temp_be_items" not in st.session_state:
    st.session_state.temp_be_items = []
if "temp_bs_items" not in st.session_state:
    st.session_state.temp_bs_items = []

# Page de Login
if not st.session_state.user:
    st.markdown("<h2 style='text-align: center;'>🔐 Connexion — Gestion de Stock Télécom</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        with st.form("login_form"):
            u = st.text_input("Identifiant")
            p = st.text_input("Mot de passe", type="password")
            submit = st.form_submit_button("Se connecter", use_container_width=True)
            if submit:
                res = query("SELECT * FROM users WHERE username=? AND password=?", (u, p))
                if res:
                    st.session_state.user = dict(res[0])
                    execute("UPDATE users SET last_login=? WHERE username=?", (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), u))
                    st.success("Connexion réussie !")
                    st.rerun()
                else:
                    st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

# User loggé
CURRENT_USER = st.session_state.user
ROLE = CURRENT_USER["role"]
ROLES = ["admin", "magasinier", "consultant"]

# Sidebar
st.sidebar.title(f"👤 {CURRENT_USER['fullname']}")
st.sidebar.caption(f"Rôle : **{ROLE.upper()}**")

if st.sidebar.button("🚪 Déconnexion"):
    st.session_state.user = None
    st.rerun()

st.sidebar.markdown("---")
CLIENTS = ["INWI", "ORANGE", "MAROC TELECOM"]
CLIENT = st.sidebar.selectbox("🎯 Sélectionner le Client / Opérateur", CLIENTS)

# Navigation principale
tabs = st.tabs([
    "📥 Bon d'Entrée (BE)",
    "📤 Bon de Sortie (BS)",
    "📊 Situation Stock",
    "📜 Historique & Impression",
    "⚙️ Configuration"
])

# =========================================================
# RUBRIQUE 1 : BON D'ENTRÉE (BE)
# =========================================================
with tabs[0]:
    if ROLE in ["admin", "magasinier"]:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Création Bon d'Entrée (BE)")

        auto_date_saisie_be = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        auto_num_be = generate_bon_number("BE", CLIENT)

        st.info(f"🕒 **Date & Heure de saisie :** {auto_date_saisie_be} | 🔢 **N° BE Généré :** `{auto_num_be}`")

        c1, c2 = st.columns(2)
        date_be = c1.date_input("Date du Bon d'Entrée*", value=date.today(), max_value=date.today(), key="be_date")
        
        fourn_list = active_names("fournisseurs")
        fournisseur_be = c2.selectbox("Fournisseur*", fourn_list + ["Autre..."])
        if fournisseur_be == "Autre...":
            fournisseur_be_custom = c2.text_input("Précisez le nom du fournisseur*")
            fournisseur_final = fournisseur_be_custom
        else:
            fournisseur_final = fournisseur_be

        lieu_livraison = st.text_input("Lieu de Livraison / Entpôt*", value="Dépôt Principal")

        st.markdown("---")
        st.markdown("##### Saisie des articles pour le BE")

        art_list = active_names("articles")
        r1, r2, r3, r4 = st.columns([3, 2, 1.5, 3])
        sel_article_be = r1.selectbox("Article*", art_list, key="be_art")
        ref_article_be = r2.text_input("Référence (SN / P/N)", key="be_ref")
        qty_article_be = r3.number_input("Quantité*", min_value=1, value=1, step=1, key="be_qty")
        rem_article_be = r4.text_input("Remarque / État", key="be_rem")

        if st.button("➕ Ajouter l'article à l'entrée", key="add_be_line"):
            found = False
            for item in st.session_state.temp_be_items:
                if item["article"] == sel_article_be and item["reference"] == ref_article_be:
                    item["quantity"] += qty_article_be
                    if rem_article_be:
                        item["remarque"] += f" | {rem_article_be}"
                    found = True
                    break
            if not found:
                st.session_state.temp_be_items.append({
                    "article": sel_article_be,
                    "reference": ref_article_be,
                    "quantity": qty_article_be,
                    "remarque": rem_article_be
                })
            st.rerun()

        if st.session_state.temp_be_items:
            st.markdown("###### Articles prêts pour l'enregistrement :")
            st.dataframe(pd.DataFrame(st.session_state.temp_be_items), use_container_width=True)

            b1, b2 = st.columns(2)
            if b1.button("💾 Valider et Enregistrer le BE", key="btn_save_be", use_container_width=True):
                if fournisseur_be == "Autre..." and not fournisseur_final.strip():
                    st.error("Veuillez indiquer le nom du fournisseur.")
                else:
                    st.session_state.confirm_be_save = True

            if st.session_state.get("confirm_be_save", False):
                st.warning("⚠️ Confirmez-vous l'enregistrement définitif de ce Bon d'Entrée ?")
                cb1, cb2 = st.columns(2)
                if cb1.button("✅ Oui, Confirmer l'enregistrement (BE)"):
                    bon_id = execute(
                        """
                        INSERT INTO bons (type,number,client,date_bon,datetime_saisie,fournisseur,lieu_livraison,created_by)
                        VALUES (?,?,?,?,?,?,?,?)
                        """,
                        (
                            "BE",
                            auto_num_be,
                            CLIENT,
                            str(date_be),
                            auto_date_saisie_be,
                            fournisseur_final,
                            lieu_livraison,
                            CURRENT_USER["username"]
                        )
                    )

                    for item in st.session_state.temp_be_items:
                        art_id = article_id_by_name(item["article"])
                        execute(
                            "INSERT INTO bon_items (bon_id, article_id, reference, quantity, remarque) VALUES (?,?,?,?,?)",
                            (bon_id, art_id, item["reference"], item["quantity"], item["remarque"])
                        )
                        c_qty = current_stock(CLIENT, art_id)
                        set_stock(CLIENT, art_id, c_qty + item["quantity"])
                        add_movement(
                            client=CLIENT,
                            article_id=art_id,
                            m_type="ENTREE",
                            qty=item["quantity"],
                            ref_bon=auto_num_be,
                            user=CURRENT_USER["username"],
                            comment=item["remarque"],
                            fournisseur=fournisseur_final
                        )

                    st.session_state.temp_be_items = []
                    st.session_state.confirm_be_save = False
                    st.success(f"✅ Bon d'Entrée {auto_num_be} enregistré avec succès !")
                    st.rerun()

                if cb2.button("❌ Annuler"):
                    st.session_state.confirm_be_save = False
                    st.rerun()

            if b2.button("🗑️ Vider la liste", key="clear_be", use_container_width=True):
                st.session_state.temp_be_items = []
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.warning("🔒 Vous n'avez pas les droits nécessaires pour saisir un Bon d'Entrée.")


# =========================================================
# RUBRIQUE 2 : BON DE SORTIE (BS)
# =========================================================
with tabs[1]:
    if ROLE in ["admin", "magasinier"]:
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("Création Bon de Sortie (BS)")

        auto_date_saisie_bs = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        auto_num_bs = generate_bon_number("BS", CLIENT)

        st.info(f"🕒 **Date & Heure de saisie :** {auto_date_saisie_bs} | 🔢 **N° BS Généré :** `{auto_num_bs}`")

        c1, c2 = st.columns(2)
        date_bs = c1.date_input("Date du Bon de Sortie*", value=date.today(), max_value=date.today(), key="bs_date_input")
        
        eq_list = active_names("equipes")
        equipe_bs = c2.selectbox("Équipe Destination*", eq_list)

        res_list = active_names("resources")
        c3, c4 = st.columns(2)
        resource_bs = c3.selectbox("Ressource / Technicien*", res_list)
        destination_bs = c4.text_input("Destination / Nom du Site*", placeholder="Ex: Site Radio Casa Centre")

        st.markdown("---")
        st.markdown("##### Saisie des articles pour le BS")

        art_list = active_names("articles")
        r1, r2, r3, r4 = st.columns([3, 1.5, 2, 3])
        sel_article_bs = r1.selectbox("Article*", art_list, key="bs_art_select")
        
        art_id_bs = article_id_by_name(sel_article_bs)
        stock_dispo = current_stock(CLIENT, art_id_bs) if art_id_bs else 0
        st.caption(f"📦 Stock actuel disponible pour **{sel_article_bs}** : `{stock_dispo}` unité(s)")

        ref_article_bs = r2.text_input("Référence (Optionnel)", key="bs_ref")
        qty_article_bs = r3.number_input("Quantité*", min_value=1, value=1, step=1, key="bs_qty")
        rem_article_bs = r4.text_input("Remarque", key="bs_rem")

        if st.button("➕ Ajouter l'article à la sortie", key="add_bs_line"):
            if qty_article_bs > stock_dispo:
                st.error(f"Impossible d'ajouter : Quantité demandée ({qty_article_bs}) supérieure au stock disponible ({stock_dispo}).")
            else:
                found = False
                for item in st.session_state.temp_bs_items:
                    if item["article"] == sel_article_bs:
                        if item["quantity"] + qty_article_bs > stock_dispo:
                            st.error("La quantité totale cumulée dépasse le stock disponible !")
                            found = True
                            break
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
                        "remarque": rem_article_bs
                    })
                st.rerun()

        if st.session_state.temp_bs_items:
            st.markdown("###### Articles réservés pour la Sortie :")
            st.dataframe(pd.DataFrame(st.session_state.temp_bs_items), use_container_width=True)

            b1, b2 = st.columns(2)
            if b1.button("💾 Valider et Enregistrer le BS", key="btn_save_bs", use_container_width=True):
                if not destination_bs.strip():
                    st.error("Le champ 'Destination / Nom du Site' est obligatoire !")
                else:
                    st.session_state.confirm_bs_save = True

            if st.session_state.get("confirm_bs_save", False):
                st.warning("⚠️ Confirmez-vous l'enregistrement définitif de ce Bon de Sortie ?")
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
                            CURRENT_USER["username"]
                        )
                    )

                    for item in st.session_state.temp_bs_items:
                        art_id = article_id_by_name(item["article"])
                        execute(
                            "INSERT INTO bon_items (bon_id, article_id, reference, quantity, remarque) VALUES (?,?,?,?,?)",
                            (bon_id, art_id, item["reference"], item["quantity"], item["remarque"])
                        )
                        c_qty = current_stock(CLIENT, art_id)
                        set_stock(CLIENT, art_id, c_qty - item["quantity"])
                        add_movement(
                            client=CLIENT,
                            article_id=art_id,
                            m_type="SORTIE",
                            qty=item["quantity"],
                            ref_bon=auto_num_bs,
                            user=CURRENT_USER["username"],
                            comment=item["remarque"],
                            equipe=equipe_bs
                        )

                    st.session_state.temp_bs_items = []
                    st.session_state.confirm_bs_save = False
                    st.success(f"✅ Bon de Sortie {auto_num_bs} enregistré avec succès !")
                    st.rerun()

                if cb2.button("❌ Annuler", key="cancel_bs"):
                    st.session_state.confirm_bs_save = False
                    st.rerun()

            if b2.button("🗑️ Vider la liste", key="clear_bs", use_container_width=True):
                st.session_state.temp_bs_items = []
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.warning("🔒 Vous n'avez pas les droits nécessaires pour saisir un Bon de Sortie.")


# =========================================================
# RUBRIQUE 3 : SITUATION DU STOCK
# =========================================================
with tabs[2]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.subheader(f"📊 État du Stock en Temps Réel — Client {CLIENT}")

    query_stock = """
        SELECT a.name AS Article, COALESCE(s.quantity, 0) AS Quantite
        FROM articles a
        LEFT JOIN stock s ON a.id = s.article_id AND s.client = ?
        WHERE a.active = 1
        ORDER BY a.name
    """
    rows_stock = query(query_stock, (CLIENT,))
    df_stock = pd.DataFrame([dict(r) for r in rows_stock])

    m1, m2, m3 = st.columns(3)
    m1.metric("Nombre total de références", len(df_stock))
    m2.metric("Total Unités en Stock", int(df_stock["Quantite"].sum()) if not df_stock.empty else 0)
    m3.metric("Articles en Rupture", len(df_stock[df_stock["Quantite"] == 0]))

    st.markdown("---")
    st.dataframe(df_stock, use_container_width=True)

    buffer_excel = BytesIO()
    with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
        df_stock.to_excel(writer, sheet_name='Stock', index=False)
    
    st.download_button(
        label="📥 Télécharger l'état du Stock (Excel)",
        data=buffer_excel.getvalue(),
        file_name=f"stock_{CLIENT}_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.ms-excel"
    )
    st.markdown('</div>', unsafe_allow_html=True)


# =========================================================
# RUBRIQUE 4 : HISTORIQUE ET ÉDITION DE DOCUMENTS
# =========================================================
with tabs[3]:
    st.markdown('<div class="glass-card">', unsafe_allow_html=True)
    st.subheader("📜 Historique des Bons & Impression")

    type_filter = st.radio("Filtrer par type de bon :", ["Tous", "BE", "BS"], horizontal=True)
    
    sql_bons = "SELECT * FROM bons WHERE client=?"
    params = [CLIENT]
    if type_filter != "Tous":
        sql_bons += " AND type=?"
        params.append(type_filter)
    sql_bons += " ORDER BY id DESC"

    bons_list = query(sql_bons, tuple(params))

    if not bons_list:
        st.info("Aucun bon trouvé pour ce filtre.")
    else:
        for b in bons_list:
            with st.expander(f"📄 {b['type']} N° {b['number']} | Date : {b['date_bon']} | Par : {b['created_by']}"):
                col_info, col_actions = st.columns([2, 1])
                
                with col_info:
                    st.write(f"**Saisie le :** {b['datetime_saisie']}")
                    if b['type'] == 'BE':
                        st.write(f"**Fournisseur :** {b['fournisseur']} | **Lieu :** {b['lieu_livraison']}")
                    else:
                        st.write(f"**Équipe :** {b['equipe']} | **Ressource :** {b['resource']} | **Site :** {b['destination']}")

                with col_actions:
                    pdf_data = generate_pdf(b['id'])
                    docx_data = generate_docx(b['id'])

                    st.download_button(
                        "📄 Télécharger PDF",
                        data=pdf_data,
                        file_name=f"{b['number']}.pdf",
                        mime="application/pdf",
                        key=f"pdf_{b['id']}"
                    )
                    st.download_button(
                        "📝 Télécharger Word",
                        data=docx_data,
                        file_name=f"{b['number']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"docx_{b['id']}"
                    )

    st.markdown('</div>', unsafe_allow_html=True)


# =========================================================
# RUBRIQUE 5 : CONFIGURATION & ADMINISTRATION
# =========================================================
with tabs[4]:
    if ROLE == "admin":
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        st.subheader("⚙️ Panneau de Configuration Système")

        tab_cfg1, tab_cfg2, tab_cfg3 = st.tabs(["📦 Articles", "🚚 Fournisseurs & Équipes", "👥 Utilisateurs"])

        with tab_cfg1:
            st.markdown("##### Ajouter un nouvel article")
            with st.form("add_art_form"):
                new_art = st.text_input("Désignation de l'article")
                if st.form_submit_button("Ajouter Article"):
                    if new_art.strip():
                        execute("INSERT OR IGNORE INTO articles(name, active) VALUES(?, 1)", (new_art.strip(),))
                        st.success("Article ajouté avec succès !")
                        st.rerun()

        with tab_cfg2:
            c_f, c_e = st.columns(2)
            with c_f:
                st.markdown("##### Ajouter un Fournisseur")
                with st.form("add_fourn_form"):
                    nf = st.text_input("Nom du fournisseur")
                    if st.form_submit_button("Ajouter"):
                        if nf.strip():
                            execute("INSERT OR IGNORE INTO fournisseurs(name, active) VALUES(?, 1)", (nf.strip(),))
                            st.rerun()

            with c_e:
                st.markdown("##### Ajouter une Équipe")
                with st.form("add_eq_form"):
                    ne = st.text_input("Nom de l'équipe")
                    if st.form_submit_button("Ajouter"):
                        if ne.strip():
                            execute("INSERT OR IGNORE INTO equipes(name, active) VALUES(?, 1)", (ne.strip(),))
                            st.rerun()

        with tab_cfg3:
            st.markdown("##### Créer un nouvel utilisateur")
            with st.form("add_user_form"):
                u_username = st.text_input("Identifiant")
                u_password = st.text_input("Mot de passe", type="password")
                u_fullname = st.text_input("Nom complet")
                u_role = st.selectbox("Rôle", ROLES)
                if st.form_submit_button("Créer l'utilisateur"):
                    if u_username and u_password and u_fullname:
                        execute(
                            "INSERT OR IGNORE INTO users VALUES (?,?,?,?,?)",
                            (u_username, u_password, u_fullname, u_role, "Jamais")
                        )
                        st.success("Utilisateur créé !")
                        st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.warning("🔒 Section réservée uniquement aux administrateurs du système.")
