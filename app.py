from datetime import date, datetime
import sqlite3
import streamlit as st

# =========================================================
# CONFIGURATION DE LA PAGE STREAMLIT
# =========================================================
st.set_page_config(
    page_title="Gestion de Stock",
    page_icon="📦",
    layout="wide"
)

# =========================================================
# FONCTIONS BASE DE DONNÉES & UTILITAIRES
# =========================================================
DB_FILE = "database.db"

def init_db():
    """Initialise la table des bons si elle n'existe pas."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT,
            date_bon TEXT,
            fournisseur TEXT,
            lieu_livraison TEXT,
            equipe TEXT,
            destination TEXT
        )
    """)
    conn.commit()
    conn.close()

def execute(query, params=()):
    """Exécute une requête d'écriture (INSERT, UPDATE, DELETE)."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(query, params)
    conn.commit()
    conn.close()

def fetch_one(query, params=()):
    """Récupère une seule ligne sous forme de dictionnaire."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute(query, params)
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None

def active_names(category):
    """Retourne des listes de choix pour les selectbox."""
    if category == "fournisseurs":
        return ["Fournisseur A", "Fournisseur B", "Fournisseur C"]
    elif category == "equipes":
        return ["Équipe Alpha", "Équipe Bêta", "Équipe Gamma"]
    return []

# Initialisation de la BDD au démarrage
init_db()

# =========================================================
# INITIALISATION DES VARIABLES SIMULÉES POUR LA DÉMO
# (À adapter selon la logique de votre application)
# =========================================================
target_type = "BE"  # Exemples : "BE" (Bon d'Entrée) ou "BS" (Bon de Sortie)
selected_bon_id = 1

# Exemple de récupération des détails d'un bon depuis la BDD
bon_detail = fetch_one("SELECT * FROM bons WHERE id = ?", (selected_bon_id,))

# Si la base est vide, on fournit une valeur par défaut de secours
if not bon_detail:
    bon_detail = {
        "id": selected_bon_id,
        "date_bon": "2026-02-15",
        "fournisseur": "Fournisseur A",
        "lieu_livraison": "Dépôt Principal",
        "equipe": "Équipe Alpha",
        "destination": "Chantier Central"
    }

# =========================================================
# APPLICATION PRINCIPALE : HISTORIQUE & ÉDITION
# =========================================================
st.title("📋 Gestion des Bons")

with st.expander("✏️ Modifier les informations du Bon"):
    with st.form("form_edit_bon"):
        # 1. Parsing sécurisé de la date existante
        try:
            parsed_date = datetime.strptime(
                bon_detail["date_bon"], "%Y-%m-%d"
            ).date()
        except (ValueError, TypeError):
            parsed_date = date.today()

        # 2. Protection contre la règle value > max_value
        today = date.today()
        safe_value = min(parsed_date, today)

        mod_date = st.date_input(
            "Nouvelle Date",
            value=safe_value,
            max_value=today,
            key=f"edit_date_{selected_bon_id}",
        )

        if target_type == "BE":
            mod_fourn = st.selectbox(
                "Fournisseur",
                active_names("fournisseurs"),
                index=0,
                key=f"edit_fourn_{selected_bon_id}",
            )
            mod_lieu = st.text_input(
                "Lieu Livraison",
                value=bon_detail.get("lieu_livraison") or "",
                key=f"edit_lieu_{selected_bon_id}",
            )
        else:
            mod_eq = st.selectbox(
                "Équipe",
                active_names("equipes"),
                index=0,
                key=f"edit_eq_{selected_bon_id}",
            )
            mod_dest = st.text_input(
                "Destination",
                value=bon_detail.get("destination") or "",
                key=f"edit_dest_{selected_bon_id}",
            )

        submitted = st.form_submit_button("Valider les modifications")
        if submitted:
            # Sauvegarde temporaire des valeurs saisies dans le session state
            st.session_state.pending_edit_date = mod_date
            if target_type == "BE":
                st.session_state.pending_edit_fourn = mod_fourn
                st.session_state.pending_edit_lieu = mod_lieu
            else:
                st.session_state.pending_edit_eq = mod_eq
                st.session_state.pending_edit_dest = mod_dest
            
            st.session_state.confirm_edit_bon = True

    # Bloc de confirmation d'édition
    if st.session_state.get("confirm_edit_bon", False):
        st.warning("Confirmez-vous la modification des informations de ce bon ?")
        if st.button("✅ Oui, Confirmer Modification"):
            saved_date = st.session_state.get("pending_edit_date")
            
            if target_type == "BE":
                saved_fourn = st.session_state.get("pending_edit_fourn")
                saved_lieu = st.session_state.get("pending_edit_lieu")
                execute(
                    "UPDATE bons SET date_bon=?, fournisseur=?, lieu_livraison=? WHERE id=?",
                    (
                        str(saved_date),
                        saved_fourn,
                        saved_lieu,
                        selected_bon_id,
                    ),
                )
            else:
                saved_eq = st.session_state.get("pending_edit_eq")
                saved_dest = st.session_state.get("pending_edit_dest")
                execute(
                    "UPDATE bons SET date_bon=?, equipe=?, destination=? WHERE id=?",
                    (
                        str(saved_date),
                        saved_eq,
                        saved_dest,
                        selected_bon_id,
                    ),
                )
            
            st.session_state.confirm_edit_bon = False
            st.success("Bon mis à jour avec succès !")
            st.rerun()
