from datetime import date, datetime
import streamlit as st

# =========================================================
# EXTRAIT CORRIGÉ : DANS L'ONGLET HISTORIQUE
# =========================================================

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
                value=bon_detail["lieu_livraison"] or "",
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
                value=bon_detail["destination"] or "",
                key=f"edit_dest_{selected_bon_id}",
            )

        submitted = st.form_submit_button("Valider les modifications")
        if submitted:
            # Stockage des valeurs saisies dans la session pour l'étape de confirmation
            st.session_state.pending_edit_date = mod_date
            if target_type == "BE":
                st.session_state.pending_edit_fourn = mod_fourn
                st.session_state.pending_edit_lieu = mod_lieu
            else:
                st.session_state.pending_edit_eq = mod_eq
                st.session_state.pending_edit_dest = mod_dest
            
            st.session_state.confirm_edit_bon = True

    if st.session_state.get("confirm_edit_bon", False):
        st.warning(
            "Confirmez-vous la modification des informations de ce bon ?"
        )
        if st.button("✅ Oui, Confirmer Modification"):
            # Récupération des valeurs enregistrées
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
            
            # Nettoyage de la session
            st.session_state.confirm_edit_bon = False
            st.success("Bon mis à jour !")
            st.rerun()
