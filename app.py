# =========================================================
# EXTRAIT CORRIGÉ : DANS L'ONGLET HISTORIQUE (REMPLACER LE FORMULAIRE D'ÉDITION)
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
