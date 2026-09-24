import pandas as pd
import streamlit as st

# Configuration de la page
st.set_page_config(
    page_title="Gestionnaire de Devis - Print & Signalétique",
    page_icon="🖨️",
    layout="wide",
)

# Sidebar : Informations Client & Expédition
st.sidebar.markdown("### 📋 Infos Client & Expédition")
nom_entreprise = st.sidebar.text_input("Nom de l'Entreprise / Société")
nom_contact = st.sidebar.text_input(
    "Nom de la personne contact (ex: Jean Dupont)"
)
adresse_client = st.sidebar.text_area("Adresse complète du Client")
siret_client = st.sidebar.text_input("SIRET du Client (optionnel)")
telephone_client = st.sidebar.text_input("Téléphone du Client")
email_client = st.sidebar.text_input("E-mail du Client")

# En-tête principal
st.markdown(
    "<h1>🖨️ Gestionnaire de Devis - Print & Signalétique</h1>",
    unsafe_allow_html=True,
)

# Navigation par onglets
onglets = [
    "Support 1",
    "Support 2",
    "Support 3",
    "Support 4",
    "Support 5",
    "Support 6",
    "Support 7",
    "Support 8",
    "Print 1",
    "Print 2",
    "Général & Devis",
    "Suivi CRM",
]
onglets_actifs = st.tabs(onglets)

# Données par défaut pour le catalogue standardisé (Print & Signalétique)
catalogue_defaut = pd.DataFrame(
    {
        "Référence": [
            "FLY-A5",
            "FLY-A4",
            "AFFICHE-A1",
            "BACHE-3X1",
            "KAKEMONO-85",
        ],
        "Désignation": [
            "Flyer A5 135g Couché Brillant",
            "Flyer A4 135g Couché Brillant",
            "Affiche A1 Papier Backlit",
            "Bâche PVC 500g M1 avec œillets",
            "Roll-up / Kakemono 85x200 cm",
        ],
        "Catégorie": [
            "Print",
            "Print",
            "Signalétique",
            "Signalétique",
            "Signalétique",
        ],
        "Prix unitaire HT (€)": [0.05, 0.09, 12.50, 45.00, 35.00],
    }
)

# Exemple de gestion pour le premier onglet (Print 1 / Article 1)
with onglets_actifs[8]:  # Onglet Print 1
    st.markdown("### 📄 Print & Signalétique - Article 1")

    # Simulation ou vérification du catalogue
    catalogue = catalogue_defaut.copy()

    if catalogue.empty:
        st.warning(
            "⚠️ Catalogue standardisé non disponible ou vide.", icon="⚠️"
        )
    else:
        st.success(
            "✅ Catalogue standardisé chargé avec succès.", icon="✅"
        )

        # Affichage du catalogue interactif
        article_selectionne = st.selectbox(
            "Sélectionnez un article du catalogue",
            catalogue["Désignation"].tolist(),
        )
        quantite = st.number_input(
            "Quantité", min_value=1, value=100, step=10
        )

        # Récupération du prix unitaire correspondant
        prix_unitaire = catalogue.loc[
            catalogue["Désignation"] == article_selectionne,
            "Prix unitaire HT (€)",
        ].values[0]
        total_ht = quantite * prix_unitaire

        st.info(
            f"**Prix unitaire :** {prix_unitaire:.2f} € HT | **Total HT estimé :** {total_ht:.2f} € HT"
        )

# Gestion des autres onglets (structure de base)
for i, tab in enumerate(onglets_actifs):
    if i != 8:
        with tab:
            st.markdown(f"### Configuration - {onglets[i]}")
            st.write(
                "Paramétrez ici les éléments de cet onglet selon vos besoins de production."
            )