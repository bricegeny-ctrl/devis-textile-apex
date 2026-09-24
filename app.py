from io import StringIO
import pandas as pd
import streamlit as st

# Configuration de la page
st.set_page_config(
    page_title="Gestionnaire de Devis - Textile & Print",
    page_icon="🖨️",
    layout="wide",
)

# --- DONNÉES DU CATALOGUE INTÉGRÉ (PRINT & SIGNALÉTIQUE) ---
catalogue_csv = """Categorie,Reference,Designation,Quantite,Prix_HT
Blocs notes,BNA6-25,Bloc Note collé - Format A6 - 25 Feuilles - 90 Gr Offset,50,1.176
Blocs notes,BNA6-25,Bloc Note collé - Format A6 - 25 Feuilles - 90 Gr Offset,100,0.735
Blocs notes,BNA6-25,Bloc Note collé - Format A6 - 25 Feuilles - 90 Gr Offset,200,0.609
Blocs notes,BNA6-25,Bloc Note collé - Format A6 - 25 Feuilles - 90 Gr Offset,500,0.567
Blocs notes,BNA6-25,Bloc Note collé - Format A6 - 25 Feuilles - 90 Gr Offset,1000,0.567
Blocs notes,BNA6-50,Bloc Note collé - Format A6 - 50 Feuilles - 90 Gr Offset,50,1.638
Blocs notes,BNA6-50,Bloc Note collé - Format A6 - 50 Feuilles - 90 Gr Offset,100,1.113
Blocs notes,BNA6-50,Bloc Note collé - Format A6 - 50 Feuilles - 90 Gr Offset,200,1.071
Blocs notes,BNA6-50,Bloc Note collé - Format A6 - 50 Feuilles - 90 Gr Offset,500,0.987
Blocs notes,BNA6-50,Bloc Note collé - Format A6 - 50 Feuilles - 90 Gr Offset,1000,0.924
Blocs notes,BNA5-25,Bloc Note collé - Format A5 - 25 Feuilles - 90 Gr Offset,50,2.352
Blocs notes,BNA5-25,Bloc Note collé - Format A5 - 25 Feuilles - 90 Gr Offset,100,1.323
Blocs notes,BNA5-25,Bloc Note collé - Format A5 - 25 Feuilles - 90 Gr Offset,200,1.176
Blocs notes,BNA5-25,Bloc Note collé - Format A5 - 25 Feuilles - 90 Gr Offset,500,1.071
Blocs notes,BNA5-25,Bloc Note collé - Format A5 - 25 Feuilles - 90 Gr Offset,1000,1.050
Chemises de présentation,CHEM3R,Chemise de présentation 3 rabats - 350g Couché Mat - Pelliculage Recto Brillant ou Mat,100,2.45
Chemises de présentation,CHEM3R,Chemise de présentation 3 rabats - 350g Couché Mat - Pelliculage Recto Brillant ou Mat,250,1.85
Flyers A5,FLA5-135,Flyer A5 - 135g Couché Brillant,500,0.06
Flyers A5,FLA5-135,Flyer A5 - 135g Couché Brillant,1000,0.04
Roll-Ups,ROLL85,Roll-Up / Kakemono 85x200 cm Structure aluminium,1,35.00
Roll-Ups,ROLL85,Roll-Up / Kakemono 85x200 cm Structure aluminium,5,30.00
Banderoles,BAND3X1,Banderole Bâche PVC 500g M1 avec œillets - 3x1m,1,45.00
Panneaux de chantier,PANCH3,Panneaux de chantier Akilux 3.5mm - 60x80cm,10,12.00
"""

# Chargement du DataFrame depuis la chaîne intégrée
df_catalogue = pd.read_csv(StringIO(catalogue_csv))

# --- SIDEBAR : INFOS CLIENT & EXPÉDITION ---
st.sidebar.markdown("### 📋 Infos Client & Expédition")
nom_entreprise = st.sidebar.text_input("Nom de l'Entreprise / Société")
nom_contact = st.sidebar.text_input(
    "Nom de la personne contact (ex: Jean Dupont)"
)
adresse_client = st.sidebar.text_area("Adresse complète du Client")
siret_client = st.sidebar.text_input("SIRET du Client (optionnel)")
telephone_client = st.sidebar.text_input("Téléphone du Client")
email_client = st.sidebar.text_input("E-mail du Client")

# --- EN-TÊTE PRINCIPAL ---
st.markdown(
    "<h1>🖨️ Gestionnaire de Devis - Textile & Solutions Print</h1>",
    unsafe_allow_html=True,
)

# --- NAVIGATION PAR ONGLETS ---
onglets = [
    "Textile 1",
    "Textile 2",
    "Textile 3",
    "Textile 4",
    "Textile 5",
    "Textile 6",
    "Textile 7",
    "Textile 8",
    "Print et Signalétique",
    "Général & Devis",
    "Suivi CRM",
]
onglets_actifs = st.tabs(onglets)

# Gestion des 8 onglets Textiles
for i in range(8):
    with onglets_actifs[i]:
        st.markdown(f"### Configuration - Textile / Article {i+1}")
        st.write(
            "Paramétrez ici les spécificités et marquages textiles habituels."
        )

# --- ONGLET DÉDIÉ : PRINT ET SIGNALÉTIQUE ---
with onglets_actifs[8]:
    st.markdown("### 📄 Catalogue Print et Signalétique")
    st.write(
        "Sélectionnez un ou plusieurs produits dans le catalogue ci-dessous, ou ajoutez un produit personnalisé à tarif libre."
    )

    # 1. Sélection multiple du catalogue intégré
    df_catalogue["Label_Complet"] = (
        "["
        + df_catalogue["Categorie"]
        + "] "
        + df_catalogue["Designation"]
        + " (Qté: "
        + df_catalogue["Quantite"].astype(str)
        + ") - "
        + df_catalogue["Prix_HT"].astype(str)
        + " € HT"
    )

    produits_selectionnes = st.multiselect(
        "Sélectionnez vos produits (choix multiples possibles)",
        options=df_catalogue["Label_Complet"].tolist(),
    )

    total_print_catalogue = 0.0

    if produits_selectionnes:
        st.markdown("#### Détail des articles sélectionnés :")
        for prod in produits_selectionnes:
            ligne_prod = df_catalogue[
                df_catalogue["Label_Complet"] == prod
            ].iloc[0]
            col_a, col_b, col_c = st.columns([3, 1, 1])
            with col_a:
                st.write(f"• **{ligne_prod['Designation']}**")
            with col_b:
                qte_choisie = st.number_input(
                    f"Qté ({ligne_prod['Reference']})",
                    min_value=1,
                    value=int(ligne_prod["Quantite"]),
                    key=f"qte_{ligne_prod['Reference']}_{ligne_prod['Quantite']}",
                )
            with col_c:
                prix_u = float(ligne_prod["Prix_HT"])
                st.write(f"Prix U: {prix_u:.2f} € HT")
            total_print_catalogue += qte_choisie * prix_u

    st.markdown("---")
    st.markdown("#### ✍️ Ajout Libre / Autre Article")

    # Champ autre avec tarif libre dans un formulaire dédié
    with st.form(key="form_autre_article"):
        col_l1, col_l2, col_l3 = st.columns([3, 1, 1])
        with col_l1:
            desc_libre = st.text_input(
                "Description du produit hors catalogue"
            )
        with col_l2:
            qte_libre = st.number_input(
                "Quantité", min_value=1, value=1, step=1
            )
        with col_l3:
            prix_libre = st.number_input(
                "Prix unitaire HT (€)",
                min_value=0.0,
                value=0.0,
                step=0.5,
                format="%.2f",
            )

        submit_autre = st.form_submit_button("Ajouter cet article libre")

    total_libre = 0.0
    if "articles_libres" not in st.session_state:
        st.session_state["articles_libres"] = []

    if submit_autre and desc_libre:
        st.session_state["articles_libres"].append(
            {
                "description": desc_libre,
                "quantite": qte_libre,
                "prix_unitaire": prix_libre,
            }
        )
        st.success(f"Article '{desc_libre}' ajouté avec succès !")

    # Affichage et gestion des articles libres enregistrés
    if st.session_state["articles_libres"]:
        st.markdown("**Articles libres ajoutés au devis :**")
        for idx, art in enumerate(st.session_state["articles_libres"]):
            tot_art = art["quantite"] * art["prix_unitaire"]
            total_libre += tot_art
            col_al1, col_al2 = st.columns([4, 1])
            with col_al1:
                st.write(
                    f"- {art['description']} | Qté : {art['quantite']} x {art['prix_unitaire']:.2f} € HT = **{tot_art:.2f} € HT**"
                )
            with col_al2:
                if st.button("Supprimer", key=f"del_libre_{idx}"):
                    st.session_state["articles_libres"].pop(idx)
                    st.rerun()

    total_general_print = total_print_catalogue + total_libre
    st.markdown("### Récapitulatif Print & Signalétique")
    st.info(
        f"**Total Général Print & Signalétique HT :** {total_general_print:.2f} € HT"
    )

# Onglets finaux (Général & Devis / Suivi CRM)
with onglets_actifs[9]:
    st.markdown("### 📊 Général & Devis")
    st.write(
        "Synthèse globale des devis textiles et print, conditions de règlement et génération du document final."
    )

with onglets_actifs[10]:
    st.markdown("### 📇 Suivi CRM")
    st.write("Historique des contacts et relances clients.")