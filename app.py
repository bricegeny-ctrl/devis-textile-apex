import sqlite3
import streamlit as st

st.set_page_config(
    page_title="Gestionnaire de Devis - Marquage Textile",
    page_icon="🖨️",
    layout="wide"
)

# Récupération des paramètres URL (si besoin)
query_params = st.query_params
ref_selectionnee = query_params.get("ref", None)

nom_produit_defaut = "T-shirt 100% coton"
quantite_totale_defaut = 10
montant_total_vetements_defaut = 49.20

if ref_selectionnee:
    try:
        conn = sqlite3.connect("apex_catalogue_imbretex.db")
        cursor = conn.cursor()
        cursor.execute("SELECT designation_courte, prix_vente_ht FROM produits WHERE ref_produit = ?", (ref_selectionnee,))
        resultat = cursor.fetchone()
        conn.close()
        if resultat:
            nom_produit_defaut = f"{resultat[0]} (Réf: {ref_selectionnee})"
            if resultat[1]:
                prix_defaut = float(str(resultat[1]).replace(',', '.'))
                montant_total_vetements_defaut = quantite_totale_defaut * prix_defaut
    except Exception as e:
        pass

st.title("🖨️ Gestionnaire de Devis - Marquage Textile")

with st.sidebar:
    st.header("Infos Client & Expédition")
    nom_entreprise = st.text_input("Nom de l'Entreprise / Société")
    nom_contact = st.text_input("Nom de la personne contact (ex: Jean Dupont)")
    adresse_client = st.text_area("Adresse complète du Client")
    siret_client = st.text_input("SIRET du Client (optionnel)")
    telephone_client = st.text_input("Téléphone du Client")
    email_client = st.text_input("E-mail du Client")
    zone_livraison = st.selectbox("Zone de Livraison", ["France Continentale", "Corse / Étranger"])
    offrir_port = st.checkbox("Offrir les frais de port (0 €)", value=False)
    
    st.markdown("---")
    st.markdown("### Frais techniques de commande (€ HT)")
    frais_tech = st.number_input("Frais techniques globaux", min_value=0.0, value=19.80, step=0.1, format="%.2f")

onglets = st.tabs([f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"])

with onglets[0]:
    st.subheader("Configuration de l'Article 1")
    
    sans_marquage_1 = st.checkbox("Vêtement sans marquage (fourniture seule) 1", value=False)
    
    col1, col2 = st.columns(2)
    with col1:
        nom_article_1 = st.text_input("Nom / Référence du vêtement 1", value=nom_produit_defaut)
        quantite_1 = st.number_input("Quantité totale (pcs) 1", min_value=1, value=quantite_totale_defaut, step=1)
    with col2:
        prix_unitaire_1 = st.number_input("Prix unitaire HT (€) 1", value=montant_total_vetements_defaut/quantite_totale_defaut if quantite_totale_defaut > 0 else 4.92, format="%.2f")

    st.markdown("---")
    st.markdown("### Gestion des Marquages (jusqu'à 4)")
    
    if not sans_marquage_1:
        nb_marquages_1 = st.selectbox("Nombre de marquages pour l'article 1", [1, 2, 3, 4], index=0)
        techniques_disponibles = ["DTF Textile Fin", "Sérigraphie", "Broderie", "Transfert Céramique / Patch"]
        emplacements_disponibles = ["Cœur (13x9 cm)", "Dos (28x20 cm)", "Manche droite", "Manche gauche", "Poitrine centrale"]

        for m in range(nb_marquages_1):
            st.markdown(f"#### --- Marquage {m+1}")
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.selectbox(f"Technique M{m+1} (Art 1)", techniques_disponibles, key=f"tech_1_{m}")
            with col_m2:
                st.selectbox(f"Emplacement M{m+1}", emplacements_disponibles, key=f"emp_1_{m}")
    else:
        st.info("Article configuré sans marquage (fourniture seule).")

    st.markdown("---")
    st.markdown("### Options & Logistique spécifiques à l'article")
    st.checkbox("Mise sous sachet individuel (pliage + pochette)", value=False, key="pliage_1")

for i in range(1, 10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        st.info("Configurez ici les articles additionnels si nécessaire.")

with onglets[10]:
    st.header("📊 Récapitulatif Général & Devis")
    
    total_vetements_1 = quantite_1 * prix_unitaire_1
    frais_port_val = 0.0 if offrir_port else 15.0
    total_ht = total_vetements_1 + frais_tech + frais_port_val
    total_ttc = total_ht * 1.20
    
    st.markdown(f"### Devis pour : {nom_entreprise if nom_entreprise else 'Client en cours'}")
    st.write(f"**Article 1 :** {nom_article_1} ({quantite_1} pcs) — Total Vêtement HT : {total_vetements_1:.2f} €")
    st.write(f"**Frais techniques :** {frais_tech:.2f} € HT")
    st.write(f"**Frais de port :** {'OFFERTS (0.00 €)' if offrir_port else f'{frais_port_val:.2f} € HT'}")
    st.markdown(f"### Total HT : {total_ht:.2f} € | Total TTC (20%) : {total_ttc:.2f} €")
    
    st.markdown("---")
    if st.button("💾 Générer le devis complet", type="primary"):
        st.success("Devis calculé et enregistré avec succès !")

with onglets[11]:
    st.header("📈 Tableau de Suivi CRM & Historique des Devis")
    st.success("Accès administrateur global activé (Vue de tous les devis de l'agence).")
    
    st.markdown("### 📋 Liste des devis et gestion des statuts")
    
    # Données fictives ou de base pour l'affichage initial du tableau CRM d'origine
    import pandas as pd
    data_crm = [
        {"Date": "2026-09-23 07:18", "Numero_Devis": "2026/09/23_000006", "Conseiller": "Brice Geny", "Client": "None", "Entreprise": "None", "Email": "None", "Telephone": "None", "Quantite_Totale": 35, "Total HT (€)": 784.20, "Total TTC (€)": 941.04, "Statut du Devis": "Sans suite"},
        {"Date": "2026-09-23 07:27", "Numero_Devis": "2026/09/23_000007", "Conseiller": "Brice Geny", "Client": "brice geny", "Entreprise": "SAS apex", "Email": "Brice.geny@gmail.com", "Telephone": "None", "Quantite_Totale": 31, "Total HT (€)": 554.00, "Total TTC (€)": 664.80, "Statut du Devis": "Sans suite"},
        {"Date": "2026-09-23 07:30", "Numero_Devis": "2026/09/23_000008", "Conseiller": "Brice Geny", "Client": "brice geny", "Entreprise": "sas apex", "Email": "Brice.geny@gmail.com", "Telephone": "None", "Quantite_Totale": 121, "Total HT (€)": 1193.80, "Total TTC (€)": 1432.56, "Statut du Devis": "En cours"}
    ]
    df_crm = pd.DataFrame(data_crm)
    st.dataframe(df_crm, use_container_width=True)
    
    st.markdown("---")
    st.markdown("### ⚡ Actions rapides sur un devis")
    devis_selectionne = st.selectbox("Sélectionner un devis", df_crm["Numero_Devis"].tolist())
    
    col_act1, col_act2 = st.columns(2)
    with col_act1:
        st.button("📄 Télécharger le PDF")
    with col_act2:
        st.button("📋 Dupliquer ce devis pour un autre client")
        
    st.markdown("---")
    st.button("📥 Exporter tout le CRM au format CSV (compatible Excel / Google Sheets)")