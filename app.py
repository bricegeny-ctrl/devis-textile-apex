import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import smtplib
from email.message import EmailMessage
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Gestionnaire de Devis Multi-Métiers", layout="wide")

# --- GESTION DU COMPTEUR DE DEVIS & CRM ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"
CATALOGUE_FILE = "catalogue_standardise.xlsx"

def obtenir_prochain_numero_devis():
    annee_courante = datetime.now().strftime("%Y")
    mois_courant = datetime.now().strftime("%m")
    jour_courant = datetime.now().strftime("%d")
    
    seq = 1
    if os.path.exists(COMPTEUR_FILE):
        try:
            with open(COMPTEUR_FILE, "r") as f:
                saved_data = json.load(f)
                dernier_num = saved_data.get("dernier_num", "")
                if "_" in dernier_num:
                    parts = dernier_num.split("_")
                    if len(parts) > 1 and parts[1].isdigit():
                        seq = int(parts[1]) + 1
        except Exception:
            seq = 1

    nouveau_num = f"{annee_courante}/{mois_courant}/{jour_courant}_{seq:06d}"
    
    with open(COMPTEUR_FILE, "w") as f:
        json.dump({"dernier_num": nouveau_num}, f)
        
    return nouveau_num

def charger_catalogue():
    if os.path.exists(CATALOGUE_FILE):
        try:
            return pd.read_excel(CATALOGUE_FILE)
        except Exception as e:
            st.error(f"Erreur lors de la lecture du catalogue Excel : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

df_catalogue = charger_catalogue()

# --- INITIALISATION ET CHARGEMENT DU CRM ---
if not os.path.exists(CRM_FILE):
    df_init = pd.DataFrame(columns=[
        "Numero_Devis", "Date", "Client", "Contact", "Email", "Telephone", 
        "Total_HT", "Total_TTC", "Statut", "Commercial", "Mode_Reglement", "Commentaires"
    ])
    df_init.to_csv(CRM_FILE, index=False)

def enregistrer_devis_crm(devis_data):
    if os.path.exists(CRM_FILE):
        df_crm = pd.read_csv(CRM_FILE)
    else:
        df_crm = pd.DataFrame(columns=list(devis_data.keys()))
    
    if not df_crm[df_crm["Numero_Devis"] == devis_data["Numero_Devis"]].empty:
        df_crm.loc[df_crm["Numero_Devis"] == devis_data["Numero_Devis"], :] = list(devis_data.values())
    else:
        df_nouveau = pd.DataFrame([devis_data])
        df_crm = pd.concat([df_crm, df_nouveau], ignore_index=True)
        
    df_crm.to_csv(CRM_FILE, index=False)

# --- SIDEBAR : CLIENT & LIVRAISON ---
st.sidebar.title("📋 Infos Client & Expédition")

client_nom = st.sidebar.text_input("Nom du Client / Société", "Client Exemple")
client_contact = st.sidebar.text_input("Nom de contact (ex: Jean)")
client_adresse = st.sidebar.text_area("Adresse complète du Client")
client_email = st.sidebar.text_input("Email du Client", "client@exemple.com")
client_tel = st.sidebar.text_input("Téléphone du Client")

st.sidebar.markdown("---")
st.sidebar.subheader("🚚 Logistique & Commercial")
zone_livraison = st.sidebar.selectbox("Zone de Livraison", ["France Continentale", "Corse, Monaco ou Andorre", "Espace UE", "DOM/TOM et pays hors UE"])
frais_port_manuel = st.sidebar.number_input("Frais de port HT (€)", min_value=0.0, value=0.0, step=5.0)

conseiller_choix = st.sidebar.selectbox(
    "Commercial / Conseiller",
    ["Arnaud Brice (Direction)", "Équipe Commerciale ABCOM", "Autre"]
)

mode_reglement = st.sidebar.selectbox(
    "Mode de Règlement",
    ["Virement bancaire 30 jours", "Comptant à la commande", "50% à la commande, 50% à 30 jours", "Prélèvement SEPA"]
)

delai_validite = st.sidebar.selectbox(
    "Validité du Devis",
    ["30 jours", "15 jours", "45 jours", "60 jours"]
)

st.sidebar.markdown("---")
logo_defaut_github = "https://raw.githubusercontent.com/abcom-visual/devis-app/main/logo.png"
st.sidebar.image(logo_defaut_github, width=150, caption="Logo actif (GitHub)")

# --- INTERFACE PRINCIPALE : ONGLETS ARTICLES ---
noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"]
onglets = st.tabs(noms_onglets)

articles_saisis = []

for i in range(10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        
        sans_marquage = st.checkbox(f"Vêtement / Support sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
        
        # Choix du mode de sélection
        mode_selection = st.radio(
            f"Mode de sélection pour l'Article {i+1}", 
            ["Catalogue Excel", "Saisie Manuelle / Libre"], 
            key=f"mode_sel_{i}"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if mode_selection == "Catalogue Excel" and not df_catalogue.empty:
                # Récupération des désignations uniques du catalogue Excel
                categories_dispo = df_catalogue['Categorie'].dropna().unique().tolist() if 'Categorie' in df_catalogue.columns else []
                if categories_dispo:
                    cat_choisie = st.selectbox(f"Catégorie catalogue {i+1}", categories_dispo, key=f"cat_choix_{i}")
                    df_filtre = df_catalogue[df_catalogue['Categorie'] == cat_choisie]
                else:
                    df_filtre = df_catalogue
                
                liste_designations = df_filtre['Designation'].dropna().unique().tolist() if 'Designation' in df_filtre.columns else []
                nom_article = st.selectbox(f"Sélectionner l'article du catalogue {i+1}", liste_designations, key=f"nom_cat_{i}")
                
                # Récupération automatique du prix si correspondance
                prix_defaut = 4.92
                if 'Designation' in df_catalogue.columns and 'Prix_HT' in df_catalogue.columns:
                    match_row = df_catalogue[df_catalogue['Designation'] == nom_article]
                    if not match_row.empty:
                        prix_defaut = float(match_row.iloc[0]['Prix_HT'])
            else:
                nom_article = st.text_input(f"Nom / Référence personnalisée {i+1}", value=f"T-Shirt 100% coton" if i==0 else f"Article {i+1}", key=f"nom_{i}")
                prix_defaut = 4.92

            qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_{i}")
            prix_vetement_ht = st.number_input(f"Prix unitaire HT du support (€) {i+1}", min_value=0.0, value=prix_defaut, key=f"px_vet_{i}")
            
        with col2:
            if not sans_marquage:
                st.markdown("**Gestion des Marquages / Impressions (jusqu'à 4)**")
                nb_marquages = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], key=f"nb_m_{i}")
            else:
                nb_marquages = 0
                st.info("ℹ️ Article sans marquage (fourniture/support seul).")

        marquages_config = []
        has_broderie = False
        
        if not sans_marquage:
            for m in range(nb_marquages):
                st.markdown(f"--- *Marquage / Impression {m+1}*")
                mc1, mc2 = st.columns(2)
                with mc1:
                    t_marq = st.selectbox(f"Technique / Procédé M{m+1} (Art {i+1})", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD", "Impression Numérique / Panneau", "Signalétique / Adhésif"], key=f"t_marq_{i}_{m}")
                with mc2:
                    if "DTF" in t_marq:
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_f_{i}_{m}")
                    elif "Broderie" in t_marq:
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20)", "Col/Signature (7x2)", "Casquettes / Bonnets", "Manche (8x5 cm)"], key=f"emp_b_{i}_{m}")
                    else:
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Recto seul", "Recto / Verso", "Sur-mesure / Pose spécifique"], key=f"emp_p_{i}_{m}")
                
                if t_marq == "Broderie HD":
                    has_broderie = True

                marquages_config.append({"technique": t_marq, "emplacement": emp})

        st.markdown("---")
        st.markdown("**Options & Logistique spécifiques à l'article**")
        oc1, oc2, oc3, oc4 = st.columns(4)
        with oc1:
            option_ensachage = st.checkbox(f"Ensachage individuel {i+1}", key=f"ens_{i}")
            type_sachet = st.selectbox(f"Type de sachet {i+1}", ["Sachet (T-shirt/Polo)", "Sachet (Veste/Sweat)"], key=f"tsach_{i}") if option_ensachage else ""
        with oc2:
            option_assurance = st.checkbox(f"Assurance MHC (Garantie) {i+1}", key=f"ass_{i}")
        with oc3:
            option_stockage = st.checkbox(f"Mise en stockage + picking {i+1}", key=f"stock_{i}")
        with oc4:
            remise_fidelite = st.number_input(f"Réduction fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"rem_{i}")

        if qte > 0:
            articles_saisis.append({
                "nom_article": nom_article,
                "quantite": qte,
                "prix_vet_unit": prix_vetement_ht,
                "sans_marquage": sans_marquage,
                "marquages_config": marquages_config,
                "has_broderie": has_broderie,
                "option_ensachage": option_ensachage,
                "type_sachet": type_sachet if option_ensachage else "",
                "option_assurance": option_assurance,
                "option_stockage": option_stockage,
                "remise_fidelite": remise_fidelite
            })

# --- CALCULS FINANCIERS ---
lignes_devis_global = []
frais_techniques_dossier = 0.0

for art in articles_saisis:
    q = art["quantite"]
    px_base = art["prix_vet_unit"]
    
    # Application de la remise fidélité sur le support
    px_support_remise = px_base * (1 - art["remise_fidelite"] / 100.0)
    total_support = q * px_support_remise
    
    total_marquages_ligne = 0.0
    if not art["sans_marquage"]:
        for m_cfg in art["marquages_config"]:
            tech = m_cfg["technique"]
            # Grille tarifaire indicative par marquage selon la quantité
            if "DTF" in tech:
                unit_m = 2.50 if q < 50 else (1.80 if q < 200 else 1.20)
            elif "Broderie" in tech:
                unit_m = 4.50 if q < 50 else (3.20 if q < 200 else 2.50)
            else:
                unit_m = 5.00 # Print / Signalétique
            total_marquages_ligne += q * unit_m

    # Options logistiques
    total_options = 0.0
    if art["option_ensachage"]:
        total_options += q * 0.35
    if art["option_assurance"]:
        total_options += q * 0.20
    if art["option_stockage"]:
        total_options += q * 0.50

    total_ligne_ht = total_support + total_marquages_ligne + total_options
    
    if art["has_broderie"]:
        frais_techniques_dossier += 35.0 # Frais de carte / digital

    lignes_devis_global.append({
        "nom_article": art["nom_article"],
        "quantite": q,
        "total_ligne_ht": total_ligne_ht,
        "sans_marquage": art["sans_marquage"],
        "marquages_config": art["marquages_config"]
    })

# --- ONGLET : GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.header("📊 Récapitulatif Général & Validation du Devis")
    
    if not lignes_devis_global:
        st.warning("⚠️ Veuillez renseigner au moins un article avec une quantité supérieure à 0 dans les onglets Articles.")
    else:
        sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
        montant_base_port = sous_total_articles + frais_techniques_dossier
        
        # Calcul frais de port
        frais_port = frais_port_manuel if frais_port_manuel > 0 else (15.0 if zone_livraison == "France Continentale" else 45.0)
        
        total_ht = sous_total_articles + frais_techniques_dossier + frais_port
        tva = total_ht * 0.20
        total_ttc = total_ht + tva

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**Client :** {client_nom}")
            st.markdown(f"**Contact :** {client_contact}")
            st.markdown(f"**Adresse :** {client_adresse}")
        with col_b:
            numero_devis_genere = obtenir_prochain_numero_devis()
            st.markdown(f"**N° de Devis :** `{numero_devis_genere}`")
            st.markdown(f"**Date :** {datetime.now().strftime('%d/%m/%Y')}")
            st.markdown(f"**Commercial :** {conseiller_choix}")

        st.markdown("---")
        st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
        if frais_techniques_dossier > 0:
            st.write(f"**Frais techniques / Broderie HT :** {frais_techniques_dossier:.2f} €")
        st.write(f"**Frais de port ({zone_livraison}) HT :** {frais_port:.2f} €")
        st.markdown(f"### **Total Général HT : {total_ht:.2f} €**")
        st.markdown(f"### **Total Général TTC (20%) : {total_ttc:.2f} €**")

        if st.button("💾 Enregistrer le devis dans le CRM"):
            devis_record = {
                "Numero_Devis": numero_devis_genere,
                "Date": datetime.now().strftime('%Y-%m-%d'),
                "Client": client_nom,
                "Contact": client_contact,
                "Email": client_email,
                "Telephone": client_tel,
                "Total_HT": round(total_ht, 2),
                "Total_TTC": round(total_ttc, 2),
                "Statut": "En attente",
                "Commercial": conseiller_choix,
                "Mode_Reglement": mode_reglement,
                "Commentaires": "Généré via l'application multi-métiers"
            }
            enregistrer_devis_crm(devis_record)
            st.success(f"✨ Devis {numero_devis_genere} enregistré avec succès dans le CRM !")

# --- ONGLET : SUIVI CRM ---
with onglets[11]:
    st.header("📈 Suivi CRM & Historique des Devis")
    if os.path.exists(CRM_FILE):
        df_crm = pd.read_csv(CRM_FILE)
        if not df_crm.empty:
            st.dataframe(df_crm, use_container_width=True)
            
            devis_selectionne = st.selectbox("Sélectionner un devis pour action", df_crm["Numero_Devis"].tolist(), key="select_devis_action")
            nouveau_statut = st.selectbox("Modifier le statut", ["En attente", "Validé", "Refusé", "Facturé"], key="nouveau_statut_crm")
            
            if st.button("Mettre à jour le statut"):
                df_crm.loc[df_crm["Numero_Devis"] == devis_selectionne, "Statut"] = nouveau_statut
                df_crm.to_csv(CRM_FILE, index=False)
                st.success("Statut mis à jour avec succès !")
                st.rerun()
        else:
            st.info("Aucun devis enregistré pour le moment dans le CRM.")
    else:
        st.info("Fichier CRM vide.")