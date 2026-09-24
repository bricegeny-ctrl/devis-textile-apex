import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import urllib.parse

st.set_page_config(page_title="Gestionnaire de Devis - Multi-Métiers", layout="wide")

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

def obtenir_prix_catalogue(df, designation, qte):
    if df.empty or 'Designation' not in df.columns or 'Prix_HT' not in df.columns:
        return 0.0
    
    df_art = df[df['Designation'] == designation]
    if df_art.empty:
        return 0.0
        
    if 'Quantite' in df_art.columns:
        quantites_dispo = sorted(df_art['Quantite'].dropna().unique().tolist())
        if quantites_dispo:
            q_choisie = quantites_dispo[0]
            for q in quantites_dispo:
                if q <= qte:
                    q_choisie = q
                else:
                    break
            match_row = df_art[df_art['Quantite'] == q_choisie]
            if not match_row.empty:
                return float(match_row.iloc[0]['Prix_HT'])
                
    return float(df_art.iloc[0]['Prix_HT'])

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
    ["Brice Geny", "Brice Bugna", "Autre"]
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
        
        metier_type = st.radio(
            f"Univers / Métier pour l'Article {i+1}",
            ["👕 Textile & Marquage (DTF / Broderie)", "📄 Print, Papeterie & Signalétique (Catalogue Excel)"],
            key=f"metier_{i}"
        )
        
        st.markdown("---")
        
        if "Textile" in metier_type:
            sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
            
            col1, col2 = st.columns(2)
            with col1:
                nom_article = st.text_input(f"Référence / Nom du vêtement {i+1}", value="T-Shirt 100% coton bio" if i==0 else f"Vêtement {i+1}", key=f"nom_textile_{i}")
                qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_textile_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT du support textile (€) {i+1}", min_value=0.0, value=4.92, format="%.2f", key=f"px_textile_{i}")
                
            with col2:
                if not sans_marquage:
                    st.markdown("**Gestion des Marquages Textiles**")
                    nb_marquages = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], key=f"nb_m_textile_{i}")
                else:
                    nb_marquages = 0
                    st.info("ℹ️ Vêtement sans marquage (fourniture seule).")

            marquages_config = []
            has_broderie = False
            
            if not sans_marquage:
                for m in range(nb_marquages):
                    st.markdown(f"--- *Marquage Textile {m+1}*")
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_textile_{i}_{m}")
                    with mc2:
                        if "Broderie" in t_marq:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20)", "Col/Signature (7x2)", "Casquettes / Bonnets", "Manche (8x5 cm)"], key=f"emp_b_{i}_{m}")
                        else:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_f_{i}_{m}")
                    
                    if t_marq == "Broderie HD":
                        has_broderie = True

                    marquages_config.append({"technique": t_marq, "emplacement": emp})

            st.markdown("---")
            st.markdown("**Options logistiques textile**")
            oc1, oc2, oc3, oc4 = st.columns(4)
            with oc1:
                option_ensachage = st.checkbox(f"Ensachage individuel {i+1}", key=f"ens_{i}")
            with oc2:
                option_assurance = st.checkbox(f"Assurance MHC {i+1}", key=f"ass_{i}")
            with oc3:
                option_stockage = st.checkbox(f"Stockage + picking {i+1}", key=f"stock_{i}")
            with oc4:
                remise_fidelite = st.number_input(f"Remise fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"rem_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "textile",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": sans_marquage,
                    "marquages_config": marquages_config,
                    "has_broderie": has_broderie,
                    "option_ensachage": option_ensachage,
                    "remise_fidelite": remise_fidelite
                })

        else:
            col1, col2 = st.columns(2)
            with col1:
                qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=100 if i==0 else 0, key=f"qte_print_{i}")

                if not df_catalogue.empty:
                    categories_dispo = df_catalogue['Categorie'].dropna().unique().tolist() if 'Categorie' in df_catalogue.columns else []
                    cat_choisie = st.selectbox(f"Catégorie Print / Signalétique {i+1}", categories_dispo, key=f"cat_print_{i}")
                    
                    df_filtre = df_catalogue[df_catalogue['Categorie'] == cat_choisie]
                    liste_designations = df_filtre['Designation'].dropna().unique().tolist() if 'Designation' in df_filtre.columns else []
                    nom_article = st.selectbox(f"Article du catalogue {i+1}", liste_designations, key=f"nom_print_{i}")
                    
                    prix_defaut = obtenir_prix_catalogue(df_catalogue, nom_article, qte)
                else:
                    nom_article = st.text_input(f"Désignation article print {i+1}", value="Flyer A5", key=f"nom_print_libre_{i}")
                    prix_defaut = 0.10

                prix_vetement_ht = st.number_input(f"Prix unitaire HT (€) {i+1}", min_value=0.0, value=float(prix_defaut), format="%.3f", key=f"px_print_{i}")

            with col2:
                st.info("ℹ️ Article Print / Signalétique géré via le catalogue tarifaire officiel.")
                remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"rem_print_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "print",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": True,
                    "marquages_config": [],
                    "has_broderie": False,
                    "option_ensachage": False,
                    "remise_fidelite": remise_fidelite
                })

# --- CALCULS FINANCIERS ---
lignes_devis_global = []
frais_techniques_dossier = 0.0

for art in articles_saisis:
    q = art["quantite"]
    px_base = art["prix_vet_unit"]
    
    px_support_remise = px_base * (1 - art["remise_fidelite"] / 100.0)
    total_support = q * px_support_remise
    
    total_marquages_ligne = 0.0
    if art["type_univers"] == "textile" and not art["sans_marquage"]:
        for m_cfg in art["marquages_config"]:
            tech = m_cfg["technique"]
            if "DTF" in tech:
                unit_m = 2.50 if q < 50 else (1.80 if q < 200 else 1.20)
            elif "Broderie" in tech:
                unit_m = 4.50 if q < 50 else (3.20 if q < 200 else 2.50)
            else:
                unit_m = 2.00
            total_marquages_ligne += q * unit_m

    total_options = 0.0
    if art.get("option_ensachage", False):
        total_options += q * 0.35

    total_ligne_ht = total_support + total_marquages_ligne + total_options
    
    if art.get("has_broderie", False):
        frais_techniques_dossier += 35.0

    lignes_devis_global.append({
        "nom_article": art["nom_article"],
        "quantite": q,
        "total_ligne_ht": total_ligne_ht,
        "type_univers": art["type_univers"]
    })

# --- ONGLET : GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.header("📊 Récapitulatif Général & Validation du Devis")
    
    if not lignes_devis_global:
        st.warning("⚠️ Veuillez renseigner au moins un article avec une quantité supérieure à 0 dans les onglets Articles.")
    else:
        sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
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

        # --- SECTION BOUTONS D'ACTION ET ENVOI MAIL ---
        st.markdown("---")
        st.subheader("📤 Actions & Envoi du Devis")

        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
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

        with col_btn2:
            sujet_mail = f"Devis {numero_devis_genere}"
            corps_mail = f"""Bonjour {client_contact or client_nom},

Veuillez trouver ci-joint votre devis n° {numero_devis_genere} d'un montant total de {total_ttc:.2f} € TTC.

Conditions de règlement : {mode_reglement}
Validité de l'offre : {delai_validite}

Restant à votre disposition pour tout renseignement complémentaire.

Bien cordialement,
{conseiller_choix}
"""
            mailto_url = f"mailto:{client_email}?subject={urllib.parse.quote(sujet_mail)}&body={urllib.parse.quote(corps_mail)}"
            
            st.markdown(
                f"""
                <a href="{mailto_url}" target="_blank">
                    <button style="width:100%; background-color:#ff4b4b; color:white; border:none; padding:10px 20px; border-radius:5px; font-weight:bold; cursor:pointer;">
                        ✉️ Envoyer le devis par e-mail (Ouvrir le client mail)
                    </button>
                </a>
                """,
                unsafe_allow_html=True
            )

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