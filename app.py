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

st.set_page_config(page_title="Gestionnaire de Devis - Marquage Textile", layout="wide")

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

def enregistrer_dans_crm(data_devis):
    df_new = pd.DataFrame([data_devis])
    if os.path.exists(CRM_FILE):
        try:
            df_exist = pd.read_csv(CRM_FILE)
            df_final = pd.concat([df_exist, df_new], ignore_index=True)
        except Exception:
            df_final = df_new
    else:
        df_final = df_new
    df_final.to_csv(CRM_FILE, index=False)

def mettre_a_jour_statut_crm(numero_devis, nouveau_statut):
    if os.path.exists(CRM_FILE):
        df = pd.read_csv(CRM_FILE)
        if "Numero_Devis" in df.columns:
            df.loc[df["Numero_Devis"] == numero_devis, "Statut"] = nouveau_statut
            df.to_csv(CRM_FILE, index=False)

# --- CHARGEMENT DU CATALOGUE EXCEL ---
@st.cache_data
def charger_catalogue():
    if os.path.exists(CATALOGUE_FILE):
        try:
            df_cat = pd.read_excel(CATALOGUE_FILE)
            return df_cat
        except Exception as e:
            st.error(f"Erreur lors du chargement du catalogue Excel : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

df_catalogue = charger_catalogue()

# --- FONCTIONS DE CALCUL DES TARIFS & TRANCHES ---
def get_tarif_dtf_fin(qte_globale, emplacement):
    if qte_globale <= 5: idx = 0
    elif qte_globale <= 9: idx = 1
    elif qte_globale <= 19: idx = 2
    elif qte_globale <= 29: idx = 3
    elif qte_globale <= 39: idx = 4
    elif qte_globale <= 49: idx = 5
    elif qte_globale <= 99: idx = 6
    elif qte_globale <= 249: idx = 7
    elif qte_globale <= 499: idx = 8
    elif qte_globale <= 999: idx = 9
    elif qte_globale <= 4999: idx = 10
    else: idx = 11

    tarifs = {
        "Cœur (13x9 cm)": [6.00, 4.50, 3.60, 2.81, 2.50, 2.40, 2.00, 1.80, 1.60, 1.40, 1.20, 0.70],
        "Dos D10 (20x13 cm)": [7.92, 6.50, 5.50, 5.00, 4.20, 4.00, 3.50, 3.00, 2.70, 2.50, 2.00, 1.00],
        "Dos D20 (28x20 cm)": [13.67, 10.00, 9.00, 6.90, 6.30, 6.00, 5.50, 4.60, 4.00, 3.50, 2.50, 1.50],
        "Format P (37x27 cm)": [17.00, 13.00, 12.00, 10.00, 9.00, 8.00, 7.50, 6.50, 6.00, 5.00, 4.00, 2.50],
        "Manche (9x8 cm)": [7.20, 5.40, 4.32, 3.37, 3.00, 2.88, 2.40, 2.16, 1.92, 1.68, 1.44, 0.84],
        "+ Perso. Nom": [4.39, 3.50, 2.50, 2.20, 1.90, 1.80, 1.70, 1.50, 1.30, 0.80, 0.30, 0.20]
    }
    return tarifs.get(emplacement, [0]*12)[idx]

def get_tarif_dtf_epais(qte_globale, emplacement):
    if qte_globale <= 5: idx = 0
    elif qte_globale <= 9: idx = 1
    elif qte_globale <= 19: idx = 2
    elif qte_globale <= 29: idx = 3
    elif qte_globale <= 39: idx = 4
    elif qte_globale <= 49: idx = 5
    elif qte_globale <= 99: idx = 6
    elif qte_globale <= 249: idx = 7
    elif qte_globale <= 499: idx = 8
    elif qte_globale <= 999: idx = 9
    elif qte_globale <= 4999: idx = 10
    else: idx = 11

    tarifs = {
        "Cœur (13x9 cm)": [6.60, 4.95, 3.96, 3.09, 2.75, 2.64, 2.20, 1.98, 1.76, 1.54, 1.32, 0.77],
        "Dos D10 (20x13 cm)": [8.71, 7.15, 6.05, 5.50, 4.62, 4.40, 3.85, 3.30, 2.97, 2.75, 2.20, 1.10],
        "Dos D20 (28x20 cm)": [15.04, 11.00, 9.90, 7.59, 6.93, 6.60, 6.05, 5.06, 4.40, 3.85, 2.75, 1.65],
        "Format P (37x27 cm)": [18.70, 14.30, 13.20, 11.00, 9.90, 8.80, 8.25, 7.15, 6.60, 5.50, 4.40, 2.75],
        "Manche (9x8 cm)": [7.92, 5.94, 4.75, 3.71, 3.30, 3.17, 2.64, 2.38, 2.11, 1.85, 1.58, 0.92],
        "+ Perso. Nom": [4.83, 3.85, 2.75, 2.42, 2.09, 1.98, 1.87, 1.65, 1.43, 0.88, 0.33, 0.22]
    }
    return tarifs.get(emplacement, [0]*12)[idx]

def get_tarif_broderie(qte_globale, emplacement):
    if qte_globale <= 3: idx = 0
    elif qte_globale <= 11: idx = 1
    elif qte_globale <= 23: idx = 2
    elif qte_globale <= 47: idx = 3
    elif qte_globale <= 95: idx = 4
    elif qte_globale <= 251: idx = 5
    elif qte_globale <= 503: idx = 6
    elif qte_globale <= 1007: idx = 7
    elif qte_globale <= 1511: idx = 8
    else: idx = 9

    tarifs = {
        "Poitrine (9x8 cm)": [15.90, 12.13, 9.20, 6.90, 5.30, 4.80, 4.50, 4.20, 3.90, 3.50],
        "Dos D10 (25x10 cm)": [18.50, 14.60, 11.30, 9.10, 7.60, 7.10, 6.70, 6.30, 5.90, 5.40],
        "Dos Large D20 (25x20)": [23.20, 18.25, 14.20, 11.90, 9.90, 9.30, 8.80, 8.30, 7.80, 7.20],
        "Col/Signature (7x2)": [10.90, 8.37, 6.20, 4.40, 3.10, 2.85, 2.65, 2.45, 2.25, 2.00],
        "Casquettes / Bonnets": [16.10, 12.27, 9.30, 7.10, 5.50, 5.05, 4.75, 4.45, 4.15, 3.75],
        "Manche (8x5 cm)": [16.10, 12.27, 9.30, 7.10, 5.50, 5.05, 4.75, 4.45, 4.15, 3.75],
        "Pantalon / Poche": [18.30, 13.97, 10.90, 8.60, 7.10, 6.60, 6.20, 5.80, 5.45, 4.95],
        "+ Perso. Nom (Cœur)": [6.00, 4.00, 4.00, 3.50, 3.00, 2.80, 2.60, 2.40, 2.20, 2.00]
    }
    return tarifs.get(emplacement, [0]*10)[idx]

def get_frais_prog_broderie(qte):
    if qte <= 1: return 41.00
    elif qte <= 3: return 41.00
    elif qte <= 11: return 23.00
    else: return 0.00

def get_tarif_assurance(qte):
    if qte <= 11: return 3.08
    elif qte <= 24: return 2.38
    elif qte <= 49: return 1.83
    elif qte <= 99: return 1.25
    elif qte <= 249: return 0.98
    elif qte <= 499: return 0.70
    elif qte <= 999: return 0.64
    elif qte <= 1999: return 0.61
    else: return 0.55

def get_tarif_ensachage(qte, type_sachet):
    if type_sachet == "Sachet (T-shirt/Polo)":
        if qte <= 11: return 1.38
        elif qte <= 24: return 1.24
        elif qte <= 49: return 1.17
        elif qte <= 99: return 1.11
        elif qte <= 249: return 1.08
        elif qte <= 499: return 1.06
        elif qte <= 1999: return 1.00
        else: return 0.98
    else:
        if qte <= 11: return 1.75
        elif qte <= 24: return 1.65
        elif qte <= 49: return 1.53
        elif qte <= 99: return 1.43
        elif qte <= 249: return 1.39
        elif qte <= 499: return 1.36
        elif qte <= 1999: return 1.34
        else: return 1.32

def get_tarif_stockage(qte):
    if qte < 50: return 0.0
    elif qte <= 99: return 1.00
    elif qte <= 249: return 0.56
    elif qte <= 499: return 0.50
    elif qte <= 1999: return 0.43
    else: return 0.30

def calculer_frais_port(montant_ht, zone):
    if montant_ht >= 1500: return 0.0
    if zone in ["France Continentale", "Corse, Monaco ou Andorre"] and montant_ht >= 1000: return 0.0
    
    if zone == "France Continentale":
        if montant_ht < 99.99: return 14.95
        elif montant_ht <= 499.99: return 20.95
        elif montant_ht <= 999.99: return 24.95
        else: return 0.0
    elif zone == "Corse, Monaco ou Andorre":
        if montant_ht < 99.99: return 19.95
        elif montant_ht <= 499.99: return 25.95
        elif montant_ht <= 999.99: return 29.95
        else: return 0.0
    elif zone == "Espace UE":
        if montant_ht < 99.99: return 25.95
        elif montant_ht <= 499.99: return 39.00
        elif montant_ht <= 999.99: return 60.00
        else: return 90.00
    else: return 0.0

# --- CHARGEMENT DU CRM POUR L'AUTO-COMPLÉTION ---
clients_connus = []
entreprises_connues = []
dict_clients = {}
if os.path.exists(CRM_FILE):
    try:
        df_crm_load = pd.read_csv(CRM_FILE)
        for _, row in df_crm_load.drop_duplicates(subset=["Client"]).iterrows():
            c_nom = str(row.get("Client", ""))
            c_ent = str(row.get("Entreprise", ""))
            if c_nom and c_nom != "nan": clients_connus.append(c_nom)
            if c_ent and c_ent != "nan": entreprises_connues.append(c_ent)
            dict_clients[c_nom] = {
                "entreprise": c_ent,
                "email": row.get("Email", ""),
                "telephone": row.get("Telephone", "")
            }
    except Exception:
        pass

# --- INTERFACE ---
st.title("🖨️ Gestionnaire de Devis - Marquage Textile")

st.sidebar.header("📋 Infos Client & Expédition")

client_entreprise = st.sidebar.text_input("Nom de l'Entreprise / Société", value="")
client_nom = st.sidebar.text_input("Nom de la personne contact (ex: Jean Dupont)", value="")

if client_nom in dict_clients and client_nom != "":
    infos = dict_clients[client_nom]
    if not client_entreprise: client_entreprise = infos["entreprise"]
    default_email = infos["email"]
    default_tel = infos["telephone"]
else:
    default_email = ""
    default_tel = ""

client_adresse = st.sidebar.text_area("Adresse complète du Client")
client_siret = st.sidebar.text_input("SIRET du Client (optionnel)")
client_contact = st.sidebar.text_input("Téléphone du Client", value=default_tel)
client_email = st.sidebar.text_input("E-mail du Client", value=default_email)
zone_livraison = st.sidebar.selectbox("Zone de Livraison", ["France Continentale", "Corse, Monaco ou Andorre", "Espace UE", "DOM/TOM et pays hors UE"])
offrir_port = st.sidebar.checkbox("Offrir les frais de port (0 €)", value=False)
frais_techniques_dossier = st.sidebar.number_input("Frais techniques de commande (€ HT)", value=19.80)

st.sidebar.markdown("---")
st.sidebar.header("👤 Conseiller émetteur")
conseiller_choix = st.sidebar.selectbox(
    "Envoyer en tant que :", 
    ["Brice Geny (brice.geny@gmail.com)", "Brice Bugna (brice.bugna@gmail.com)"]
)

if "Brice Geny" in conseiller_choix:
    conseiller_nom = "Brice Geny"
    conseiller_email = "brice.geny@gmail.com"
    conseiller_tel = "06 32 69 73 28"
    autre_email = "brice.bugna@gmail.com"
    gmail_password = st.secrets.get("EMAIL_PASSWORD_GENY", st.secrets.get("EMAIL_PASSWORD", ""))
else:
    conseiller_nom = "Brice Bugna"
    conseiller_email = "brice.bugna@gmail.com"
    conseiller_tel = "06 29 92 94 74"
    autre_email = "brice.geny@gmail.com"
    gmail_password = st.secrets.get("EMAIL_PASSWORD_BUGNA", st.secrets.get("EMAIL_PASSWORD", ""))

st.sidebar.markdown("---")
st.sidebar.header("🖼️ Logo de l'entreprise")
logo_file = st.sidebar.file_uploader("Importer un autre logo (PNG/JPG)", type=["png", "jpg", "jpeg"])

logo_defaut_github = "Gemini_Generated_Image_mxbmbrmxbmbrmxbm.jpeg"
if logo_file is None and os.path.exists(logo_defaut_github):
    st.sidebar.image(logo_defaut_github, width=150, caption="Logo actif (GitHub)")

# --- CONFIGURATION DES ARTICLES DEPUIS LE CATALOGUE EXCEL ---
st.header("🛍️ Sélection des Articles (Catalogue par Catégories)")

if df_catalogue.empty:
    st.warning(f"⚠️ Le fichier `{CATALOGUE_FILE}` est introuvable ou vide. Veuillez vérifier sa présence dans le dossier.")
    categories_disponibles = []
else:
    # On détecte automatiquement les noms de colonnes (ajustez si besoin selon vos en-têtes exacts dans Excel)
    cols = df_catalogue.columns.tolist()
    col_cat = cols[0] if len(cols) > 0 else "Catégorie"
    col_nom = cols[1] if len(cols) > 1 else "Article"
    col_prix = cols[2] if len(cols) > 2 else "Prix HT"
    
    categories_disponibles = df_catalogue[col_cat].dropna().unique().tolist()

nb_articles_config = st.number_input("Nombre d'articles différents à ajouter au devis", min_value=1, max_value=20, value=3)

articles_saisis = []

for i in range(int(nb_articles_config)):
    st.markdown(markdown_text := f"### Article {i+1}")
    
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        if categories_disponibles:
            cat_choisie = st.selectbox(f"Catégorie {i+1}", categories_disponibles, key=f"cat_{i}")
            # Filtrer les articles de cette catégorie
            articles_cat = df_catalogue[df_catalogue[col_cat] == cat_choisie][col_nom].dropna().unique().tolist()
            article_choisi = st.selectbox(f"Référence / Article {i+1}", articles_cat, key=f"art_ref_{i}")
            
            # Récupérer le prix unitaire par défaut depuis le catalogue si disponible
            ligne_art = df_catalogue[(df_catalogue[col_cat] == cat_choisie) & (df_catalogue[col_nom] == article_choisi)]
            prix_defaut = float(ligne_art[col_prix].values[0]) if not ligne_art.empty and col_prix in ligne_art.columns else 4.92
        else:
            cat_choisie = "Standard"
            article_choisi = st.text_input(f"Nom de l'article {i+1}", value=f"Article {i+1}", key=f"art_txt_{i}")
            prix_defaut = 4.92

    with col_sel2:
        nom_article_final = f"{cat_choisie} - {article_choisi}" if categories_disponibles else article_choisi
        nom_article = st.text_input(f"Libellé personnalisé sur devis {i+1}", value=nom_article_final, key=f"nom_perso_{i}")
        qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_{i}")
        prix_vetement_ht = st.number_input(f"Prix unitaire HT (€) {i+1}", min_value=0.0, value=float(prix_defaut), key=f"px_vet_{i}")

    sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
    
    nb_marquages = 0
    marquages_config = []
    has_broderie = False
    
    if not sans_marquage:
        nb_marquages = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], key=f"nb_m_{i}")
        for m in range(nb_marquages):
            st.markdown(f"--- *Marquage {m+1} pour Article {i+1}*")
            mc1, mc2 = st.columns(2)
            with mc1:
                t_marq = st.selectbox(f"Technique M{m+1} (Art {i+1})", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
            with mc2:
                if t_marq == "DTF Textile Fin":
                    emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_f_{i}_{m}")
                elif t_marq == "DTF Textile Épais":
                    emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_e_{i}_{m}")
                else:
                    emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20)", "Col/Signature (7x2)", "Casquettes / Bonnets", "Manche (8x5 cm)", "Pantalon / Poche", "+ Perso. Nom (Cœur)"], key=f"emp_b_{i}_{m}")
            
            if t_marq == "Broderie HD":
                has_broderie = True

            marquages_config.append({"technique": t_marq, "emplacement": emp})

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
            "type_sachet": type_sachet,
            "option_assurance": option_assurance,
            "option_stockage": option_stockage,
            "remise_fidelite": remise_fidelite
        })
    st.markdown("---")

# --- ÉTAPE 2 : CALCUL DES QUANTITÉS GLOBALES PAR MARQUAGE ---
compteur_global_marquages = {}
for art in articles_saisis:
    if not art["sans_marquage"]:
        qte_art = art["quantite"]
        for m in art["marquages_config"]:
            cle = (m["technique"], m["emplacement"])
            compteur_global_marquages[cle] = compteur_global_marquages.get(cle, 0) + qte_art

# --- ÉTAPE 3 : CONSOLIDATION DU DEVIS ---
lignes_devis_global = []
for art in articles_saisis:
    qte = art["quantite"]
    marquages_calcules = []
    total_marquage_unit = 0.0
    
    if not art["sans_marquage"]:
        for m in art["marquages_config"]:
            cle = (m["technique"], m["emplacement"])
            qte_globale_marq = compteur_global_marquages.get(cle, qte)
            
            if m["technique"] == "DTF Textile Fin":
                tarif_m = get_tarif_dtf_fin(qte_globale_marq, m["emplacement"])
            elif m["technique"] == "DTF Textile Épais":
                tarif_m = get_tarif_dtf_epais(qte_globale_marq, m["emplacement"])
            else:
                tarif_m = get_tarif_broderie(qte_globale_marq, m["emplacement"])
                
            total_marquage_unit += tarif_m
            marquages_calcules.append({"nom": f"{m['technique']} - {m['emplacement']}", "tarif": tarif_m})

    frais_prog_broderie = get_frais_prog_broderie(qte) if art["has_broderie"] else 0.0
    coût_ensachage_unit = get_tarif_ensachage(qte, art["type_sachet"]) if art["option_ensachage"] else 0.0
    coût_assurance_unit = get_tarif_assurance(qte) if art["option_assurance"] else 0.0
    coût_stockage_unit = get_tarif_stockage(qte) if art["option_stockage"] else 0.0

    total_unit_ht = art["prix_vet_unit"] + total_marquage_unit + coût_ensachage_unit + coût_assurance_unit + coût_stockage_unit
    total_ligne_brut = (total_unit_ht * qte) + frais_prog_broderie
    montant_remise = total_ligne_brut * (art["remise_fidelite"] / 100.0)
    total_ligne_ht = total_ligne_brut - montant_remise
    
    lignes_devis_global.append({
        "nom_article": art["nom_article"],
        "quantite": qte,
        "prix_vet_unit": art["prix_vet_unit"],
        "sans_marquage": art["sans_marquage"],
        "marquages": marquages_calcules,
        "frais_prog_broderie": frais_prog_broderie,
        "option_ensachage": art["option_ensachage"],
        "type_sachet": art["type_sachet"],
        "coût_ensachage_unit": coût_ensachage_unit,
        "option_assurance": art["option_assurance"],
        "coût_assurance_unit": coût_assurance_unit,
        "option_stockage": art["option_stockage"],
        "coût_stockage_unit": coût_stockage_unit,
        "remise_fidelite": art["remise_fidelite"],
        "total_ligne_ht": total_ligne_ht
    })

# --- SECTION RÉCAPITULATIF & DEVIS ---
st.subheader("📊 Récapitulatif Général & Génération du Devis Professionnel")

if not lignes_devis_global:
    st.warning("Veuillez renseigner au moins un article avec une quantité supérieure à 0.")
else:
    sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
    montant_base_port = sous_total_articles + frais_techniques_dossier
    
    if offrir_port:
        frais_port = 0.0
    else:
        frais_port = calculer_frais_port(montant_base_port, zone_livraison)
    
    total_general_ht = montant_base_port + frais_port
    tva = total_general_ht * 0.20
    total_ttc = total_general_ht + tva
    
    quantite_globale_totale = sum([item["quantite"] for item in lignes_devis_global])
    cout_unitaire_moyen = total_general_ht / quantite_globale_totale if quantite_globale_totale > 0 else 0

    st.write(f"**Quantité globale pièces :** {quantite_globale_totale}")
    st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
    st.write(f"**Frais techniques :** {frais_techniques_dossier:.2f} € HT")
    st.write(f"**Frais de port ({zone_livraison}) :** {frais_port:.2f} € HT" if not offrir_port else "**Frais de port :** Offerts (0.00 €)")
    st.markdown(f"### **Total Général HT : {total_general_ht:.2f} €** | **TOTAL TTC (20%) : {total_ttc:.2f} €**")

    if st.button("📄 Générer le numéro de devis et le PDF"):
        num_devis = obtenir_prochain_numero_devis()
        pdf_filename = f"Devis_{num_devis.replace('/', '_')}.pdf"
        st.session_state['dernier_pdf'] = pdf_filename
        st.session_state['dernier_num'] = num_devis
        st.session_state['total_ttc_cache'] = total_ttc
        st.session_state['total_ht_cache'] = total_general_ht

        data_crm = {
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Numero_Devis": num_devis,
            "Conseiller": conseiller_nom,
            "Client": client_nom,
            "Entreprise": client_entreprise,
            "Email": client_email,
            "Telephone": client_contact,
            "Quantite_Totale": quantite_globale_totale,
            "Total_HT": round(total_general_ht, 2),
            "Total_TTC": round(total_ttc, 2),
            "Statut": "En cours"
        }
        enregistrer_dans_crm(data_crm)
        st.success("✅ Données enregistrées dans le CRM avec succès !")

    if 'dernier_pdf' in st.session_state:
        pdf_filename = st.session_state['dernier_pdf']
        num_devis = st.session_state['dernier_num']
        total_ttc = st.session_state['total_ttc_cache']
        total_general_ht = st.session_state['total_ht_cache']

        # --- GÉNÉRATION DU PDF PROFESSIONNEL ---
        doc = SimpleDocTemplate(pdf_filename, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        story = []
        styles = getSampleStyleSheet()

        style_sub = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#4a5568'))
        style_cell = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=9, leading=11)
        style_cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold')
        style_right_bold = ParagraphStyle('RightBold', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=2)
        style_right_normal = ParagraphStyle('RightNormal', parent=styles['Normal'], fontSize=9, leading=11, alignment=2)

        logo_path = None
        if logo_file is not None:
            logo_path = "temp_logo.png"
            with open(logo_path, "wb") as f:
                f.write(logo_file.getbuffer())
        elif os.path.exists(logo_defaut_github):
            logo_path = logo_defaut_github

        header_text = Paragraph(
            "<b>APEX BUSINESS & COM - SOLUTIONS TEXTILE & MARQUAGE</b><br/>"
            "70 - Marnay<br/>"
            "Tél (Brice Geny) : 06 32 69 73 28 &nbsp;|&nbsp; Tél (Brice Bugna) : 06 29 92 94 74<br/>"
            "Email : brice.geny@gmail.com", 
            style_sub
        )
        
        if logo_path and os.path.exists(logo_path):
            img_logo = RLImage(logo_path, width=110, height=45)
            t_header = Table([[img_logo, header_text]], colWidths=[120, 420])
            t_header.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
            story.append(t_header)
        else:
            story.append(header_text)

        story.append(Spacer(1, 10))
        story.append(Paragraph(f"<b>N° Devis :</b> {num_devis} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Date :</b> {datetime.now().strftime('%d/%m/%Y')} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Validité :</b> 30 Jours", style_sub))
        story.append(Spacer(1, 10))

        siret_txt = f"<br/>SIRET : {client_siret}" if client_siret else ""
        contact_txt = f"<br/>Contact : {client_contact}" if client_contact else ""
        nom_aff_client = f"<b>{client_entreprise}</b><br/>À l'attention de : {client_nom}" if client_entreprise else f"<b>{client_nom or 'Client'}</b>"
        client_info_text = f"<b>CLIENT / DESTINATAIRE :</b><br/>{nom_aff_client}<br/>{client_adresse.replace(chr(10), '<br/>')}{siret_txt}{contact_txt}<br/>Email : {client_email}"
        order_info_text = f"<b>DÉTAILS DE LA COMMANDE :</b><br/>Quantité globale : {quantite_globale_totale} pièces<br/>Délai estimé : 8 à 10 jours ouvrés<br/>Conseiller : {conseiller_nom}"
        
        t_info = Table([[Paragraph(client_info_text, style_cell), Paragraph(order_info_text, style_cell)]], colWidths=[270, 270])
        t_info.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f7fafc')),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e0')),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('PADDING', (0,0), (-1,-1), 8),
        ]))
        story.append(t_info)
        story.append(Spacer(1, 15))

        table_data = [[
            Paragraph("<b>DÉSIGNATION & CARACTÉRISTIQUES</b>", style_cell_bold), 
            Paragraph("<b>QTÉ</b>", style_cell_bold), 
            Paragraph("<b>PRIX UNIT. HT</b>", style_cell_bold), 
            Paragraph("<b>TOTAL HT</b>", style_cell_bold)
        ]]

        for item in lignes_devis_global:
            libelle_support = f"<b>Support (Sans marquage) : {item['nom_article']}</b>" if item['sans_marquage'] else f"<b>Support : {item['nom_article']}</b>"
            table_data.append([
                Paragraph(libelle_support, style_cell), 
                str(item['quantite']), 
                f"{item['prix_vet_unit']:.2f} €", 
                f"{item['prix_vet_unit']*item['quantite']:.2f} €"
            ])
            for m in item['marquages']:
                table_data.append([
                    Paragraph(f"&nbsp;&nbsp;&bull; Marquage : {m['nom']}", style_cell), 
                    str(item['quantite']), 
                    f"{m['tarif']:.2f} €", 
                    f"{m['tarif']*item['quantite']:.2f} €"
                ])
            if item['frais_prog_broderie'] > 0:
                table_data.append([
                    Paragraph("&nbsp;&nbsp;&bull; Frais de technique & programme Broderie", style_cell), 
                    "1", 
                    f"{item['frais_prog_broderie']:.2f} €", 
                    f"{item['frais_prog_broderie']:.2f} €"
                ])
            if item['option_ensachage']:
                table_data.append([
                    Paragraph(f"&nbsp;&nbsp;&bull; Option : {item['type_sachet']}", style_cell), 
                    str(item['quantite']), 
                    f"{item['coût_ensachage_unit']:.2f} €", 
                    f"{item['coût_ensachage_unit']*item['quantite']:.2f} €"
                ])
            if item['option_assurance']:
                table_data.append([
                    Paragraph("&nbsp;&nbsp;&bull; Option : Assurance MHC (Garantie textile)", style_cell), 
                    str(item['quantite']), 
                    f"{item['coût_assurance_unit']:.2f} €", 
                    f"{item['coût_assurance_unit']*item['quantite']:.2f} €"
                ])
            if item['option_stockage']:
                table_data.append([
                    Paragraph("&nbsp;&nbsp;&bull; Option : Mise en stockage + picking", style_cell), 
                    str(item['quantite']), 
                    f"{item['coût_stockage_unit']:.2f} €", 
                    f"{item['coût_stockage_unit']*item['quantite']:.2f} €"
                ])
            if item['remise_fidelite'] > 0:
                brut_calc = (item['prix_vet_unit']*item['quantite']) + sum([m['tarif']*item['quantite'] for m in item['marquages']]) + item['frais_prog_broderie'] + (item['coût_ensachage_unit']*item['quantite'] if item['option_ensachage'] else 0) + (item['coût_assurance_unit']*item['quantite'] if item['option_assurance'] else 0) + (item['coût_stockage_unit']*item['quantite'] if item['option_stockage'] else 0)
                montant_remise_ligne = brut_calc * (item['remise_fidelite']/100.0)
                table_data.append([
                    Paragraph(f"&nbsp;&nbsp;&bull; <b>Réduction fidélité ({item['remise_fidelite']}%)</b>", style_cell), 
                    "1", 
                    "-", 
                    f"-{montant_remise_ligne:.2f} €"
                ])

        table_data.append([Paragraph("Frais techniques de dossier", style_cell), "1", f"{frais_techniques_dossier:.2f} €", f"{frais_techniques_dossier:.2f} €"])
        if frais_port > 0 or offrir_port:
            port_libelle = f"Frais d'envoi ({zone_livraison})" if not offrir_port else f"Frais d'envoi ({zone_livraison}) - Offerts"
            port_val = f"{frais_port:.2f} €"
            table_data.append([Paragraph(port_libelle, style_cell), "1", port_val, port_val])

        t_main = Table(table_data, colWidths=[260, 45, 115, 120])
        t_main.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#edf2f7')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('PADDING', (0,0), (-1,-1), 6),
        ]))
        story.append(t_main)
        story.append(Spacer(1, 10))

        totaux_data = [
            ["", Paragraph("Sous-Total HT :", style_right_normal), Paragraph(f"{total_general_ht:.2f} €", style_right_normal)],
            ["", Paragraph("TVA (20%) :", style_right_normal), Paragraph(f"{tva:.2f} €", style_right_normal)],
            ["", Paragraph("TOTAL TTC :", style_right_bold), Paragraph(f"{total_ttc:.2f} €", style_right_bold)],
            ["", Paragraph("Coût unitaire HT / pièce :", style_right_normal), Paragraph(f"{cout_unitaire_moyen:.2f} €", style_right_normal)]
        ]
        t_totaux = Table(totaux_data, colWidths=[240, 160, 140])
        t_totaux.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LINEABOVE', (1, 2), (-1, 2), 1, colors.black),
            ('PADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_totaux)
        story.append(Spacer(1, 15))

        conditions_text = "<b>Conditions de règlement & Bon pour accord :</b><br/>• Acompte de 50% à la commande, solde à la livraison.<br/>• Fichiers vectoriels fournis (.AI, .EPS, .PDF).<br/>• Bon pour accord daté et signé requis."
        story.append(Paragraph(conditions_text, style_sub))

        doc.build(story)
        st.success(f"Devis PDF professionnel généré sous le numéro : **{num_devis}** (`{pdf_filename}`)")

        # --- ENVOI DIRECT GMAIL ---
        st.markdown(f"### ✉️ Envoi direct par E-mail (via {conseiller_email})")
        
        if st.button("🚀 Envoyer le devis par e-mail maintenant"):
            if not client_email:
                st.error("Veuillez renseigner l'e-mail du client dans la barre latérale.")
            elif not gmail_password:
                st.error(f"Veuillez renseigner le mot de passe d'application Gmail pour {conseiller_email} dans les secrets Streamlit Cloud.")
            else:
                try:
                    msg = EmailMessage()
                    msg['Subject'] = f"Devis {num_devis} - Marquage Textile"
                    msg['From'] = conseiller_email
                    msg['To'] = client_email
                    msg['Cc'] = autre_email
                    
                    corps_mail = f"""Bonjour {client_nom or 'Client'},

Veuillez trouver ci-joint votre devis n° {num_devis} d'un montant total de {total_ttc:.2f} € TTC.

Restant à votre disposition pour toute information complémentaire.

Cordialement,
{conseiller_nom}
APEX BUSINESS & COM
Tél : {conseiller_tel}
Email : {conseiller_email}

---
Application créée par APEX - Tous droits réservés
"""
                    msg.set_content(corps_mail)

                    with open(pdf_filename, 'rb') as f:
                        file_data = f.read()
                        file_name = os.path.basename(pdf_filename)
                    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(conseiller_email, gmail_password)
                        smtp.send_message(msg)
                    
                    st.success(f"✅ E-mail envoyé avec succès à {client_email} depuis la boîte de {conseiller_nom} (avec copie à {autre_email}) !")
                except Exception as e:
                    st.error(f"❌ Erreur lors de l'envoi de l'e-mail : {e}")

# --- SECTION SUIVI CRM ---
st.markdown("---")
st.subheader("📈 Tableau de Suivi CRM & Historique des Devis")

if os.path.exists(CRM_FILE):
    df_crm = pd.read_csv(CRM_FILE)
    if not df_crm.empty:
        edited_df = st.data_editor(
            df_crm,
            column_config={
                "Statut": st.column_config.SelectboxColumn(
                    "Statut du Devis",
                    options=["En cours", "Accepté", "Refusé", "Sans suite"],
                    required=True
                ),
                "Total_HT": st.column_config.NumberColumn("Total HT (€)", format="%.2f €"),
                "Total_TTC": st.column_config.NumberColumn("Total TTC (€)", format="%.2f €"),
            },
            disabled=["Date", "Numero_Devis", "Conseiller", "Client", "Entreprise", "Email", "Telephone", "Quantite_Totale"],
            use_container_width=True,
            key="crm_editor"
        )
        if not edited_df.equals(df_crm):
            edited_df.to_csv(CRM_FILE, index=False)
            st.toast("✅ Statuts mis à jour avec succès !", icon="💾")
    else:
        st.info("Aucun devis enregistré pour le moment.")
else:
    st.info("Aucun historique CRM disponible.")