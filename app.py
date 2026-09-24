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

st.set_page_config(page_title="Gestionnaire de Devis - ABCOM", layout="wide")

# --- GESTION DU COMPTEUR DE DEVIS & CRM ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"

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

def get_tarif_carte_visite(qte, type_finition, recto_verso):
    # type_finition: "Sans pelliculage", "Avec pelliculage (Recto)", "Avec pelliculage (Recto/Verso)" ou similaire
    # Simplification grille : 100, 250, 500, 1000
    if qte <= 150: q_idx = 0
    elif qte <= 375: q_idx = 1
    elif qte <= 750: q_idx = 2
    else: q_idx = 3

    if "Sans pelliculage" in type_finition:
        if recto_verso == "Recto":
            grille = [19.9, 29.9, 39.9, 45.9]
        else:
            grille = [29.9, 39.9, 49.9, 59.9]
    else: # Avec pelliculage
        if recto_verso == "Recto":
            grille = [34.9, 45.9, 55.9, 64.9]
        else:
            grille = [39.9, 49.9, 59.9, 69.9]
    return grille[q_idx]

def get_tarif_calendrier_magnetique(qte):
    if qte <= 75: return 108.0 * (qte/50.0)
    elif qte <= 150: return 172.0 * (qte/100.0)
    elif qte <= 225: return 302.0 * (qte/200.0)
    elif qte <= 275: return 366.0 * (qte/250.0)
    elif qte <= 350: return 432.0 * (qte/300.0)
    elif qte <= 450: return 506.0 * (qte/400.0)
    else: return 628.0 * (qte/500.0)

def get_tarif_calendrier_papier(qte, grammage_format):
    # grammage_format ex: "250g - A4", "350g - A5", etc.
    if qte <= 175: idx = 0
    elif qte <= 375: idx = 1
    elif qte <= 750: idx = 2
    elif qte <= 1750: idx = 3
    else: idx = 4

    # Grilles simplifiées issues du tableau
    if "250g" in grammage_format:
        if "A4" in grammage_format: tarifs = [45.36, 74.898, 123.516, 176.4, 207.61]
        elif "A5" in grammage_format: tarifs = [35.28, 50.4, 73.58, 122.20, 126.14]
        else: tarifs = [30.24, 36.54, 49.14, 77.52, 85.41]
    else: # 350g
        if "A4" in grammage_format: tarifs = [62.9, 125.8, 168.3, 236.3, 450.5]
        elif "A5" in grammage_format: tarifs = [45.9, 68.0, 81.6, 112.2, 207.4]
        else: tarifs = [35.0, 50.0, 65.0, 90.0, 150.0]
    return tarifs[idx]

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

# --- CHARGEMENT DU CRM ---
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
st.title("🖨️ ABCOM - Gestionnaire de Devis & CRM")

st.sidebar.header("📋 Infos Client & Expédition")
client_entreprise = st.sidebar.text_input("Nom de l'Entreprise / Société", value="")
client_nom = st.sidebar.text_input("Nom du contact (ex: Brice Geny)", value="")

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
st.sidebar.header("👤 Gestionnaire émetteur")
conseiller_choix = st.sidebar.selectbox("Émis par :", ["Brice Geny (ABCOM)"])
conseiller_nom = "Brice Geny"
conseiller_email = "brice.geny@gmail.com"
conseiller_tel = "03 84 70 00 00"
gmail_password = st.secrets.get("EMAIL_PASSWORD_GENY", st.secrets.get("EMAIL_PASSWORD", ""))

st.sidebar.markdown("---")
st.sidebar.header("🖼️ Logo de l'entreprise")
logo_file = st.sidebar.file_uploader("Importer un logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
logo_defaut_github = "Gemini_Generated_Image_mxbmbrmxbmbrmxbm.jpeg"

noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"]
onglets = st.tabs(noms_onglets)

articles_saisis = []

for i in range(10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        
        type_produit = st.selectbox(f"Type de produit {i+1}", ["Textile / Vêtement", "Cartes de visite", "Calendriers"], key=f"t_prod_{i}")
        
        if type_produit == "Textile / Vêtement":
            sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
            col1, col2 = st.columns(2)
            with col1:
                nom_article = st.text_input(f"Nom / Référence du vêtement {i+1}", value=f"T-Shirt 100% coton" if i==0 else f"Article {i+1}", key=f"nom_{i}")
                qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT du vêtement (€) {i+1}", min_value=0.0, value=4.92, key=f"px_vet_{i}")
            with col2:
                if not sans_marquage:
                    nb_marquages = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], key=f"nb_m_{i}")
                else:
                    nb_marquages = 0

            marquages_config = []
            has_broderie = False
            if not sans_marquage:
                for m in range(nb_marquages):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
                    with mc2:
                        if t_marq == "DTF Textile Fin":
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_f_{i}_{m}")
                        elif t_marq == "DTF Textile Épais":
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_e_{i}_{m}")
                        else:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20)", "Col/Signature (7x2)", "Casquettes / Bonnets", "Manche (8x5 cm)", "Pantalon / Poche", "+ Perso. Nom (Cœur)"], key=f"emp_b_{i}_{m}")
                    if t_marq == "Broderie HD": has_broderie = True
                    marquages_config.append({"technique": t_marq, "emplacement": emp})

            oc1, oc2, oc3, oc4 = st.columns(4)
            with oc1:
                option_ensachage = st.checkbox(f"Ensachage individuel {i+1}", key=f"ens_{i}")
                type_sachet = st.selectbox(f"Type de sachet {i+1}", ["Sachet (T-shirt/Polo)", "Sachet (Veste/Sweat)"], key=f"tsach_{i}") if option_ensachage else ""
            with oc2: option_assurance = st.checkbox(f"Assurance MHC {i+1}", key=f"ass_{i}")
            with oc3: option_stockage = st.checkbox(f"Stockage + picking {i+1}", key=f"stock_{i}")
            with oc4: remise_fidelite = st.number_input(f"Réduction / Promo (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"rem_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type": "Textile",
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

        elif type_produit == "Cartes de visite":
            nom_article = st.text_input(f"Libellé Cartes de visite {i+1}", value="Cartes de visite standard 350g", key=f"nom_cv_{i}")
            qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=250, key=f"qte_cv_{i}")
            cc1, cc2 = st.columns(2)
            with cc1:
                finition_cv = st.selectbox(f"Finition {i+1}", ["Sans pelliculage", "Avec pelliculage (Brillant / Mat / Velours)"], key=f"fin_cv_{i}")
            with cc2:
                recto_verso_cv = st.selectbox(f"Impression {i+1}", ["Recto", "Recto/Verso"], key=f"rv_cv_{i}")
            remise_fidelite = st.number_input(f"Réduction / Promo (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"rem_cv_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type": "Cartes de visite",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "finition": finition_cv,
                    "recto_verso": recto_verso_cv,
                    "remise_fidelite": remise_fidelite
                })

        else: # Calendriers
            nom_article = st.text_input(f"Libellé Calendrier {i+1}", value="Calendrier publicitaire", key=f"nom_cal_{i}")
            qte = st.number_input(f"Quantité {i+1}", min_value=0, value=100, key=f"qte_cal_{i}")
            type_cal = st.selectbox(f"Type de calendrier {i+1}", ["Papier (couché brillant/mat)", "Magnétique A5"], key=f"t_cal_{i}")
            
            grammage_format = ""
            if type_cal == "Papier (couché brillant/mat)":
                grammage_format = st.selectbox(f"Format & Grammage {i+1}", ["250g - A4", "250g - A5", "250g - A6", "350g - A4", "350g - A5"], key=f"gf_cal_{i}")
            
            remise_fidelite = st.number_input(f"Réduction / Promo (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, step=1.0, key=f"rem_cal_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type": "Calendrier",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "type_cal": type_cal,
                    "grammage_format": grammage_format,
                    "remise_fidelite": remise_fidelite
                })

# --- CALCULS GLOBAUX ---
compteur_global_marquages = {}
for art in articles_saisis:
    if art["type"] == "Textile" and not art["sans_marquage"]:
        qte_art = art["quantite"]
        for m in art["marquages_config"]:
            cle = (m["technique"], m["emplacement"])
            compteur_global_marquages[cle] = compteur_global_marquages.get(cle, 0) + qte_art

lignes_devis_global = []
for art in articles_saisis:
    if art["type"] == "Textile":
        qte = art["quantite"]
        marquages_calcules = []
        total_marq_unit = 0.0
        if not art["sans_marquage"]:
            for m in art["marquages_config"]:
                cle = (m["technique"], m["emplacement"])
                qte_globale_marq = compteur_global_marquages.get(cle, qte)
                if m["technique"] == "DTF Textile Fin": tarif_m = get_tarif_dtf_fin(qte_globale_marq, m["emplacement"])
                elif m["technique"] == "DTF Textile Épais": tarif_m = get_tarif_dtf_epais(qte_globale_marq, m["emplacement"])
                else: tarif_m = get_tarif_broderie(qte_globale_marq, m["emplacement"])
                total_marq_unit += tarif_m
                marquages_calcules.append({"nom": f"{m['technique']} - {m['emplacement']}", "tarif": tarif_m})

        frais_prog = get_frais_prog_broderie(qte) if art["has_broderie"] else 0.0
        c_ens = get_tarif_ensachage(qte, art["type_sachet"]) if art["option_ensachage"] else 0.0
        c_ass = get_tarif_assurance(qte) if art["option_assurance"] else 0.0
        c_stk = get_tarif_stockage(qte) if art["option_stockage"] else 0.0

        total_unit = art["prix_vet_unit"] + total_marq_unit + c_ens + c_ass + c_stk
        brut = (total_unit * qte) + frais_prog
        remise = brut * (art["remise_fidelite"] / 100.0)
        total_ht = brut - remise

        lignes_devis_global.append({
            "nom_article": art["nom_article"], "quantite": qte, "type": "Textile",
            "prix_unit": art["prix_vet_unit"], "sans_marquage": art["sans_marquage"],
            "marquages": marquages_calcules, "frais_prog": frais_prog,
            "option_ensachage": art["option_ensachage"], "type_sachet": art["type_sachet"], "coût_ensachage_unit": c_ens,
            "option_assurance": art["option_assurance"], "coût_assurance_unit": c_ass,
            "option_stockage": art["option_stockage"], "coût_stockage_unit": c_stk,
            "remise_fidelite": art["remise_fidelite"], "total_ligne_ht": total_ht
        })

    elif art["type"] == "Cartes de visite":
        qte = art["quantite"]
        prix_total = get_tarif_carte_visite(qte, art["finition"], art["recto_verso"])
        brut = prix_total
        remise = brut * (art["remise_fidelite"] / 100.0)
        total_ht = brut - remise
        prix_unit = total_ht / qte if qte > 0 else 0

        lignes_devis_global.append({
            "nom_article": f"{art['nom_article']} ({art['finition']}, {art['recto_verso']})",
            "quantite": qte, "type": "Cartes de visite", "prix_unit": prix_unit,
            "remise_fidelite": art["remise_fidelite"], "total_ligne_ht": total_ht
        })

    else: # Calendrier
        qte = art["quantite"]
        if art["type_cal"] == "Magnétique A5":
            brut = get_tarif_calendrier_magnetique(qte)
        else:
            brut = get_tarif_calendrier_papier(qte, art["grammage_format"])
        remise = brut * (art["remise_fidelite"] / 100.0)
        total_ht = brut - remise
        prix_unit = total_ht / qte if qte > 0 else 0

        lignes_devis_global.append({
            "nom_article": f"{art['nom_article']} ({art['type_cal']} {art['grammage_format']})",
            "quantite": qte, "type": "Calendrier", "prix_unit": prix_unit,
            "remise_fidelite": art["remise_fidelite"], "total_ligne_ht": total_ht
        })

# --- ONGLET GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.subheader("📊 Récapitulatif Général & Devis Professionnel ABCOM")

    if not lignes_devis_global:
        st.warning("Veuillez renseigner au moins un article avec une quantité supérieure à 0.")
    else:
        sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
        montant_base_port = sous_total_articles + frais_techniques_dossier
        
        frais_port = 0.0 if offrir_port else calculer_frais_port(montant_base_port, zone_livraison)
        total_general_ht = montant_base_port + frais_port
        tva = total_general_ht * 0.20
        total_ttc = total_general_ht + tva
        quantite_globale_totale = sum([item["quantite"] for item in lignes_devis_global])

        st.write(f"**Quantité globale pièces / articles :** {quantite_globale_totale}")
        st.write(f"**Sous-Total Articles & Imprimés HT :** {sous_total_articles:.2f} €")
        st.write(f"**Frais techniques :** {frais_techniques_dossier:.2f} € HT")
        st.write(f"**Frais de port ({zone_livraison}) :** {frais_port:.2f} € HT" if not offrir_port else "**Frais de port :** Offerts (0.00 €)")
        st.markdown(f"### **Total Général HT : {total_general_ht:.2f} €** | **TOTAL TTC (20%) : {total_ttc:.2f} €**")

        if st.button("📄 Générer le numéro de devis et le PDF ABCOM"):
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
                with open(logo_path, "wb") as f: f.write(logo_file.getbuffer())
            elif os.path.exists(logo_defaut_github):
                logo_path = logo_defaut_github

            header_text = Paragraph(
                "<b>SAS ABCOM - AGENCE DE COMMUNICATION VISUELLE</b><br/>"
                "Fondée en 1997 • Print, Signalétique, Marquage Textile & Objets Publicitaires<br/>"
                "Plasne (Jura) • Tél : 03 84 70 00 00<br/>"
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
            order_info_text = f"<b>DÉTAILS DE LA COMMANDE :</b><br/>Quantité globale : {quantite_globale_totale} unités<br/>Émetteur : Brice Geny (ABCOM)"
            
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
                if item["type"] == "Textile":
                    lib = f"<b>Support : {item['nom_article']}</b>" if not item['sans_marquage'] else f"<b>Support (Sans marquage) : {item['nom_article']}</b>"
                    table_data.append([Paragraph(lib, style_cell), str(item['quantite']), f"{item['prix_unit']:.2f} €", f"{item['prix_unit']*item['quantite']:.2f} €"])
                    for m in item.get('marquages', []):
                        table_data.append([Paragraph(f"&nbsp;&nbsp;&bull; Marquage : {m['nom']}", style_cell), str(item['quantite']), f"{m['tarif']:.2f} €", f"{m['tarif']*item['quantite']:.2f} €"])
                    if item.get('frais_prog', 0) > 0:
                        table_data.append([Paragraph("&nbsp;&nbsp;&bull; Frais techniques & programme Broderie", style_cell), "1", f"{item['frais_prog']:.2f} €", f"{item['frais_prog']:.2f} €"])
                    if item.get('option_ensachage'):
                        table_data.append([Paragraph(f"&nbsp;&nbsp;&bull; Option : {item['type_sachet']}", style_cell), str(item['quantite']), f"{item['coût_ensachage_unit']:.2f} €", f"{item['coût_ensachage_unit']*item['quantite']:.2f} €"])
                    if item.get('option_assurance'):
                        table_data.append([Paragraph("&nbsp;&nbsp;&bull; Option : Assurance MHC", style_cell), str(item['quantite']), f"{item['coût_assurance_unit']:.2f} €", f"{item['coût_assurance_unit']*item['quantite']:.2f} €"])
                    if item.get('option_stockage'):
                        table_data.append([Paragraph("&nbsp;&nbsp;&bull; Option : Stockage + picking", style_cell), str(item['quantite']), f"{item['coût_stockage_unit']:.2f} €", f"{item['coût_stockage_unit']*item['quantite']:.2f} €"])
                else:
                    table_data.append([Paragraph(f"<b>{item['nom_article']}</b>", style_cell), str(item['quantite']), f"{item['prix_unit']:.2f} €", f"{item['total_ligne_ht']:.2f} €"])

                if item.get('remise_fidelite', 0) > 0:
                    brut_L = item['total_ligne_ht'] / (1.0 - item['remise_fidelite']/100.0)
                    montant_rem_l = brut_L * (item['remise_fidelite']/100.0)
                    table_data.append([Paragraph(f"&nbsp;&nbsp;&bull; <b>Remise / Promo ({item['remise_fidelite']}%)</b>", style_cell), "1", "-", f"-{montant_rem_l:.2f} €"])

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
            st.success(f"Devis PDF ABCOM généré sous le numéro : **{num_devis}** (`{pdf_filename}`)")

            # --- ENVOI GMAIL ---
            st.markdown(f"### ✉️ Envoi direct par E-mail (via {conseiller_email})")
            if st.button("🚀 Envoyer le devis par e-mail"):
                if not client_email: st.error("Veuillez renseigner l'e-mail du client.")
                elif not gmail_password: st.error("Mot de passe d'application Gmail manquant dans les secrets.")
                else:
                    try:
                        msg = EmailMessage()
                        msg['Subject'] = f"Devis ABCOM n° {num_devis}"
                        msg['From'] = conseiller_email
                        msg['To'] = client_email
                        
                        corps_mail = f"""Bonjour {client_nom or 'Client'},

Veuillez trouver ci-joint votre devis n° {num_devis} émis par la SAS ABCOM d'un montant total de {total_ttc:.2f} € TTC.

Restant à votre disposition,

Cordialement,
Brice Geny
SAS ABCOM
Tél : 03 84 70 00 00
Email : brice.geny@gmail.com
"""
                        msg.set_content(corps_mail)
                        with open(pdf_filename, 'rb') as f:
                            msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(pdf_filename))

                        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                            smtp.login(conseiller_email, gmail_password)
                            smtp.send_message(msg)
                        st.success(f"✅ E-mail envoyé avec succès à {client_email} !")
                    except Exception as e:
                        st.error(f"Erreur d'envoi : {e}")

# --- ONGLET SUIVI CRM ---
with onglets[11]:
    st.subheader("📈 Tableau de Suivi CRM ABCOM")
    if os.path.exists(CRM_FILE):
        df_crm = pd.read_csv(CRM_FILE)
        if not df_crm.empty:
            edited_df = st.data_editor(
                df_crm,
                column_config={
                    "Statut": st.column_config.SelectboxColumn("Statut du Devis", options=["En cours", "Accepté", "Refusé", "Sans suite"], required=True),
                    "Total_HT": st.column_config.NumberColumn("Total HT (€)", format="%.2f €"),
                    "Total_TTC": st.column_config.NumberColumn("Total TTC (€)", format="%.2f €"),
                },
                disabled=["Date", "Numero_Devis", "Conseiller", "Client", "Entreprise", "Email", "Telephone", "Quantite_Totale"],
                use_container_width=True,
                key="crm_editor"
            )
            if not edited_df.equals(df_crm):
                edited_df.to_csv(CRM_FILE, index=False)
                st.toast("✅ Statuts mis à jour !", icon="💾")
        else:
            st.info("Aucun devis enregistré pour le moment.")
    else:
        st.info("Le CRM est vide.")