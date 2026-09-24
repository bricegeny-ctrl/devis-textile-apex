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

st.set_page_config(page_title="Gestionnaire de Devis - SAS ABCOM", layout="wide")

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
                dernier_num = saved_data.get("dernier_num", f"DEV-{annee_courante}-0000")
                parts = dernier_num.split("-")
                if len(parts) == 3 and parts[1] == annee_courante:
                    seq = int(parts[2]) + 1
        except:
            seq = 1
            
    nouveau_num = f"DEV-{annee_courante}-{seq:04d}"
    with open(COMPTEUR_FILE, "w") as f:
        json.dump({"dernier_num": nouveau_num}, f)
    return nouveau_num

def enregistrer_dans_crm(data_devis):
    df_new = pd.DataFrame([data_devis])
    if os.path.exists(CRM_FILE):
        df_exist = pd.read_csv(CRM_FILE)
        df_combined = pd.concat([df_exist, df_new], ignore_index=True)
    else:
        df_combined = df_new
    df_combined.to_csv(CRM_FILE, index=False)

def mettre_a_jour_statut_crm(numero_devis, nouveau_statut):
    if os.path.exists(CRM_FILE):
        df_crm = pd.read_csv(CRM_FILE)
        if "Numero_Devis" in df_crm.columns and "Statut" in df_crm.columns:
            df_crm.loc[df_crm["Numero_Devis"] == numero_devis, "Statut"] = nouveau_statut
            df_crm.to_csv(CRM_FILE, index=False)

# --- FONCTIONS DE CALCUL DES TARIFS & TRANCHES (TEXTILE) ---
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
        "Dos Large D20 (25x20)": [23.20, 19.30, 15.90, 12.80, 10.90, 10.20, 9.60, 9.10, 8.60, 8.00],
        "Col/Signature (7x2)": [11.50, 8.80, 6.50, 4.80, 3.60, 3.30, 3.10, 2.90, 2.70, 2.40],
        "Casquettes / Bonnets": [13.20, 10.10, 7.60, 5.70, 4.40, 4.00, 3.70, 3.50, 3.20, 2.90],
        "Manche (8x5 cm)": [13.20, 10.10, 7.60, 5.70, 4.40, 4.00, 3.70, 3.50, 3.20, 2.90],
        "Pantalon / Poche": [13.20, 10.10, 7.60, 5.70, 4.40, 4.00, 3.70, 3.50, 3.20, 2.90],
        "+ Perso. Nom (Cœur)": [6.80, 5.20, 4.00, 3.00, 2.30, 2.10, 2.00, 1.80, 1.70, 1.50]
    }
    return tarifs.get(emplacement, [0]*10)[idx]

def get_frais_prog_broderie(qte):
    return 0.0 if qte >= 25 else 35.0

def get_tarif_assurance(qte):
    if qte <= 50: return qte * 0.20
    elif qte <= 200: return qte * 0.15
    else: return qte * 0.10

def get_tarif_ensachage(qte, type_sachet):
    tarif_unit = 0.35 if "T-shirt" in type_sachet else 0.45
    return qte * tarif_unit

def get_tarif_stockage(qte):
    return qte * 0.10

def calculer_frais_port(montant_ht, zone):
    if zone == "France métropolitaine":
        if montant_ht < 150: return 15.0
        elif montant_ht < 400: return 25.0
        else: return 0.0
    elif zone == "Europe":
        return 45.0
    else:
        return 0.0

# --- CHARGEMENT DU CRM POUR L'AUTO-COMPLÉTION ---
clients_connus = {}
if os.path.exists(CRM_FILE):
    try:
        df_crm_load = pd.read_csv(CRM_FILE)
        for _, row in df_crm_load.iterrows():
            nom_cli = row.get("Client", "")
            if nom_cli and nom_cli not in clients_connus:
                clients_connus[nom_cli] = {
                    "entreprise": row.get("Entreprise", ""),
                    "email": row.get("Email", ""),
                    "telephone": row.get("Telephone", ""),
                    "adresse": row.get("Adresse", "")
                }
    except:
        pass

st.title("🖨️ SAS ABCOM - Gestionnaire Professionnel de Devis")
st.markdown("Interface unifiée : **Textile**, **Print & Petits Formats**, et **Signalétique & Grands Formats**.")

# --- NAVIGATION / ONGLES ---
noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"]
onglets = st.tabs(noms_onglets)

articles_saisis = []

for i in range(10):
    with onglets[i]:
        st.header(f"Configuration Article {i+1}")
        
        type_produit = st.selectbox(
            f"Type de produit (Article {i+1})",
            ["Marquage Textile", "Flyers", "Dépliants", "Cartes de Visite", "Blocs-notes", "Chemises de présentation", "Calendriers", "Sous-bocks", "Adhésifs", "Banderoles", "Panneaux Akilux", "Roll-Up"],
            key=f"type_prod_{i}"
        )
        
        art_data = {"id": i+1, "type_produit": type_produit}
        
        if type_produit == "Marquage Textile":
            col1, col2 = st.columns(2)
            with col1:
                art_data["nom"] = st.text_input(f"Désignation / Modèle textile {i+1}", key=f"nom_textile_{i}")
                art_data["qte"] = st.number_input(f"Quantité articles {i+1}", min_value=1, value=10, key=f"qte_{i}")
                art_data["prix_unit_support"] = st.number_input(f"Prix unitaire HT du support vierge {i+1} (€)", min_value=0.0, value=3.50, step=0.1, key=f"pus_{i}")
            
            with col2:
                nb_marquages = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], key=f"nb_m_{i}")
                art_data["marquages"] = []
                for m in range(nb_marquages):
                    st.markdown(f"**Marquage M{m+1}**")
                    t_marq = st.selectbox(f"Technique M{m+1} (Art {i+1})", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
                    if t_marq == "DTF Textile Fin":
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_f_{i}_{m}")
                    elif t_marq == "DTF Textile Épais":
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_e_{i}_{m}")
                    else:
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20)", "Col/Signature (7x2)", "Casquettes / Bonnets", "Manche (8x5 cm)", "Pantalon / Poche", "+ Perso. Nom (Cœur)"], key=f"emp_b_{i}_{m}")
                    art_data["marquages"].append({"technique": t_marq, "emplacement": emp})
            
            # Options additionnelles textile
            with st.expander(f"Options & Finitions Textile (Article {i+1})"):
                art_data["opt_ensachage"] = st.checkbox(f"Ensachage individuel {i+1}", key=f"ens_{i}")
                art_data["type_sachet"] = st.selectbox(f"Type de sachet {i+1}", ["Sachet (T-shirt/Polo)", "Sachet (Veste/Sweat)"], key=f"tsach_{i}") if art_data["opt_ensachage"] else ""
                art_data["opt_assurance"] = st.checkbox(f"Assurance Sérénité (100% Satisfait ou Refait) {i+1}", key=f"ass_{i}")
                art_data["opt_stockage"] = st.checkbox(f"Stockage déporté sécurisé {i+1}", key=f"stk_{i}")

            articles_saisis.append(art_data)

        elif type_produit == "Flyers":
            col1, col2 = st.columns(2)
            with col1:
                art_data["format"] = st.selectbox(f"Format {i+1}", ["A6", "A5"], key=f"fly_fmt_{i}")
                art_data["papier"] = st.selectbox(f"Papier {i+1}", ["135g couché brillant", "170g couché demi mat", "250g couché brillant ou demi mat", "350g couché demi mat", "115g recyclé"], key=f"fly_pap_{i}")
            with col2:
                art_data["faces"] = st.selectbox(f"Impression {i+1}", ["Recto", "Recto/verso"], key=f"fly_fac_{i}")
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [100, 250, 500, 1000, 2500, 5000], key=f"fly_qte_{i}")
            
            # Grille tarifaire indicative Flyers
            # Logique de prix fixe selon grille ou prix unitaire approximatif
            base_prix = 35.0
            if art_data["format"] == "A5": base_prix = 40.0
            if "verso" in art_data["faces"]: base_prix *= 1.2
            art_data["prix_total_ht"] = base_prix * (art_data["quantite"] / 100)
            articles_saisis.append(art_data)

        elif type_produit == "Dépliants":
            col1, col2 = st.columns(2)
            with col1:
                art_data["pliage"] = st.selectbox(f"Type de dépliant {i+1}", [
                    "A6 fermé/A5 ouvert (1 pli)", 
                    "A5 fermé/A4 ouvert (1 pli)", 
                    "A4 fermé/A3 ouvert (1 pli)", 
                    "A4 fermé/A3 ouvert (2 plis rouleau/accordéon)"
                ], key=f"dep_pli_{i}")
                art_data["papier"] = st.selectbox(f"Papier {i+1}", ["135g couché brillant", "170g couché demi mat", "250g couché brillant ou demi mat", "130g recyclé"], key=f"dep_pap_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [100, 250, 500, 1000], key=f"dep_qte_{i}")
            art_data["prix_total_ht"] = 75.0 * (art_data["quantite"] / 100)
            articles_saisis.append(art_data)

        elif type_produit == "Cartes de Visite":
            col1, col2 = st.columns(2)
            with col1:
                art_data["pelliculage"] = st.selectbox(f"Pelliculage {i+1}", ["Sans pelliculage", "Avec pelliculage (Brillant/Mat/Velours)"], key=f"cv_pel_{i}")
                art_data["faces"] = st.selectbox(f"Faces {i+1}", ["Recto", "Recto/Verso"], key=f"cv_fac_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [100, 250, 500, 1000], key=f"cv_qte_{i}")
            
            p_base = 19.9 if "Recto" in art_data["faces"] else 29.9
            if "Avec" in art_data["pelliculage"]: p_base += 15.0
            art_data["prix_total_ht"] = p_base * (art_data["quantite"] / 100)
            articles_saisis.append(art_data)

        elif type_produit == "Blocs-notes":
            col1, col2 = st.columns(2)
            with col1:
                art_data["modele"] = st.selectbox(f"Modèle Bloc-note {i+1}", [
                    "BNA6-25 (A6 - 25 Feuillets)", "BNA6-50 (A6 - 50 Feuillets)",
                    "BNA5-25 (A5 - 25 Feuillets)", "BNA5-50 (A5 - 50 Feuillets)",
                    "BNA4-25 (A4 - 25 Feuillets)", "BNA4-50 (A4 - 50 Feuillets)"
                ], key=f"bn_mod_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [50, 100, 200, 500, 1000], key=f"bn_qte_{i}")
            
            pu_bn = 1.20
            if "50" in art_data["modele"]: pu_bn = 1.60
            if "A5" in art_data["modele"]: pu_bn = 2.30
            if "A4" in art_data["modele"]: pu_bn = 3.50
            art_data["prix_total_ht"] = pu_bn * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Chemises de présentation":
            col1, col2 = st.columns(2)
            with col1:
                art_data["modele"] = st.selectbox(f"Modèle Chemise {i+1}", [
                    "CPA5R-SR (A4 Ouvert / A5 Fermé - Recto)",
                    "CPA5RV-SR (A4 Ouvert / A5 Fermé - Recto/Verso)",
                    "CPA4R-SR (A3 Ouvert / A4 Fermé - Recto)",
                    "CPA4RV-SR (A3 Ouvert / A4 Fermé - Recto/Verso)"
                ], key=f"chem_mod_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [100, 250, 500, 1000], key=f"chem_qte_{i}")
            pu_chem = 1.0 if "Recto" in art_data["modele"] else 1.3
            art_data["prix_total_ht"] = pu_chem * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Calendriers":
            col1, col2 = st.columns(2)
            with col1:
                art_data["modele"] = st.selectbox(f"Modèle Calendrier {i+1}", ["Calendrier Mural A4 / A5 / A6", "Calendrier Magnétique A5"], key=f"cal_mod_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [50, 100, 250, 500, 1000], key=f"cal_qte_{i}")
            art_data["prix_total_ht"] = 1.50 * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Sous-bocks":
            col1, col2 = st.columns(2)
            with col1:
                art_data["forme"] = st.selectbox(f"Forme Sous-bock {i+1}", ["9,3x9,3 cm carré", "10x10 cm coins arrondis", "Rond (10 cm diam.)"], key=f"sb_form_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [10, 25, 50, 100, 250], key=f"sb_qte_{i}")
            art_data["prix_total_ht"] = 4.10 * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Adhésifs":
            col1, col2 = st.columns(2)
            with col1:
                art_data["support"] = st.selectbox(f"Support Adhésif {i+1}", ["Adhésif vinyl classique", "Adhésif papier couché chrome brillant 80g"], key=f"adh_sup_{i}")
                art_data["taille"] = st.selectbox(f"Format Adhésif {i+1}", ["10x10 cm", "20x20 cm", "29,5x29,5 cm"], key=f"adh_tail_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité {i+1}", [1, 10, 50, 100], key=f"adh_qte_{i}")
            art_data["prix_total_ht"] = 35.0 * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Banderoles":
            col1, col2 = st.columns(2)
            with col1:
                art_data["format"] = st.selectbox(f"Format Banderole {i+1} (510g M1)", [
                    "200 x 80 cm (75,38 € HT)",
                    "200 x 100 cm (87,50 € HT)",
                    "300 x 100 cm (124,25 € HT)",
                    "400 x 100 cm (150,50 € HT)",
                    "500 x 100 cm (176,75 € HT)"
                ], key=f"band_fmt_{i}")
            with col2:
                art_data["quantite"] = st.number_input(f"Quantité Banderoles {i+1}", min_value=1, value=1, key=f"band_qte_{i}")
            
            prix_base_band = 75.38
            if "200 x 100" in art_data["format"]: prix_base_band = 87.50
            elif "300 x 100" in art_data["format"]: prix_base_band = 124.25
            elif "400 x 100" in art_data["format"]: prix_base_band = 150.50
            elif "500 x 100" in art_data["format"]: prix_base_band = 176.75
            
            art_data["prix_total_ht"] = prix_base_band * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Panneaux Akilux":
            col1, col2 = st.columns(2)
            with col1:
                art_data["format"] = st.selectbox(f"Format Panneau {i+1}", [
                    "60 x 40 cm (ép. 3,5 mm)",
                    "80 x 60 cm (ép. 3,5 mm)",
                    "120 x 80 cm (ép. 3,5 mm)"
                ], key=f"aky_fmt_{i}")
                art_data["faces"] = st.selectbox(f"Impression Panneau {i+1}", ["Recto", "Recto-verso"], key=f"aky_fac_{i}")
            with col2:
                art_data["quantite"] = st.selectbox(f"Quantité Panneaux {i+1}", [1, 5, 10, 25, 50], key=f"aky_qte_{i}")
            
            p_aky = 9.20
            if "80 x 60" in art_data["format"]: p_aky = 12.40
            elif "120 x 80" in art_data["format"]: p_aky = 20.00
            if "verso" in art_data["faces"]: p_aky *= 1.4
            
            art_data["prix_total_ht"] = p_aky * art_data["quantite"]
            articles_saisis.append(art_data)

        elif type_produit == "Roll-Up":
            col1, col2 = st.columns(2)
            with col1:
                art_data["modele"] = st.selectbox(f"Modèle Roll-Up {i+1}", [
                    "RU-ECOB (Eco - PVC 510g M1 - Recto)",
                    "RU-ECOT (Eco - Polyester 420g M1 - Recto)",
                    "RU-PROB (Pro - PVC Mat 510g M1 - Recto)",
                    "RU-PROT (Pro - Polyester 420g M1 - Recto)"
                ], key=f"rup_mod_{i}")
            with col2:
                art_data["quantite"] = st.number_input(f"Quantité Roll-Up {i+1}", min_value=1, value=1, key=f"rup_qte_{i}")
            
            p_rup = 64.80
            if "ECOT" in art_data["modele"]: p_rup = 93.60
            elif "PROB" in art_data["modele"]: p_rup = 96.00
            elif "PROT" in art_data["modele"]: p_rup = 141.60
            
            art_data["prix_total_ht"] = p_rup * art_data["quantite"]
            articles_saisis.append(art_data)

# --- ONGLET 10 : GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.header("📊 Récapitulatif Général & Édition du Devis")
    
    st.subheader("1. Informations Client")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        nom_client_input = st.selectbox("Sélectionner un client existant ou saisir", [""] + list(clients_connus.keys()), key="select_client_crm")
        if nom_client_input and nom_client_input in clients_connus:
            infos_defaut = clients_connus[nom_client_input]
        else:
            infos_defaut = {"entreprise": "", "email": "", "telephone": "", "adresse": ""}
            
        client_nom = st.text_input("Nom du client", value=nom_client_input if nom_client_input else "", key="client_nom_field")
        client_entreprise = st.text_input("Entreprise", value=infos_defaut["entreprise"], key="client_ent_field")
    with col_c2:
        client_email = st.text_input("Email", value=infos_defaut["email"], key="client_email_field")
        client_telephone = st.text_input("Téléphone", value=infos_defaut["telephone"], key="client_tel_field")
        client_adresse = st.text_area("Adresse de livraison / facturation", value=infos_defaut["adresse"], key="client_adr_field")

    st.markdown("---")
    st.subheader("2. Paramètres Logistiques & Remises")
    col_l1, col_l2, col_l3 = st.columns(3)
    with col_l1:
        zone_port = st.selectbox("Zone de Livraison", ["France métropolitaine", "Europe", "Retrait atelier (Plasne)"])
    with col_l2:
        remise_globale_pct = st.number_input("Remise globale (%)", min_value=0.0, max_value=100.0, value=0.0, step=1.0)
    with col_l3:
        mode_reglement = st.selectbox("Modalités de règlement", ["30 jours fin de mois", "Comptant à la commande", "50% à la commande, 50% à livraison", "Virement bancaire 10 jours"])

    # --- CALCULS GLOBAUX ---
    total_ht_global = 0.0
    lignes_devis_finales = []
    
    for art in articles_saisis:
        if art["type_produit"] == "Marquage Textile":
            qte = art["qte"]
            pus = art["prix_unit_support"]
            frais_support_tot = qte * pus
            total_marquages_art = 0.0
            
            desc_marq = []
            for m in art["marquages"]:
                tech = m["technique"]
                emp = m["emplacement"]
                if "DTF Fin" in tech:
                    t_unit = get_tarif_dtf_fin(qte, emp)
                elif "DTF Épais" in tech:
                    t_unit = get_tarif_dtf_epais(qte, emp)
                else:
                    t_unit = get_tarif_broderie(qte, emp)
                total_marquages_art += t_unit * qte
                desc_marq.append(f"{tech} ({emp}) : {t_unit:.2f}€/u")
            
            frais_prog = get_frais_prog_broderie(qte) if any("Broderie" in m["technique"] for m in art["marquages"]) else 0.0
            frais_ass = get_tarif_assurance(qte) if art["opt_assurance"] else 0.0
            frais_ens = get_tarif_ensachage(qte, art["type_sachet"]) if art["opt_ensachage"] else 0.0
            frais_stk = get_tarif_stockage(qte) if art["opt_stockage"] else 0.0
            
            total_art_ht = frais_support_tot + total_marquages_art + frais_prog + frais_ass + frais_ens + frais_stk
            total_ht_global += total_art_ht
            
            lignes_devis_finales.append({
                "Article": f"Textile {art.get('nom', '')} (Qte: {qte})",
                "Details": " + ".join(desc_marq),
                "Total HT": total_art_ht
            })
        else:
            t_ht = art.get("prix_total_ht", 0.0)
            total_ht_global += t_ht
            lignes_devis_finales.append({
                "Article": f"{art['type_produit']} (Qte: {art.get('quantite', art.get('qte', 1))})",
                "Details": str(art.get("format", art.get("modele", art.get("support", "")))),
                "Total HT": t_ht
            })

    # Application frais de port et remise
    frais_port = calculer_frais_port(total_ht_global, zone_port)
    montant_remise = total_ht_global * (remise_globale_pct / 100.0)
    total_net_ht = total_ht_global - montant_remise + frais_port
    montant_tva = total_net_ht * 0.20
    total_ttc = total_net_ht + montant_tva

    st.markdown("---")
    st.subheader("3. Synthèse Financière")
    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
    col_s1.metric("Total HT", f"{total_ht_global:.2f} €")
    col_s2.metric("Frais de port", f"{frais_port:.2f} €")
    col_s3.metric("TVA (20%)", f"{montant_tva:.2f} €")
    col_s4.metric("Total TTC", f"{total_ttc:.2f} €")

    if st.button("✨ Générer le Numéro de Devis & Enregistrer", type="primary"):
        numero_devis = obtenir_prochain_numero_devis()
        st.success(f"Devis **{numero_devis}** généré avec succès !")
        
        data_crm = {
            "Numero_Devis": numero_devis,
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Client": client_nom,
            "Entreprise": client_entreprise,
            "Email": client_email,
            "Telephone": client_telephone,
            "Adresse": client_adresse,
            "Total_TTC": round(total_ttc, 2),
            "Statut": "En attente"
        }
        enregistrer_dans_crm(data_crm)
        st.info("💾 Devis enregistré dans le CRM interactif.")

# --- ONGLET 11 : SUIVI CRM INTERACTIF ---
with onglets[11]:
    st.header("📈 Suivi CRM & Historique des Devis")
    if os.path.exists(CRM_FILE):
        df_crm = pd.read_csv(CRM_FILE)
        st.dataframe(df_crm, use_container_width=True)
        
        st.markdown("---")
        st.subheader("Mise à jour des statuts")
        col_u1, col_u2 = st.columns(2)
        with col_u1:
            devis_sel = st.selectbox("Sélectionner un devis", df_crm["Numero_Devis"].tolist() if "Numero_Devis" in df_crm.columns else [])
        with col_u2:
            nouveau_st = st.selectbox("Nouveau statut", ["En attente", "Validé", "Facturé", "Annulé"])
            if st.button("Mettre à jour le statut"):
                mettre_a_jour_statut_crm(devis_sel, nouveau_st)
                st.success("Statut mis à jour ! Rechargez la page si nécessaire.")
    else:
        st.info("Aucun devis enregistré pour le moment dans le CRM.")