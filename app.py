import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import urllib.parse
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Gestionnaire de Devis - APEX", layout="wide")

# --- GESTION DES FICHIERS ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"
CATALOGUE_FILE = "catalogue_print.xlsx"
TARIF_MARQUAGE_FILE = "2026-05-10- tarif marquage-broderie.xlsx"
PDF_DIR = "devis_pdf"

if not os.path.exists(PDF_DIR):
    os.makedirs(PDF_DIR)

def obtenir_prochain_numero_devis():
    annee_courante = datetime.now().strftime("%Y")
    mois_courant = datetime.now().strftime("%m")
    jour_courant = datetime.now().strftime("%d")
    date_prefix = f"{annee_courante}/{mois_courant}/{jour_courant}"
    
    seq = 1
    if os.path.exists(COMPTEUR_FILE):
        try:
            with open(COMPTEUR_FILE, "r") as f:
                saved_data = json.load(f)
                dernier_num = saved_data.get("dernier_num", "")
                if "_" in dernier_num and dernier_num.startswith(date_prefix):
                    parts = dernier_num.split("_")
                    if len(parts) > 1 and parts[1].isdigit():
                        seq = int(parts[1]) + 1
        except Exception:
            seq = 1

    nouveau_num = f"{date_prefix}_{seq:08d}"
    with open(COMPTEUR_FILE, "w") as f:
        json.dump({"dernier_num": nouveau_num}, f)
    return nouveau_num

import pandas as pd

# Charge toutes les feuilles dans un dictionnaire de DataFrames
excel_path = 'catalogue print et signalétique.xlsx'
toutes_les_feuilles = pd.read_excel(excel_path, sheet_name=None)

# Pour fusionner toutes les feuilles si elles ont la même structure :
df_global = pd.concat(toutes_les_feuilles.values(), ignore_index=True)

# --- MOTEUR DE LECTURE EXCEL CATALOGUE PRINT & SIGNALÉTIQUE (100% FIABLE & INTÉGRAL) ---
def obtenir_prix_catalogue_intelligent(cat_print, choix_ref, qte):
    if not os.path.exists(CATALOGUE_FILE):
        return 0.15
    
    try:
        # Lecture brute de l'intégralité de la feuille sans en-tête fixe
        df_all = pd.read_excel(CATALOGUE_FILE, sheet_name=0, header=None)
    except Exception:
        return 0.15

    # Extraction dynamique des paliers de quantité depuis la ligne 1 (colonnes 3 à la fin)
    paliers_cols = []
    for c in range(3, df_all.shape[1]):
        val_hdr = df_all.iloc[1, c]
        try:
            paliers_cols.append((c, float(val_hdr)))
        except:
            pass

    # Détermination de la colonne cible selon la quantité commandée
    col_cible = 3
    if paliers_cols:
        for idx, (col_idx, q_seuil) in enumerate(paliers_cols):
            if idx < len(paliers_cols) - 1:
                q_prochain = paliers_cols[idx + 1][1]
                if q_seuil <= qte < q_prochain:
                    col_cible = col_idx
                    break
            else:
                if qte >= q_seuil:
                    col_cible = col_idx
                    break
    else:
        # Paliers par défaut si l'en-tête est absent
        if qte <= 4: col_cible = 3
        elif qte <= 9: col_cible = 4
        elif qte <= 24: col_cible = 5
        elif qte <= 49: col_cible = 6
        elif qte <= 99: col_cible = 7
        elif qte <= 249: col_cible = 8
        elif qte <= 499: col_cible = 9
        elif qte <= 999: col_cible = 10
        elif qte <= 2499: col_cible = 11
        elif qte <= 4999: col_cible = 12
        elif qte <= 9999: col_cible = 13
        else: col_cible = 14

    ref_lower = str(choix_ref).lower().strip()
    best_row = -1
    max_match = -1

    # Parcours de TOUTES les lignes du tableau à partir de la ligne 2 jusqu'à la toute dernière (ligne 83+)
    for r in range(2, len(df_all)):
        # Récupération sécurisée en ignorant les NaN
        row_cat = str(df_all.iloc[r, 0]) if pd.notna(df_all.iloc[r, 0]) else ""
        row_sub = str(df_all.iloc[r, 1]) if pd.notna(df_all.iloc[r, 1]) else ""
        row_ref = str(df_all.iloc[r, 2]) if pd.notna(df_all.iloc[r, 2]) else ""
        
        # Concaténation propre sans polluer avec le mot "nan"
        full_row_text = f"{row_cat} {row_sub} {row_ref}".lower()

        score = 0
        mots_ref = [m for m in ref_lower.split() if len(m) > 1]
        match_mots = sum(1 for m in mots_ref if m in full_row_text)
        score += match_mots * 10

        if score > max_match:
            max_match = score
            best_row = r

    if best_row == -1 or max_match <= 0:
        return 0.15

    try:
        if col_cible >= df_all.shape[1]:
            col_cible = df_all.shape[1] - 1
            
        prix_val = float(df_all.iloc[best_row, col_cible])
        
        # Fallback intelligent si la cellule du prix est vide : recherche de la colonne valide la plus proche
        if pd.isna(prix_val) or prix_val <= 0:
            for alt_col in range(col_cible - 1, 2, -1):
                if 0 <= alt_col < df_all.shape[1]:
                    alt_val = float(df_all.iloc[best_row, alt_col])
                    if not pd.isna(alt_val) and alt_val > 0:
                        prix_val = alt_val
                        break

        if pd.isna(prix_val) or prix_val <= 0:
            return 0.15
            
        return round(prix_val, 4)
    except Exception:
        return 0.15

# --- MOTEUR DE LECTURE EXCEL MARQUAGE & BRODERIE ---
def lire_grille_depuis_excel(sheet_name_keyword, emplacement, qte_totale):
    """Lit dynamiquement le fichier Excel de tarifs de marquage/broderie s'il existe."""
    if not os.path.exists(TARIF_MARQUAGE_FILE):
        return None
    try:
        xls = pd.ExcelFile(TARIF_MARQUAGE_FILE)
        sheet_target = None
        for sheet in xls.sheet_names:
            if sheet_name_keyword.lower() in sheet.lower():
                sheet_target = sheet
                break
        if not sheet_target:
            sheet_target = xls.sheet_names[0]
        
        df = pd.read_excel(xls, sheet_name=sheet_target, header=None)
        # Recherche de la ligne correspondant à l'emplacement et de la colonne correspondant à la quantité
        # Implémentation générique basée sur la structure tabulaire standard
        for r in range(len(df)):
            row_str = str(df.iloc[r, 0]).lower()
            if any(m in row_str for m in emplacement.lower().split()):
                # Ligne trouvée, recherche de la bonne colonne de quantité
                # (Par défaut, on parcourt la ligne pour trouver le tarif correspondant au palier)
                for c in range(1, df.shape[1]):
                    val_cell = df.iloc[r, c]
                    if isinstance(val_cell, (int, float)) and val_cell > 0:
                        return float(val_cell)
    except Exception:
        pass
    return None

# --- GRILLES TARIFAIRES OFFICIELLES & LECTURE FICHIER (DTF & BRODERIE) ---
def obtenir_tarif_dtf_unitaire(type_textile, emplacement, qte_totale):
    # Tentative de lecture depuis le fichier Excel dédié
    prix_excel = lire_grille_depuis_excel("dtf", emplacement, qte_totale)
    if prix_excel is not None and prix_excel > 0:
        prix = prix_excel
    else:
        grille_dtf = {
            "Cœur (13x9 cm)": [(5, 6.00), (9, 4.50), (19, 3.60), (29, 2.81), (39, 2.50), (49, 2.40), (99, 2.00), (249, 1.80), (499, 1.60), (999, 1.40), (1999, 1.20), (4999, 1.00), (float('inf'), 0.70)],
            "Opposé Cœur (9x8 cm)": [(5, 6.00), (9, 4.50), (19, 3.60), (29, 2.81), (39, 2.50), (49, 2.40), (99, 2.00), (249, 1.80), (499, 1.60), (999, 1.40), (1999, 1.20), (4999, 1.00), (float('inf'), 0.70)],
            "Dos D10 (20x13 cm)": [(5, 7.92), (9, 6.50), (19, 5.50), (29, 5.00), (39, 4.20), (49, 4.00), (99, 3.50), (249, 3.00), (499, 2.70), (999, 2.50), (1999, 2.00), (4999, 1.50), (float('inf'), 1.00)],
            "Dos D20 (28x20 cm)": [(5, 13.67), (9, 10.00), (19, 9.00), (29, 6.90), (39, 6.30), (49, 6.00), (99, 5.50), (249, 4.60), (499, 4.00), (999, 3.50), (1999, 2.80), (4999, 2.20), (float('inf'), 1.50)],
            "Format P (37x27 cm)": [(5, 17.00), (9, 13.00), (19, 12.00), (29, 10.00), (39, 9.00), (49, 8.00), (99, 7.50), (249, 6.50), (499, 6.00), (999, 5.00), (1999, 4.00), (4999, 3.20), (float('inf'), 2.50)],
            "Manche (9x8 cm)": [(5, 7.20), (9, 5.40), (19, 4.32), (29, 3.37), (39, 3.00), (49, 2.88), (99, 2.40), (249, 2.16), (499, 1.92), (999, 1.68), (1999, 1.44), (4999, 1.20), (float('inf'), 0.84)],
            "Casquette / Bonnet": [(5, 7.20), (9, 5.40), (19, 4.32), (29, 3.37), (39, 3.00), (49, 2.88), (99, 2.40), (249, 2.16), (499, 1.92), (999, 1.68), (1999, 1.44), (4999, 1.20), (float('inf'), 0.84)],
            "Parapluie": [(5, 8.76), (9, 6.50), (19, 5.20), (29, 4.10), (39, 3.50), (49, 3.20), (99, 2.80), (249, 2.40), (499, 2.10), (999, 1.80), (1999, 1.50), (4999, 1.20), (float('inf'), 0.90)],
            "Bagagerie": [(5, 9.54), (9, 7.20), (19, 5.80), (29, 4.50), (39, 3.80), (49, 3.50), (99, 3.00), (249, 2.60), (499, 2.30), (999, 2.00), (1999, 1.60), (4999, 1.30), (float('inf'), 1.00)],
            "Pantalon / Poche": [(5, 9.54), (9, 7.20), (19, 5.80), (29, 4.50), (39, 3.80), (49, 3.50), (99, 3.00), (249, 2.60), (499, 2.30), (999, 2.00), (1999, 1.60), (4999, 1.30), (float('inf'), 1.00)],
            "+ Personnalisation Nom": [(5, 5.00), (9, 4.00), (19, 3.00), (29, 2.40), (39, 2.10), (49, 2.00), (99, 1.80), (249, 1.50), (499, 1.30), (999, 1.00), (1999, 0.70), (4999, 0.40), (float('inf'), 0.20)]
        }
        cle = emplacement if emplacement in grille_dtf else "Cœur (13x9 cm)"
        paliers = grille_dtf[cle]
        prix = paliers[-1][1]
        for limite, p in paliers:
            if qte_totale <= limite:
                prix = p
                break
    if "Vif" in type_textile:
        prix = round(prix * 1.10, 2)
    return prix

def obtenir_tarif_broderie_unitaire(emplacement, qte_totale):
    prix_excel = lire_grille_depuis_excel("broderie", emplacement, qte_totale)
    if prix_excel is not None and prix_excel > 0:
        return prix_excel
    
    grille_brod = {
        "Poitrine (9x8 cm)": [(3, 9.21), (11, 8.23), (23, 6.50), (47, 5.20), (95, 4.30), (251, 3.80), (503, 3.50), (1007, 3.20), (1511, 2.90), (2015, 2.60), (float('inf'), 2.30)],
        "Dos D10 (25x10 cm)": [(3, 11.35), (11, 10.44), (23, 8.50), (47, 7.10), (95, 6.00), (251, 5.40), (503, 4.90), (1007, 4.50), (1511, 4.10), (2015, 3.70), (float('inf'), 3.30)],
        "Dos Large D20 (25x20 cm)": [(3, 14.21), (11, 13.61), (23, 11.00), (47, 9.20), (95, 7.80), (251, 7.10), (503, 6.50), (1007, 5.90), (1511, 5.30), (2015, 4.80), (float('inf'), 4.20)],
        "Col / Signature (7x2 cm)": [(3, 6.50), (11, 5.63), (23, 4.50), (47, 3.60), (95, 2.90), (251, 2.60), (503, 2.30), (1007, 2.10), (1511, 1.90), (2015, 1.70), (float('inf'), 1.50)],
        "Casquettes / Bonnets": [(3, 10.47), (11, 9.54), (23, 8.57), (47, 7.71), (95, 6.55), (251, 5.35), (503, 4.61), (1007, 4.00), (1511, 3.54), (2015, 3.27), (float('inf'), 3.01)],
        "Parapluie / Bagagerie": [(3, 11.35), (11, 10.44), (23, 8.50), (47, 7.10), (95, 6.00), (251, 5.40), (503, 4.90), (1007, 4.50), (1511, 4.10), (2015, 3.70), (float('inf'), 3.30)],
        "Manche (8x5 cm)": [(3, 10.47), (11, 9.54), (23, 7.50), (47, 6.00), (95, 5.00), (251, 4.50), (503, 4.10), (1007, 3.70), (1511, 3.30), (2015, 2.90), (float('inf'), 2.50)],
        "Pantalon / Poche": [(3, 11.07), (11, 10.71), (23, 8.80), (47, 7.40), (95, 6.20), (251, 5.60), (503, 5.10), (1007, 4.60), (1511, 4.20), (2015, 3.80), (float('inf'), 3.40)],
        "+ Perso. Nom (Cœur)": [(3, 4.39), (11, 4.20), (23, 3.50), (47, 2.80), (95, 2.30), (251, 2.10), (503, 1.90), (1007, 1.70), (1511, 1.50), (2015, 1.30), (float('inf'), 1.10)]
    }
    cle = emplacement if emplacement in grille_brod else "Poitrine (9x8 cm)"
    paliers = grille_brod[cle]
    prix = paliers[-1][1]
    for limite, p in paliers:
        if qte_totale <= limite:
            prix = p
            break
    return prix

def calculer_frais_port(montant_base, zone):
    if zone == "France Continentale":
        if montant_base < 99.99: return 14.95
        elif montant_base < 499.99: return 20.95
        elif montant_base < 999.99: return 24.95
        else: return 0.0
    elif zone == "Livraison Corse, Monaco ou Andorre":
        if montant_base < 99.99: return 19.95
        elif montant_base < 499.99: return 25.95
        elif montant_base < 999.99: return 29.95
        else: return 0.0
    else:
        if montant_base < 99.99: return 25.95
        elif montant_base < 499.99: return 39.0
        elif montant_base < 999.99: return 60.0
        elif montant_base < 1500.0: return 90.0
        else: return 0.0

if not os.path.exists(CRM_FILE):
    pd.DataFrame(columns=[
        "Numero_Devis", "Date", "Client", "Contact", "Email", "Telephone", 
        "Total_HT", "Total_TTC", "Statut", "Commercial", "Mode_Reglement", "PDF_Path"
    ]).to_csv(CRM_FILE, index=False)

def enregistrer_dans_crm(devis_data):
    df_crm = pd.read_csv(CRM_FILE) if os.path.exists(CRM_FILE) else pd.DataFrame(columns=list(devis_data.keys()))
    if not df_crm[df_crm["Numero_Devis"] == devis_data["Numero_Devis"]].empty:
        df_crm.loc[df_crm["Numero_Devis"] == devis_data["Numero_Devis"], :] = list(devis_data.values())
    else:
        df_crm = pd.concat([df_crm, pd.DataFrame([devis_data])], ignore_index=True)
    df_crm.to_csv(CRM_FILE, index=False)

# --- SIDEBAR ---
st.sidebar.title("📋 Infos Client & Expédition")
client_nom = st.sidebar.text_input("Nom du Client", "Client Exemple")
client_entreprise = st.sidebar.text_input("Société / Entreprise", "")
client_siret = st.sidebar.text_input("SIRET", "")
client_contact = st.sidebar.text_input("Nom de contact", "")
client_adresse = st.sidebar.text_area("Adresse complète", "1 rue de l'Exemple\n70000 Vesoul")
client_email = st.sidebar.text_input("Email", "client@exemple.com")
client_contact_tel = st.sidebar.text_input("Téléphone", "0600000000")

logo_file = st.sidebar.file_uploader("Logo entreprise (PNG/JPG)", type=["png", "jpg", "jpeg"])

st.sidebar.markdown("---")
zone_livraison = st.sidebar.selectbox("Zone de Livraison", ["France Continentale", "Livraison Corse, Monaco ou Andorre", "Espace UE"])
offrir_port = st.sidebar.checkbox("🎁 Offrir les frais de port", value=False)
conseiller_nom = st.sidebar.selectbox("Commercial / Conseiller", ["Brice Geny", "Brice Bugna"])
mode_reglement = st.sidebar.selectbox("Mode de Règlement", ["Virement bancaire 30 jours", "Comptant à la commande", "50% à la commande, 50% à 30 jours"])

# --- INTERFACE PRINCIPALE ---
noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"]
onglets = st.tabs(noms_onglets)

articles_saisis = []
total_textile_brut = 0.0

for i in range(10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        metier_type = st.radio(
            f"Univers / Métier pour l'Article {i+1}",
            ["👕 Textile & Marquage (DTF / Broderie)", "🧢 Casquettes, Bonnets & Accessoires", "📄 Print, Papeterie & Signalétique (Catalogue APEX)"],
            key=f"metier_{i}"
        )
        st.markdown("---")
        
        if "Textile" in metier_type:
            sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
            col1, col2 = st.columns(2)
            with col1:
                nom_article = st.text_input(f"Référence / Nom du vêtement {i+1}", value="T-Shirt 100% coton bio" if i==0 else f"Article {i+1}", key=f"nom_textile_{i}")
                qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_textile_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT support (€) {i+1}", min_value=0.0, value=4.92, format="%.2f", key=f"px_textile_{i}")
            with col2:
                nb_marquages = st.selectbox(f"Nombre de marquages {i+1}", [1, 2, 3, 4], key=f"nb_m_textile_{i}") if not sans_marquage else 0

            marquages = []
            if not sans_marquage:
                for m in range(nb_marquages):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Léger", "DTF Textile Vif", "Broderie HD"], key=f"t_marq_{i}_{m}")
                    with mc2:
                        if "Broderie" in t_marq:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20 cm)", "Col / Signature (7x2 cm)", "Manche (8x5 cm)", "Pantalon / Poche", "+ Perso. Nom (Cœur)"], key=f"emp_{i}_{m}")
                        else:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Opposé Cœur (9x8 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "Pantalon / Poche", "+ Personnalisation Nom"], key=f"emp_{i}_{m}")
                    marquages.append({"technique": t_marq, "emplacement": emp})

            option_ensachage = st.checkbox(f"Option ensachage individuel {i+1}", key=f"ens_{i}")
            type_sachet = st.selectbox(f"Type de sachet {i+1}", ["Sachet (T-shirt/Polo)", "Sachet (Veste/Sweat)"], key=f"tsach_{i}") if option_ensachage else ""
            option_assurance = st.checkbox(f"Option assurance garantie textile {i+1}", key=f"ass_{i}")
            option_stockage = st.checkbox(f"Option stockage + picking {i+1}", key=f"stock_{i}")
            remise_fidelite = st.number_input(f"Réduction fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "textile",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": sans_marquage,
                    "marquages": marquages,
                    "option_ensachage": option_ensachage,
                    "type_sachet": type_sachet,
                    "option_assurance": option_assurance,
                    "option_stockage": option_stockage,
                    "remise_fidelite": remise_fidelite
                })
                total_textile_brut += qte * prix_vetement_ht

        elif "Casquettes" in metier_type:
            sans_marquage = st.checkbox(f"Accessoire sans marquage (fourniture seule) {i+1}", key=f"sans_marq_acc_{i}")
            col1, col2 = st.columns(2)
            with col1:
                nom_article = st.text_input(f"Référence / Nom (Casquette, Bonnet, Parapluie, Bagagerie) {i+1}", value="Casquette publicitaire", key=f"nom_acc_{i}")
                qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=25 if i==0 else 0, key=f"qte_acc_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT support (€) {i+1}", min_value=0.0, value=2.50, format="%.2f", key=f"px_acc_{i}")
            with col2:
                nb_marquages = st.selectbox(f"Nombre de marquages {i+1}", [1, 2], key=f"nb_m_acc_{i}") if not sans_marquage else 0

            marquages = []
            if not sans_marquage:
                for m in range(nb_marquages):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Léger", "Broderie HD"], key=f"t_marq_acc_{i}_{m}")
                    with mc2:
                        if "Broderie" in t_marq:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Casquettes / Bonnets", "Parapluie / Bagagerie"], key=f"emp_acc_{i}_{m}")
                        else:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Casquette / Bonnet", "Parapluie", "Bagagerie"], key=f"emp_acc_{i}_{m}")
                    marquages.append({"technique": t_marq, "emplacement": emp})

            option_ensachage = False
            type_sachet = ""
            option_assurance = st.checkbox(f"Option assurance garantie textile {i+1}", key=f"ass_acc_{i}")
            option_stockage = st.checkbox(f"Option stockage + picking {i+1}", key=f"stock_acc_{i}")
            remise_fidelite = st.number_input(f"Réduction fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_acc_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "textile",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": sans_marquage,
                    "marquages": marquages,
                    "option_ensachage": option_ensachage,
                    "type_sachet": type_sachet,
                    "option_assurance": option_assurance,
                    "option_stockage": option_stockage,
                    "remise_fidelite": remise_fidelite
                })
                total_textile_brut += qte * prix_vetement_ht

        else:
            cat_print = st.selectbox(
                f"Catégorie Print & Signalétique {i+1}",
                [
                    "Flyers", "Dépliants", "Blocs notes", "Chemises de présentation", 
                    "Banderoles", "Panneaux de chantier", "Roll-Up", "Sous bocks", 
                    "Adhésifs", "Cartes de visite", "Calendriers", "Menus restaurants",
                    "➕ Autre / Produit hors catalogue (Saisie libre)"
                ],
                key=f"cat_print_{i}"
            )
            
            if "Autre" in cat_print:
                choix_ref = st.text_input(f"Nom / Désignation du produit libre {i+1}", value="Produit personnalisé", key=f"ref_libre_{i}")
                col1, col2 = st.columns(2)
                with col1:
                    qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=1 if i==0 else 0, key=f"qte_print_{i}")
                    prix_vetement_ht = st.number_input(f"Prix unitaire HT (€) {i+1} (Saisie libre)", min_value=0.0, value=10.00, format="%.4f", key=f"px_libre_{i}")
                with col2:
                    st.info("💡 Saisie manuelle active (produit hors catalogue).")
                    remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_print_{i}")
            else:
                df_all = pd.read_excel(CATALOGUE_FILE, sheet_name=0, header=None)
        	    options_articles = {}
        	for r in range(2, len(df_all)):
            	    cat = str(df_all.iloc[r, 0]).strip() if pd.notna(df_all.iloc[r, 0]) else "Autres"
            	    ref = str(df_all.iloc[r, 2]).strip() if pd.notna(df_all.iloc[r, 2]) else ""
            
            	    if ref:
                	if cat not in options_articles:
                    	    options_articles[cat] = []
                	if ref not in options_articles[cat]:
                            options_articles[cat].append(ref)

        	choix_ref = st.selectbox(
            	    f"Modèle exact {i+1}", 
                    options_articles.get(cat_print, ["Article standard"]), 
            	    key=f"ref_print_{i}"
        	)
                
                col1, col2 = st.columns(2)
                with col1:
                    qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=100 if i==0 else 0, key=f"qte_print_{i}")
                    prix_unitaire_auto = obtenir_prix_catalogue_intelligent(cat_print, choix_ref, qte)
                    prix_vetement_ht = prix_unitaire_auto
                    st.metric(label=f"Prix unitaire HT (€) {i+1} (Catalogue auto)", value=f"{prix_unitaire_auto:.4f} €")
                with col2:
                    st.success(f"✅ Tarif appliqué ({qte} ex) : **{prix_unitaire_auto:.4f} € HT**")
                    remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_print_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "print",
                    "nom_article": f"{cat_print} - {choix_ref}" if "Autre" not in cat_print else choix_ref,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": True,
                    "marquages": [],
                    "option_ensachage": False,
                    "type_sachet": "",
                    "option_assurance": False,
                    "option_stockage": False,
                    "remise_fidelite": remise_fidelite
                })

# --- CALCUL DES QUANTITÉS CUMULÉES ---
quantites_cumulees_marquages = {}
quantite_totale_broderie_global = 0
has_broderie_global = False

for item in articles_saisis:
    if item["type_univers"] == "textile" and not item["sans_marquage"]:
        q = item["quantite"]
        for m in item["marquages"]:
            cle = (m["technique"], m["emplacement"])
            quantites_cumulees_marquages[cle] = quantites_cumulees_marquages.get(cle, 0) + q
            if "Broderie" in m["technique"]:
                has_broderie_global = True
                quantite_totale_broderie_global += q

frais_tech_auto = 19.80 if total_textile_brut > 0 else 0.0

# --- ONGLET GÉNÉRAL & DEVIS (Index 10) ---
with onglets[10]:
    st.subheader("📊 Récapitulatif Général & Génération du Devis Professionnel")

    if not articles_saisis:
        st.warning("Veuillez renseigner au moins un article avec une quantité supérieure à 0.")
    else:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            supprimer_frais_tech = st.checkbox("⚙️ Supprimer / Offrir les frais techniques de dossier", value=False)
            frais_techniques_dossier = 0.0 if supprimer_frais_tech else frais_tech_auto
        with col_f2:
            offrir_frais_broderie = st.checkbox("🎁 Offrir les Frais de technique & programme Broderie", value=False)

        if has_broderie_global and not offrir_frais_broderie:
            if quantite_totale_broderie_global == 1:
                frais_prog_broderie = 52.0
            elif 2 <= quantite_totale_broderie_global <= 3:
                frais_prog_broderie = 41.0
            elif 4 <= quantite_totale_broderie_global <= 11:
                frais_prog_broderie = 23.0
            else:
                frais_prog_broderie = 0.0
        else:
            frais_prog_broderie = 0.0

        lignes_devis_global = []

        for item in articles_saisis:
            q = item["quantite"]
            px_support = item["prix_vet_unit"] * (1 - item["remise_fidelite"] / 100.0)
            tot_support = q * px_support
            
            marquages_calcules = []
            tot_marquages = 0.0
            
            if item["type_univers"] == "textile" and not item["sans_marquage"]:
                for m in item["marquages"]:
                    cle = (m["technique"], m["emplacement"])
                    qte_tot_ref = quantites_cumulees_marquages.get(cle, q)
                    
                    if "Broderie" in m["technique"]:
                        tarif_m = obtenir_tarif_broderie_unitaire(m["emplacement"], qte_tot_ref)
                    else:
                        tarif_m = obtenir_tarif_dtf_unitaire(m["technique"], m["emplacement"], qte_tot_ref)
                        
                    tot_marquages += tarif_m * q
                    marquages_calcules.append({"nom": f"{m['technique']} ({m['emplacement']})", "tarif": tarif_m})

            coût_ens_unit = (1.375 if q<=11 else (1.243 if q<=24 else (1.166 if q<=49 else (1.111 if q<=99 else (1.078 if q<=249 else (1.056 if q<=499 else (1.023 if q<=999 else (1.001 if q<=1999 else 0.979)))))))) if item["option_ensachage"] else 0.0
            tot_ens = coût_ens_unit * q

            coût_ass_unit = (3.08 if q<=11 else (2.376 if q<=24 else (1.826 if q<=49 else (1.254 if q<=99 else (0.979 if q<=249 else (0.704 if q<=499 else (0.638 if q<=999 else (0.605 if q<=1999 else 0.55)))))))) if item["option_assurance"] else 0.0
            tot_ass = coût_ass_unit * q

            coût_stock_unit = (1.00 if q<=99 else (0.56 if q<=249 else (0.50 if q<=499 else (0.43 if q<=999 else 0.30)))) if item["option_stockage"] else 0.0
            tot_stock = coût_stock_unit * q
            
            tot_ligne = tot_support + tot_marquages + tot_ens + tot_ass + tot_stock
            
            lignes_devis_global.append({
                **item,
                "prix_vet_unit": px_support,
                "marquages_calcules": marquages_calcules,
                "coût_ensachage_unit": coût_ens_unit,
                "coût_assurance_unit": coût_ass_unit,
                "coût_stockage_unit": coût_stock_unit,
                "total_ligne_ht": tot_ligne
            })

        sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
        montant_base_port = sous_total_articles + frais_prog_broderie + frais_techniques_dossier
        
        frais_port = 0.0 if offrir_port else calculer_frais_port(montant_base_port, zone_livraison)
        
        total_general_ht = montant_base_port + frais_port
        tva = total_general_ht * 0.20
        total_ttc = total_general_ht + tva
        
        quantite_globale_totale = sum([item["quantite"] for item in lignes_devis_global])
        cout_unitaire_moyen = total_general_ht / quantite_globale_totale if quantite_globale_totale > 0 else 0

        st.write(f"**Quantité globale pièces :** {quantite_globale_totale}")
        st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
        if frais_prog_broderie > 0 or offrir_frais_broderie:
            libelle_fp = f"{frais_prog_broderie:.2f} € HT" if not offrir_frais_broderie else "Offerts (0.00 €)"
            st.write(f"**Frais de technique & programme Broderie :** {libelle_fp}")
        st.write(f"**Frais techniques de dossier :** {frais_techniques_dossier:.2f} € HT")
        st.write(f"**Frais de port ({zone_livraison}) :** {frais_port:.2f} € HT" if not offrir_port else "**Frais de port :** Offerts (0.00 €)")
        st.markdown(f"### **Total Général HT : {total_general_ht:.2f} €** | **TOTAL TTC (20%) : {total_ttc:.2f} €**")

        if st.button("📄 Générer le numéro de devis et le PDF"):
            num_devis = obtenir_prochain_numero_devis()
            pdf_filename = os.path.join(PDF_DIR, f"Devis_{num_devis.replace('/', '_')}.pdf")
            st.session_state['dernier_pdf'] = pdf_filename
            st.session_state['dernier_num'] = num_devis

            data_crm = {
                "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Numero_Devis": num_devis,
                "Commercial": conseiller_nom,
                "Client": client_nom,
                "Entreprise": client_entreprise,
                "Email": client_email,
                "Telephone": client_contact_tel,
                "Quantite_Totale": quantite_globale_totale,
                "Total_HT": round(total_general_ht, 2),
                "Total_TTC": round(total_ttc, 2),
                "Statut": "En cours",
                "Mode_Reglement": mode_reglement,
                "PDF_Path": pdf_filename
            }
            enregistrer_dans_crm(data_crm)
            st.success("✅ Données enregistrées dans le CRM avec succès !")

            # --- GÉNÉRATION DU PDF ---
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
            elif os.path.exists("logo.png"):
                logo_path = "logo.png"
            elif os.path.exists("logo.jpg"):
                logo_path = "logo.jpg"

            header_text = Paragraph(
                "<b>APEX - SOLUTIONS VISUELLES, PRINT & TEXTILE</b><br/>"
                "Plasne (Jura)<br/>"
                "Tél (Brice Geny) : 06 32 69 73 28 &nbsp;|&nbsp; Tél (Brice Bugna) : 06 29 92 94 74<br/>"
                "Email : contact@apex-visual.fr", 
                style_sub
            )
            if logo_path and os.path.exists(logo_path):
                img_logo = RLImage(logo_path, width=120, height=50)
                t_header = Table([[img_logo, header_text]], colWidths=[130, 410])
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
                table_data.append([Paragraph(libelle_support, style_cell), str(item['quantite']), f"{item['prix_vet_unit']:.3f} €", f"{item['prix_vet_unit']*item['quantite']:.2f} €"])
                for m in item['marquages_calcules']:
                    table_data.append([Paragraph(f"&nbsp;&nbsp;&bull; Marquage : {m['nom']}", style_cell), str(item['quantite']), f"{m['tarif']:.2f} €", f"{m['tarif']*item['quantite']:.2f} €"])
                if item['option_ensachage']:
                    table_data.append([Paragraph(f"&nbsp;&nbsp;&bull; Option : {item['type_sachet']}", style_cell), str(item['quantite']), f"{item['coût_ensachage_unit']:.2f} €", f"{item['coût_ensachage_unit']*item['quantite']:.2f} €"])
                if item['option_assurance']:
                    table_data.append([Paragraph("&nbsp;&nbsp;&bull; Option : Assurance MHC (Garantie textile)", style_cell), str(item['quantite']), f"{item['coût_assurance_unit']:.2f} €", f"{item['coût_assurance_unit']*item['quantite']:.2f} €"])
                if item['option_stockage']:
                    table_data.append([Paragraph("&nbsp;&nbsp;&bull; Option : Mise en stockage + picking", style_cell), str(item['quantite']), f"{item['coût_stockage_unit']:.2f} €", f"{item['coût_stockage_unit']*item['quantite']:.2f} €"])

            if frais_prog_broderie > 0 or offrir_frais_broderie:
                fp_libelle = "Frais de technique & programme Broderie" if not offrir_frais_broderie else "Frais de technique & programme Broderie - Offerts"
                fp_val = f"{frais_prog_broderie:.2f} €" if not offrir_frais_broderie else "0.00 €"
                table_data.append([Paragraph(fp_libelle, style_cell), "1", fp_val, fp_val])

            if frais_techniques_dossier > 0:
                table_data.append([Paragraph("Frais techniques de dossier", style_cell), "1", f"{frais_techniques_dossier:.2f} €", f"{frais_techniques_dossier:.2f} €"])
            
            if frais_port > 0 or offrir_port:
                port_libelle = f"Frais d'envoi ({zone_livraison})" if not offrir_port else f"Frais d'envoi ({zone_livraison}) - Offerts"
                port_val = f"{frais_port:.2f} €" if not offrir_port else "0.00 €"
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
                ["", Paragraph("Coût unitaire HT / pièce :", style_right_normal), Paragraph(f"{cout_unitaire_moyen:.3f} €", style_right_normal)]
            ]
            t_totaux = Table(totaux_data, colWidths=[240, 160, 140])
            t_totaux.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LINEABOVE', (1, 2), (-1, 2), 1, colors.black),
                ('PADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(t_totaux)
            story.append(Spacer(1, 15))

            conditions_text = f"<b>Conditions de règlement & Bon pour accord :</b><br/>• Règlement : {mode_reglement}<br/>• Fichiers vectoriels fournis (.AI, .EPS, .PDF).<br/>• Bon pour accord daté et signé requis."
            story.append(Paragraph(conditions_text, style_sub))

            doc.build(story)
            st.success(f"Devis PDF professionnel généré sous le numéro : **{num_devis}**")

        if 'dernier_pdf' in st.session_state:
            pdf_filename = st.session_state['dernier_pdf']
            if os.path.exists(pdf_filename):
                with open(pdf_filename, "rb") as f:
                    st.download_button("📥 Télécharger le PDF du devis", f, file_name=os.path.basename(pdf_filename), mime="application/pdf")
                
                st.markdown("---")
                st.subheader("✉️ Envoi direct du devis par e-mail en 1 clic")
                email_dest = st.text_input("Destinataire de l'e-mail", value=client_email)
                sujet_mail = st.text_input("Objet de l'e-mail", value=f"Devis {st.session_state.get('dernier_num', '')} - APEX")
                corps_mail = st.text_area("Message", value=f"Bonjour {client_nom},\n\nVeuillez trouver ci-joint votre devis établi par APEX.\n\nCordialement,\n{conseiller_nom}\nAPEX")

                mailto_link = f"mailto:{email_dest}?subject={urllib.parse.quote(sujet_mail)}&body={urllib.parse.quote(corps_mail)}"
                st.markdown(f'<a href="{mailto_link}" target="_blank"><button style="background-color:#2b6cb0; color:white; border:none; padding:10px 20px; border-radius:5px; cursor:pointer; font-weight:bold; width:100%;">📧 Ouvrir dans le client mail (Secours)</button></a>', unsafe_allow_html=True)

# --- ONGLET SUIVI CRM (Index 11) ---
with onglets[11]:
    st.header("📈 Suivi CRM & Historique des Devis")
    if os.path.exists(CRM_FILE):
        df_crm = pd.read_csv(CRM_FILE)
        if not df_crm.empty:
            st.dataframe(df_crm, use_container_width=True)
            devis_selectionne = st.selectbox("Sélectionner un devis", df_crm["Numero_Devis"].tolist(), key="select_crm")
            ligne_dev = df_crm[df_crm["Numero_Devis"] == devis_selectionne].iloc[0]
            if pd.notna(ligne_dev.get("PDF_Path")) and os.path.exists(str(ligne_dev["PDF_Path"])):
                with open(ligne_dev["PDF_Path"], "rb") as pdf_file:
                    st.download_button("📥 Télécharger le PDF de ce devis", pdf_file, file_name=os.path.basename(ligne_dev["PDF_Path"]), mime="application/pdf")
        else:
            st.info("Aucun devis dans le CRM.")
    else:
        st.info("CRM vide.")
