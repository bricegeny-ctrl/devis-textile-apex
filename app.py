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
import urllib.parse

st.set_page_config(page_title="Gestionnaire de Devis - APEX", layout="wide")

# --- GESTION DES FICHIERS ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"
CATALOGUE_FILE = "catalogue print et signalétique.xlsx"
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

# --- MOTEUR DE LECTURE EXCEL : GESTION DES PRIX UNIQUES ET SÉCURITÉ GLOBALE ---
def obtenir_prix_catalogue_intelligent(cat_print, choix_ref, qte):
    if not os.path.exists(CATALOGUE_FILE):
        return 0.15
    
    try:
        df_all = pd.read_excel(CATALOGUE_FILE, sheet_name=0, header=None)
    except Exception:
        return 0.15

    cat_lower = str(cat_print).lower().strip()
    ref_lower = str(choix_ref).lower().strip()

    # Détermination de la colonne cible idéale selon la quantité
    if qte >= 10000:
        col_cible = 14
    elif qte >= 5000:
        col_cible = 13
    elif qte >= 2500:
        col_cible = 12
    elif qte >= 1000:
        col_cible = 11
    elif qte >= 500:
        col_cible = 10
    elif qte >= 250:
        col_cible = 9
    elif qte >= 100:
        col_cible = 8
    elif qte >= 50:
        col_cible = 7
    elif qte >= 25:
        col_cible = 6
    elif qte >= 10:
        col_cible = 5
    elif qte >= 5:
        col_cible = 4
    else:
        col_cible = 3

    # Recherche de la ligne exacte du produit
    best_row = -1
    max_match = -1

    for r in range(2, len(df_all)):
        row_cat = str(df_all.iloc[r, 0]).lower().strip()
        row_ref = str(df_all.iloc[r, 2]).lower().strip()

        score = 0
        if cat_lower in row_cat or row_cat in cat_lower:
            score += 10
        
        if ref_lower == row_ref:
            score += 50
        elif ref_lower in row_ref or row_ref in ref_lower:
            score += 25

        if score > max_match:
            max_match = score
            best_row = r

    if best_row == -1:
        return 0.15

    # Extraction sécurisée avec repli intelligent vers la gauche si la colonne ciblée est vide (cas des prix uniques comme les banderoles)
    try:
        prix_val = -1.0
        
        # On essaie de lire la colonne cible, ou les colonnes juste avant (vers la gauche) si elle est vide
        for c in range(col_cible, 2, -1):
            if c < df_all.shape[1]:
                val = df_all.iloc[best_row, c]
                if not pd.isna(val) and str(val).strip() != "":
                    try:
                        temp_val = float(str(val).replace('€', '').replace('EUR', '').replace(' ', '').replace(',', '.'))
                        if temp_val > 0:
                            prix_val = temp_val
                            break
                    except Exception:
                        continue

        # Si vraiment rien n'est trouvé sur la ligne, on cherche un prix de secours à l'index 3 (colonne D)
        if prix_val <= 0:
            val_secours = df_all.iloc[best_row, 3]
            if not pd.isna(val_secours) and str(val_secours).strip() != "":
                prix_val = float(str(val_secours).replace('€', '').replace('EUR', '').replace(' ', '').replace(',', '.'))

        if prix_val <= 0:
            return 0.15
            
        return round(prix_val, 4)
    except Exception:
        return 0.15

# --- GRILLES TARIFAIRES OFFICIELLES (MARQUAGE & BRODERIE) ---
def obtenir_tarif_dtf_unitaire(type_textile, emplacement, qte_totale):
    grille_fin = {
        "Cœur (13x9 cm)": [(5, 6.00), (9, 4.50), (19, 3.60), (29, 2.81), (39, 2.50), (49, 2.40), (99, 2.00), (249, 1.80), (499, 1.60), (999, 1.40), (5000, 1.20), (float('inf'), 0.70)],
        "Dos D10 (20x13 cm)": [(5, 7.92), (9, 6.50), (19, 5.50), (29, 5.00), (39, 4.20), (49, 4.00), (99, 3.50), (249, 3.00), (499, 2.70), (999, 2.50), (5000, 2.00), (float('inf'), 1.00)],
        "Dos D20 (28x20 cm)": [(5, 13.67), (9, 10.00), (19, 9.00), (29, 6.90), (39, 6.30), (49, 6.00), (99, 5.50), (249, 4.60), (499, 4.00), (999, 3.50), (5000, 2.50), (float('inf'), 1.50)],
        "Format P (37x27 cm)": [(5, 17.00), (9, 13.00), (19, 12.00), (29, 10.00), (39, 9.00), (49, 8.00), (99, 7.50), (249, 6.50), (499, 6.00), (999, 5.00), (5000, 4.00), (float('inf'), 2.50)],
        "Manche (9x8 cm)": [(5, 7.20), (9, 5.40), (19, 4.32), (29, 3.37), (39, 3.00), (49, 2.88), (99, 2.40), (249, 2.16), (499, 1.92), (999, 1.68), (5000, 1.44), (float('inf'), 0.84)],
        "+ Personnalisation Nom": [(5, 4.39), (9, 3.50), (19, 2.50), (29, 2.20), (39, 1.90), (49, 1.80), (99, 1.70), (249, 1.50), (499, 1.30), (999, 0.80), (5000, 0.30), (float('inf'), 0.20)]
    }
    cle = emplacement if emplacement in grille_fin else "Cœur (13x9 cm)"
    paliers = grille_fin[cle]
    prix = paliers[-1][1]
    for limite, p in paliers:
        if qte_totale <= limite:
            prix = p
            break
    if "Épais" in type_textile:
        prix = round(prix * 1.10, 2)
    return prix

def obtenir_tarif_broderie_unitaire(emplacement, qte_totale):
    grille_brod = {
        "Poitrine (9x8 cm)": [(3, 15.90), (11, 12.13), (23, 9.20), (47, 6.90), (95, 5.30), (251, 4.80), (503, 4.50), (1007, 4.20), (1511, 3.90), (float('inf'), 3.50)],
        "Dos D10 (25x10 cm)": [(3, 18.50), (11, 14.60), (23, 11.30), (47, 9.10), (95, 7.60), (251, 7.10), (503, 6.70), (1007, 6.30), (1511, 5.90), (float('inf'), 5.40)],
        "Dos Large D20 (25x20 cm)": [(3, 23.20), (11, 18.25), (23, 14.20), (47, 11.90), (95, 9.90), (251, 9.30), (503, 8.80), (1007, 8.30), (1511, 7.80), (float('inf'), 7.20)],
        "Col / Signature (7x2 cm)": [(3, 10.90), (11, 8.37), (23, 6.20), (47, 4.40), (95, 3.10), (251, 2.85), (503, 2.65), (1007, 2.45), (1511, 2.25), (float('inf'), 2.00)],
        "Casquettes / Bonnets": [(3, 16.10), (11, 12.27), (23, 9.30), (47, 7.10), (95, 5.50), (251, 5.05), (503, 4.75), (1007, 4.45), (1511, 4.15), (float('inf'), 3.75)],
        "Manche (8x5 cm)": [(3, 16.10), (11, 12.27), (23, 9.30), (47, 7.10), (95, 5.50), (251, 5.05), (503, 4.75), (1007, 4.45), (1511, 4.15), (float('inf'), 3.75)],
        "Pantalon / Poche": [(3, 18.30), (11, 13.97), (23, 10.90), (47, 8.60), (95, 7.10), (251, 6.60), (503, 6.20), (1007, 5.80), (1511, 5.45), (float('inf'), 4.95)],
        "+ Perso. Nom (Cœur)": [(3, 6.00), (11, 4.00), (23, 4.00), (47, 3.50), (95, 3.00), (251, 2.80), (503, 2.60), (1007, 2.40), (1511, 2.20), (float('inf'), 2.00)]
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
logo_defaut_github = "logo.png"

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
            ["👕 Textile & Marquage (DTF / Broderie)", "📄 Print, Papeterie & Signalétique (Catalogue APEX)"],
            key=f"metier_{i}"
        )
        st.markdown("---")
        
        if "Textile" in metier_type:
            sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
            col1, col2 = st.columns(2)
            with col1:
                nom_article = st.text_input(f"Référence / Nom du vêtement {i+1}", value="T-Shirt 100% coton bio" if i==0 else f"Vêtement {i+1}", key=f"nom_textile_{i}")
                qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_textile_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT support (€) {i+1}", min_value=0.0, value=4.92, format="%.2f", key=f"px_textile_{i}")
            with col2:
                nb_marquages = st.selectbox(f"Nombre de marquages {i+1}", [1, 2, 3, 4], key=f"nb_m_textile_{i}") if not sans_marquage else 0

            marquages = []
            if not sans_marquage:
                for m in range(nb_marquages):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
                    with mc2:
                        if "Broderie" in t_marq:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20 cm)", "Col / Signature (7x2 cm)", "Casquettes / Bonnets", "Manche (8x5 cm)", "Pantalon / Poche", "+ Perso. Nom (Cœur)"], key=f"emp_{i}_{m}")
                        else:
                            emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Personnalisation Nom"], key=f"emp_{i}_{m}")
                    marquages.append({"technique": t_marq, "emplacement": emp})

            option_ensachage = st.checkbox(f"Option ensachage individuel {i+1}", key=f"ens_{i}")
            type_sachet = st.selectbox(f"Type de sachet {i+1}", ["Sachet (T-shirt/Polo)", "Sachet (Veste/Sweat)"], key=f"tsach_{i}") if option_ensachage else ""
            
            option_assurance = st.checkbox(f"Option assurance garantie textile {i+1}", key=f"ass_{i}")
            option_stockage = st.checkbox(f"Option stockage + picking {i+1}", key=f"stock_{i}")
            remise_fidelite = st.number_input(f"Réduction fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_{i}")

# Saisie du nombre d'articles global
nb_articles = st.number_input("Nombre d'articles dans le devis", min_value=1, value=1, step=1, key="nb_articles_global")

for i in range(nb_articles):
    st.markdown(f"### Configuration de l'Article {i+1}")
    
    # Choix de l'univers pour cet article
    univers = st.radio(
        f"Univers / Métier pour l'Article {i+1}", 
        ["Textile & Marquage (DTF / Broderie)", "Print, Papeterie & Signalétique (Catalogue APEX)"], 
        key=f"univers_{i}"
    )

    if "Textile" in univers:
        # --- VOTRE PARTIE TEXTILE ---
        nom_article = st.text_input(f"Référence / Nom du vêtement {i+1}", value="T-Shirt 100% coton bio", key=f"txt_ref_{i}")
        qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10, key=f"txt_qte_{i}")
        prix_vetement_ht = st.number_input(f"Prix unitaire HT support (€) {i+1}", min_value=0.0, value=4.92, format="%.2f", key=f"txt_px_{i}")
        
        option_ensachage = st.checkbox(f"Option ensachage individuel {i+1}", key=f"txt_ens_{i}")
        option_assurance = st.checkbox(f"Option assurance garantie textile {i+1}", key=f"txt_ass_{i}")
        option_stockage = st.checkbox(f"Option stockage + picking {i+1}", key=f"txt_stock_{i}")
        remise_fidelite = st.number_input(f"Réduction fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"txt_rem_{i}")

        if qte > 0:
            articles_saisis.append({
                "type_univers": "textile",
                "nom_article": nom_article,
                "quantite": qte,
                "prix_vet_unit": prix_vetement_ht,
                "option_ensachage": option_ensachage,
                "option_assurance": option_assurance,
                "option_stockage": option_stockage,
                "remise_fidelite": remise_fidelite
            })
            total_textile_brut += qte * prix_vetement_ht

    else:
        # --- PARTIE PRINT & SIGNALETIQUE ---
        cat_print = st.selectbox(
            f"Catégorie Print & Signalétique {i+1}",
            [
                "Flyers", "Dépliants", "Blocs notes", "Chemises de présentation",
                "Banderoles", "Panneaux de chantier", "Roll-up", "Sous bocks",
                "Adhésifs", "Cartes de visite", "Calendriers", "Menus restaurants",
                "➕ Autre / Produit hors catalogue (Saisie libre)"
            ],
            key=f"print_cat_{i}"
        )

        if "Autre" in cat_print:
            choix_ref = st.text_input(f"Nom / Désignation du produit libre {i+1}", value="Produit personnalisé", key=f"print_ref_libre_{i}")
            coll, col2 = st.columns(2)
            with coll:
                qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=100, key=f"print_qte_libre_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT (€) {i+1} (Saisie libre)", min_value=0.0, value=10.00, format="%.4f", key=f"print_px_libre_{i}")
            with col2:
                st.info("💡 Saisie manuelle active (produit hors catalogue).")
                remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"print_rem_libre_{i}")
        else:
            options_trouvees = obtenir_modeles_pour_categorie(cat_print)
            if not options_trouvees:
                options_trouvees = ["Article standard"]

            choix_ref = st.selectbox(f"Modèle exact {i+1}", options_trouvees, key=f"print_modele_{i}")

            coll, col2 = st.columns(2)
            with coll:
                qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=100, key=f"print_qte_cat_{i}")
                
                # Calcul automatique du prix par tranche
                prix_unitaire_auto = obtenir_prix_catalogue_intelligent(cat_print, choix_ref, qte)
                prix_vetement_ht = prix_unitaire_auto  

                st.metric(label=f"Prix unitaire HT (€) {i+1} (Catalogue auto)", value=f"{prix_unitaire_auto:.4f} €")

            with col2:
                st.success(f"✅ Tarif appliqué ({qte} ex) : **{prix_unitaire_auto:.4f} € HT**")
                remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"print_rem_cat_{i}")

        if qte > 0:
            articles_saisis.append({
                "type_univers": "print",
                "nom_article": f"{cat_print} - {choix_ref}" if "Autre" not in cat_print else choix_ref,
                "quantite": qte,
                "prix_vet_unit": prix_vetement_ht,
                "remise_fidelite": remise_fidelite
            })

# --- CALCUL DES QUANTITÉS CUMULÉES ---
quantites_cumulees_marquages = {}
for item in articles_saisis:
    if item["type_univers"] == "textile" and not item["sans_marquage"]:
        q = item["quantite"]
        for m in item["marquages"]:
            cle = (m["technique"], m["emplacement"])
            quantites_cumulees_marquages[cle] = quantites_cumulees_marquages.get(cle, 0) + q

frais_tech_auto = 19.80 if total_textile_brut > 0 else 0.0

# --- ONGLET GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.subheader("📊 Récapitulatif Général & Génération du Devis Professionnel")

    if not articles_saisis:
        st.warning("Veuillez renseigner au moins un article avec une quantité supérieure à 0.")
    else:
        supprimer_frais_tech = st.checkbox("⚙️ Supprimer / Offrir les frais techniques de dossier", value=False)
        frais_techniques_dossier = 0.0 if supprimer_frais_tech else frais_tech_auto

        lignes_devis_global = []
        has_broderie_global = False
        quantite_totale_broderie = 0

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
                        has_broderie_global = True
                        quantite_totale_broderie += q
                    else:
                        tarif_m = obtenir_tarif_dtf_unitaire(m["technique"], m["emplacement"], qte_tot_ref)
                        
                    tot_marquages += tarif_m * q
                    marquages_calcules.append({"nom": f"{m['technique']} ({m['emplacement']})", "tarif": tarif_m})

            if has_broderie_global:
                if 2 <= quantite_totale_broderie <= 3:
                    frais_prog_broderie = 41.0
                elif 4 <= quantite_totale_broderie <= 11:
                    frais_prog_broderie = 23.0
                else:
                    frais_prog_broderie = 0.0
            else:
                frais_prog_broderie = 0.0

            if item["option_ensachage"]:
                coût_ens_unit = 1.38 if q<=11 else (1.24 if q<=24 else (1.17 if q<=49 else (1.11 if q<=99 else (1.08 if q<=249 else (1.06 if q<=499 else 1.00)))))
            else:
                coût_ens_unit = 0.0
            tot_ens = coût_ens_unit * q

            if item["option_assurance"]:
                coût_ass_unit = 3.08 if q<=11 else (2.38 if q<=24 else (1.83 if q<=49 else (1.25 if q<=99 else (0.98 if q<=249 else (0.70 if q<=499 else (0.64 if q<=999 else 0.61))))))
            else:
                coût_ass_unit = 0.0
            tot_ass = coût_ass_unit * q

            if item["option_stockage"]:
                coût_stock_unit = 1.00 if q<=99 else (0.56 if q<=249 else (0.50 if q<=499 else (0.43 if q<=999 else 0.30)))
            else:
                coût_stock_unit = 0.0
            tot_stock = coût_stock_unit * q
            
            tot_ligne = tot_support + tot_marquages + tot_ens + tot_ass + tot_stock
            
            lignes_devis_global.append({
                **item,
                "prix_vet_unit": px_support,
                "marquages_calcules": marquages_calcules,
                "frais_prog_broderie": frais_prog_broderie,
                "coût_ensachage_unit": coût_ens_unit,
                "coût_assurance_unit": coût_ass_unit,
                "coût_stockage_unit": coût_stock_unit,
                "total_ligne_ht": tot_ligne
            })

        sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
        frais_prog_total = lignes_devis_global[0]["frais_prog_broderie"] if lignes_devis_global else 0.0
        montant_base_port = sous_total_articles + frais_prog_total + frais_techniques_dossier
        
        frais_port = 0.0 if offrir_port else calculer_frais_port(montant_base_port, zone_livraison)
        
        total_general_ht = montant_base_port + frais_port
        tva = total_general_ht * 0.20
        total_ttc = total_general_ht + tva
        
        quantite_globale_totale = sum([item["quantite"] for item in lignes_devis_global])
        cout_unitaire_moyen = total_general_ht / quantite_globale_totale if quantite_globale_totale > 0 else 0

        st.write(f"**Quantité globale pièces :** {quantite_globale_totale}")
        st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
        if frais_prog_total > 0:
            st.write(f"**Frais de technique & programme Broderie :** {frais_prog_total:.2f} € HT")
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

            # --- GÉNÉRATION DU PDF (MENTION APEX UNIQUEMENT) ---
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

            if frais_prog_total > 0:
                table_data.append([Paragraph("Frais de technique & programme Broderie", style_cell), "1", f"{frais_prog_total:.2f} €", f"{frais_prog_total:.2f} €"])

            if frais_techniques_dossier > 0:
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

# --- ONGLET SUIVI CRM ---
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