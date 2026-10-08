from datetime import datetime
import json
import os
import urllib.parse
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Gestionnaire de Devis - APEX", layout="wide")

# --- GESTION DES FICHIERS ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"
CATALOGUE_FILE = "catalogue_print.xlsx"
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
  try:
    with open(COMPTEUR_FILE, "w") as f:
      json.dump({"dernier_num": nouveau_num}, f)
  except Exception:
    pass
  return nouveau_num


# --- MOTEUR DE LECTURE EXCEL : CATALOGUE PRINT & SIGNALÉTIQUE ---
def obtenir_prix_catalogue_intelligent(cat_print, choix_ref, qte):
  if not os.path.exists(CATALOGUE_FILE):
    return 0.15

  try:
    df_all = pd.read_excel(CATALOGUE_FILE, sheet_name=0, header=None)
  except Exception as e:
    return 0.15

  qtys = []
  for c in range(3, df_all.shape[1]):
    try:
      val = float(df_all.iloc[0, c])
      qtys.append((c, val))
    except:
      pass

  target_col = qtys[0][0] if qtys else 3
  for col_idx, q_val in qtys:
    if qte >= q_val:
      target_col = col_idx

  cat_lower = str(cat_print).lower().strip()
  ref_lower = str(choix_ref).lower().strip()

  best_row = -1
  max_match = -1

  for r in range(1, len(df_all)):
    row_cat = (
        str(df_all.iloc[r, 0]).lower().strip()
        if pd.notna(df_all.iloc[r, 0])
        else ""
    )
    row_sub = (
        str(df_all.iloc[r, 1]).lower().strip()
        if pd.notna(df_all.iloc[r, 1])
        else ""
    )
    row_ref = (
        str(df_all.iloc[r, 2]).lower().strip()
        if pd.notna(df_all.iloc[r, 2])
        else ""
    )

    score = 0
    if cat_lower in row_cat or row_cat in cat_lower:
      score += 10

    mots_ref = [m for m in ref_lower.split() if len(m) > 2]
    match_mots = sum(
        1 for m in mots_ref if m in row_sub or m in row_ref or row_cat in m
    )
    score += match_mots * 5

    if score > max_match:
      max_match = score
      best_row = r

  if best_row == -1 or max_match < 5:
    for r in range(1, len(df_all)):
      if cat_lower in str(df_all.iloc[r, 0]).lower():
        best_row = r
        break

  if best_row == -1:
    return 0.15

  try:
    prix_val = float(df_all.iloc[best_row, target_col])
    if pd.isna(prix_val) or prix_val <= 0:
      for alt_col in range(df_all.shape[1] - 1, 2, -1):
        alt_val = float(df_all.iloc[best_row, alt_col])
        if not pd.isna(alt_val) and alt_val > 0:
          prix_val = alt_val
          break

    if pd.isna(prix_val) or prix_val <= 0:
      return 0.15

    return round(prix_val, 4)
  except Exception:
    return 0.15


# --- GRILLES TARIFAIRES OFFICIELLES (MARQUAGE & BRODERIE) ---
def obtenir_tarif_dtf_unitaire(type_textile, emplacement, qte_totale):
  grille_fin = {
      "Cœur (13x9 cm)": [
          (5, 6.00),
          (9, 4.50),
          (19, 3.60),
          (29, 2.81),
          (39, 2.50),
          (49, 2.40),
          (99, 2.00),
          (249, 1.80),
          (499, 1.60),
          (999, 1.40),
          (5000, 1.20),
          (float("inf"), 0.70),
      ],
      "Dos D10 (20x13 cm)": [
          (5, 7.92),
          (9, 6.50),
          (19, 5.50),
          (29, 5.00),
          (39, 4.20),
          (49, 4.00),
          (99, 3.50),
          (249, 3.00),
          (499, 2.70),
          (999, 2.50),
          (5000, 2.00),
          (float("inf"), 1.00),
      ],
      "Dos D20 (28x20 cm)": [
          (5, 13.67),
          (9, 10.00),
          (19, 9.00),
          (29, 6.90),
          (39, 6.30),
          (49, 6.00),
          (99, 5.50),
          (249, 4.60),
          (499, 4.00),
          (999, 3.50),
          (5000, 2.50),
          (float("inf"), 1.50),
      ],
      "Format P (37x27 cm)": [
          (5, 17.00),
          (9, 13.00),
          (19, 12.00),
          (29, 10.00),
          (39, 9.00),
          (49, 8.00),
          (99, 7.50),
          (249, 6.50),
          (499, 6.00),
          (999, 5.00),
          (5000, 4.00),
          (float("inf"), 2.50),
      ],
      "Manche (9x8 cm)": [
          (5, 7.20),
          (9, 5.40),
          (19, 4.32),
          (29, 3.37),
          (39, 3.00),
          (49, 2.88),
          (99, 2.40),
          (249, 2.16),
          (499, 1.92),
          (999, 1.68),
          (5000, 1.44),
          (float("inf"), 0.84),
      ],
      "+ Personnalisation Nom": [
          (5, 4.39),
          (9, 3.50),
          (19, 2.50),
          (29, 2.20),
          (39, 1.90),
          (49, 1.80),
          (99, 1.70),
          (249, 1.50),
          (499, 1.30),
          (999, 0.80),
          (5000, 0.30),
          (float("inf"), 0.20),
      ],
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
      "Poitrine (9x8 cm)": [
          (3, 15.90),
          (11, 12.13),
          (23, 9.20),
          (47, 6.90),
          (95, 5.30),
          (251, 4.80),
          (503, 4.50),
          (1007, 4.20),
          (1511, 3.90),
          (float("inf"), 3.50),
      ],
      "Dos D10 (25x10 cm)": [
          (3, 18.50),
          (11, 14.60),
          (23, 11.30),
          (47, 9.10),
          (95, 7.60),
          (251, 7.10),
          (503, 6.70),
          (1007, 6.30),
          (1511, 5.90),
          (float("inf"), 5.40),
      ],
      "Dos Large D20 (25x20 cm)": [
          (3, 23.20),
          (11, 18.25),
          (23, 14.20),
          (47, 11.90),
          (95, 9.90),
          (251, 9.30),
          (503, 8.80),
          (1007, 8.30),
          (1511, 7.80),
          (float("inf"), 7.20),
      ],
      "Col / Signature (7x2 cm)": [
          (3, 10.90),
          (11, 8.37),
          (23, 6.20),
          (47, 4.40),
          (95, 3.10),
          (251, 2.85),
          (503, 2.65),
          (1007, 2.45),
          (1511, 2.25),
          (float("inf"), 2.00),
      ],
      "Casquettes / Bonnets": [
          (3, 16.10),
          (11, 12.27),
          (23, 9.30),
          (47, 7.10),
          (95, 5.50),
          (251, 5.05),
          (503, 4.75),
          (1007, 4.45),
          (1511, 4.15),
          (float("inf"), 3.75),
      ],
      "Manche (8x5 cm)": [
          (3, 16.10),
          (11, 12.27),
          (23, 9.30),
          (47, 7.10),
          (95, 5.50),
          (251, 5.05),
          (503, 4.75),
          (1007, 4.45),
          (1511, 4.15),
          (float("inf"), 3.75),
      ],
      "Pantalon / Poche": [
          (3, 18.30),
          (11, 13.97),
          (23, 10.90),
          (47, 8.60),
          (95, 7.10),
          (251, 6.60),
          (503, 6.20),
          (1007, 5.80),
          (1511, 5.45),
          (float("inf"), 4.95),
      ],
      "+ Perso. Nom (Cœur)": [
          (3, 6.00),
          (11, 4.00),
          (23, 4.00),
          (47, 3.50),
          (95, 3.00),
          (251, 2.80),
          (503, 2.60),
          (1007, 2.40),
          (1511, 2.20),
          (float("inf"), 2.00),
      ],
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
    if montant_base < 99.99:
      return 14.95
    elif montant_base < 499.99:
      return 20.95
    elif montant_base < 999.99:
      return 24.95
    else:
      return 0.0
  elif zone == "Livraison Corse, Monaco ou Andorre":
    if montant_base < 99.99:
      return 19.95
    elif montant_base < 499.99:
      return 25.95
    elif montant_base < 999.99:
      return 29.95
    else:
      return 0.0
  else:
    if montant_base < 99.99:
      return 25.95
    elif montant_base < 499.99:
      return 39.0
    elif montant_base < 999.99:
      return 60.0
    elif montant_base < 1500.0:
      return 90.0
    else:
      return 0.0


if not os.path.exists(CRM_FILE):
  try:
    pd.DataFrame(
        columns=[
            "Numero_Devis",
            "Date",
            "Client",
            "Contact",
            "Email",
            "Telephone",
            "Total_HT",
            "Total_TTC",
            "Statut",
            "Commercial",
            "Mode_Reglement",
            "PDF_Path",
        ]
    ).to_csv(CRM_FILE, index=False)
  except Exception:
    pass


def enregistrer_dans_crm(devis_data):
  try:
    if os.path.exists(CRM_FILE):
      df_crm = pd.read_csv(CRM_FILE)
    else:
      df_crm = pd.DataFrame(columns=list(devis_data.keys()))

    if not df_crm[df_crm["Numero_Devis"] == devis_data["Numero_Devis"]].empty:
      df_crm.loc[
          df_crm["Numero_Devis"] == devis_data["Numero_Devis"], :
      ] = list(devis_data.values())
    else:
      df_crm = pd.concat([df_crm, pd.DataFrame([devis_data])], ignore_index=True)
    df_crm.to_csv(CRM_FILE, index=False)
  except Exception as e:
    st.error(f"Erreur CRM : {e}")


# --- SIDEBAR ---
st.sidebar.title("📋 Infos Client & Expédition")
client_nom = st.sidebar.text_input("Nom du Client", "Client Exemple")
client_entreprise = st.sidebar.text_input("Société / Entreprise", "")
client_siret = st.sidebar.text_input("SIRET", "")
client_contact = st.sidebar.text_input("Nom de contact", "")
client_adresse = st.sidebar.text_area(
    "Adresse complète", "1 rue de l'Exemple\n70000 Vesoul"
)
client_email = st.sidebar.text_input("Email", "client@exemple.com")
client_contact_tel = st.sidebar.text_input("Téléphone", "0600000000")

logo_file = st.sidebar.file_uploader(
    "Logo entreprise (PNG/JPG)", type=["png", "jpg", "jpeg"]
)

st.sidebar.markdown("---")
zone_livraison = st.sidebar.selectbox(
    "Zone de Livraison",
    [
        "France Continentale",
        "Livraison Corse, Monaco ou Andorre",
        "Espace UE",
    ],
)
offrir_port = st.sidebar.checkbox("🎁 Offrir les frais de port", value=False)
conseiller_nom = st.sidebar.selectbox(
    "Commercial / Conseiller", ["Brice Geny", "Brice Bugna"]
)
mode_reglement = st.sidebar.selectbox(
    "Mode de Règlement",
    [
        "Virement bancaire 30 jours",
        "Comptant à la commande",
        "50% à la commande, 50% à 30 jours",
    ],
)

# --- INTERFACE PRINCIPALE ---
noms_onglets = [f"Article {i+1}" for i in range(10)] + [
    "📊 Général & Devis",
    "📈 Suivi CRM",
]
onglets = st.tabs(noms_onglets)

articles_saisis = []
total_textile_brut = 0.0

for i in range(10):
  with onglets[i]:
    st.subheader(f"Configuration de l'Article {i+1}")
    metier_type = st.radio(
        f"Univers / Métier pour l'Article {i+1}",
        [
            "👕 Textile & Marquage (DTF / Broderie)",
            "📄 Print, Papeterie & Signalétique (Catalogue APEX)",
        ],
        key=f"metier_{i}",
    )
    st.markdown("---")

    if "Textile" in metier_type:
      sans_marquage = st.checkbox(
          f"Vêtement sans marquage (fourniture seule) {i+1}",
          key=f"sans_marq_{i}",
      )
      col1, col2 = st.columns(2)
      with col1:
        nom_article = st.text_input(
            f"Référence / Nom du vêtement {i+1}",
            value="T-Shirt 100% coton bio" if i == 0 else f"Vêtement {i+1}",
            key=f"nom_textile_{i}",
        )
        qte = st.number_input(
            f"Quantité (pcs) {i+1}",
            min_value=0,
            value=10 if i == 0 else 0,
            key=f"qte_textile_{i}",
        )
        prix_vetement_ht = st.number_input(
            f"Prix unitaire HT support (€) {i+1}",
            min_value=0.0,
            value=4.92,
            format="%.2f",
            key=f"px_textile_{i}",
        )
      with col2:
        nb_marquages = (
            st.selectbox(
                f"Nombre de marquages {i+1}", [1, 2, 3, 4], key=f"nb_m_textile_{i}"
            )
            if not sans_marquage
            else 0
        )

      marquages = []
      if not sans_marquage:
        for m in range(nb_marquages):
          mc1, mc2 = st.columns(2)
          with mc1:
            t_marq = st.selectbox(
                f"Technique M{m+1}",
                ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"],
                key=f"t_marq_{i}_{m}",
            )
          with mc2:
            if "Broderie" in t_marq:
              emp = st.selectbox(
                  f"Emplacement M{m+1}",
                  [
                      "Poitrine (9x8 cm)",
                      "Dos D10 (25x10 cm)",
                      "Dos Large D20 (25x20 cm)",
                      "Col / Signature (7x2 cm)",
                      "Casquettes / Bonnets",
                      "Manche (8x5 cm)",
                      "Pantalon / Poche",
                      "+ Perso. Nom (Cœur)",
                  ],
                  key=f"emp_{i}_{m}",
              )
            else:
              emp = st.selectbox(
                  f"Emplacement M{m+1}",
                  [
                      "Cœur (13x9 cm)",
                      "Dos D10 (20x13 cm)",
                      "Dos D20 (28x20 cm)",
                      "Format P (37x27 cm)",
                      "Manche (9x8 cm)",
                      "+ Personnalisation Nom",
                  ],
                  key=f"emp_{i}_{m}",
              )
          marquages.append({"technique": t_marq, "emplacement": emp})

      option_ensachage = st.checkbox(
          f"Option ensachage individuel {i+1}", key=f"ens_{i}"
      )
      if option_ensachage:
        type_sachet = st.selectbox(
            f"Type de sachet {i+1}",
            ["Sachet (T-shirt/Polo)", "Sachet (Veste/Sweat)"],
            key=f"tsach_{i}",
        )
      else:
        type_sachet = ""

      option_assurance = st.checkbox(
          f"Option assurance garantie textile {i+1}", key=f"ass_{i}"
      )
      option_stockage = st.checkbox(
          f"Option stockage + picking {i+1}", key=f"stock_{i}"
      )
      remise_fidelite = st.number_input(
          f"Réduction fidélité (%) {i+1}",
          min_value=0.0,
          max_value=100.0,
          value=0.0,
          key=f"rem_{i}",
      )

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
            "remise_fidelite": remise_fidelite,
        })
        total_textile_brut += qte * prix_vetement_ht
    else:
      cat_print = st.selectbox(
          f"Catégorie Print & Signalétique {i+1}",
          [
              "Flyers",
              "Dépliants",
              "Blocs notes",
              "Chemises de présentation",
              "Banderoles",
              "Panneaux de chantier",
              "Roll-Up",
              "Sous bocks",
              "Adhésifs",
              "Cartes de visite",
              "Calendriers",
              "Menus restaurants",
              "➕ Autre / Produit hors catalogue (Saisie libre)",
          ],
          key=f"cat_print_{i}",
      )

      if "Autre" in cat_print:
        choix_ref = st.text_input(
            f"Nom / Désignation du produit libre {i+1}",
            value="Produit personnalisé",
            key=f"ref_libre_{i}",
        )
        col1, col2 = st.columns(2)
        with col1:
          qte = st.number_input(
              f"Quantité (exemplaires) {i+1}",
              min_value=0,
              value=1 if i == 0 else 0,
              key=f"qte_print_{i}",
          )
          prix_vetement_ht = st.number_input(
              f"Prix unitaire HT (€) {i+1} (Saisie libre)",
              min_value=0.0,
              value=10.00,
              format="%.4f",
              key=f"px_libre_{i}",
          )
        with col2:
          st.info("💡 Saisie manuelle active (produit hors catalogue).")
          remise_fidelite = st.number_input(
              f"Remise commerciale (%) {i+1}",
              min_value=0.0,
              max_value=100.0,
              value=0.0,
              key=f"rem_print_{i}",
          )
      else:
        options_articles = {
            "Flyers": [
                "Flyer A6 - 135g couché brillant - Recto",
                "Flyer A6 - 135g couché brillant - Recto/Verso",
                "Flyer A6 - 170g couché demi mat - Recto",
                "Flyer A6 - 170g couché demi mat - Recto/Verso",
                "Flyer A6 - 250g couché - Recto",
                "Flyer A6 - 250g couché - Recto/Verso",
                "Flyer A6 - 350g couché - Recto",
                "Flyer A6 - 350g couché - Recto/Verso",
                "Flyer A6 - 115g recyclé - Recto",
                "Flyer A6 - 115g recyclé - Recto/Verso",
                "Flyer A5 - 135g couché brillant - Recto",
                "Flyer A5 - 135g couché brillant - Recto/Verso",
                "Flyer A5 - 170g couché demi mat - Recto",
                "Flyer A5 - 170g couché demi mat - Recto/Verso",
            ],
            "Dépliants": [
                (
                    "Dépliant A6 fermé / A5 ouvert (1 pli) - 135g couché"
                    " brillant"
                ),
                (
                    "Dépliant A5 fermé / A4 ouvert (1 pli) - 135g couché"
                    " brillant"
                ),
            ],
            "Blocs notes": [
                "Bloc Note collé - Format A6 - 25 Feuilles - 90 Gr Offset",
                "Bloc Note collé - Format A5 - 50 Feuilles - 90 Gr Offset",
            ],
            "Chemises de présentation": [
                "Chemise de présentation A4 - 300g - 2 rabats"
            ],
            "Banderoles": [
                "Banderole 200 x 80 cm - 510g M1 avec œillets",
                "Banderole 300 x 100 cm - 510g M1 avec œillets",
            ],
            "Panneaux de chantier": [
                "Panneau Akylux 60 x 40 cm - 3,5mm",
                "Panneau Akylux 80 x 60 cm - 3,5mm",
            ],
            "Roll-Up": ["Roll-Up Eco - Bâche PVC 510g M1 - 85x200cm"],
            "Sous bocks": ["Sous bock carton 580g - 9,3x9,3 cm"],
            "Adhésifs": ["Adhésif vinyl classique 10x10cm"],
            "Cartes de visite": ["Carte de visite standard - 350g - Recto/Verso"],
            "Calendriers": ["Calendrier A4 - 250g couché brillant"],
            "Menus restaurants": ["Menu restaurant indéchirable 300g - A5"],
        }

        choix_ref = st.selectbox(
            f"Modèle exact {i+1}",
            options_articles.get(cat_print, ["Article standard"]),
            key=f"ref_print_{i}",
        )

        col1, col2 = st.columns(2)
        with col1:
          qte = st.number_input(
              f"Quantité (exemplaires) {i+1}",
              min_value=0,
              value=100 if i == 0 else 0,
              key=f"qte_print_{i}",
          )

          prix_unitaire_auto = obtenir_prix_catalogue_intelligent(
              cat_print, choix_ref, qte
          )
          prix_vetement_ht = prix_unitaire_auto

          st.metric(
              label=f"Prix unitaire HT (€) {i+1} (Catalogue auto)",
              value=f"{prix_unitaire_auto:.4f} €",
          )

        with col2:
          st.success(
              f"✅ Tarif appliqué ({qte} ex) : **{prix_unitaire_auto:.4f} € HT**"
          )
          remise_fidelite = st.number_input(
              f"Remise commerciale (%) {i+1}",
              min_value=0.0,
              max_value=100.0,
              value=0.0,
              key=f"rem_print_{i}",
          )

      if qte > 0:
        articles_saisis.append({
            "type_univers": "print",
            "nom_article": (
                f"{cat_print} - {choix_ref}"
                if "Autre" not in cat_print
                else choix_ref
            ),
            "quantite": qte,
            "prix_vet_unit": prix_vetement_ht,
            "sans_marquage": True,
            "marquages": [],
            "option_ensachage": False,
            "type_sachet": "",
            "option_assurance": False,
            "option_stockage": False,
            "remise_fidelite": remise_fidelite,
        })

# --- CALCUL DES QUANTITÉS CUMULÉES ---
quantites_cumulees_marquages = {}
for item in articles_saisis:
  if item["type_univers"] == "textile" and not item["sans_marquage"]:
    q = item["quantite"]
    for m in item["marquages"]:
      cle = (m["technique"], m["emplacement"])
      quantites_cumulees_marquages[cle] = (
          quantites_cumulees_marquages.get(cle, 0) + q
      )

frais_tech_auto = 19.80 if total_textile_brut > 0 else 0.0

# --- ONGLET GÉNÉRAL & DEVIS ---
with onglets[10]:
  st.subheader(
      "📊 Récapitulatif Général & Génération du Devis Professionnel"
  )

  if not articles_saisis:
    st.warning(
        "Veuillez renseigner au moins un article avec une quantité supérieure à"
        " 0."
    )
  else:
    supprimer_frais_tech = st.checkbox(
        "⚙️ Supprimer / Offrir les frais techniques de dossier", value=False
    )
    frais_techniques_dossier = (
        0.0 if supprimer_frais_tech else frais_tech_auto
    )

    lignes_devis_global = []
    has_broderie_global = False
    quantite_totale_broderie = 0

    for item in articles_saisis:
      q = item["quantite"]
      px_support = item["prix_vet_unit"] * (
          1 - item["remise_fidelite"] / 100.0
      )
      tot_support = q * px_support

      marquages_calcules = []
      tot_marquages = 0.0

      if item["type_univers"] == "textile" and not item["sans_marquage"]:
        for m in item["marquages"]:
          cle = (m["technique"], m["emplacement"])
          qte_tot_ref = quantites_cumulees_marquages.get(cle, q)

          if "Broderie" in m["technique"]:
            tarif_m = obtenir_tarif_broderie_unitaire(
                m["emplacement"], qte_tot_ref
            )
            has_broderie_global = True
            quantite_totale_broderie += q
          else:
            tarif_m = obtenir_tarif_dtf_unitaire(
                m["technique"], m["emplacement"], qte_tot_ref
            )

          tot_marquages += tarif_m * q
          marquages_calcules.append({
              "nom": f"{m['technique']} ({m['emplacement']})",
              "tarif": tarif_m,
          })

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
        coût_ens_unit = (
            1.38
            if q <= 11
            else (
                1.24
                if q <= 24
                else (
                    1.17
                    if q <= 49
                    else (
                        1.11
                        if q <= 99
                        else (
                            1.08
                            if q <= 249
                            else (1.06 if q <= 499 else 1.00)
                        )
                    )
                )
            )
        )
      else:
        coût_ens_unit = 0.0
      tot_ens = coût_ens_unit * q

      if item["option_assurance"]:
        coût_ass_unit = (
            3.08
            if q <= 11
            else (
                2.38
                if q <= 24
                else (
                    1.83
                    if q <= 49
                    else (
                        1.25
                        if q <= 99
                        else (
                            0.98
                            if q <= 249
                            else (
                                0.70
                                if q <= 499
                                else (0.64 if q <= 999 else 0.61)
                            )
                        )
                    )
                )
            )
        )
      else:
        coût_ass_unit = 0.0
      tot_ass = coût_ass_unit * q

      if item["option_stockage"]:
        coût_stock_unit = (
            1.00
            if q <= 99
            else (
                0.56
                if q <= 249
                else (0.50 if q <= 499 else (0.43 if q <= 999 else 0.30))
            )
        )
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
          "total_ligne_ht": tot_ligne,
      })

    sous_total_articles = sum(
        [item["total_ligne_ht"] for item in lignes_devis_global]
    )
    frais_prog_total = (
        lignes_devis_global[0]["frais_prog_broderie"]
        if lignes_devis_global
        else 0.0
    )
    montant_base_port = (
        sous_total_articles + frais_prog_total + frais_techniques_dossier
    )

    frais_port = (
        0.0
        if offrir_port
        else calculer_frais_port(montant_base_port, zone_livraison)
    )

    total_general_ht = montant_base_port + frais_port
    tva = total_general_ht * 0.20
    total_ttc = total_general_ht + tva

    quantite_globale_totale = sum(
        [item["quantite"] for item in lignes_devis_global]
    )
    cout_unitaire_moyen = (
        total_general_ht / quantite_globale_totale
        if quantite_globale_totale > 0
        else 0
    )

    st.write(f"**Quantité globale pièces :** {quantite_globale_totale}")
    st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
    if frais_prog_total > 0:
      st.write(
          f"**Frais de technique & programme Broderie :**"
          f" {frais_prog_total:.2f} € HT"
      )
    st.write(
        f"**Frais techniques de dossier :** {frais_techniques_dossier:.2f} € HT"
    )
    st.write(
        f"**Frais de port ({zone_livraison}) :** {frais_port:.2f} € HT"
        if not offrir_port
        else "**Frais de port :** Offerts (0.00 €)"
    )
    st.markdown(
        f"### **Total Général HT : {total_general_ht:.2f} €** | **TOTAL TTC"
        f" (20%) : {total_ttc:.2f} €**"
    )

    if st.button("📄 Générer le numéro de devis et le PDF"):
      num_devis = obtenir_prochain_numero_devis()
      pdf_filename = os.path.join(
          PDF_DIR, f"Devis_{num_devis.replace('/', '_')}.pdf"
      )
      st.session_state["dernier_pdf"] = pdf_filename
      st.session_state["dernier_num"] = num_devis

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
          "PDF_Path": pdf_filename,
      }
      enregistrer_dans_crm(data_crm)
      st.success("✅ Données enregistrées dans le CRM avec succès !")

      try:
        doc = SimpleDocTemplate(
            pdf_filename,
            pagesize=A4,
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30,
        )
        story = []
        styles = getSampleStyleSheet()

        style_sub = ParagraphStyle(
            "SubStyle",
            parent=styles["Normal"],
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#4a5568"),
        )