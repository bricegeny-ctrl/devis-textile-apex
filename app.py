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
  annee = datetime.now().strftime("%Y")
  mois = datetime.now().strftime("%m")
  jour = datetime.now().strftime("%d")
  prefix = f"{annee}/{mois}/{jour}"
  seq = 1
  if os.path.exists(COMPTEUR_FILE):
    try:
      with open(COMPTEUR_FILE, "r") as f:
        data = json.load(f)
        dernier = data.get("dernier_num", "")
        if "_" in dernier and dernier.startswith(prefix):
          parts = dernier.split("_")
          if len(parts) > 1 and parts[1].isdigit():
            seq = int(parts[1]) + 1
    except:
      seq = 1
  nouveau = f"{prefix}_{seq:08d}"
  try:
    with open(COMPTEUR_FILE, "w") as f:
      json.dump({"dernier_num": nouveau}, f)
  except:
    pass
  return nouveau


def obtenir_prix_catalogue_intelligent(cat_print, choix_ref, qte):
  if not os.path.exists(CATALOGUE_FILE):
    return 0.15
  try:
    df_all = pd.read_excel(CATALOGUE_FILE, sheet_name=0, header=None)
  except:
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
    rc = (
        str(df_all.iloc[r, 0]).lower().strip()
        if pd.notna(df_all.iloc[r, 0])
        else ""
    )
    rs = (
        str(df_all.iloc[r, 1]).lower().strip()
        if pd.notna(df_all.iloc[r, 1])
        else ""
    )
    rf = (
        str(df_all.iloc[r, 2]).lower().strip()
        if pd.notna(df_all.iloc[r, 2])
        else ""
    )
    score = 0
    if cat_lower in rc or rc in cat_lower:
      score += 10
    mots = [m for m in ref_lower.split() if len(m) > 2]
    match_mots = sum(1 for m in mots if m in rs or m in rf or rc in m)
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
    prix = float(df_all.iloc[best_row, target_col])
    if pd.isna(prix) or prix <= 0:
      for alt in range(df_all.shape[1] - 1, 2, -1):
        val_alt = float(df_all.iloc[best_row, alt])
        if not pd.isna(val_alt) and val_alt > 0:
          prix = val_alt
          break
    if pd.isna(prix) or prix <= 0:
      return 0.15
    return round(prix, 4)
  except:
    return 0.15


st.sidebar.title("Infos Client & Expédition")
client_nom = st.sidebar.text_input("Nom du Client", "Client Exemple")
client_entreprise = st.sidebar.text_input("Société / Entreprise", "")
client_siret = st.sidebar.text_input("SIRET", "")
client_contact = st.sidebar.text_input("Nom de contact", "")
client_adresse = st.sidebar.text_area(
    "Adresse", "1 rue de l'Exemple\n70000 Vesoul"
)
client_email = st.sidebar.text_input("Email", "client@exemple.com")
client_contact_tel = st.sidebar.text_input("Téléphone", "0600000000")
logo_file = st.sidebar.file_uploader("Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
st.sidebar.markdown("---")
zone_livraison = st.sidebar.selectbox(
    "Zone de Livraison",
    [
        "France Continentale",
        "Livraison Corse, Monaco ou Andorre",
        "Espace UE",
    ],
)
offrir_port = st.sidebar.checkbox("Offrir les frais de port", value=False)
conseiller_nom = st.sidebar.selectbox(
    "Commercial", ["Brice Geny", "Brice Bugna"]
)
mode_reglement = st.sidebar.selectbox(
    "Mode de Règlement",
    [
        "Virement bancaire 30 jours",
        "Comptant à la commande",
        "50% à la commande, 50% à 30 jours",
    ],
)

st.title("Gestionnaire de Devis - APEX")
st.write(
    "Application prête. Veuillez configurer vos articles et générer vos devis"
    " en toute sérénité."
)