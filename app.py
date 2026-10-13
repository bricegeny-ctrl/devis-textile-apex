from datetime import datetime
import json
import os
import urllib.parse
from google.oauth2.service_account import Credentials
import gspread
from gspread.exceptions import SpreadsheetNotFound
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import requests
import streamlit as st

CATALOGUE_FILE = "catalogue_print.xlsx"
COMPTEUR_FILE = "compteur_devis.json"
PDF_DIR = "devis_pdf"

st.set_page_config(page_title="Gestionnaire de Devis - APEX", layout="wide")


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


# --- CONNEXION GOOGLE SHEETS DIRECTE (GSPREAD) ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def obtenir_client_gspread():
  try:
    if "gcp_service_account" in st.secrets:
      creds_dict = dict(st.secrets["gcp_service_account"])
      creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
      return gspread.authorize(creds)
  except Exception as e:
    st.error(f"Erreur d'authentification Google Sheets : {e}")
  return None


def charger_crm_gsheets():
  client = obtenir_client_gspread()
  if client:
    try:
      sheet = client.open("CRM_Apex").sheet1
      data = sheet.get_all_records()
      if data:
        return pd.DataFrame(data)
    except Exception as e:
      pass
  return pd.DataFrame(
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
          "Deal_ID",
      ]
  )


def enregistrer_dans_crm_gsheets(devis_data):
  if "Lignes_Articles" in devis_data:
    devis_data.pop("Lignes_Articles")

  devis_data_propre = {
      k: ("" if pd.isna(v) else v) for k, v in devis_data.items()
  }

  client = obtenir_client_gspread()
  if not client:
    st.error(
        "❌ Erreur : Impossible de se connecter à Google Sheets (Vérifiez les"
        " secrets)."
    )
    return

  try:
    sh = client.open("CRM_Apex")
    sheet = sh.sheet1

    entetes_actuels = sheet.row_values(1)

    if not entetes_actuels:
      entetes_actuels = list(devis_data_propre.keys())
      sheet.append_row(entetes_actuels)

    ligne_a_ecrire = [devis_data_propre.get(col, "") for col in entetes_actuels]

    cellule_num = None
    try:
      cellule_num = sheet.find(str(devis_data_propre.get("Numero_Devis", "")))
    except Exception:
      pass

    if cellule_num:
      sheet.update(f"A{cellule_num.row}", [ligne_a_ecrire])
    else:
      sheet.append_row(ligne_a_ecrire)

    st.success("✅ Devis enregistré avec succès dans le CRM Google Sheets !")
  except Exception as e:
    st.error(f"❌ Erreur détaillée lors de l'enregistrement CRM : {str(e)}")


def mettre_a_jour_statut_gsheets(num_devis, nouveau_statut):
  client = obtenir_client_gspread()
  if client:
    try:
      sheet = client.open("CRM_Apex").sheet1
      cellule_num = sheet.find(str(num_devis))
      if cellule_num:
        entetes = sheet.row_values(1)
        col_idx = (
            entetes.index("Statut") + 1
            if "Statut" in entetes
            else (
                entetes.index("Statut du Devis") + 1
                if "Statut du Devis" in entetes
                else 9
            )
        )
        sheet.update_cell(cellule_num.row, col_idx, nouveau_statut)
    except Exception as e:
      st.error(f"❌ Erreur lors de la mise à jour du statut dans GSheets : {e}")


# --- MOTEUR DE LECTURE EXCEL INTELLIGENT ---
def charger_catalogue_print():
  if not os.path.exists(CATALOGUE_FILE):
    return None
  try:
    df_all = pd.read_excel(CATALOGUE_FILE, sheet_name=0, header=0)
    for col in [0, 1, 2]:
      if col < df_all.shape[1]:
        df_all.iloc[:, col] = (
            df_all.iloc[:, col].astype(str).str.replace("\n", " ").str.strip()
        )
    return df_all
  except Exception as e:
    st.error(f"Erreur lors de la lecture du catalogue Excel : {e}")
    return None


def parse_fourchette_quantite(texte_entete):
  texte = str(texte_entete).strip()
  if "à" in texte:
    parts = texte.split("à")
    try:
      return float(parts[0].strip()), float(parts[1].strip())
    except:
      return None, None
  elif "et plus" in texte:
    num_str = texte.replace("et plus", "").strip()
    try:
      return float(num_str), float("inf")
    except:
      return None, None
  else:
    try:
      val = float(texte)
      return val, val
    except:
      return None, None


def obtenir_prix_catalogue_exact_robuste(
    cat_choisie, sub_choisie, ref_choisie, qte
):
  df_all = charger_catalogue_print()
  if df_all is None:
    return 0.15

  col_ranges = {}
  for c in range(3, df_all.shape[1]):
    r_min, r_max = parse_fourchette_quantite(df_all.columns[c])
    if r_min is not None:
      col_ranges[c] = (r_min, r_max)

  best_row = -1
  for r in range(len(df_all)):
    c_val = str(df_all.iloc[r, 0]).strip()
    s_val = str(df_all.iloc[r, 1]).strip()
    r_val = str(df_all.iloc[r, 2]).strip()

    if (
        c_val.lower() == str(cat_choisie).lower().strip()
        and s_val.lower() == str(sub_choisie).lower().strip()
        and r_val.lower() == str(ref_choisie).lower().strip()
    ):
      best_row = r
      break

  if best_row == -1:
    for r in range(len(df_all)):
      if (
          str(cat_choisie).lower().strip()
          in str(df_all.iloc[r, 0]).lower().strip()
      ):
        best_row = r
        break

  if best_row == -1 or not col_ranges:
    return 0.15

  premier_col_valide = None
  palier_min_qte = 0
  prix_palier_min = 0.15

  cols_tries = sorted(col_ranges.keys(), key=lambda x: col_ranges[x][0])

  for c in cols_tries:
    val = df_all.iloc[best_row, c]
    if pd.notna(val) and float(val) > 0:
      premier_col_valide = c
      palier_min_qte, _ = col_ranges[c]
      prix_palier_min = float(val)
      break

  if premier_col_valide is None:
    return 0.15

  if qte < palier_min_qte and qte > 0:
    prix_ajuste = (palier_min_qte * prix_palier_min) / qte
    return round(prix_ajuste, 4)

  target_col = None
  for c, (r_min, r_max) in col_ranges.items():
    if r_min <= qte <= r_max:
      target_col = c
      break

  if target_col is None and cols_tries:
    max_col = cols_tries[-1]
    if qte >= col_ranges[max_col][0]:
      target_col = max_col

  if target_col is None:
    return round(prix_palier_min, 4)

  try:
    val_prix = df_all.iloc[best_row, target_col]
    if pd.notna(val_prix) and float(val_prix) > 0:
      return round(float(val_prix), 4)
    else:
      return round(prix_palier_min, 4)
  except:
    return round(prix_palier_min, 4)


# --- SYNCHRONISATION HUBSPOT & RECHERCHE PAR SIRET ---
def chercher_entreprise_hubspot_par_siret(siret):
  siret_nettoye = "".join(filter(str.isdigit, str(siret)))
  if not siret_nettoye:
    return None

  hubspot_token = st.secrets.get("HUBSPOT_TOKEN", "")
  if not hubspot_token:
    return None

  headers = {
      "Authorization": f"Bearer {hubspot_token}",
      "Content-Type": "application/json",
  }
  url = "https://api.hubapi.com/crm/v3/objects/companies?limit=100&properties=name,address,city,zip,phone,siret"

  try:
    response = requests.get(url, headers=headers, timeout=5)
    if response.status_code == 200:
      entreprises = response.json().get("results", [])
      siret_9 = siret_nettoye[:9]
      for ent in entreprises:
        props = ent.get("properties", {})
        siret_hs = "".join(filter(str.isdigit, str(props.get("siret", ""))))
        if siret_9 in siret_hs or siret_nettoye in siret_hs:
          return props
  except:
    pass
  return None


def synchroniser_avec_hubspot(client_data):
  hubspot_token = st.secrets.get("HUBSPOT_TOKEN", "")
  if not hubspot_token or hubspot_token == "VOTRE_TOKEN_ACCES_PRIVE_HUBSPOT":
    return None

  headers = {
      "Authorization": f"Bearer {hubspot_token}",
      "Content-Type": "application/json",
  }

  siret_9 = client_data.get("SIRET_9", "")
  entreprise_nom = client_data.get("Entreprise", "").strip()
  nom_client = client_data.get("Client", "")
  email = client_data.get("Email", "")
  telephone = client_data.get("Telephone", "")
  commercial = client_data.get("Commercial", "")

  company_id = None
  if entreprise_nom and siret_9:
    search_url = "https://api.hubapi.com/crm/v3/objects/companies/search"
    search_payload = {
        "filterGroups": [{
            "filters": [{
                "propertyName": "siret",
                "operator": "EQ",
                "value": siret_9,
            }]
        }]
    }
    try:
      response = requests.post(
          search_url, headers=headers, json=search_payload, timeout=5
      )
      if response.status_code == 200:
        results = response.json().get("results", [])
        if results:
          company_id = results[0]["id"]
    except:
      pass

    if not company_id:
      create_company_url = "https://api.hubapi.com/crm/v3/objects/companies"
      company_payload = {
          "properties": {
              "name": entreprise_nom,
              "siret": siret_9,
              "description": f"Propriétaire / Commercial : {commercial}",
          }
      }
      try:
        resp_comp = requests.post(
            create_company_url,
            headers=headers,
            json=company_payload,
            timeout=5,
        )
        if resp_comp.status_code in [200, 201]:
          company_id = resp_comp.json().get("id")
      except:
        pass

  contact_id = None
  if email:
    contact_search_url = (
        "https://api.hubapi.com/crm/v3/objects/contacts/search"
    )
    contact_search_payload = {
        "filterGroups": [{
            "filters": [{
                "propertyName": "email",
                "operator": "EQ",
                "value": email,
            }]
        }]
    }
    try:
      resp_c_search = requests.post(
          contact_search_url,
          headers=headers,
          json=contact_search_payload,
          timeout=5,
      )
      if resp_c_search.status_code == 200:
        c_results = resp_c_search.json().get("results", [])
        if c_results:
          contact_id = c_results[0]["id"]
    except:
      pass

  parts_nom = nom_client.split(" ", 1)
  prenom = parts_nom[0]
  nom_famille = parts_nom[1] if len(parts_nom) > 1 else ""

  if not contact_id:
    create_contact_url = "https://api.hubapi.com/crm/v3/objects/contacts"
    contact_payload = {
        "properties": {
            "firstname": prenom,
            "lastname": nom_famille,
            "email": email,
            "phone": telephone,
        }
    }
    try:
      resp_contact = requests.post(
          create_contact_url, headers=headers, json=contact_payload, timeout=5
      )
      if resp_contact.status_code in [200, 201]:
        contact_id = resp_contact.json().get("id")
    except:
      pass

  if contact_id and company_id:
    assoc_url = f"https://api.hubapi.com/crm/v3/objects/contacts/{contact_id}/associations/companies/{company_id}/contact_to_company"
    try:
      requests.put(assoc_url, headers=headers, timeout=5)
    except:
      pass

  deal_id = None
  if company_id:
    deal_url = "https://api.hubapi.com/crm/v3/objects/deals"
    deal_payload = {
        "properties": {
            "dealname": (
                f"Devis {client_data.get('Numero_Devis')} -"
                f" {client_data.get('Entreprise')}"
            ),
            "amount": str(client_data.get("Total_HT", 0)),
            "dealstage": "appointmentscheduled",
            "pipeline": "default",
        }
    }
    try:
      resp_deal = requests.post(
          deal_url, headers=headers, json=deal_payload, timeout=5
      )
      if resp_deal.status_code in [200, 201]:
        deal_id = resp_deal.json().get("id")
        assoc_deal_url = f"https://api.hubapi.com/crm/v3/objects/deals/{deal_id}/associations/companies/{company_id}/deal_to_company"
        requests.put(assoc_deal_url, headers=headers, timeout=5)
    except:
      pass

  return deal_id


def mettre_a_jour_statut_deal_hubspot(deal_id, statut_str):
  if not deal_id or pd.isna(deal_id) or str(deal_id).strip() == "":
    return
  hubspot_token = st.secrets.get("HUBSPOT_TOKEN", "")
  if not hubspot_token:
    return

  mapping_stages = {
      "En cours": "appointmentscheduled",
      "Envoyé": "qualifiedtobuy",
      "Accepté (Gagné)": "closedwon",
      "Refusé (Perdu)": "closedlost",
  }
  stage_hs = mapping_stages.get(statut_str, "appointmentscheduled")

  url = f"https://api.hubapi.com/crm/v3/objects/deals/{deal_id}"
  headers = {
      "Authorization": f"Bearer {hubspot_token}",
      "Content-Type": "application/json",
  }
  payload = {"properties": {"dealstage": stage_hs}}
  try:
    requests.patch(url, headers=headers, json=payload, timeout=5)
  except:
    pass


# --- GRILLES TARIFAIRES OFFICIELLES ---
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


# --- INITIALISATION SESSION STATE DUPLICATION ---
if "client
