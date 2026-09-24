from datetime import datetime
from email.message import EmailMessage
import json
import os
import smtplib
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Image as RLImage, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

st.set_page_config(
    page_title="SAS ABCOM - Gestionnaire Global & Configurateur", layout="wide"
)

# --- NAVIGATION PRINCIPALE DANS LA BARRE LATÉRALE ---
st.sidebar.title("🖨️ SAS ABCOM - Portail")
module_principal = st.sidebar.selectbox(
    "Choisissez le module :",
    [
        "👕 Marquage Textile & Broderie",
        "🖨️ Solutions Print & Signalétique",
    ],
)
st.sidebar.markdown("---")

# =====================================================================
# MODULE 1 : MARQUAGE TEXTILE & BRODERIE
# =====================================================================
if module_principal == "👕 Marquage Textile & Broderie":

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


  def mettre_a_jour_statut_crm(numero_devis, nouveau_statut):
    if os.path.exists(CRM_FILE):
      df = pd.read_csv(CRM_FILE)
      if "Numero_Devis" in df.columns:
        df.loc[df["Numero_Devis"] == numero_devis, "Statut"] = nouveau_statut
        df.to_csv(CRM_FILE, index=False)


  # --- FONCTIONS DE CALCUL TEXTILE ---
  def get_tarif_dtf_fin(qte_globale, emplacement):
    if qte_globale <= 5:
      idx = 0
    elif qte_globale <= 9:
      idx = 1
    elif qte_globale <= 19:
      idx = 2
    elif qte_globale <= 29:
      idx = 3
    elif qte_globale <= 39:
      idx = 4
    elif qte_globale <= 49:
      idx = 5
    elif qte_globale <= 99:
      idx = 6
    elif qte_globale <= 249:
      idx = 7
    elif qte_globale <= 499:
      idx = 8
    elif qte_globale <= 999:
      idx = 9
    elif qte_globale <= 4999:
      idx = 10
    else:
      idx = 11

    tarifs = {
        "Cœur (13x9 cm)": [
            6.00,
            4.50,
            3.60,
            2.81,
            2.50,
            2.40,
            2.00,
            1.80,
            1.60,
            1.40,
            1.20,
            0.70,
        ],
        "Dos D10 (20x13 cm)": [
            7.92,
            6.50,
            5.50,
            5.00,
            4.20,
            4.00,
            3.50,
            3.00,
            2.70,
            2.50,
            2.00,
            1.00,
        ],
        "Dos D20 (28x20 cm)": [
            13.67,
            10.00,
            9.00,
            6.90,
            6.30,
            6.00,
            5.50,
            4.60,
            4.00,
            3.50,
            2.50,
            1.50,
        ],
        "Format P (37x27 cm)": [
            17.00,
            13.00,
            12.00,
            10.00,
            9.00,
            8.00,
            7.50,
            6.50,
            6.00,
            5.00,
            4.00,
            2.50,
        ],
        "Manche (9x8 cm)": [
            7.20,
            5.40,
            4.32,
            3.37,
            3.00,
            2.88,
            2.40,
            2.16,
            1.92,
            1.68,
            1.44,
            0.84,
        ],
        "+ Perso. Nom": [
            4.39,
            3.50,
            2.50,
            2.20,
            1.90,
            1.80,
            1.70,
            1.50,
            1.30,
            0.80,
            0.30,
            0.20,
        ],
    }
    return tarifs.get(emplacement, [0] * 12)[idx]


  def get_tarif_dtf_epais(qte_globale, emplacement):
    if qte_globale <= 5:
      idx = 0
    elif qte_globale <= 9:
      idx = 1
    elif qte_globale <= 19:
      idx = 2
    elif qte_globale <= 29:
      idx = 3
    elif qte_globale <= 39:
      idx = 4
    elif qte_globale <= 49:
      idx = 5
    elif qte_globale <= 99:
      idx = 6
    elif qte_globale <= 249:
      idx = 7
    elif qte_globale <= 499:
      idx = 8
    elif qte_globale <= 999:
      idx = 9
    elif qte_globale <= 4999:
      idx = 10
    else:
      idx = 11

    tarifs = {
        "Cœur (13x9 cm)": [
            6.75,
            5.06,
            4.05,
            3.16,
            2.81,
            2.70,
            2.25,
            2.03,
            1.80,
            1.58,
            1.35,
            0.79,
        ],
        "Dos D10 (20x13 cm)": [
            8.91,
            7.31,
            6.19,
            5.63,
            4.73,
            4.50,
            3.94,
            3.38,
            3.04,
            2.81,
            2.25,
            1.13,
        ],
        "Dos D20 (28x20 cm)": [
            15.38,
            11.25,
            10.13,
            7.76,
            7.09,
            6.75,
            6.19,
            5.18,
            4.50,
            3.94,
            2.81,
            1.69,
        ],
        "Format P (37x27 cm)": [
            19.13,
            14.63,
            13.50,
            11.25,
            10.13,
            9.00,
            8.44,
            7.31,
            6.75,
            5.63,
            4.50,
            2.81,
        ],
        "Manche (9x8 cm)": [
            8.10,
            6.08,
            4.86,
            3.79,
            3.38,
            3.24,
            2.70,
            2.43,
            2.16,
            1.89,
            1.62,
            0.95,
        ],
        "+ Perso. Nom": [
            4.94,
            3.94,
            2.81,
            2.48,
            2.14,
            2.03,
            1.91,
            1.69,
            1.46,
            0.90,
            0.34,
            0.23,
        ],
    }
    return tarifs.get(emplacement, [0] * 12)[idx]


  # Interface saisie client textile
  st.title("👕 SAS ABCOM - Devis & Marquage Textile")
  st.sidebar.header("📋 Infos Client")
  client_entreprise = st.sidebar.text_input(
      "Nom de l'Entreprise / Société",
      value=st.session_state.get("duplique_entreprise", ""),
  )
  client_nom = st.sidebar.text_input(
      "Nom du contact", value=st.session_state.get("duplique_client", "")
  )
  client_adresse = st.sidebar.text_area("Adresse complète")
  client_email = st.sidebar.text_input(
      "E-mail", value=st.session_state.get("duplique_email", "")
  )

  st.info(
      "Module textile actif : configurez vos articles, marquages DTF ou"
      " broderies."
  )

  # --- GESTION CRM & HISTORIQUE TEXTILE ---
  st.markdown("---")
  st.header("📊 CRM & Historique des Devis")
  if os.path.exists(CRM_FILE):
    df_crm = pd.read_csv(CRM_FILE)
    if not df_crm.empty:
      st.dataframe(df_crm, use_container_width=True)
    else:
      st.info("Aucun devis enregistré pour le moment.")
  else:
    st.info("Le CRM est vide.")

# =====================================================================
# MODULE 2 : SOLUTIONS PRINT & SIGNALÉTIQUE (CONFIGURATEUR MENUS DÉROULANTS)
# =====================================================================
else:
  st.title("🖨️ SAS ABCOM - Configurateur & Devis Print")
  st.markdown(
      "Sélectionnez une famille de produits et configurez vos options via les"
      " menus déroulants."
  )

  EXCEL_FILE = r"C:\Users\setup\OneDrive\OneDrive - IRIS - AB Com\Bureau\reprise\site et appli\tarifs print-panneaux-banderoles.xlsx"


  @st.cache_data
  def load_print_data():
    try:
      if not os.path.exists(EXCEL_FILE):
        return None
      return pd.read_excel(EXCEL_FILE, sheet_name="Feuil1")
    except Exception:
      return None


  df_print = load_print_data()

  if df_print is None:
    st.error(
        f"⚠️ Fichier introuvable ou erreur de lecture au chemin : {EXCEL_FILE}"
    )
  else:
    famille = st.selectbox(
        "📌 Choisissez une catégorie de produits :",
        [
            "Blocs-Notes",
            "Chemises de présentation",
            "Banderoles",
            "Panneaux de chantier",
            "Roll-Ups",
            "Flyers",
            "Dépliants",
            "Menus restaurants",
            "Sous-bocks",
            "Adhésifs",
            "Cartes de visite",
            "Calendriers",
        ],
    )

    st.markdown("---")

    if famille == "Blocs-Notes":
      st.subheader("📝 Blocs-Notes (Papier 90g Offset)")
      sous_df = df_print.iloc[3:9].copy()
      sous_df.columns = [
          "Réf",
          "Désignation",
          "P1",
          "50 ex",
          "100 ex",
          "200 ex",
          "500 ex",
          "1000 ex",
          "C8",
          "C9",
          "C10",
          "C11",
      ]
      choix_produit = st.selectbox(
          "Sélectionnez le format et nombre de feuilles :",
          sous_df["Désignation"].tolist(),
      )
      ligne = sous_df[sous_df["Désignation"] == choix_produit].iloc[0]
      st.info(f"**Référence :** {ligne['Réf']}")
      st.write(f"**Description :** {choix_produit}")

      col1, col2, col3, col4 = st.columns(4)
      with col1:
        if pd.notna(ligne["50 ex"]):
          st.metric("50 ex", f"{ligne['50 ex']} € HT /u")
      with col2:
        if pd.notna(ligne["100 ex"]):
          st.metric("100 ex", f"{ligne['100 ex']} € HT /u")
      with col3:
        if pd.notna(ligne["200 ex"]):
          st.metric("200 ex", f"{ligne['200 ex']} € HT /u")
      with col4:
        if pd.notna(ligne["500 ex"]):
          st.metric("500 ex", f"{ligne['500 ex']} € HT /u")

    elif famille == "Chemises de présentation":
      st.subheader("📂 Chemises de présentation (300g)")
      sous_df = df_print.iloc[16:20].copy()
      sous_df.columns = [
          "Réf",
          "Désignation",
          "PU1",
          "100 ex",
          "250 ex",
          "500 ex",
          "1000 ex",
          "2500 ex",
          "5000 ex",
          "C9",
          "C10",
          "C11",
      ]
      choix_produit = st.selectbox(
          "Sélectionnez le modèle de chemise :", sous_df["Désignation"].tolist()
      )
      ligne = sous_df[sous_df["Désignation"] == choix_produit].iloc[0]
      st.info(f"**Référence :** {ligne['Réf']}")
      st.write(f"**Modèle :** {choix_produit}")

      c1, c2, c3, c4 = st.columns(4)
      with c1:
        if pd.notna(ligne["100 ex"]):
          st.metric("100 ex", f"{ligne['100 ex']} € HT /u")
      with c2:
        if pd.notna(ligne["250 ex"]):
          st.metric("250 ex", f"{ligne['250 ex']} € HT /u")
      with c3:
        if pd.notna(ligne["500 ex"]):
          st.metric("500 ex", f"{ligne['500 ex']} € HT /u")
      with c4:
        if pd.notna(ligne["1000 ex"]):
          st.metric("1000 ex", f"{ligne['1000 ex']} € HT /u")

    elif famille == "Banderoles":
      st.subheader("🚩 Banderoles (M1 510g/m² avec œillets)")
      sous_df = df_print.iloc[29:34].dropna(how="all").copy()
      sous_df.columns = ["Format / Finition", "Prix HT (€)", "C2", "C3", "C4"]
      choix_format = st.selectbox(
          "Sélectionnez le format de la banderole :",
          sous_df["Format / Finition"].tolist(),
      )
      ligne = sous_df[sous_df["Format / Finition"] == choix_format].iloc[0]
      st.success(
          f"### Tarif pour {choix_format} : **{ligne['Prix HT (€)']} € HT**"
      )

    elif famille == "Panneaux de chantier":
      st.subheader("🚧 Panneaux de chantier (Akylux)")
      sous_df = df_print.iloc[39:45].dropna(how="all").copy()
      sous_df.columns = ["Réf", "Désignation", "Prix"]
      choix_panneau = st.selectbox(
          "Sélectionnez le panneau de chantier :", sous_df["Désignation"].tolist()
      )
      ligne = sous_df[sous_df["Désignation"] == choix_panneau].iloc[0]
      st.info(f"**Référence :** {ligne['Réf']}")
      st.success(f"**Tarif de base :** {ligne['Prix']} € HT")

    elif famille == "Roll-Ups":
      st.subheader("📌 Roll-Ups (Structures enroulables)")
      sous_df = df_print.iloc[53:57].dropna(how="all").copy()
      sous_df.columns = ["Réf", "Désignation"]
      choix_ru = st.selectbox(
          "Sélectionnez le type de Roll-Up :", sous_df["Désignation"].tolist()
      )
      ligne = sous_df[sous_df["Désignation"] == choix_ru].iloc[0]
      st.info(f"**Référence :** {ligne['Réf']}")
      st.write(f"**Modèle choisi :** {choix_ru}")

    else:
      st.subheader(f"📄 {famille}")
      mapping_lignes = {
          "Flyers": (59, 84),
          "Dépliants": (86, 111),
          "Menus restaurants": (113, 117),
          "Sous-bocks": (119, 123),
          "Adhésifs": (127, 137),
          "Cartes de visite": (141, 150),
          "Calendriers": (151, 182),
      }
      if famille in mapping_lignes:
        debut, fin = mapping_lignes[famille]
        sous_df = df_print.iloc[debut:fin].dropna(how="all").copy()
        texte_cols = [
            col
            for col in sous_df.columns
            if sous_df[col].dtype == object and sous_df[col].nunique() > 1
        ]
        if texte_cols:
          col_choix = texte_cols[0]
          options_produits = sous_df[col_choix].dropna().unique().tolist()
          if options_produits:
            selection = st.selectbox(
                "Affiner le produit :", options_produits, key=f"sel_{famille}"
            )
            ligne_sel = sous_df[sous_df[col_choix] == selection]
            st.write("Détails et tarifs correspondants :")
            st.write(ligne_sel.dropna(axis=1, how="all").to_dict(orient="records"))