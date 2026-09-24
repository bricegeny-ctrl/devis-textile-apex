import os
import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="SAS ABCOM - Configurateur Print", page_icon="🖨️", layout="wide"
)

# Nom du fichier Excel (placé dans le même dossier que app.py)
EXCEL_FILE = "tarifs print-panneaux-banderoles.xlsx"


@st.cache_data
def load_data():
  if not os.path.exists(EXCEL_FILE):
    return None
  return pd.read_excel(EXCEL_FILE, sheet_name="Feuil1")


df = load_data()

# Titre principal
st.title("🖨️ SAS ABCOM - Configurateur & Devis Print")
st.markdown(
    "Sélectionnez une famille de produits et configurez vos options via les"
    " menus déroulants."
)

if df is None:
  st.error(
      f"⚠️ Fichier introuvable : '{EXCEL_FILE}'. Assurez-vous qu'il est bien"
      " dans le même dossier que app.py."
  )
else:
  # Menu déroulant principal pour choisir la catégorie de produits
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

  # --- CONFIGURATEUR : BLOCS-NOTES ---
  if famille == "Blocs-Notes":
    st.subheader("📝 Blocs-Notes (Papier 90g Offset)")
    sous_df = df.iloc[3:9].copy()
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

  # --- CONFIGURATEUR : CHEMISES DE PRÉSENTATION ---
  elif famille == "Chemises de présentation":
    st.subheader("📂 Chemises de présentation (300g)")
    sous_df = df.iloc[16:20].copy()
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

  # --- CONFIGURATEUR : BANDEROLES ---
  elif famille == "Banderoles":
    st.subheader("🚩 Banderoles (M1 510g/m² avec œillets)")
    sous_df = df.iloc[29:34].dropna(how="all").copy()
    sous_df.columns = ["Format / Finition", "Prix HT (€)", "C2", "C3", "C4"]

    choix_format = st.selectbox(
        "Sélectionnez le format de la banderole :",
        sous_df["Format / Finition"].tolist(),
    )
    ligne = sous_df[sous_df["Format / Finition"] == choix_format].iloc[0]

    st.success(f"### Tarif pour {choix_format} : **{ligne['Prix HT (€)']} € HT**")

  # --- CONFIGURATEUR : PANNEAUX DE CHANTIER ---
  elif famille == "Panneaux de chantier":
    st.subheader("🚧 Panneaux de chantier (Akylux)")
    sous_df = df.iloc[39:45].dropna(how="all").copy()
    sous_df.columns = ["Réf", "Désignation", "Prix"]

    choix_panneau = st.selectbox(
        "Sélectionnez le panneau de chantier :", sous_df["Désignation"].tolist()
    )
    ligne = sous_df[sous_df["Désignation"] == choix_panneau].iloc[0]

    st.info(f"**Référence :** {ligne['Réf']}")
    st.success(f"**Tarif de base :** {ligne['Prix']} € HT")

  # --- CONFIGURATEUR : ROLL-UPS ---
  elif famille == "Roll-Ups":
    st.subheader("📌 Roll-Ups (Structures enroulables)")
    sous_df = df.iloc[53:57].dropna(how="all").copy()
    sous_df.columns = ["Réf", "Désignation"]

    choix_ru = st.selectbox(
        "Sélectionnez le type de Roll-Up :", sous_df["Désignation"].tolist()
    )
    ligne = sous_df[sous_df["Désignation"] == choix_ru].iloc[0]

    st.info(f"**Référence :** {ligne['Réf']}")
    st.write(f"**Modèle choisi :** {choix_ru}")

  # --- AUTRES FAMILLES (FLYERS, DÉPLIANTS, CARTES, ETC.) ---
  else:
    st.subheader(f"📄 {famille}")
    st.info(
        "Utilisez les options ci-dessous pour consulter les déclinaisons de"
        f" cette gamme ({famille})."
    )

    # Sélection de la plage de lignes correspondante selon la famille choisie
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
      sous_df = df.iloc[debut:fin].dropna(how="all").copy()
      # On cherche une colonne de désignation ou texte pour le selectbox
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