import os
import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="SAS ABCOM - Configurateur de Tarifs", page_icon="🖨️", layout="wide"
)

# Chemin complet de votre fichier Excel
EXCEL_FILE = r"C:\Users\setup\OneDrive\OneDrive - IRIS - AB Com\Bureau\reprise\site et appli\tarifs print-panneaux-banderoles.xlsx"


@st.cache_data
def load_data():
  try:
    if not os.path.exists(EXCEL_FILE):
      st.error(f"Fichier introuvable au chemin : {EXCEL_FILE}")
      return None
    df = pd.read_excel(EXCEL_FILE, sheet_name="Feuil1")
    return df
  except Exception as e:
    st.error(f"Erreur lors du chargement du fichier Excel : {e}")
    return None


df = load_data()

# Titre principal
st.title("🖨️ SAS ABCOM - Devis & Tarifs Interactifs")
st.markdown(
    "Naviguez entre les onglets ci-dessous et utilisez les menus déroulants pour"
    " sélectionner vos produits."
)

if df is not None:
  # Création des onglets principaux (onglets à la place d'une liste latérale)
  onglets = st.tabs([
      "Blocs-Notes",
      "Chemises",
      "Banderoles",
      "Panneaux",
      "Roll-Ups",
      "Flyers & Dépliants",
      "Menus & Sous-bocks",
      "Adhésifs",
      "Cartes & Calendriers",
  ])

  # --- ONGLET 1 : BLOCS-NOTES ---
  with onglets[0]:
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
    produit_choisi = st.selectbox(
        "Sélectionnez votre bloc-note :",
        sous_df["Désignation"].tolist(),
        key="sel_bloc",
    )
    ligne = sous_df[sous_df["Désignation"] == produit_choisi].iloc[0]

    st.info(f"**Référence :** {ligne['Réf']} | **Produit :** {ligne['Désignation']}")
    st.write("Grille tarifaire détaillée pour ce produit :")
    st.dataframe(
        sous_df[sous_df["Désignation"] == produit_choisi].dropna(
            axis=1, how="all"
        ),
        use_container_width=True,
    )

  # --- ONGLET 2 : CHEMISES ---
  with onglets[1]:
    st.subheader("📂 Chemises de présentation")
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
    produit_choisi = st.selectbox(
        "Sélectionnez votre modèle de chemise :",
        sous_df["Désignation"].tolist(),
        key="sel_chemise",
    )
    ligne = sous_df[sous_df["Désignation"] == produit_choisi].iloc[0]

    st.info(f"**Référence :** {ligne['Réf']}")
    st.dataframe(
        sous_df[sous_df["Désignation"] == produit_choisi].dropna(
            axis=1, how="all"
        ),
        use_container_width=True,
    )

  # --- ONGLET 3 : BANDEROLES ---
  with onglets[2]:
    st.subheader("🚩 Banderoles (M1 510g/m²)")
    sous_df = df.iloc[29:34].dropna(how="all").copy()
    sous_df.columns = ["Format / Finition", "Prix HT (€)", "C2", "C3", "C4"]
    format_choisi = st.selectbox(
        "Sélectionnez le format de banderole :",
        sous_df["Format / Finition"].tolist(),
        key="sel_banderole",
    )
    ligne = sous_df[sous_df["Format / Finition"] == format_choisi].iloc[0]

    st.metric(
        label=f"Tarif - {format_choisi}", value=f"{ligne['Prix HT (€, d)'] } €"
        if "Prix HT (€, d)" in ligne
        else f"{ligne['Prix HT (€)']} €"
    )

  # --- ONGLET 4 : PANNEAUX ---
  with onglets[3]:
    st.subheader("🚧 Panneaux de chantier (Akylux)")
    sous_df = df.iloc[39:45].dropna(how="all").copy()
    sous_df.columns = ["Réf", "Désignation", "Prix"]
    produit_choisi = st.selectbox(
        "Sélectionnez votre panneau :",
        sous_df["Désignation"].tolist(),
        key="sel_panneau",
    )
    st.dataframe(
        sous_df[sous_df["Désignation"] == produit_choisi],
        use_container_width=True,
    )

  # --- ONGLET 5 : ROLL-UPS ---
  with onglets[4]:
    st.subheader("📌 Roll-Ups (Structures enroulables)")
    sous_df = df.iloc[53:57].dropna(how="all").copy()
    sous_df.columns = ["Réf", "Désignation"]
    produit_choisi = st.selectbox(
        "Sélectionnez votre Roll-Up :",
        sous_df["Désignation"].tolist(),
        key="sel_rollup",
    )
    ligne = sous_df[sous_df["Désignation"] == produit_choisi].iloc[0]
    st.success(f"**Référence :** {ligne['Réf']} — **Modèle :** {produit_choisi}")

  # --- ONGLET 6 : FLYERS & DÉPLIANTS ---
  with onglets[5]:
    st.subheader("📄 Flyers & Dépliants")
    type_choix = st.radio(
        "Choisissez la gamme :", ["Flyers", "Dépliants"], horizontal=True
    )
    if type_choix == "Flyers":
      sous_df = df.iloc[59:84].dropna(how="all").copy()
    else:
      sous_df = df.iloc[86:111].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  # --- ONGLET 7 : MENUS & SOUS-BOCKS ---
  with onglets[6]:
    st.subheader("🍽️ Menus restaurants & Sous-bocks")
    produit_type = st.selectbox(
        "Sélectionnez le produit :",
        ["Menus restaurants (300g indéchirable)", "Sous-bocks (Carton 580g)"],
        key="sel_menu_bock",
    )
    if "Menus" in produit_type:
      sous_df = df.iloc[113:117].dropna(how="all").copy()
    else:
      sous_df = df.iloc[119:123].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  # --- ONGLET 8 : ADHÉSIFS ---
  with onglets[7]:
    st.subheader("🏷️ Adhésifs")
    sous_df = df.iloc[127:137].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  # --- ONGLET 9 : CARTES & CALENDRIERS ---
  with onglets[8]:
    st.subheader("📇 Cartes de visite & Calendriers")
    choix_cc = st.selectbox(
        "Sélectionnez la catégorie :",
        ["Cartes de visite", "Calendriers & Magnétiques"],
        key="sel_cc",
    )
    if "Cartes" in choix_cc:
      sous_df = df.iloc[141:150].dropna(how="all").copy()
    else:
      sous_df = df.iloc[151:182].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)