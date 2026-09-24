import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Tarifs SAS ABCOM - Impression & Communication",
    page_icon="🖨️",
    layout="wide",
)

# Chargement du fichier Excel
EXCEL_FILE = "tarifs print-panneaux-banderoles.xlsx"


@st.cache_data
def load_data():
  try:
    df = pd.read_excel(EXCEL_FILE, sheet_name="Feuil1")
    return df
  except Exception as e:
    st.error(f"Erreur lors du chargement du fichier Excel : {e}")
    return None


df = load_data()

# Titre principal
st.title("🖨️ SAS ABCOM - Grille Tarifaire")
st.markdown(
    "Consultez l'ensemble de nos tarifs par catégorie de produits (Impression,"
    " Signalétique, Grands Formats, Goodies)."
)

if df is not None:
  # Menu latéral pour filtrer les catégories
  st.sidebar.header("Navigation par Catégorie")
  categorie = st.sidebar.selectbox(
      "Sélectionnez une gamme :",
      [
          "Toutes les catégories",
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
          "Calendriers & Calendriers magnétiques",
      ],
  )

  if categorie == "Toutes les catégories" or categorie == "Blocs-Notes":
    st.header("📝 Blocs-Notes")
    st.markdown(
        "*Papier 90g Offset de qualité, collés en tête et dos en carton.*"
    )
    st.dataframe(df.iloc[2:9].reset_index(drop=True), use_container_width=True)
    st.markdown("---")

  if (
      categorie == "Toutes les catégories"
      or categorie == "Chemises de présentation"
  ):
    st.header("📂 Chemises de présentation")
    st.markdown(
        "*Format 300g avec encoche pour carte de visite, 2 rabats, simple ou"
        " double rainage.*"
    )
    st.dataframe(df.iloc[14:20].reset_index(drop=True), use_container_width=True)
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Banderoles":
    st.header("🚩 Banderoles")
    st.markdown("*Banderole normée M1, 510 g/m2, avec œillet et ourlet.*")
    st.dataframe(df.iloc[28:34].reset_index(drop=True), use_container_width=True)
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Panneaux de chantier":
    st.header("🚧 Panneaux de chantier")
    st.markdown(
        "*Panneaux alvéolaires en polypropylène (Akylux) - Œillets aux 4 coins.*"
    )
    st.dataframe(df.iloc[37:45].reset_index(drop=True), use_container_width=True)
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Roll-Ups":
    st.header("📌 Roll-Ups (Structures enroulables)")
    st.dataframe(df.iloc[51:57].reset_index(drop=True), use_container_width=True)
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Flyers":
    st.header("📄 Flyers")
    st.dataframe(df.iloc[59:84].reset_index(drop=True), use_container_width=True)
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Dépliants":
    st.header("📑 Dépliants")
    st.dataframe(
        df.iloc[86:111].reset_index(drop=True), use_container_width=True
    )
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Menus restaurants":
    st.header("🍽️ Menus restaurants")
    st.markdown("*Papier 300g indéchirable.*")
    st.dataframe(
        df.iloc[113:117].reset_index(drop=True), use_container_width=True
    )
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Sous-bocks":
    st.header("🍺 Sous-bocks")
    st.markdown("*Carton épais 580g.*")
    st.dataframe(
        df.iloc[119:123].reset_index(drop=True), use_container_width=True
    )
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Adhésifs":
    st.header("🏷️ Adhésifs")
    st.dataframe(
        df.iloc[127:137].reset_index(drop=True), use_container_width=True
    )
    st.markdown("---")

  if categorie == "Toutes les catégories" or categorie == "Cartes de visite":
    st.header("📇 Cartes de visite")
    st.dataframe(
        df.iloc[141:150].reset_index(drop=True), use_container_width=True
    )
    st.markdown("---")

  if (
      categorie == "Toutes les catégories"
      or categorie == "Calendriers & Calendriers magnétiques"
  ):
    st.header("📅 Calendriers & Magnétiques")
    st.dataframe(
        df.iloc[151:182].reset_index(drop=True), use_container_width=True
    )
    st.markdown("---")

  # Option pour afficher le tableau brut complet si besoin
  with st.expander("🔍 Afficher le tableau brut complet du fichier Excel"):
    st.dataframe(df, use_container_width=True)