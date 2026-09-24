import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="SAS ABCOM - Configurateur de Tarifs", page_icon="🖨️", layout="wide"
)

# Chargement du fichier Excel
EXCEL_FILE = "tarifs print-panneaux-banderoles.xlsx"


@st.cache_data
def load_data():
  try:
    df = pd.read_excel(EXCEL_FILE, sheet_name="Feuil1")
    return df
  except Exception as e:
    st.error(
        f"Erreur lors du chargement du fichier Excel : {e}. Vérifiez que le"
        f" fichier '{EXCEL_FILE}' est bien présent dans le dossier du projet."
    )
    return None


df = load_data()

# Titre principal
st.title("🖨️ SAS ABCOM - Devis & Tarifs Interactifs")
st.markdown(
    "Sélectionnez une catégorie de produits puis affinez votre choix via le"
    " menu déroulant pour consulter les caractéristiques et tarifs."
)

if df is not None:
  # Menu de sélection de la catégorie principale
  categorie = st.selectbox(
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
          "Calendriers & Magnétiques",
      ],
  )

  st.markdown("---")

  # Affichage dynamique selon la catégorie sélectionnée avec menu déroulant de produit
  if categorie == "Blocs-Notes":
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
        "Col8",
        "Col9",
        "Col10",
        "Col11",
    ]
    produit_choisi = st.selectbox(
        "Sélectionnez le bloc-note :", sous_df["Désignation"].tolist()
    )
    ligne = sous_df[sous_df["Désignation"] == produit_choisi].iloc[0]

    col1, col2 = st.columns(2)
    with col1:
      st.info(f"**Référence :** {ligne['Réf']}")
      st.write(f"**Description :** {ligne['Désignation']}")
    with col2:
      st.success(
          "Tarifs unitaires dégressifs disponibles dans la grille complète."
      )
      st.dataframe(
          sous_df[sous_df["Désignation"] == produit_choisi].dropna(
              axis=1, how="all"
          ),
          use_container_width=True,
      )

  elif categorie == "Chemises de présentation":
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
        "Col9",
        "Col10",
        "Col11",
    ]
    produit_choisi = st.selectbox(
        "Sélectionnez le modèle de chemise :", sous_df["Désignation"].tolist()
    )
    ligne = sous_df[sous_df["Désignation"] == produit_choisi].iloc[0]

    st.info(f"**Référence :** {ligne['Réf']}")
    st.dataframe(
        sous_df[sous_df["Désignation"] == produit_choisi].dropna(
            axis=1, how="all"
        ),
        use_container_width=True,
    )

  elif categorie == "Banderoles":
    st.subheader("🚩 Banderoles (M1 510g/m² avec œillets)")
    sous_df = df.iloc[29:34].dropna(how="all").copy()
    sous_df.columns = ["Format / Finition", "Prix HT (€)", "C2", "C3", "C4"]
    format_choisi = st.selectbox(
        "Sélectionnez le format de banderole :",
        sous_df["Format / Finition"].tolist(),
    )
    ligne = sous_df[sous_df["Format / Finition"] == format_choisi].iloc[0]

    st.metric(
        label=f"Tarif pour : {format_choisi}", value=f"{ligne['Prix HT (€)']} €"
    )

  elif categorie == "Panneaux de chantier":
    st.subheader("🚧 Panneaux de chantier (Akylux)")
    sous_df = df.iloc[39:45].dropna(how="all").copy()
    sous_df.columns = ["Réf", "Désignation", "Prix"]
    produit_choisi = st.selectbox(
        "Sélectionnez le panneau :", sous_df["Désignation"].tolist()
    )
    st.dataframe(
        sous_df[sous_df["Désignation"] == produit_choisi],
        use_container_width=True,
    )

  elif categorie == "Roll-Ups":
    st.subheader("📌 Roll-Ups (Structures enroulables)")
    sous_df = df.iloc[53:57].dropna(how="all").copy()
    sous_df.columns = ["Réf", "Désignation"]
    produit_choisi = st.selectbox(
        "Sélectionnez le type de Roll-Up :", sous_df["Désignation"].tolist()
    )
    st.write(f"**Modèle :** {produit_choisi}")

  elif categorie == "Flyers":
    st.subheader("📄 Flyers")
    st.info(
        "Gamme complète de flyers disponible. Sélectionnez un format ou un"
        " grammage :"
    )
    sous_df = df.iloc[59:84].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  elif categorie == "Dépliants":
    st.subheader("📑 Dépliants")
    sous_df = df.iloc[86:111].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  elif categorie == "Menus restaurants":
    st.subheader("🍽️ Menus restaurants indéchirables")
    sous_df = df.iloc[113:117].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  elif categorie == "Sous-bocks":
    st.subheader("🍺 Sous-bocks (Carton 580g)")
    sous_df = df.iloc[119:123].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  elif categorie == "Adhésifs":
    st.subheader("🏷️ Adhésifs")
    sous_df = df.iloc[127:137].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  elif categorie == "Cartes de visite":
    st.subheader("📇 Cartes de visite")
    sous_df = df.iloc[141:150].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)

  elif categorie == "Calendriers & Magnétiques":
    st.subheader("📅 Calendriers & Magnétiques")
    sous_df = df.iloc[151:182].dropna(how="all").copy()
    st.dataframe(sous_df.dropna(axis=1, how="all"), use_container_width=True)