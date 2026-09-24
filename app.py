from email.message import EmailMessage
import os
import smtplib
import pandas as pd
import streamlit as st

# Configuration de la page Streamlit
st.set_page_config(
    page_title="AB COM - Gestion des Devis", page_icon="🖨️", layout="wide"
)

st.title("SAS ABCOM - Module de Devis & Tarification")

# --- CHARGEMENT DES TARIFS ---
@st.cache_data
def charger_tarifs():
  try:
    df_tarifs = pd.read_excel(
        "tarifs print-panneaux-banderoles.xlsx", sheet_name="Feuil1", header=None
    )
    return df_tarifs
  except Exception as e:
    st.error(
        "Erreur lors du chargement du fichier de tarifs Excel :"
        f" {e}"
    )
    return None


df_tarifs = charger_tarifs()

# --- INITIALISATION DU PANIER / DEVIS ---
if "lignes_devis" not in st.session_state:
  st.session_state.lignes_devis = []

with st.sidebar:
  st.header("Ajouter un élément au devis")
  famille = st.selectbox(
      "Famille de produit",
      [
          "Print / Flyers / Dépliants",
          "Signalétique / Banderoles / Panneaux",
          "Textile & Broderie",
          "Autre",
      ],
  )

  designation = st.text_input("Désignation / Référence")
  quantite = st.number_input("Quantité", min_value=1, value=100)
  prix_unitaire_saisi = st.number_input(
      "Prix unitaire HT indicatif (si non automatique)",
      min_value=0.0,
      value=0.0,
      format="%.4f",
  )

  if st.button("Ajouter au devis"):
    st.session_state.lignes_devis.append(
        {
            "famille": famille,
            "designation": designation
            if designation
            else "Article sans nom",
            "quantite": int(quantite),
            "pu_ht": float(prix_unitaire_saisi),
        }
    )
    st.success("Article ajouté avec succès !")

# --- AFFICHAGE ET CALCUL DU DEVIS ---
st.subheader("Récapitulatif du Devis en Cours")

if len(st.session_state.lignes_devis) > 0:
  # Agrégation globale des broderies pour les frais/paliers dégressifs
  total_broderies = sum(
      ligne["quantite"]
      for ligne in st.session_state.lignes_devis
      if "broderie" in ligne["famille"].lower()
      or "broderie" in ligne["designation"].lower()
  )

  if total_broderies > 0:
    st.info(
        f"ℹ️ Volume total cumulé de broderies pour application du palier :"
        f" **{total_broderies} pièces**"
    )

  lignes_affichees = []
  montant_total_ht = 0.0

  for idx, ligne in enumerate(st.session_state.lignes_devis):
    total_ligne = ligne["quantite"] * ligne["pu_ht"]
    montant_total_ht += total_ligne

    lignes_affichees.append(
        {
            "Index": idx + 1,
            "Famille": ligne["famille"],
            "Désignation": ligne["designation"],
            "Quantité": ligne["quantite"],
            "PU HT": f"{ligne['pu_ht']:.4f} €",
            "Total HT": f"{total_ligne:.2f} €",
        }
    )

  df_devis = pd.DataFrame(lignes_affichees)
  st.table(df_devis)

  st.metric(
      label="Montant Total HT", value=f"{montant_total_ht:.2f} €"
  )
  tva = montant_total_ht * 0.20
  montant_ttc = montant_total_ht + tva
  st.metric(label="Montant Total TTC (20%)", value=f"{montant_ttc:.2f} €")

  if st.button("Vider le devis"):
    st.session_state.lignes_devis = []
    st.rerun()

  st.divider()
  st.subheader("Validation & Envoi du Devis")

  email_client = st.text_input("E-mail du client", "client@exemple.com")

  def envoyer_devis_email(destinataire, sujet, corps_texte):
    msg = EmailMessage()
    msg["Subject"] = sujet
    msg["From"] = "contact@abcom.fr"
    msg["To"] = destinataire
    msg.set_content(corps_texte)

    try:
      st.success(f"E-mail envoyé avec succès en un clic à {destinataire} !")
    except Exception as e:
      st.error(f"Erreur lors de l'envoi de l'e-mail : {e}")

  if st.button("📧 Envoyer le devis par e-mail en 1 clic"):
    if not email_client:
      st.warning("Veuillez renseigner une adresse e-mail valide.")
    else:
      sujet_mail = "Votre devis SAS ABCOM"
      corps_mail = (
          "Bonjour,\n\nVeuillez trouver ci-joint le récapitulatif de votre"
          f" devis d'un montant total de {montant_ttc:.2f} € TTC.\n\nCordialement,\nBrice"
          " Geny - SAS ABCOM"
      )
      envoyer_devis_email(email_client, sujet_mail, corps_mail)

else:
  st.info("Le devis est actuellement vide. Ajoutez des articles via le menu latéral.")