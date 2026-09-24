import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

st.set_page_config(page_title="Gestionnaire de Devis - Multi-Métiers", layout="wide")

# --- GESTION DES FICHIERS & SECRETS ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"
CATALOGUE_FILE = "catalogue_standardise.xlsx"
PDF_DIR = "devis_pdf"

if not os.path.exists(PDF_DIR):
    os.makedirs(PDF_DIR)

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

def charger_catalogue():
    if os.path.exists(CATALOGUE_FILE):
        try:
            return pd.read_excel(CATALOGUE_FILE)
        except Exception as e:
            st.error(f"Erreur catalogue : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

df_catalogue = charger_catalogue()

def obtenir_prix_catalogue(df, designation, qte):
    """Recherche exacte du tarif unitaire dans le catalogue selon le palier de quantité."""
    if df.empty or 'Designation' not in df.columns or 'Prix_HT' not in df.columns:
        return 0.0
    
    df_art = df[df['Designation'] == designation]
    if df_art.empty:
        return 0.0
        
    if 'Quantite' in df_art.columns:
        paliers = sorted(df_art['Quantite'].dropna().unique().tolist())
        if paliers:
            palier_choisi = paliers[0]
            for p in paliers:
                if p <= qte:
                    palier_choisi = p
                else:
                    break
            match_row = df_art[df_art['Quantite'] == palier_choisi]
            if not match_row.empty:
                return float(match_row.iloc[0]['Prix_HT'])
                
    return float(df_art.iloc[0]['Prix_HT'])

# --- CRM INITIALISATION ---
if not os.path.exists(CRM_FILE):
    df_init = pd.DataFrame(columns=[
        "Numero_Devis", "Date", "Client", "Contact", "Email", "Telephone", 
        "Total_HT", "Total_TTC", "Statut", "Commercial", "Mode_Reglement", "PDF_Path"
    ])
    df_init.to_csv(CRM_FILE, index=False)

def enregistrer_devis_crm(devis_data):
    df_crm = pd.read_csv(CRM_FILE) if os.path.exists(CRM_FILE) else pd.DataFrame(columns=list(devis_data.keys()))
    if not df_crm[df_crm["Numero_Devis"] == devis_data["Numero_Devis"]].empty:
        df_crm.loc[df_crm["Numero_Devis"] == devis_data["Numero_Devis"], :] = list(devis_data.values())
    else:
        df_crm = pd.concat([df_crm, pd.DataFrame([devis_data])], ignore_index=True)
    df_crm.to_csv(CRM_FILE, index=False)

# --- GENERATEUR DE PDF ---
def generer_pdf_devis(numero_devis, client_infos, lignes, totaux, commercial_infos):
    filename = os.path.join(PDF_DIR, f"Devis_{numero_devis.replace('/', '_')}.pdf")
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph(f"<b>DEVIS N° {numero_devis}</b>", styles['Heading1']))
    elements.append(Spacer(1, 10))
    
    texte_client = f"<b>Client :</b> {client_infos['nom']}<br/><b>Contact :</b> {client_infos['contact']}<br/><b>Adresse :</b> {client_infos['adresse']}<br/><b>Email :</b> {client_infos['email']}"
    texte_comm = f"<b>Date :</b> {datetime.now().strftime('%d/%m/%Y')}<br/><b>Commercial :</b> {commercial_infos['nom']}<br/><b>Règlement :</b> {commercial_infos['reglement']}"
    
    table_infos = Table([[Paragraph(texte_client, styles['Normal']), Paragraph(texte_comm, styles['Normal'])]], colWidths=[250, 250])
    elements.append(table_infos)
    elements.append(Spacer(1, 20))

    data_table = [["Désignation", "Qté", "P.U. HT (€)", "Total HT (€)"]]
    for l in lignes:
        data_table.append([l['nom'], str(l['qte']), f"{l['pu']:.2f}", f"{l['total']:.2f}"])
    
    t = Table(data_table, colWidths=[230, 60, 100, 110])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 15))

    texte_totaux = f"""
    <b>Sous-Total Articles HT :</b> {totaux['sous_total']:.2f} €<br/>
    <b>Frais techniques HT :</b> {totaux['frais_techniques']:.2f} €<br/>
    <b>Frais de port HT :</b> {totaux['port']:.2f} €<br/>
    <b>Total HT :</b> {totaux['total_ht']:.2f} €<br/>
    <b>TVA (20%) :</b> {totaux['tva']:.2f} €<br/>
    <b>TOTAL TTC :</b> {totaux['total_ttc']:.2f} €
    """
    elements.append(Table([[Paragraph("", styles['Normal']), Paragraph(texte_totaux, styles['Normal'])]], colWidths=[250, 250]))
    
    doc.build(elements)
    return filename

# --- ENVOI MAIL SMTP (Via Secrets.toml transparent) ---
def envoyer_email_smtp(destinataire, sujet, corps, pdf_path):
    try:
        smtp_user = st.secrets["smtp"]["email"]
        smtp_password = st.secrets["smtp"]["password"]
    except Exception as e:
        return False, fnu"Erreur de lecture des secrets Streamlit (secrets.toml) : {e}"

    expediteur = smtp_user
    bcc = "Brice.geny@gmail.com"
    
    msg = MIMEMultipart()
    msg['From'] = expediteur
    msg['To'] = destinataire
    msg['Bcc'] = bcc
    msg['Subject'] = sujet

    msg.attach(MIMEText(corps, 'plain'))

    with open(pdf_path, "rb") as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(pdf_path)}")
    msg.attach(part)

    destinataires_finaux = [destinataire, bcc]
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(expediteur, destinataires_finaux, msg.as_string())
        server.quit()
        return True, "E-mail envoyé avec succès avec copie à Brice.geny@gmail.com !"
    except Exception as e:
        return False, f"Erreur lors de l'envoi SMTP : {e}"

# --- SIDEBAR ---
st.sidebar.title("📋 Infos Client & Expédition")
client_nom = st.sidebar.text_input("Nom du Client / Société", "Client Exemple")
client_contact = st.sidebar.text_input("Nom de contact")
client_adresse = st.sidebar.text_area("Adresse complète du Client")
client_email = st.sidebar.text_input("Email du Client", "client@exemple.com")
client_tel = st.sidebar.text_input("Téléphone du Client")

st.sidebar.markdown("---")
st.sidebar.subheader("🚚 Logistique & Commercial")
zone_livraison = st.sidebar.selectbox("Zone de Livraison", ["France Continentale", "Corse, Monaco ou Andorre", "Espace UE"])

conseiller_choix = st.sidebar.selectbox("Commercial / Conseiller", ["Brice Geny", "Brice Bugna"])
mode_reglement = st.sidebar.selectbox("Mode de Règlement", ["Virement bancaire 30 jours", "Comptant à la commande", "50% à la commande, 50% à 30 jours"])
delai_validite = st.sidebar.selectbox("Validité du Devis", ["30 jours", "15 jours", "45 jours"])

# --- INTERFACE PRINCIPALE ---
noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"]
onglets = st.tabs(noms_onglets)

articles_saisis = []

for i in range(10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        metier_type = st.radio(
            f"Univers / Métier pour l'Article {i+1}",
            ["👕 Textile & Marquage (DTF / Broderie)", "📄 Print, Papeterie & Signalétique (Catalogue Excel)"],
            key=f"metier_{i}"
        )
        st.markdown("---")
        
        if "Textile" in metier_type:
            sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", key=f"sans_marq_{i}")
            col1, col2 = st.columns(2)
            with col1:
                nom_article = st.text_input(f"Référence / Nom du vêtement {i+1}", value="T-Shirt 100% coton bio" if i==0 else f"Vêtement {i+1}", key=f"nom_textile_{i}")
                qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_textile_{i}")
                prix_vetement_ht = st.number_input(f"Prix unitaire HT du support textile (€) {i+1}", min_value=0.0, value=4.92, format="%.2f", key=f"px_textile_{i}")
            with col2:
                nb_marquages = st.selectbox(f"Nombre de marquages {i+1}", [1, 2, 3, 4], key=f"nb_m_textile_{i}") if not sans_marquage else 0

            marquages_config = []
            has_broderie = False
            if not sans_marquage:
                for m in range(nb_marquages):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
                    with mc2:
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur", "Dos", "Manche", "Poitrine"], key=f"emp_{i}_{m}")
                    if "Broderie" in t_marq:
                        has_broderie = True
                    marquages_config.append({"technique": t_marq, "emplacement": emp})

            remise_fidelite = st.number_input(f"Remise fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "textile",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": sans_marquage,
                    "marquages_config": marquages_config,
                    "has_broderie": has_broderie,
                    "remise_fidelite": remise_fidelite
                })
        else:
            col1, col2 = st.columns(2)
            with col1:
                qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=100 if i==0 else 0, key=f"qte_print_{i}")
                
                # Récupération automatique du prix unitaire depuis le catalogue selon la quantité
                if not df_catalogue.empty:
                    categories_dispo = df_catalogue['Categorie'].dropna().unique().tolist() if 'Categorie' in df_catalogue.columns else []
                    cat_choisie = st.selectbox(f"Catégorie Print {i+1}", categories_dispo, key=f"cat_print_{i}")
                    
                    df_filtre = df_catalogue[df_catalogue['Categorie'] == cat_choisie]
                    liste_designations = df_filtre['Designation'].dropna().unique().tolist() if 'Designation' in df_filtre.columns else []
                    nom_article = st.selectbox(f"Article du catalogue {i+1}", liste_designations, key=f"nom_print_{i}")
                    
                    prix_defaut = obtenir_prix_catalogue(df_catalogue, nom_article, qte)
                else:
                    nom_article = st.text_input(f"Désignation article print {i+1}", value="Flyer A5", key=f"nom_print_libre_{i}")
                    prix_defaut = 0.10

                prix_vetement_ht = st.number_input(f"Prix unitaire HT (€) {i+1}", min_value=0.0, value=float(prix_defaut), format="%.3f", key=f"px_print_{i}")
            with col2:
                st.info("ℹ️ Prix unitaire calculé automatiquement depuis la grille de tarifs par quantité du catalogue.")
                remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_print_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "print",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": True,
                    "marquages_config": [],
                    "has_broderie": False,
                    "remise_fidelite": remise_fidelite
                })

# --- CALCULS FINANCIERS ---
lignes_devis_global = []
total_textile_brut = 0.0

for art in articles_saisis:
    q = art["quantite"]
    px_base = art["prix_vet_unit"]
    px_support_remise = px_base * (1 - art["remise_fidelite"] / 100.0)
    total_support = q * px_support_remise
    
    total_marquages_ligne = 0.0
    if art["type_univers"] == "textile" and not art["sans_marquage"]:
        for m_cfg in art["marquages_config"]:
            unit_m = 2.50 if q < 50 else (1.80 if q < 200 else 1.20)
            total_marquages_ligne += q * unit_m

    total_ligne_ht = total_support + total_marquages_ligne
    
    if art["type_univers"] == "textile":
        total_textile_brut += total_ligne_ht

    lignes_devis_global.append({
        "nom": art["nom_article"],
        "qte": q,
        "pu": px_support_remise,
        "total": total_ligne_ht,
        "type_univers": art["type_univers"]
    })

# --- ONGLET : GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.header("📊 Récapitulatif Général & Validation du Devis")
    
    if not lignes_devis_global:
        st.warning("⚠️ Veuillez renseigner au moins un article avec une quantité > 0.")
    else:
        sous_total_articles = sum([item["total"] for item in lignes_devis_global])
        
        # Frais techniques textile automatiques : Commande textile < 800 € HT = 24.90 €
        frais_techniques_dossier = 24.90 if (total_textile_brut > 0 and total_textile_brut < 800.0) else 0.0

        # Calcul automatique des frais de port selon la grille officielle
        # Détermination de la base de calcul pour le port (Sous-total + frais techniques)
        montant_base_port = sous_total_articles + frais_techniques_dossier

        if zone_livraison == "France Continentale":
            if montant_base_port < 99.99: port_auto = 14.95
            elif montant_base_port < 499.99: port_auto = 20.95
            elif montant_base_port < 999.99: port_auto = 24.95
            else: port_auto = 0.0 # Franco de port
        elif zone_livraison == "Livraison Corse, Monaco ou Andorre":
            if montant_base_port < 99.99: port_auto = 19.95
            elif montant_base_port < 499.99: port_auto = 25.95
            elif montant_base_port < 999.99: port_auto = 29.95
            else: port_auto = 0.0 # Franco de port
        else: # Espace UE
            if montant_base_port < 99.99: port_auto = 25.95
            elif montant_base_port < 499.99: port_auto = 39.0
            elif montant_base_port < 999.99: port_auto = 60.0
            elif montant_base_port < 1500.0: port_auto = 90.0
            else: port_auto = 0.0

        st.markdown("---")
        col_port1, col_port2 = st.columns(2)
        with col_port1:
            annuler_port = st.checkbox("🎁 Offrir les frais de port (Franco total)", value=False)
        with col_port2:
            frais_port = 0.0 if annuler_port else port_auto

        total_ht = sous_total_articles + frais_techniques_dossier + frais_port
        tva = total_ht * 0.20
        total_ttc = total_ht + tva

        numero_devis_genere = obtenir_prochain_numero_devis()

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown(f"**Client :** {client_nom}")
            st.markdown(f"**Contact :** {client_contact}")
            st.markdown(f"**Adresse :** {client_adresse}")
        with col_b:
            st.markdown(f"**N° de Devis :** `{numero_devis_genere}`")
            st.markdown(f"**Date :** {datetime.now().strftime('%d/%m/%Y')}")
            st.markdown(f"**Commercial :** {conseiller_choix}")

        st.markdown("---")
        st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
        if frais_techniques_dossier > 0:
            st.write(f"**Frais techniques textile (Commande < 800€ HT) :** {frais_techniques_dossier:.2f} €")
        st.write(f"**Frais de port ({zone_livraison}) HT :** {frais_port:.2f} €")
        st.markdown(f"### **Total HT : {total_ht:.2f} €**")
        st.markdown(f"### **Total TTC (20%) : {total_ttc:.2f} €**")

        st.markdown("---")
        st.subheader("📤 Validation, Enregistrement PDF & Envoi Mail Automatique")

        if st.button("💾 Générer le PDF, enregistrer dans le CRM & Envoyer l'e-mail automatique"):
            client_infos = {"nom": client_nom, "contact": client_contact, "adresse": client_adresse, "email": client_email}
            totaux_dict = {"sous_total": sous_total_articles, "frais_techniques": frais_techniques_dossier, "port": frais_port, "total_ht": total_ht, "tva": tva, "total_ttc": total_ttc}
            commercial_infos = {"nom": conseiller_choix, "reglement": mode_reglement}
            
            pdf_path = generer_pdf_devis(numero_devis_genere, client_infos, lignes_devis_global, totaux_dict, commercial_infos)
            
            devis_record = {
                "Numero_Devis": numero_devis_genere,
                "Date": datetime.now().strftime('%Y-%m-%d'),
                "Client": client_nom,
                "Contact": client_contact,
                "Email": client_email,
                "Telephone": client_tel,
                "Total_HT": round(total_ht, 2),
                "Total_TTC": round(total_ttc, 2),
                "Statut": "En attente",
                "Commercial": conseiller_choix,
                "Mode_Reglement": mode_reglement,
                "PDF_Path": pdf_path
            }
            enregistrer_devis_crm(devis_record)
            st.success(f"✨ PDF généré avec succès ({pdf_path}) et enregistré dans le CRM !")

            # Envoi automatique via secrets.toml sans demande de code
            sujet = f"Votre devis n° {numero_devis_genere}"
            corps = f"""Bonjour {client_contact or client_nom},

Veuillez trouver ci-joint votre devis n° {numero_devis_genere} d'un montant total de {total_ttc:.2f} € TTC.

Conditions de règlement : {mode_reglement}
Validité de l'offre : {delai_validite}

Restant à votre disposition pour tout renseignement complémentaire.

Bien cordialement,
{conseiller_choix}
"""
            succes, message = envoyer_email_smtp(client_email, sujet, corps, pdf_path)
            if succes:
                st.success(message)
            else:
                st.error(message)

            with open(pdf_path, "rb") as f:
                st.download_button("📥 Télécharger le PDF du devis", f, file_name=os.path.basename(pdf_path), mime="application/pdf")

# --- ONGLET : SUIVI CRM ---
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
                st.info("Aucun fichier PDF enregistré pour ce devis.")
        else:
            st.info("Aucun devis dans le CRM.")
    else:
        st.info("CRM vide.")