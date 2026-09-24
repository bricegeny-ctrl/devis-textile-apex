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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import toml

st.set_page_config(page_title="Gestionnaire de Devis - Multi-Métiers", layout="wide")

# --- GESTION DES FICHIERS & SECRETS ---
COMPTEUR_FILE = "compteur_devis.json"
CRM_FILE = "crm_devis.csv"
CATALOGUE_FILE = "catalogue_standardise.xlsx"
PDF_DIR = "devis_pdf"

if not os.path.exists(PDF_DIR):
    os.makedirs(PDF_DIR)

def charger_secrets_smtp():
    # 1. Tentative via st.secrets natif de Streamlit Cloud
    try:
        if "smtp" in st.secrets:
            return st.secrets["smtp"]["email"], st.secrets["smtp"]["password"]
    except Exception:
        pass
    
    # 2. Tentative via le chemin local spécifique
    chemin_local = r"C:\Users\setup\OneDrive\OneDrive - IRIS - AB Com\Bureau\reprise\site et appli\.streamlit\secrets.toml"
    if os.path.exists(chemin_local):
        try:
            data = toml.load(chemin_local)
            if "smtp" in data:
                return data["smtp"]["email"], data["smtp"]["password"]
        except Exception:
            pass
            
    # 3. Tentative dans le répertoire courant
    for chemin_relatif in [".streamlit/secrets.toml", "secrets.toml"]:
        if os.path.exists(chemin_relatif):
            try:
                data = toml.load(chemin_relatif)
                if "smtp" in data:
                    return data["smtp"]["email"], data["smtp"]["password"]
            except Exception:
                pass
                
    return None, None

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
    with open(COMPTEUR_FILE, "w") as f:
        json.dump({"dernier_num": nouveau_num}, f)
    return nouveau_num

def charger_catalogue():
    if os.path.exists(CATALOGUE_FILE):
        try:
            return pd.read_excel(CATALOGUE_FILE)
        except Exception as e:
            return pd.DataFrame()
    return pd.DataFrame()

df_catalogue = charger_catalogue()

def obtenir_prix_catalogue(df, designation, qte):
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

def calculer_frais_port(montant_base, zone):
    if zone == "France Continentale":
        if montant_base < 99.99: return 14.95
        elif montant_base < 499.99: return 20.95
        elif montant_base < 999.99: return 24.95
        else: return 0.0
    elif zone == "Livraison Corse, Monaco ou Andorre":
        if montant_base < 99.99: return 19.95
        elif montant_base < 499.99: return 25.95
        elif montant_base < 999.99: return 29.95
        else: return 0.0
    else:
        if montant_base < 99.99: return 25.95
        elif montant_base < 499.99: return 39.0
        elif montant_base < 999.99: return 60.0
        elif montant_base < 1500.0: return 90.0
        else: return 0.0

if not os.path.exists(CRM_FILE):
    df_init = pd.DataFrame(columns=[
        "Numero_Devis", "Date", "Client", "Contact", "Email", "Telephone", 
        "Total_HT", "Total_TTC", "Statut", "Commercial", "Mode_Reglement", "PDF_Path"
    ])
    df_init.to_csv(CRM_FILE, index=False)

def enregistrer_dans_crm(devis_data):
    df_crm = pd.read_csv(CRM_FILE) if os.path.exists(CRM_FILE) else pd.DataFrame(columns=list(devis_data.keys()))
    if not df_crm[df_crm["Numero_Devis"] == devis_data["Numero_Devis"]].empty:
        df_crm.loc[df_crm["Numero_Devis"] == devis_data["Numero_Devis"], :] = list(devis_data.values())
    else:
        df_crm = pd.concat([df_crm, pd.DataFrame([devis_data])], ignore_index=True)
    df_crm.to_csv(CRM_FILE, index=False)

def envoyer_email_smtp(destinataire, sujet, corps, pdf_path):
    smtp_user, smtp_password = charger_secrets_smtp()
    if not smtp_user or not smtp_password:
        return False, "⚠️ E-mail non envoyé : Identifiants SMTP non configurés dans secrets.toml (génération du devis PDF réussie)."

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

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(expediteur, [destinataire, bcc], msg.as_string())
        server.quit()
        return True, "✅ E-mail envoyé avec succès avec copie à Brice.geny@gmail.com !"
    except Exception as e:
        return False, f"Erreur SMTP : {e}"

# --- SIDEBAR ---
st.sidebar.title("📋 Infos Client & Expédition")
client_nom = st.sidebar.text_input("Nom du Client", "Client Exemple")
client_entreprise = st.sidebar.text_input("Société / Entreprise", "")
client_siret = st.sidebar.text_input("SIRET", "")
client_contact = st.sidebar.text_input("Nom de contact", "")
client_adresse = st.sidebar.text_area("Adresse complète", "1 rue de l'Exemple\n70000 Vesoul")
client_email = st.sidebar.text_input("Email", "client@exemple.com")
client_contact_tel = st.sidebar.text_input("Téléphone", "0600000000")

logo_file = st.sidebar.file_uploader("Logo entreprise (PNG/JPG)", type=["png", "jpg", "jpeg"])
logo_defaut_github = "logo.png"

st.sidebar.markdown("---")
zone_livraison = st.sidebar.selectbox("Zone de Livraison", ["France Continentale", "Livraison Corse, Monaco ou Andorre", "Espace UE"])
offrir_port = st.sidebar.checkbox("🎁 Offrir les frais de port", value=False)
conseiller_nom = st.sidebar.selectbox("Commercial / Conseiller", ["Brice Geny", "Brice Bugna"])
mode_reglement = st.sidebar.selectbox("Mode de Règlement", ["Virement bancaire 30 jours", "Comptant à la commande", "50% à la commande, 50% à 30 jours"])

# --- INTERFACE PRINCIPALE ---
noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"]
onglets = st.tabs(noms_onglets)

articles_saisis = []
total_textile_brut = 0.0

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
                prix_vetement_ht = st.number_input(f"Prix unitaire HT support (€) {i+1}", min_value=0.0, value=4.92, format="%.2f", key=f"px_textile_{i}")
            with col2:
                nb_marquages = st.selectbox(f"Nombre de marquages {i+1}", [1, 2, 3, 4], key=f"nb_m_textile_{i}") if not sans_marquage else 0

            marquages = []
            if not sans_marquage:
                for m in range(nb_marquages):
                    mc1, mc2 = st.columns(2)
                    with mc1:
                        t_marq = st.selectbox(f"Technique M{m+1}", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
                    with mc2:
                        emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur", "Dos", "Manche", "Poitrine"], key=f"emp_{i}_{m}")
                    
                    # Tarif dynamique selon technique et quantité
                    if "Broderie" in t_marq:
                        tarif_m = 6.50 if qte < 20 else (5.00 if qte < 50 else 3.80)
                    else:
                        tarif_m = 2.50 if qte < 50 else (1.80 if qte < 200 else 1.20)
                        
                    marquages.append({"nom": f"{t_marq} ({emp})", "tarif": tarif_m})

            frais_prog_broderie = 35.0 if any("Broderie" in m["nom"] for m in marquages) else 0.0
            option_ensachage = st.checkbox(f"Option ensachage individuel {i+1}", key=f"ens_{i}")
            coût_ensachage_unit = 0.35 if option_ensachage else 0.0
            type_sachet = "Sachet individuel transparent" if option_ensachage else ""
            
            option_assurance = st.checkbox(f"Option assurance garantie textile {i+1}", key=f"ass_{i}")
            coût_assurance_unit = 0.15 if option_assurance else 0.0
            
            option_stockage = st.checkbox(f"Option stockage + picking {i+1}", key=f"stock_{i}")
            coût_stockage_unit = 0.50 if option_stockage else 0.0
            
            remise_fidelite = st.number_input(f"Réduction fidélité (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "textile",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": sans_marquage,
                    "marquages": marquages,
                    "frais_prog_broderie": frais_prog_broderie,
                    "option_ensachage": option_ensachage,
                    "coût_ensachage_unit": coût_ensachage_unit,
                    "type_sachet": type_sachet,
                    "option_assurance": option_assurance,
                    "coût_assurance_unit": coût_assurance_unit,
                    "option_stockage": option_stockage,
                    "coût_stockage_unit": coût_stockage_unit,
                    "remise_fidelite": remise_fidelite
                })
                total_textile_brut += qte * prix_vetement_ht
        else:
            col1, col2 = st.columns(2)
            with col1:
                qte = st.number_input(f"Quantité (exemplaires) {i+1}", min_value=0, value=100 if i==0 else 0, key=f"qte_print_{i}")
                
                if not df_catalogue.empty:
                    categories_dispo = df_catalogue['Categorie'].dropna().unique().tolist() if 'Categorie' in df_catalogue.columns else []
                    cat_choisie = st.selectbox(f"Catégorie Print {i+1}", categories_dispo, key=f"cat_print_{i}")
                    
                    df_filtre = df_catalogue[df_catalogue['Categorie'] == cat_choisie]
                    liste_designations = df_filtre['Designation'].dropna().unique().tolist() if 'Designation' in df_filtre.columns else []
                    nom_article = st.selectbox(f"Article du catalogue {i+1}", liste_designations, key=f"nom_print_{i}")
                    
                    # Récupération automatique du prix selon la catégorie, désignation et quantité
                    prix_defaut = obtenir_prix_catalogue(df_filtre, nom_article, qte)
                else:
                    nom_article = st.text_input(f"Désignation article print {i+1}", value="Flyer A5", key=f"nom_print_libre_{i}")
                    prix_defaut = 0.10

                prix_vetement_ht = st.number_input(f"Prix unitaire HT (€) {i+1}", min_value=0.0, value=float(prix_defaut), format="%.3f", key=f"px_print_{i}")
            with col2:
                st.info("ℹ️ Prix unitaire calculé dynamiquement depuis la grille de tarifs du catalogue Excel par catégorie.")
                remise_fidelite = st.number_input(f"Remise commerciale (%) {i+1}", min_value=0.0, max_value=100.0, value=0.0, key=f"rem_print_{i}")

            if qte > 0:
                articles_saisis.append({
                    "type_univers": "print",
                    "nom_article": nom_article,
                    "quantite": qte,
                    "prix_vet_unit": prix_vetement_ht,
                    "sans_marquage": True,
                    "marquages": [],
                    "frais_prog_broderie": 0.0,
                    "option_ensachage": False,
                    "coût_ensachage_unit": 0.0,
                    "type_sachet": "",
                    "option_assurance": False,
                    "coût_assurance_unit": 0.0,
                    "option_stockage": False,
                    "coût_stockage_unit": 0.0,
                    "remise_fidelite": remise_fidelite
                })

# --- CALCUL DES FRAIS TECHNIQUES ---
frais_tech_auto = 24.90 if (total_textile_brut > 0 and total_textile_brut < 800.0) else 0.0

# --- ONGLET 10 : GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.subheader("📊 Récapitulatif Général & Génération du Devis Professionnel")

    if not articles_saisis:
        st.warning("Veuillez renseigner au moins un article avec une quantité supérieure à 0.")
    else:
        supprimer_frais_tech = st.checkbox("⚙️ Supprimer / Offrir les frais techniques de dossier", value=False)
        frais_techniques_dossier = 0.0 if supprimer_frais_tech else frais_tech_auto

        lignes_devis_global = []
        for item in articles_saisis:
            q = item["quantite"]
            px_support = item["prix_vet_unit"] * (1 - item["remise_fidelite"] / 100.0)
            tot_support = q * px_support
            tot_marquages = sum([m["tarif"] * q for m in item["marquages"]])
            tot_ens = (item["coût_ensachage_unit"] * q) if item["option_ensachage"] else 0
            tot_ass = (item["coût_assurance_unit"] * q) if item["option_assurance"] else 0
            tot_stock = (item["coût_stockage_unit"] * q) if item["option_stockage"] else 0
            
            tot_ligne = tot_support + tot_marquages + item["frais_prog_broderie"] + tot_ens + tot_ass + tot_stock
            
            lignes_devis_global.append({
                **item,
                "prix_vet_unit": px_support,
                "total_ligne_ht": tot_ligne
            })

        sous_total_articles = sum([item["total_ligne_ht"] for item in lignes_devis_global])
        montant_base_port = sous_total_articles + frais_techniques_dossier
        
        if offrir_port:
            frais_port = 0.0
        else:
            frais_port = calculer_frais_port(montant_base_port, zone_livraison)
        
        total_general_ht = montant_base_port + frais_port
        tva = total_general_ht * 0.20
        total_ttc = total_general_ht + tva
        
        quantite_globale_totale = sum([item["quantite"] for item in lignes_devis_global])
        cout_unitaire_moyen = total_general_ht / quantite_globale_totale if quantite_globale_totale > 0 else 0

        st.write(f"**Quantité globale pièces :** {quantite_globale_totale}")
        st.write(f"**Sous-Total Articles HT :** {sous_total_articles:.2f} €")
        st.write(f"**Frais techniques :** {frais_techniques_dossier:.2f} € HT")
        st.write(f"**Frais de port ({zone_livraison}) :** {frais_port:.2f} € HT" if not offrir_port else "**Frais de port :** Offerts (0.00 €)")
        st.markdown(f"### **Total Général HT : {total_general_ht:.2f} €** | **TOTAL TTC (20%) : {total_ttc:.2f} €**")

        if st.button("📄 Générer le numéro de devis et le PDF"):
            num_devis = obtenir_prochain_numero_devis()
            pdf_filename = os.path.join(PDF_DIR, f"Devis_{num_devis.replace('/', '_')}.pdf")
            st.session_state['dernier_pdf'] = pdf_filename
            st.session_state['dernier_num'] = num_devis
            st.session_state['total_ttc_cache'] = total_ttc
            st.session_state['total_ht_cache'] = total_general_ht

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
                "PDF_Path": pdf_filename
            }
            enregistrer_dans_crm(data_crm)
            st.success("✅ Données enregistrées dans le CRM avec succès !")

            # --- GÉNÉRATION DU PDF PROFESSIONNEL ---
            doc = SimpleDocTemplate(pdf_filename, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            story = []
            styles = getSampleStyleSheet()

            style_sub = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#4a5568'))
            style_cell = ParagraphStyle('Cell', parent=styles['Normal'], fontSize=9, leading=11)
            style_cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold')
            style_right_bold = ParagraphStyle('RightBold', parent=styles['Normal'], fontSize=9, leading=11, fontName='Helvetica-Bold', alignment=2)
            style_right_normal = ParagraphStyle('RightNormal', parent=styles['Normal'], fontSize=9, leading=11, alignment=2)

            logo_path = None
            if logo_file is not None:
                logo_path = "temp_logo.png"
                with open(logo_path, "wb") as f:
                    f.write(logo_file.getbuffer())
            elif os.path.exists(logo_defaut_github):
                logo_path = logo_defaut_github

            header_text = Paragraph(
                "<b>APEX BUSINESS & COM - SOLUTIONS TEXTILE & MARQUAGE</b><br/>"
                "70 - Marnay<br/>"
                "Tél (Brice Geny) : 06 32 69 73 28 &nbsp;|&nbsp; Tél (Brice Bugna) : 06 29 92 94 74<br/>"
                "Email : brice.geny@gmail.com", 
                style_sub
            )
            
            if logo_path and os.path.exists(logo_path):
                img_logo = RLImage(logo_path, width=110, height=45)
                t_header = Table([[img_logo, header_text]], colWidths=[120, 420])
                t_header.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                story.append(t_header)
            else:
                story.append(header_text)

            story.append(Spacer(1, 10))
            story.append(Paragraph(f"<b>N° Devis :</b> {num_devis} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Date :</b> {datetime.now().strftime('%d/%m/%Y')} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Validité :</b> 30 Jours", style_sub))
            story.append(Spacer(1, 10))

            siret_txt = f"<br/>SIRET : {client_siret}" if client_siret else ""
            contact_txt = f"<br/>Contact : {client_contact}" if client_contact else ""
            nom_aff_client = f"<b>{client_entreprise}</b><br/>À l'attention de : {client_nom}" if client_entreprise else f"<b>{client_nom or 'Client'}</b>"
            client_info_text = f"<b>CLIENT / DESTINATAIRE :</b><br/>{nom_aff_client}<br/>{client_adresse.replace(chr(10), '<br/>')}{siret_txt}{contact_txt}<br/>Email : {client_email}"
            order_info_text = f"<b>DÉTAILS DE LA COMMANDE :</b><br/>Quantité globale : {quantite_globale_totale} pièces<br/>Délai estimé : 8 à 10 jours ouvrés<br/>Conseiller : {conseiller_nom}"
            
            t_info = Table([[Paragraph(client_info_text, style_cell), Paragraph(order_info_text, style_cell)]], colWidths=[270, 270])
            t_info.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f7fafc')),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#cbd5e0')),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('PADDING', (0,0), (-1,-1), 8),
            ]))
            story.append(t_info)
            story.append(Spacer(1, 15))

            table_data = [[
                Paragraph("<b>DÉSIGNATION & CARACTÉRISTIQUES</b>", style_cell_bold), 
                Paragraph("<b>QTÉ</b>", style_cell_bold), 
                Paragraph("<b>PRIX UNIT. HT</b>", style_cell_bold), 
                Paragraph("<b>TOTAL HT</b>", style_cell_bold)
            ]]

            for item in lignes_devis_global:
                libelle_support = f"<b>Support (Sans marquage) : {item['nom_article']}</b>" if item['sans_marquage'] else f"<b>Support : {item['nom_article']}</b>"
                table_data.append([
                    Paragraph(libelle_support, style_cell), 
                    str(item['quantite']), 
                    f"{item['prix_vet_unit']:.2f} €", 
                    f"{item['prix_vet_unit']*item['quantite']:.2f} €"
                ])
                for m in item['marquages']:
                    table_data.append([
                        Paragraph(f"&nbsp;&nbsp;&bull; Marquage : {m['nom']}", style_cell), 
                        str(item['quantite']), 
                        f"{m['tarif']:.2f} €", 
                        f"{m['tarif']*item['quantite']:.2f} €"
                    ])
                if item['frais_prog_broderie'] > 0:
                    table_data.append([
                        Paragraph("&nbsp;&nbsp;&bull; Frais de technique & programme Broderie", style_cell), 
                        "1", 
                        f"{item['frais_prog_broderie']:.2f} €", 
                        f"{item['frais_prog_broderie']:.2f} €"
                    ])
                if item['option_ensachage']:
                    table_data.append([
                        Paragraph(f"&nbsp;&nbsp;&bull; Option : {item['type_sachet']}", style_cell), 
                        str(item['quantite']), 
                        f"{item['coût_ensachage_unit']:.2f} €", 
                        f"{item['coût_ensachage_unit']*item['quantite']:.2f} €"
                    ])
                if item['option_assurance']:
                    table_data.append([
                        Paragraph("&nbsp;&nbsp;&bull; Option : Assurance MHC (Garantie textile)", style_cell), 
                        str(item['quantite']), 
                        f"{item['coût_assurance_unit']:.2f} €", 
                        f"{item['coût_assurance_unit']*item['quantite']:.2f} €"
                    ])
                if item['option_stockage']:
                    table_data.append([
                        Paragraph("&nbsp;&nbsp;&bull; Option : Mise en stockage + picking", style_cell), 
                        str(item['quantite']), 
                        f"{item['coût_stockage_unit']:.2f} €", 
                        f"{item['coût_stockage_unit']*item['quantite']:.2f} €"
                    ])

            if frais_techniques_dossier > 0:
                table_data.append([Paragraph("Frais techniques de dossier", style_cell), "1", f"{frais_techniques_dossier:.2f} €", f"{frais_techniques_dossier:.2f} €"])
            
            if frais_port > 0 or offrir_port:
                port_libelle = f"Frais d'envoi ({zone_livraison})" if not offrir_port else f"Frais d'envoi ({zone_livraison}) - Offerts"
                port_val = f"{frais_port:.2f} €"
                table_data.append([Paragraph(port_libelle, style_cell), "1", port_val, port_val])

            t_main = Table(table_data, colWidths=[260, 45, 115, 120])
            t_main.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#edf2f7')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#e2e8f0')),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('PADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(t_main)
            story.append(Spacer(1, 10))

            totaux_data = [
                ["", Paragraph("Sous-Total HT :", style_right_normal), Paragraph(f"{total_general_ht:.2f} €", style_right_normal)],
                ["", Paragraph("TVA (20%) :", style_right_normal), Paragraph(f"{tva:.2f} €", style_right_normal)],
                ["", Paragraph("TOTAL TTC :", style_right_bold), Paragraph(f"{total_ttc:.2f} €", style_right_bold)],
                ["", Paragraph("Coût unitaire HT / pièce :", style_right_normal), Paragraph(f"{cout_unitaire_moyen:.2f} €", style_right_normal)]
            ]
            t_totaux = Table(totaux_data, colWidths=[240, 160, 140])
            t_totaux.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LINEABOVE', (1, 2), (-1, 2), 1, colors.black),
                ('PADDING', (0,0), (-1,-1), 4),
            ]))
            story.append(t_totaux)
            story.append(Spacer(1, 15))

            conditions_text = f"<b>Conditions de règlement & Bon pour accord :</b><br/>• Règlement : {mode_reglement}<br/>• Fichiers vectoriels fournis (.AI, .EPS, .PDF).<br/>• Bon pour accord daté et signé requis."
            story.append(Paragraph(conditions_text, style_sub))

            doc.build(story)
            
            # Envoi automatique par mail
            sujet = f"Votre devis n° {num_devis}"
            corps = f"""Bonjour {client_contact or client_nom},

Veuillez trouver ci-joint votre devis n° {num_devis} d'un montant total de {total_ttc:.2f} € TTC.

Conditions de règlement : {mode_reglement}
Validité de l'offre : 30 jours

Restant à votre disposition pour tout renseignement complémentaire.

Bien cordialement,
{conseiller_nom}
"""
            succes_mail, msg_mail = envoyer_email_smtp(client_email, sujet, corps, pdf_filename)
            if succes_mail:
                st.success(msg_mail)
            else:
                st.warning(msg_mail)

            st.success(f"Devis PDF professionnel généré sous le numéro : **{num_devis}** (`{pdf_filename}`)")

        if 'dernier_pdf' in st.session_state:
            pdf_filename = st.session_state['dernier_pdf']
            if os.path.exists(pdf_filename):
                with open(pdf_filename, "rb") as f:
                    st.download_button("📥 Télécharger le PDF du devis", f, file_name=os.path.basename(pdf_filename), mime="application/pdf")

# --- ONGLET 11 : SUIVI CRM ---
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