import streamlit as st
import pandas as pd
from datetime import datetime
import os
import json
import smtplib
from email.message import EmailMessage
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Gestionnaire de Devis - Marquage Textile", layout="wide")

# --- GESTION DU COMPTEUR DE DEVIS AUTOMATIQUE ---
COMPTEUR_FILE = "compteur_devis.json"

def obtenir_prochain_numero_devis():
    date_jour = datetime.now().strftime("%Y_%m_%d")
    data = {"date": date_jour, "seq": 1}
    
    if os.path.exists(COMPTEUR_FILE):
        try:
            with open(COMPTEUR_FILE, "r") as f:
                saved_data = json.load(f)
                if saved_data.get("date") == date_jour:
                    data["seq"] = saved_data.get("seq", 0) + 1
        except Exception:
            pass
            
    with open(COMPTEUR_FILE, "w") as f:
        json.dump(data, f)
        
    return f"{datetime.now().strftime('%Y/%m/%d')}_{data['seq']:06d}"

# --- FONCTIONS DE CALCUL DES TARIFS (AVEC QUANTITÉ GLOBALE) ---

def get_tarif_dtf_fin(qte_globale, emplacement):
    if qte_globale <= 5: idx = 0
    elif qte_globale <= 9: idx = 1
    elif qte_globale <= 19: idx = 2
    elif qte_globale <= 29: idx = 3
    elif qte_globale <= 39: idx = 4
    elif qte_globale <= 49: idx = 5
    elif qte_globale <= 99: idx = 6
    elif qte_globale <= 249: idx = 7
    elif qte_globale <= 499: idx = 8
    elif qte_globale <= 999: idx = 9
    elif qte_globale <= 4999: idx = 10
    else: idx = 11

    tarifs = {
        "Cœur (13x9 cm)": [6.00, 4.50, 3.60, 2.81, 2.50, 2.40, 2.00, 1.80, 1.60, 1.40, 1.20, 0.70],
        "Dos D10 (20x13 cm)": [7.92, 6.50, 5.50, 5.00, 4.20, 4.00, 3.50, 3.00, 2.70, 2.50, 2.00, 1.00],
        "Dos D20 (28x20 cm)": [13.67, 10.00, 9.00, 6.90, 6.30, 6.00, 5.50, 4.60, 4.00, 3.50, 2.50, 1.50],
        "Format P (37x27 cm)": [17.00, 13.00, 12.00, 10.00, 9.00, 8.00, 7.50, 6.50, 6.00, 5.00, 4.00, 2.50],
        "Manche (9x8 cm)": [7.20, 5.40, 4.32, 3.37, 3.00, 2.88, 2.40, 2.16, 1.92, 1.68, 1.44, 0.84],
        "+ Perso. Nom": [4.39, 3.50, 2.50, 2.20, 1.90, 1.80, 1.70, 1.50, 1.30, 0.80, 0.30, 0.20]
    }
    return tarifs.get(emplacement, [0]*12)[idx]

def get_tarif_dtf_epais(qte_globale, emplacement):
    if qte_globale <= 5: idx = 0
    elif qte_globale <= 9: idx = 1
    elif qte_globale <= 19: idx = 2
    elif qte_globale <= 29: idx = 3
    elif qte_globale <= 39: idx = 4
    elif qte_globale <= 49: idx = 5
    elif qte_globale <= 99: idx = 6
    elif qte_globale <= 249: idx = 7
    elif qte_globale <= 499: idx = 8
    elif qte_globale <= 999: idx = 9
    elif qte_globale <= 4999: idx = 10
    else: idx = 11

    tarifs = {
        "Cœur (13x9 cm)": [6.60, 4.95, 3.96, 3.09, 2.75, 2.64, 2.20, 1.98, 1.76, 1.54, 1.32, 0.77],
        "Dos D10 (20x13 cm)": [8.71, 7.15, 6.05, 5.50, 4.62, 4.40, 3.85, 3.30, 2.97, 2.75, 2.20, 1.10],
        "Dos D20 (28x20 cm)": [15.04, 11.00, 9.90, 7.59, 6.93, 6.60, 6.05, 5.06, 4.40, 3.85, 2.75, 1.65],
        "Format P (37x27 cm)": [18.70, 14.30, 13.20, 11.00, 9.90, 8.80, 8.25, 7.15, 6.60, 5.50, 4.40, 2.75],
        "Manche (9x8 cm)": [7.92, 5.94, 4.75, 3.71, 3.30, 3.17, 2.64, 2.38, 2.11, 1.85, 1.58, 0.92],
        "+ Perso. Nom": [4.83, 3.85, 2.75, 2.42, 2.09, 1.98, 1.87, 1.65, 1.43, 0.88, 0.33, 0.22]
    }
    return tarifs.get(emplacement, [0]*12)[idx]

def get_tarif_broderie(qte_globale, emplacement):
    if qte_globale <= 3: idx = 0
    elif qte_globale <= 11: idx = 1
    elif qte_globale <= 23: idx = 2
    elif qte_globale <= 47: idx = 3
    elif qte_globale <= 95: idx = 4
    elif qte_globale <= 251: idx = 5
    elif qte_globale <= 503: idx = 6
    elif qte_globale <= 1007: idx = 7
    elif qte_globale <= 1511: idx = 8
    else: idx = 9

    tarifs = {
        "Poitrine (9x8 cm)": [15.90, 12.13, 9.20, 6.90, 5.30, 4.80, 4.50, 4.20, 3.90, 3.50],
        "Dos D10 (25x10 cm)": [18.50, 14.60, 11.30, 9.10, 7.60, 7.10, 6.70, 6.30, 5.90, 5.40],
        "Dos Large D20 (25x20)": [23.20, 18.25, 14.20, 11.90, 9.90, 9.30, 8.80, 8.30, 7.80, 7.20],
        "Col/Signature (7x2)": [10.90, 8.37, 6.20, 4.40, 3.10, 2.85, 2.65, 2.45, 2.25, 2.00],
        "Casquettes / Bonnets": [16.10, 12.27, 9.30, 7.10, 5.50, 5.05, 4.75, 4.45, 4.15, 3.75],
        "Manche (8x5 cm)": [16.10, 12.27, 9.30, 7.10, 5.50, 5.05, 4.75, 4.45, 4.15, 3.75],
        "Pantalon / Poche": [18.30, 13.97, 10.90, 8.60, 7.10, 6.60, 6.20, 5.80, 5.45, 4.95],
        "+ Perso. Nom (Cœur)": [6.00, 4.00, 4.00, 3.50, 3.00, 2.80, 2.60, 2.40, 2.20, 2.00]
    }
    return tarifs.get(emplacement, [0]*10)[idx]

def calculer_frais_port(montant_ht, zone):
    if montant_ht >= 1500: return 0.0
    if zone in ["France Continentale", "Corse, Monaco ou Andorre"] and montant_ht >= 1000: return 0.0
    
    if zone == "France Continentale":
        if montant_ht < 99.99: return 14.95
        elif montant_ht <= 499.99: return 20.95
        elif montant_ht <= 999.99: return 24.95
        else: return 0.0
    elif zone == "Corse, Monaco ou Andorre":
        if montant_ht < 99.99: return 19.95
        elif montant_ht <= 499.99: return 25.95
        elif montant_ht <= 999.99: return 29.95
        else: return 0.0
    elif zone == "Espace UE":
        if montant_ht < 99.99: return 25.95
        elif montant_ht <= 499.99: return 39.00
        elif montant_ht <= 999.99: return 60.00
        else: return 90.00
    else: return 0.0

# --- INTERFACE ---
st.title("🖨️ Gestionnaire de Devis - Marquage Textile")

st.sidebar.header("📋 Infos Client & Expédition")
client_nom = st.sidebar.text_input("Nom du Client / Entreprise")
client_adresse = st.sidebar.text_area("Adresse complète du Client")
client_siret = st.sidebar.text_input("SIRET du Client (optionnel)")
client_contact = st.sidebar.text_input("Contact / Téléphone")
client_email = st.sidebar.text_input("E-mail du Client")
zone_livraison = st.sidebar.selectbox("Zone de Livraison", ["France Continentale", "Corse, Monaco ou Andorre", "Espace UE", "DOM/TOM et pays hors UE"])
offrir_port = st.sidebar.checkbox("Offrir les frais de port (0 €)", value=False)
frais_techniques_dossier = st.sidebar.number_input("Frais techniques de commande (€ HT)", value=19.80)

st.sidebar.markdown("---")
st.sidebar.header("🖼️ Logo de l'entreprise")
logo_file = st.sidebar.file_uploader("Importer le logo Apex (PNG/JPG)", type=["png", "jpg", "jpeg"])

st.sidebar.markdown("---")
# Récupération automatique et invisible du mot de passe stocké dans les secrets Streamlit Cloud
gmail_password = st.secrets.get("EMAIL_PASSWORD", "")

noms_onglets = [f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis"]
onglets = st.tabs(noms_onglets)

# --- ÉTAPE 1 : COLLECTE DES DONNÉES BRUTES DE CHAQUE ARTICLE ---
articles_saisis = []

for i in range(10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        col1, col2 = st.columns(2)
        with col1:
            nom_article = st.text_input(f"Nom / Référence du vêtement {i+1}", value=f"T-Shirt 100% coton" if i==0 else f"Article {i+1}", key=f"nom_{i}")
            qte = st.number_input(f"Quantité (pcs) {i+1}", min_value=0, value=10 if i==0 else 0, key=f"qte_{i}")
            prix_vetement_ht = st.number_input(f"Prix unitaire HT du vêtement (€) {i+1}", min_value=0.0, value=4.92, key=f"px_vet_{i}")
            
        with col2:
            st.markdown("**Gestion des Marquages (jusqu'à 4)**")
            nb_marquages = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], key=f"nb_m_{i}")

        marquages_config = []
        for m in range(nb_marquages):
            st.markdown(f"--- *Marquage {m+1}*")
            mc1, mc2 = st.columns(2)
            with mc1:
                t_marq = st.selectbox(f"Technique M{m+1} (Art {i+1})", ["DTF Textile Fin", "DTF Textile Épais", "Broderie HD"], key=f"t_marq_{i}_{m}")
            with mc2:
                if t_marq == "DTF Textile Fin":
                    emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_f_{i}_{m}")
                elif t_marq == "DTF Textile Épais":
                    emp = st.selectbox(f"Emplacement M{m+1}", ["Cœur (13x9 cm)", "Dos D10 (20x13 cm)", "Dos D20 (28x20 cm)", "Format P (37x27 cm)", "Manche (9x8 cm)", "+ Perso. Nom"], key=f"emp_e_{i}_{m}")
                else:
                    emp = st.selectbox(f"Emplacement M{m+1}", ["Poitrine (9x8 cm)", "Dos D10 (25x10 cm)", "Dos Large D20 (25x20)", "Col/Signature (7x2)", "Casquettes / Bonnets", "Manche (8x5 cm)", "Pantalon / Poche", "+ Perso. Nom (Cœur)"], key=f"emp_b_{i}_{m}")
            
            marquages_config.append({"technique": t_marq, "emplacement": emp})

        if qte > 0:
            articles_saisis.append({
                "nom_article": nom_article,
                "quantite": qte,
                "prix_vet_unit": prix_vetement_ht,
                "marquages_config": marquages_config
            })

# --- ÉTAPE 2 : CALCUL DES QUANTITÉS GLOBALES PAR MARQUAGE POUR TARIFICATION DÉGRESSIVE ---
compteur_global_marquages = {}
for art in articles_saisis:
    qte_art = art["quantite"]
    for m in art["marquages_config"]:
        cle = (m["technique"], m["emplacement"])
        compteur_global_marquages[cle] = compteur_global_marquages.get(cle, 0) + qte_art

# --- ÉTAPE 3 : CONSOLIDATION DU DEVIS ---
lignes_devis_global = []
for art in articles_saisis:
    qte = art["quantite"]
    marquages_calcules = []
    total_marquage_unit = 0.0
    
    for m in art["marquages_config"]:
        cle = (m["technique"], m["emplacement"])
        qte_globale_marq = compteur_global_marquages.get(cle, qte)
        
        if m["technique"] == "DTF Textile Fin":
            tarif_m = get_tarif_dtf_fin(qte_globale_marq, m["emplacement"])
        elif m["technique"] == "DTF Textile Épais":
            tarif_m = get_tarif_dtf_epais(qte_globale_marq, m["emplacement"])
        else:
            tarif_m = get_tarif_broderie(qte_globale_marq, m["emplacement"])
            
        total_marquage_unit += tarif_m
        marquages_calcules.append({"nom": f"{m['technique']} - {m['emplacement']}", "tarif": tarif_m})

    total_unit_ht = art["prix_vet_unit"] + total_marquage_unit
    total_ligne_ht = total_unit_ht * qte
    
    lignes_devis_global.append({
        "nom_article": art["nom_article"],
        "quantite": qte,
        "prix_vet_unit": art["prix_vet_unit"],
        "marquages": marquages_calcules,
        "total_ligne_ht": total_ligne_ht
    })

# --- ONGLET GÉNÉRAL & DEVIS ---
with onglets[10]:
    st.subheader("📊 Récapitulatif Général & Génération du Devis Professionnel")

    if not lignes_devis_global:
        st.warning("Veuillez renseigner au moins un article avec une quantité supérieure à 0.")
    else:
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
            pdf_filename = f"Devis_{num_devis.replace('/', '_')}.pdf"
            st.session_state['dernier_pdf'] = pdf_filename
            st.session_state['dernier_num'] = num_devis
            st.session_state['total_ttc_cache'] = total_ttc

        if 'dernier_pdf' in st.session_state:
            pdf_filename = st.session_state['dernier_pdf']
            num_devis = st.session_state['dernier_num']
            total_ttc = st.session_state['total_ttc_cache']

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
            else:
                # Logo par défaut intégré s'il n'est pas uploadé
                logo_path = "logo_apex.png"
                if not os.path.exists(logo_path):
                    # Création d'une image de secours propre si le fichier physique n'est pas là
                    from PIL import Image as PILImage, ImageDraw
                    img_secours = PILImage.new('RGB', (300, 120), color=(240, 240, 240))
                    d = ImageDraw.Draw(img_secours)
                    d.text((20, 40), "APEX Business & COM", fill=(200, 30, 30))
                    img_secours.save(logo_path)

            # En-tête avec téléphones sur une ligne et e-mail en dessous
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
            client_info_text = f"<b>CLIENT / DESTINATAIRE :</b><br/><b>{client_nom or 'Client'}</b><br/>{client_adresse.replace(chr(10), '<br/>')}{siret_txt}{contact_txt}<br/>Email : {client_email}"
            order_info_text = f"<b>DÉTAILS DE LA COMMANDE :</b><br/>Quantité globale : {quantite_globale_totale} pièces<br/>Délai estimé : 8 à 10 jours ouvrés<br/>Conseiller : Brice Geny"
            
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
                table_data.append([
                    Paragraph(f"<b>Support : {item['nom_article']}</b>", style_cell), 
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

            conditions_text = "<b>Conditions de règlement & Bon pour accord :</b><br/>• Acompte de 50% à la commande, solde à la livraison.<br/>• Fichiers vectoriels fournis (.AI, .EPS, .PDF).<br/>• Bon pour accord daté et signé requis."
            story.append(Paragraph(conditions_text, style_sub))

            doc.build(story)
            st.success(f"Devis PDF professionnel généré sous le numéro : **{num_devis}** (`{pdf_filename}`)")

            # --- ENVOI DIRECT GMAIL ---
            st.markdown("### ✉️ Envoi direct par E-mail (via brice.geny@gmail.com)")
            
            if st.button("🚀 Envoyer le devis par e-mail maintenant"):
                if not client_email:
                    st.error("Veuillez renseigner l'e-mail du client dans la barre latérale.")
                elif not gmail_password:
                    st.error("Veuillez renseigner votre mot de passe d'application Gmail dans la barre latérale.")
                else:
                    try:
                        msg = EmailMessage()
                        msg['Subject'] = f"Devis {num_devis} - Marquage Textile"
                        msg['From'] = "brice.geny@gmail.com"
                        msg['To'] = client_email
                        msg['Cc'] = "brice.geny@gmail.com"
                        
                        corps_mail = f"""Bonjour {client_nom or 'Client'},

Veuillez trouver ci-joint votre devis n° {num_devis} d'un montant total de {total_ttc:.2f} € TTC.

Restant à votre disposition pour toute information complémentaire.

Cordialement,
Brice Geny
"""
                        msg.set_content(corps_mail)

                        with open(pdf_filename, 'rb') as f:
                            file_data = f.read()
                            file_name = os.path.basename(pdf_filename)
                        msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

                        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                            smtp.login("brice.geny@gmail.com", gmail_password)
                            smtp.send_message(msg)
                        
                        st.success(f"✅ E-mail envoyé avec succès à {client_email} (avec copie à vous-même) !")
                    except Exception as e:
                        st.error(f"❌ Erreur lors de l'envoi de l'e-mail : {e}")