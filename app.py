import sqlite3
import json
import streamlit as st
from datetime import datetime

st.set_page_config(
    page_title="Gestionnaire de Devis - Marquage Textile",
    page_icon="🖨️",
    layout="wide"
)

# Récupération des paramètres URL (si un panier ou une référence est transmis)
query_params = st.query_params
cart_brut = query_params.get("cart", None)
ref_selectionnee = query_params.get("ref", None)

articles_panier = []
if cart_brut:
    try:
        articles_panier = json.loads(cart_brut)
    except Exception as e:
        st.sidebar.error(f"Erreur de lecture du panier : {e}")

st.title("🖨️ Gestionnaire de Devis - Marquage Textile")

# Barre latérale : Infos Client & Expédition
with st.sidebar:
    st.header("Infos Client & Expédition")
    nom_entreprise = st.text_input("Nom de l'Entreprise / Société")
    nom_contact = st.text_input("Nom de la personne contact (ex: Jean Dupont)")
    adresse_client = st.text_area("Adresse complète du Client")
    siret_client = st.text_input("SIRET du Client (optionnel)")
    telephone_client = st.text_input("Téléphone du Client")
    email_client = st.text_input("E-mail du Client")
    zone_livraison = st.selectbox("Zone de Livraison", ["France Continentale", "Corse / Étranger"])
    offrir_port = st.checkbox("Offrir les frais de port (0 €)", value=False)
    
    st.markdown("---")
    st.subheader("Frais techniques de commande (€ HT)")
    frais_tech = st.number_input("Frais techniques globaux", min_value=0.0, value=19.80, step=0.1, format="%.2f")

# Création des onglets (10 articles + Récapitulatif + Suivi CRM)
onglets = st.tabs([f"Article {i+1}" for i in range(10)] + ["📊 Général & Devis", "📈 Suivi CRM"])

configs_articles = []
techniques_disponibles = ["DTF Textile Fin", "Broderie"]
emplacements_par_technique = {
    "DTF Textile Fin": ["Cœur (13x9 cm)", "Dos (28x20 cm)", "Manche droite", "Manche gauche", "Poitrine centrale"],
    "Broderie": ["Cœur (10x10 cm)", "Dos (25x20 cm)", "Casquette (Frontal)", "Pantalon poche", "Manche"]
}

for i in range(10):
    with onglets[i]:
        st.subheader(f"Configuration de l'Article {i+1}")
        
        # Récupération des données du panier web si disponibles pour cet article
        art_donnees = articles_panier[i] if i < len(articles_panier) else None
        
        default_nom = f"{art_donnees['designation']} (Réf: {art_donnees['reference']})" if art_donnees else (f"Article {i+1}" if i > 0 else "T-shirt 100% coton")
        default_q = 0
        default_prix = 4.92
        details_texte = ""
        
        if art_donnees:
            lignes_det = []
            t_q = 0
            t_p = 0.0
            for l in art_donnees['lignes']:
                t_q += l['quantite']
                t_p += l['quantite'] * l['prix']
                lignes_det.append(f"- {l['quantite']}x {l['coloris']} (Taille: {l['taille']}) à {l['prix']:.2f}€ HT/p")
            default_q = t_q
            default_prix = t_p / t_q if t_q > 0 else 4.92
            details_texte = "\n".join(lignes_det)
        elif i == 0 and ref_selectionnee:
            try:
                conn = sqlite3.connect("apex_catalogue_imbretex.db")
                cursor = conn.cursor()
                cursor.execute("SELECT designation_courte, prix_vente_ht FROM produits WHERE ref_produit = ?", (ref_selectionnee,))
                res = cursor.fetchone()
                conn.close()
                if res:
                    default_nom = f"{res[0]} (Réf: {ref_selectionnee})"
                    if res[1]:
                        default_prix = float(str(res[1]).replace(',', '.'))
            except Exception:
                pass

        sans_marquage = st.checkbox(f"Vêtement sans marquage (fourniture seule) {i+1}", value=False, key=f"sans_m_{i}")
        
        col1, col2 = st.columns(2)
        with col1:
            nom_art = st.text_input(f"Nom / Référence du vêtement {i+1}", value=default_nom, key=f"nom_{i}")
            qte = st.number_input(f"Quantité totale (pcs) {i+1}", min_value=0, value=default_q if default_q > 0 else (10 if i==0 and not art_donnees else 0), step=1, key=f"qte_{i}")
        with col2:
            prix_u = st.number_input(f"Prix unitaire HT (€) {i+1}", value=default_prix, format="%.2f", key=f"prix_{i}")

        if details_texte:
            st.markdown("**Détail des tailles et coloris sélectionnés :**")
            st.text(details_texte)

        st.markdown("---")
        st.markdown(f"### Gestion des Marquages (Article {i+1})")
        
        nb_maq = 0
        marqs = []
        if not sans_marquage and qte > 0:
            # Si des marquages ont été transmis depuis le panier web, on les pré-remplit
            default_nb_maq = len(art_donnees['marquages']) if (art_donnees and art_donnees.get('marquages')) else 1
            nb_maq = st.selectbox(f"Nombre de marquages pour l'article {i+1}", [1, 2, 3, 4], index=default_nb_maq-1 if default_nb_maq<=4 else 0, key=f"nb_m_{i}")
            
            for m in range(nb_maq):
                col_m1, col_m2 = st.columns(2)
                maq_web = art_donnees['marquages'][m] if (art_donnees and art_donnees.get('marquages') and m < len(art_donnees['marquages'])) else {}
                
                with col_m1:
                    tech_def = maq_web.get('technique', techniques_disponibles[0])
                    tech_idx = techniques_disponibles.index(tech_def) if tech_def in techniques_disponibles else 0
                    tech = st.selectbox(f"Technique M{m+1} (Art {i+1})", techniques_disponibles, index=tech_idx, key=f"tech_{i}_{m}")
                with col_m2:
                    emplacements_disponibles = emplacements_par_technique.get(tech, ["Cœur (13x9 cm)"])
                    emp_def = maq_web.get('emplacement', emplacements_disponibles[0])
                    emp_idx = emplacements_disponibles.index(emp_def) if emp_def in emplacements_disponibles else 0
                    emp = st.selectbox(f"Emplacement M{m+1} (Art {i+1})", emplacements_disponibles, index=emp_idx, key=f"emp_{i}_{m}")
                marqs.append({"technique": tech, "emplacement": emp})
        else:
            st.info("Aucun marquage pour cet article (ou quantité à 0).")

        configs_articles.append({
            "nom": nom_art,
            "quantite": qte,
            "prix_unitaire": prix_u,
            "details": details_texte,
            "sans_marquage": sans_marquage,
            "marquages": marqs
        })

# Onglet Récapitulatif Général & Devis
with onglets[10]:
    st.header("📊 Récapitulatif Général & Devis")
    
    total_vetements = 0.0
    for idx, art in enumerate(configs_articles):
        if art['quantite'] > 0:
            t_art = art['quantite'] * art['prix_unitaire']
            total_vetements += t_art
            st.write(f"**Article {idx+1} :** {art['nom']} — {art['quantite']} pcs × {art['prix_unitaire']:.2f} € = {t_art:.2f} € HT")
            if art['details']:
                st.text(art['details'])
    
    st.markdown("---")
    frais_port_val = 0.0 if offrir_port else 15.0
    total_net_ht = total_vetements + frais_tech + frais_port_val
    
    st.write(f"**Total Vêtements HT :** {total_vetements:.2f} €")
    st.write(f"**Frais techniques :** {frais_tech:.2f} €")
    st.write(f"**Frais de port :** {'OFFERTS (0.00 €)' if offrir_port else f'{frais_port_val:.2f} € HT'}")
    st.subheader(f"💰 Total Net HT : {total_net_ht:.2f} €")
    
    if st.button("💾 Générer le devis complet", type="primary"):
        st.success("Devis calculé et validé avec succès !")

# Onglet Suivi CRM : Historique et Demandes Web entrantes
with onglets[11]:
    st.header("📈 Tableau de Suivi CRM & Historique des Devis")
    st.write("Retrouvez ci-dessous les demandes de devis transmises par les clients depuis le site web.")
    
    st.markdown("### 🌐 Demandes Web entrantes (Statut : WEB - A traiter)")
    
    try:
        conn = sqlite3.connect("apex_catalogue_imbretex.db")
        cursor = conn.cursor()
        
        # Vérification et lecture de la table devis_web
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='devis_web'")
        if cursor.fetchone():
            cursor.execute("SELECT id, date, nom_client, email_client, panier_json, statut FROM devis_web WHERE statut = 'WEB - A traiter' ORDER BY id DESC")
            demandes_web = cursor.fetchall()
            
            if not demandes_web:
                st.info("Aucune nouvelle demande web en attente.")
            else:
                for d in demandes_web:
                    id_web, date_web, nom_web, email_web, panier_json_str, statut_web = d
                    
                    with st.expander(f"📁 Demande Web #{id_web} — Client : {nom_web} ({email_web}) [Reçue le {date_web}]"):
                        st.write(f"**Statut actuel :** `{statut_web}`")
                        st.write(f"**E-mail de contact :** {email_web}")
                        
                        try:
                            panier_data = json.loads(panier_json_str)
                            st.markdown("**Détail du panier et des options :**")
                            for idx, article in enumerate(panier_data):
                                st.write(f"* **Article {idx+1} :** {article.get('designation')} (Réf: {article.get('reference')})")
                                for ligne in article.get('lignes', []):
                                    st.write(f"  - {ligne.get('quantite')}x {ligne.get('coloris')} (Taille: {ligne.get('taille')}) à {ligne.get('prix')}€ HT")
                                for maq in article.get('marquages', []):
                                    if maq.get('technique') and maq.get('emplacement'):
                                        st.write(f"  > 🪡 **Marquage :** {maq.get('technique')} sur {maq.get('emplacement')}")
                        except Exception as json_err:
                            st.error(f"Erreur de lecture du panier : {json_err}")
                        
                        if st.button(f"Prendre en charge la demande #{id_web}", key=f"traiter_{id_web}"):
                            cursor.execute("UPDATE devis_web SET statut = 'En cours' WHERE id = ?", (id_web,))
                            conn.commit()
                            st.rerun()
        else:
            st.info("Aucune table de devis web initialisée pour l'instant.")
        
        conn.close()
    except Exception as e:
        st.error(f"Erreur lors du chargement des demandes web : {e}")