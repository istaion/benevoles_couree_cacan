import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv

load_dotenv()

# Configuration
PRODUCTION = True  # Changez à True pour envoyer aux vraies adresses
FAKE_EMAIL = "dihime1249@acedby.com"  # Adresse de test quand PRODUCTION = False
GMAIL_USERNAME = os.getenv("GMAIL_USERNAME")
QUESTIONNAIRE_LINK = "https://docs.google.com/forms/d/e/1FAIpQLSfzb-O1Z1EBBYcQHdWzx7w3XRvD2IxANkdS9oUnK6l97W7IbA/viewform?usp=header"  # Remplacez par le vrai lien du questionnaire

def load_volunteers_from_questionnaire(questionnaire_file_path):
    """
    Charge la liste des bénévoles depuis le fichier questionnaire Excel
    
    Args:
        questionnaire_file_path (str): Chemin vers le fichier Excel du questionnaire
        
    Returns:
        list: Liste des dictionnaires avec les informations des bénévoles
    """
    # Lire le fichier Excel
    df = pd.read_excel(questionnaire_file_path)
    
    volunteers = []
    
    # Parcourir chaque ligne du questionnaire
    for index, row in df.iterrows():
        # Adapter ces noms de colonnes selon votre fichier Excel
        # Vous devrez peut-être ajuster les noms des colonnes
        try:
            name = ""
            email = ""
            
            # Essayer différentes combinaisons de noms de colonnes possibles
            # Adaptez selon les vrais noms de colonnes de votre fichier
            if "Prénom" in df.columns and "NOM" in df.columns:
                prenom = str(row["Prénom"]).strip() if pd.notna(row["Prénom"]) else ""
                nom = str(row["NOM"]).strip() if pd.notna(row["NOM"]) else ""
                name = f"{prenom} {nom}".strip()
            elif "Nom complet" in df.columns:
                name = str(row["Nom complet"]).strip() if pd.notna(row["Nom complet"]) else ""
            elif "Nom" in df.columns:
                name = str(row["Nom"]).strip() if pd.notna(row["Nom"]) else ""
            
            # Pour l'email
            if "Adresse mail" in df.columns:
                email = str(row["Adresse mail"]).strip() if pd.notna(row["Adresse mail"]) else ""
            elif "Email" in df.columns:
                email = str(row["Email"]).strip() if pd.notna(row["Email"]) else ""
            elif "Adresse e-mail" in df.columns:
                email = str(row["Adresse e-mail"]).strip() if pd.notna(row["Adresse e-mail"]) else ""
            
            # Vérifier que nous avons au moins un nom et un email
            if name and email and email != "nan":
                volunteers.append({
                    "name": name,
                    "email": email
                })
            else:
                print(f"⚠️  Ligne {index + 1}: Données manquantes - Nom: '{name}', Email: '{email}'")
                
        except Exception as e:
            print(f"❌ Erreur ligne {index + 1}: {str(e)}")
    
    return volunteers

def create_thank_you_email_content(volunteer_name, questionnaire_link):
    """
    Crée le contenu de l'email de remerciement
    
    Args:
        volunteer_name (str): Nom du bénévole
        questionnaire_link (str): Lien vers le questionnaire
        
    Returns:
        str: Contenu de l'email formaté
    """
    
    email_content = f"""Chèr·e {volunteer_name},

Tout d'abord, une myriade de merci ! Ton aide a été précieuse tout au long de cette journée, et c'est grâce à des personnes comme toi que cette fête a pu avoir lieu dans la joie et la bonne humeur.

On te propose maintenant un petit questionnaire de retour :
👉 Il te permet de partager tes impressions,
👉 et aussi de t'inscrire à notre mailing liste pour qu'on puisse te recontacter l'année prochaine – que ce soit comme bénévole, ou même pour rejoindre la team d'organisation !

Voici le questionnaire : {questionnaire_link}

Encore un immense merci, et passe une très belle semaine ✨

À bientôt,
L'équipe d'organisation"""

    return email_content

def send_thank_you_email(smtp_server, sender_email, recipient_email, volunteer_name, questionnaire_link):
    """
    Envoie un email de remerciement à un bénévole
    
    Args:
        smtp_server: Serveur SMTP connecté
        sender_email (str): Adresse email d'envoi
        recipient_email (str): Adresse email du destinataire
        volunteer_name (str): Nom du bénévole
        questionnaire_link (str): Lien vers le questionnaire
        
    Returns:
        bool: True si envoi réussi, False sinon
    """
    try:
        # Créer le message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Merci pour ton aide ! 💛"
        
        # Ajouter le contenu
        body = create_thank_you_email_content(volunteer_name, questionnaire_link)
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Envoyer l'email
        text = msg.as_string()
        smtp_server.sendmail(sender_email, recipient_email, text)
        
        return True
    except Exception as e:
        print(f"Erreur lors de l'envoi à {recipient_email}: {str(e)}")
        return False

def send_thank_you_emails_to_all_volunteers(volunteers_info, questionnaire_link, production=PRODUCTION):
    """
    Envoie des emails de remerciement à tous les bénévoles
    
    Args:
        volunteers_info (list): Liste des informations des bénévoles
        questionnaire_link (str): Lien vers le questionnaire
        production (bool): Si True, envoie aux vraies adresses, sinon à l'adresse test
    """
    # Configuration Gmail
    gmail_user = GMAIL_USERNAME
    
    # Récupérer le mot de passe depuis les variables d'environnement
    gmail_password = os.getenv("GMAIL_PASSWORD")
    
    if not gmail_password:
        print("❌ Mot de passe Gmail non trouvé dans les variables d'environnement")
        print("Veuillez définir GMAIL_PASSWORD dans votre fichier .env")
        return
    
    try:
        # Se connecter au serveur Gmail
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(gmail_user, gmail_password)
        
        # Compteurs
        sent_count = 0
        failed_count = 0
        skipped_count = 0
        
        print(f"\nMode: {'PRODUCTION' if production else 'TEST'}")
        print(f"Nombre de bénévoles à traiter: {len(volunteers_info)}")
        print(f"Lien questionnaire: {questionnaire_link}")
        print("-" * 50)
        
        for volunteer in volunteers_info:
            name = volunteer['name']
            email = volunteer['email']
            
            # Vérifier si l'email existe
            if not email or email.strip() == '':
                print(f"❌ {name}: Pas d'adresse email")
                skipped_count += 1
                continue
            
            # Déterminer l'adresse de destination
            if production:
                recipient_email = email.strip()
            else:
                recipient_email = FAKE_EMAIL
            
            # Envoyer l'email
            print(f"📧 Envoi à {name} ({recipient_email})...")
            
            if send_thank_you_email(server, gmail_user, recipient_email, name, questionnaire_link):
                print(f"✅ {name}: Email envoyé avec succès")
                sent_count += 1
            else:
                print(f"❌ {name}: Échec de l'envoi")
                failed_count += 1
        
        # Fermer la connexion
        server.quit()
        
        # Résumé
        print("\n" + "="*50)
        print("RÉSUMÉ DE L'ENVOI - EMAILS DE REMERCIEMENT")
        print("="*50)
        print(f"✅ Emails envoyés avec succès: {sent_count}")
        print(f"❌ Emails échoués: {failed_count}")
        print(f"⏭️  Bénévoles ignorés (pas d'email): {skipped_count}")
        print(f"📊 Total traité: {len(volunteers_info)}")
        
    except Exception as e:
        print(f"Erreur de connexion au serveur Gmail: {str(e)}")
        print("\nVérifiez:")
        print("1. Que votre nom d'utilisateur Gmail est correct")
        print("2. Que vous utilisez un mot de passe d'application si l'authentification à 2 facteurs est activée")
        print("3. Que l'accès aux applications moins sécurisées est autorisé (si applicable)")

def preview_thank_you_email(volunteer_name, questionnaire_link):
    """
    Prévisualise l'email de remerciement
    
    Args:
        volunteer_name (str): Nom du bénévole
        questionnaire_link (str): Lien vers le questionnaire
    """
    print("="*60)
    print("PRÉVISUALISATION DE L'EMAIL DE REMERCIEMENT")
    print("="*60)
    print(f"Objet: Merci pour ton aide ! 💛")
    print("-"*60)
    print(create_thank_you_email_content(volunteer_name, questionnaire_link))
    print("="*60)

def display_volunteers_summary(volunteers_info):
    """
    Affiche un résumé des bénévoles chargés
    
    Args:
        volunteers_info (list): Liste des informations des bénévoles
    """
    print(f"\n📊 RÉSUMÉ DES BÉNÉVOLES CHARGÉS")
    print("="*50)
    print(f"Nombre total de bénévoles: {len(volunteers_info)}")
    
    with_email = sum(1 for v in volunteers_info if v['email'])
    without_email = len(volunteers_info) - with_email
    
    print(f"✅ Avec adresse email: {with_email}")
    print(f"❌ Sans adresse email: {without_email}")
    print("="*50)
    
    # Afficher les premiers bénévoles pour vérification
    print("\n🔍 APERÇU DES PREMIERS BÉNÉVOLES:")
    for i, volunteer in enumerate(volunteers_info[:5]):
        print(f"{i+1}. {volunteer['name']} - {volunteer['email']}")
    
    if len(volunteers_info) > 5:
        print(f"... et {len(volunteers_info) - 5} autres")

# Utilisation du script
if __name__ == "__main__":
    print("🎉 SCRIPT D'ENVOI D'EMAILS DE REMERCIEMENT AUX BÉNÉVOLES")
    print("="*60)
    
    # Fichier du questionnaire
    questionnaire_file = "Questionnaire bénévoles 2025 - 19e édition (réponses).xlsx"
    
    # Vérifier que le fichier existe
    if not os.path.exists(questionnaire_file):
        print(f"❌ Erreur: Le fichier '{questionnaire_file}' n'existe pas.")
        print("Veuillez vérifier le nom et l'emplacement du fichier.")
        exit(1)
    
    try:
        # Charger les bénévoles depuis le questionnaire
        print(f"📂 Chargement des bénévoles depuis: {questionnaire_file}")
        volunteers_info = load_volunteers_from_questionnaire(questionnaire_file)
        
        # Afficher le résumé
        display_volunteers_summary(volunteers_info)
        
        # Vérifier qu'on a des bénévoles
        if not volunteers_info:
            print("❌ Aucun bénévole trouvé dans le fichier.")
            print("Vérifiez les noms des colonnes dans votre fichier Excel.")
            exit(1)
        
        # Prévisualiser un email
        if volunteers_info:
            print(f"\n🔍 PRÉVISUALISATION pour {volunteers_info[0]['name']}:")
            preview_thank_you_email(volunteers_info[0]['name'], QUESTIONNAIRE_LINK)
        
        # Demander confirmation
        print(f"\n⚠️  MODE ACTUEL: {'PRODUCTION' if PRODUCTION else 'TEST'}")
        if not PRODUCTION:
            print(f"🔄 Les emails seront envoyés à l'adresse test: {FAKE_EMAIL}")
        else:
            print("🚨 Les emails seront envoyés aux vraies adresses des bénévoles!")
        
        confirmation = input("\n✅ Voulez-vous continuer l'envoi ? (oui/non): ").lower()
        
        if confirmation in ['oui', 'o', 'yes', 'y']:
            # Vérifier que le lien du questionnaire est défini
            if QUESTIONNAIRE_LINK == "https://forms.gle/VOTRE_LIEN_ICI":
                print("⚠️  ATTENTION: N'oubliez pas de remplacer QUESTIONNAIRE_LINK par le vrai lien!")
                use_placeholder = input("Continuer avec le lien de placeholder ? (oui/non): ").lower()
                if use_placeholder not in ['oui', 'o', 'yes', 'y']:
                    print("❌ Envoi annulé. Modifiez QUESTIONNAIRE_LINK et relancez le script.")
                    exit(1)
            
            # Envoyer les emails
            send_thank_you_emails_to_all_volunteers(volunteers_info, QUESTIONNAIRE_LINK, PRODUCTION)
        else:
            print("❌ Envoi annulé par l'utilisateur.")
    
    except FileNotFoundError:
        print(f"❌ Erreur: Impossible de trouver le fichier '{questionnaire_file}'")
    except Exception as e:
        print(f"❌ Erreur: {str(e)}")
        print("\nVérifiez:")
        print("1. Que le fichier Excel existe et est accessible")
        print("2. Que les noms des colonnes correspondent à votre fichier")
        print("3. Que le fichier n'est pas ouvert dans Excel")