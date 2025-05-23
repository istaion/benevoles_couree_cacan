import pandas as pd
import re
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv

load_dotenv()

# Configuration
PRODUCTION = False  # Changez à True pour envoyer aux vraies adresses
FAKE_EMAIL = "nokos46841@betzenn.com"  # Adresse de test quand PRODUCTION = False
GMAIL_USERNAME = os.getenv("GMAIL_USERNAME")
LIEN_TABLEAU = "https://docs.google.com/spreadsheets/d/1xkpUqZJvz-IIqCZ5M4Uk_Cio48ZIapwGYy1kILx76Ng/edit?gid=443923923#gid=443923923"  # Remplacez par votre lien

def process_schedule_data(schedule_file_path, email_df=None):
    """
    Process the schedule Excel file and create a list of dictionaries with volunteer information.
    
    Args:
        schedule_file_path (str): Path to the Excel file containing schedule data
        email_df (pandas.DataFrame, optional): DataFrame containing volunteer email information
        
    Returns:
        list: List of dictionaries with volunteer information
    """
    # Read the Excel file
    df = pd.read_excel(schedule_file_path)
    
    # Dictionary to store volunteer information
    volunteers = {}
    
    # Time slots are in the column headers, columns 2 onwards
    # Let's be more flexible and get all columns after the second one
    all_columns = df.columns.tolist()
    time_slots = [col for col in all_columns[2:] if col and not col.startswith('Unnamed')]
    
    print(f"Colonnes trouvées: {all_columns}")
    print(f"Créneaux horaires détectés: {time_slots}")
    print(f"Nombre de créneaux: {len(time_slots)}")
    
    # Iterate through the rows to find missions and volunteers
    for row_idx in range(0, df.shape[0]):
        # Skip empty rows
        if row_idx < len(df) and (pd.isna(df.iloc[row_idx, 0]) and all(pd.isna(df.iloc[row_idx, 2:2+len(time_slots)]))):
            continue
            
        # Get the mission name (column 0)
        mission = df.iloc[row_idx, 0]
        
        # Skip rows that are just "Responsable" or configuration rows
        if mission in ['Responsable', 'Responsable benevole', 'Responsable caisse', 'Responsable Artistes']:
            continue
            
        # Skip if no mission name (might be a continuation of previous mission)
        if pd.isna(mission) or mission == '':
            # Try to get the mission from the previous non-empty row
            prev_row = row_idx - 1
            while prev_row >= 0:
                prev_mission = df.iloc[prev_row, 0]
                if not pd.isna(prev_mission) and prev_mission != '' and prev_mission not in ['Responsable', 'Responsable benevole', 'Responsable caisse', 'Responsable Artistes']:
                    mission = prev_mission
                    break
                prev_row -= 1
            
            if pd.isna(mission) or mission == '':
                continue
        
        # Iterate through time slots (starting from column 2)
        for col_idx in range(2, min(len(all_columns), 2 + len(time_slots))):
            # Calculate the time_slot index safely
            time_slot_idx = col_idx - 2
            if time_slot_idx >= len(time_slots):
                continue
                
            time_slot = time_slots[time_slot_idx]
            
            # Skip if time slot is empty or invalid
            if not time_slot or pd.isna(time_slot) or time_slot == '':
                continue
                
            # Get volunteer name at this position
            if col_idx >= len(df.columns):
                continue
                
            volunteer_name = df.iloc[row_idx, col_idx]
            
            # Skip if no volunteer assigned or if it's a number (indicating required people count)
            if pd.isna(volunteer_name) or volunteer_name == '' or is_numeric(volunteer_name):
                continue
                
            # Skip if it's a responsible person name (these appear in some cells but aren't volunteers)
            if volunteer_name in ['Ludal', 'ABDELKRIM BELALEM', 'Lulu', 'Laura', 'Tim Grenier']:
                continue
                
            # Normalize volunteer name (remove extra spaces)
            volunteer_name = str(volunteer_name).strip()
            
            # Add volunteer to dictionary if not already present
            if volunteer_name not in volunteers:
                volunteers[volunteer_name] = {
                    "name": volunteer_name,
                    "email": "",
                    "missions": []
                }
                
            # Add mission and time slot to volunteer
            volunteers[volunteer_name]["missions"].append((mission, time_slot))
    
    # If email dataframe is provided, match emails to volunteers
    if email_df is not None:
        for volunteer_name, volunteer_info in volunteers.items():
            email = find_email_for_volunteer(volunteer_name, email_df)
            volunteer_info["email"] = email
    
    # Convert dictionary to list
    result = list(volunteers.values())
    
    return result

def is_numeric(value):
    """Check if a value is numeric (integer)"""
    try:
        int(value)
        return True
    except (ValueError, TypeError):
        return False

def find_email_for_volunteer(volunteer_name, email_df):
    """
    Find email for a volunteer by matching name in email_df
    
    Args:
        volunteer_name (str): Volunteer name
        email_df (pandas.DataFrame): DataFrame containing email information
        
    Returns:
        str: Email address if found, otherwise empty string
    """
    # You'll need to implement this based on your email_df format
    # This is just a placeholder function
    
    # Example implementation might look like:
    # Normalize the volunteer name
    volunteer_name = volunteer_name.strip()
    
    # Try to split the name into first and last name
    name_parts = volunteer_name.split()
    
    # If the name has at least two parts
    if len(name_parts) >= 2:
        first_name = name_parts[0].strip()
        last_name = " ".join(name_parts[1:]).strip()
        
        # Look for a match in the email_df
        matches = email_df[(email_df["Prénom"].str.strip() == first_name) & 
                           (email_df["NOM"].str.strip() == last_name)]
        
        if not matches.empty:
            return matches.iloc[0]["Adresse mail"]
    
    # If we couldn't find an exact match, try a more flexible approach
    # This will need to be customized based on your specific data
    
    return ""  # Return empty string if no match found

def format_missions(missions):
    """
    Formate la liste des missions pour l'affichage dans l'email
    
    Args:
        missions (list): Liste de tuples (mission, horaire)
        
    Returns:
        str: Missions formatées pour l'email
    """
    if not missions:
        return "Aucune mission assignée pour le moment."
    
    formatted_missions = []
    for mission, horaire in missions:
        formatted_missions.append(f"• {mission} : {horaire}")
    
    return "\n".join(formatted_missions)
def create_email_content(volunteer_info, lien_tableau):
    """
    Crée le contenu de l'email pour un bénévole
    
    Args:
        volunteer_info (dict): Informations du bénévole
        lien_tableau (str): Lien vers le tableau d'échanges
        
    Returns:
        str: Contenu de l'email formaté
    """
    name = volunteer_info['name']
    missions_formatted = format_missions(volunteer_info['missions'])
    
    email_content = f"""Chers bénévoles,

Tout d'abord, un immense merci pour avoir répondu à notre appel ! Vous êtes super, et grâce à vous, ça va être une chouette fête !

Ensuite, un petit désolé ! On a eu du mal à faire notre petit planning, et si vous avez répondu "oui j'adore le travail", il est possible que vous ayez un peu trop d'heures...

Il est possible que de nouveaux bénévoles s'inscrivent d'ici vendredi, donc checkez votre mail samedi matin si jamais.

---

👉 Voici tes missions pour le jour J :

{missions_formatted}

---

🔁 Et le petit tableau pour pouvoir échanger avec tes amis le cas échéant :
{lien_tableau}

---

📋 **DESCRIPTION DES MISSIONS** :

**SERVICE AU BAR**  
Une équipe à la tireuse et une équipe au comptoir.  
Si vous êtes à la tireuse et que vous ne savez pas comment faire des bières sans mousse, n’hésitez pas à demander au responsable !  
Si vous êtes au service, commencez par vous renseigner sur les prix et sur les boissons disponibles.  
On vous paye en cacoin (pas d’argent au bar), qu’il faut mettre dans un carton que la team caisse peut venir chercher de temps en temps.

**VAISSELLE BAR**  
Faire la vaisselle des écocups.  
Si vous avez tout lavé, vous pouvez faire un petit tour dans la fête pour récupérer les écocups abandonnés.

**CATERING**  
Voir avec Laura à la maison de l’Étrange !

**FILTRAGE ACCÈS CATERING**  
Vérifier qu’il n’y ait que des personnes autorisées qui accèdent au catering.

**VENTE CACOIN ET ADHÉSION**  
C’est ici que vous vendez les cacoins ! Un cacoin = un euro.  
Un responsable passera de temps en temps vider la caisse des gros billets.  
Vous pouvez aussi proposer aux gens de laisser de l'argent dans le chapeau !  
Il y aura également la trousse de premiers secours à cet endroit.  
Et des casques antibruit pour les enfants (à échanger contre une pièce d'identité).  
Lorsque vous quittez votre poste, pensez à dire au suivant où sont les pièces d'identité.

**ACCUEIL BÉNÉVOLES**  
Lorsque vous n'avez rien à faire, vous pouvez aider la caisse.  
Sinon, il faut donner aux bénévoles et artistes leurs tickets.  
Pour les questions, rediriger vers Laura la responsable (il paraît qu'elle aura une énorme flèche verte sur la tête).

**TOILETTES**  
Vérifier l'état des toilettes, remettre du papier toilette et jeter un seau d'eau de temps en temps.  
Vérifier que les gens qui se dirigent vers les loges ont un bracelet artiste/bénévole.

**CHAPEAU**  
Votre unique objectif : ramasser plein de moula pour les artistes !  
Gérez votre temps comme vous le sentez…

**ATELIER**  
Trouvez votre atelier et voir avec le responsable !

**FREE SHOP**  
Installer les fringues, animer l'endroit, vérifier qu'il n'y ait pas d'abus (max 5 pièces par personne sauf exception).  
Ce poste, c’est aussi les canards volants : on peut venir vous demander de l'aide sur une autre mission.  
S'il n'y a pas de filtrage catering, vérifiez que les personnes accédant à la zone loge ont un bracelet artiste/bénévole.

**CANARDS COSTAUDS**  
Safe team.

**TECHNIQUE SON**  
Aider l’ingé son à faire ses trucs magiques avec des câbles et des boutons.  
Surveiller le matos à côté du plateau.

---

Encore mille mercis ! Nous t'envoyons des torrents de gratitude !  
Et hâte d'être à samedi 🤩

La team bénévole
"""

    return email_content

def send_email(smtp_server, sender_email, recipient_email, volunteer_info, lien_tableau):
    """
    Envoie un email à un bénévole
    
    Args:
        smtp_server: Serveur SMTP connecté
        sender_email (str): Adresse email d'envoi
        recipient_email (str): Adresse email du destinataire
        volunteer_info (dict): Informations du bénévole
        lien_tableau (str): Lien vers le tableau
        
    Returns:
        bool: True si envoi réussi, False sinon
    """
    try:
        # Créer le message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = "Courée Cacan, un grand merci chers bénévoles ! Voici vos missions"
        
        # Ajouter le contenu
        body = create_email_content(volunteer_info, lien_tableau)
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Envoyer l'email
        text = msg.as_string()
        smtp_server.sendmail(sender_email, recipient_email, text)
        
        return True
    except Exception as e:
        print(f"Erreur lors de l'envoi à {recipient_email}: {str(e)}")
        return False

def send_emails_to_volunteers(volunteers_info, production=PRODUCTION):
    """
    Envoie des emails à tous les bénévoles
    
    Args:
        volunteers_info (list): Liste des informations des bénévoles
        production (bool): Si True, envoie aux vraies adresses, sinon à l'adresse test
    """
    # Configuration Gmail
    gmail_user = GMAIL_USERNAME
    
    # Demander le mot de passe de manière sécurisée
    print("Veuillez entrer votre mot de passe Gmail (ou mot de passe d'application) :")
    gmail_password = os.getenv("GMAIL_PASSWORD")
    
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
            
            if send_email(server, gmail_user, recipient_email, volunteer, LIEN_TABLEAU):
                print(f"✅ {name}: Email envoyé avec succès")
                sent_count += 1
            else:
                print(f"❌ {name}: Échec de l'envoi")
                failed_count += 1
        
        # Fermer la connexion
        server.quit()
        
        # Résumé
        print("\n" + "="*50)
        print("RÉSUMÉ DE L'ENVOI")
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

def preview_email(volunteer_info):
    """
    Prévisualise l'email pour un bénévole
    
    Args:
        volunteer_info (dict): Informations du bénévole
    """
    print("="*60)
    print("PRÉVISUALISATION DE L'EMAIL")
    print("="*60)
    print(f"À: {volunteer_info['email'] if volunteer_info['email'] else 'PAS EMAIL'}")
    print(f"Objet: Courée Cacan, un grand merci chers bénévoles ! Voici vos missions")
    print("-"*60)
    print(create_email_content(volunteer_info, LIEN_TABLEAU))
    print("="*60)


# Example usage:
if __name__ == "__main__":
    # Example schedule data (replace with your actual data)
    schedule_csv = "Tableau rempli(1).xlsx"
    
    # Process the schedule data
    volunteers_info = process_schedule_data(schedule_csv, pd.read_excel('Questionnaire bénévoles 2025 - 19e édition (réponses).xlsx'))
    print(volunteers_info)
    for item in volunteers_info:
        print(f"{item["name"]} a {len(item["missions"])}")
        if item["name"] == "jean christophe bessagnet":
            item["email"]="jisse_2882@yahoo.fr"
        if item["name"] == "Jean p Dellon":
            item["email"]="jphidellon@gmail.com"
        if item["name"] == "Damien Baleux":
            item["email"]="Damien.baleux@gmail.com"
        if item["name"] == "Anaïs":
            item["email"]="anaisoph@gmail.com"
        if item["name"] == "Miran le bg Loof":
            item["email"]="miranlud@hotmail.com"
        if item["name"] == "Jeanne Clément":
            item["email"]="jeanne.clementlp2i@gmail.com"

    with open("volunteers_schedule.txt", "w", encoding="utf-8") as f:
        for volunteer in volunteers_info:
            f.write(f"{volunteer['name']}:\n")
            for mission, time in volunteer['missions']:
                f.write(f"{time}    {mission}\n")
            f.write("\n\n")

    send_emails_to_volunteers(volunteers_info, production=False)