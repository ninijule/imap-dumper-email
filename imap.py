import imaplib
import email
from email.header import decode_header
import os
import re

# Informations de connexion (à remplacer par vos propres informations)
IMAP_SERVER = 'imap.example.com'
EMAIL_ACCOUNT = 'your_email@example.com'
PASSWORD = 'your_password'

def decode_folder_name(encoded_name):
    decoded_bytes, encoding = decode_header(encoded_name)[0]
    if isinstance(decoded_bytes, bytes):
        return decoded_bytes.decode(encoding or 'utf-8')
    return decoded_bytes

def encode_folder_name(folder_name):
    if any(ord(char) > 127 for char in folder_name):
        return folder_name.encode('imap_utf7').decode('ascii')
    return folder_name

# Fonction pour télécharger les pièces jointes
def download_attachments(msg, download_folder):
    if not os.path.isdir(download_folder):
        os.makedirs(download_folder)

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if filename:
            filepath = os.path.join(download_folder, filename)
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f"Téléchargé: {filepath}")

# Connexion au serveur IMAP
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL_ACCOUNT, PASSWORD)

# Récupération de tous les dossiers
status, folders = mail.list()
if status == 'OK':
    print("Dossiers:")
    for folder in folders:
        folder_info = folder.decode()
        match = re.search(r'\((.*?)\) "(.*?)" (.*)', folder_info)
        if match:
            folder_name = match.group(3).strip('"')
            folder_name = decode_folder_name(folder_name)
            print(folder_name)

            # Encode the folder name correctly for IMAP
            folder_name = encode_folder_name(folder_name)

            # Sélectionner le dossier
            status, data = mail.select('"%s"' % folder_name)
            if status == 'OK':
                # Recherche de tous les emails dans le dossier sélectionné
                status, messages = mail.search(None, 'ALL')
                if status == 'OK':
                    mail_ids = messages[0].split()
                    print(f"\nDossier: {folder_name}")
                    print(f"Nombre d'emails: {len(mail_ids)}")

                    # Parcours des IDs de mails et affichage des informations de l'en-tête et du corps
                    for mail_id in mail_ids:
                        status, msg_data = mail.fetch(mail_id, '(RFC822)')
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])

                                # Décode l'en-tête du sujet en toute sécurité
                                subject = msg['subject']
                                if subject:
                                    subject = decode_header(subject)[0][0]
                                    if isinstance(subject, bytes):
                                        subject = subject.decode()

                                from_ = msg['from']
                                to = msg['to']
                                date = msg['date']

                                print(f"From: {from_}")
                                print(f"To: {to}")
                                print(f"Subject: {subject}")
                                print(f"Date: {date}")

                                # Télécharger les pièces jointes
                                download_attachments(msg, "attachments")

                                # Extraire le corps du message
                                if msg.is_multipart():
                                    for part in msg.walk():
                                        if part.get_content_type() == "text/plain":
                                            body = part.get_payload(decode=True).decode()
                                            print(f"Body:\n{body}")
                                else:
                                    body = msg.get_payload(decode=True).decode()
                                    print(f"Body:\n{body}\n")

                else:
                    print(f"Erreur lors de la recherche des emails dans le dossier {folder_name}")
            else:
                print(f"Erreur lors de la sélection du dossier {folder_name}")
else:
    print("Impossible de récupérer les dossiers.")

# Déconnexion
mail.logout()
