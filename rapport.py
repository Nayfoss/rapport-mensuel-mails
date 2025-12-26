import imaplib
import email
from email.header import decode_header
import csv
import smtplib
from email.message import EmailMessage
import os

# Secrets GitHub
EMAIL_USER = os.environ["EMAIL_USER"]
EMAIL_PASSWORD = os.environ["EMAIL_PASSWORD"]
IMAP_SERVER = os.environ["IMAP_SERVER"]
SMTP_SERVER = os.environ["SMTP_SERVER"]
SMTP_PORT = int(os.environ["SMTP_PORT"])

# Les 3 objets que tu veux filtrer
TARGET_SUBJECTS = [
    "Bon de Don – Association AUBE Ait Bouyahia",
    "Nouvelle demande d'aide",
    "Nouvelle inscription bénévole"
]

def decode_subject(raw_subject):
    subject, encoding = decode_header(raw_subject)[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8")
    return subject

def read_sent_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASSWORD)

    # Ouvrir le dossier "Sent"
    mail.select('"[Gmail]/Sent Mail"')

    results = {}

    for subject in TARGET_SUBJECTS:

        # Recherche IMAP simple (Gmail-friendly)
        status, messages = mail.search(None, "ALL")
        mail_ids = messages[0].split()

        rows = []

        for mail_id in mail_ids:
            status, msg_data = mail.fetch(mail_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            decoded_subject = decode_subject(msg["Subject"])

            # Filtrer en Python (car IMAP ne supporte pas UTF‑8)
            if subject not in decoded_subject:
                continue

            date = msg["Date"]

            # Récupérer le contenu
            if msg.is_multipart():
                content = ""
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        content += part.get_payload(decode=True).decode("utf-8", errors="ignore")
            else:
                content = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

            rows.append([date, decoded_subject, content])

        results[subject] = rows

    mail.logout()
    return results



def generate_csv(data):
    filenames = []

    for subject, rows in data.items():
        filename = f"{subject.replace(' ', '_')}.csv"
        filenames.append(filename)

        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Date", "Objet", "Contenu"])
            writer.writerows(rows)

    return filenames

def send_email_with_csv(files):
    msg = EmailMessage()
    msg["Subject"] = "Rapport mensuel"
    msg["From"] = EMAIL_USER
    msg["To"] = EMAIL_USER
    msg.set_content("Voici les rapports mensuels en pièces jointes.")

    for file in files:
        with open(file, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="text",
                subtype="csv",
                filename=file
            )

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def main():
    data = read_sent_emails()
    files = generate_csv(data)
    send_email_with_csv(files)

if __name__ == "__main__":
    main()
