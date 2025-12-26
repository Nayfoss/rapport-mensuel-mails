import imaplib
import email
from email.header import decode_header
import csv
import smtplib
from email.message import EmailMessage
import os
from pdfminer.high_level import extract_text
import re

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
    if raw_subject is None:
        return ""
    subject, encoding = decode_header(raw_subject)[0]
    if isinstance(subject, bytes):
        subject = subject.decode(encoding or "utf-8", errors="ignore")
    return subject

def read_sent_emails():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_USER, EMAIL_PASSWORD)

    mail.select('"[Gmail]/Sent Mail"')

    results = {}

    status, messages = mail.search(None, "ALL")
    mail_ids = messages[0].split()

    for subject in TARGET_SUBJECTS:
        rows = []

        for mail_id in mail_ids:
            status, msg_data = mail.fetch(mail_id, "(RFC822)")
            msg = email.message_from_bytes(msg_data[0][1])

            decoded_subject = decode_subject(msg["Subject"])
            if subject not in decoded_subject:
                continue

            date = msg["Date"]

            # Contenu texte
            content = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        content += part.get_payload(decode=True).decode("utf-8", errors="ignore")
            else:
                content = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

            if subject == "Nouvelle demande d'aide":
                nom, prenom, tel, message = parse_demande_aide(content)
                rows.append([date, nom, prenom, tel, message])

            elif subject == "Nouvelle inscription bénévole":
                nom, prenom, tel, groupe, aides, autre = parse_inscription_benevole(content)
                rows.append([date, nom, prenom, tel, groupe, aides, autre])

            elif subject == "Bon de Don – Association AUBE Ait Bouyahia":
                pdf_path = None
                for part in msg.walk():
                    if part.get_content_type() == "application/pdf":
                        pdf_path = f"bon_{mail_id.decode()}.pdf"
                        with open(pdf_path, "wb") as f:
                            f.write(part.get_payload(decode=True))

                if pdf_path:
                    numero, nom_benef, prenom_benef, nom_donateur, prenom_donateur, biens = parse_bon_de_don_pdf(pdf_path)
                    rows.append([date, numero, nom_benef, prenom_benef, nom_donateur, prenom_donateur, biens])
                    os.remove(pdf_path)

        results[subject] = rows

    mail.logout()
    return results

def parse_bon_de_don_pdf(pdf_path):
    text = extract_text(pdf_path)

    # Nettoyage encodage
    text = text.replace("Ã©", "é").replace("Ã¨", "è").replace("Ã", "à").replace("â€™", "'")

    # Numéro du bon
    numero = re.search(r"Num[eé]ro du bon\s*:\s*(.*)", text)
    numero = numero.group(1).strip() if numero else ""

    # Fonction pour récupérer Nom / Prénom d'une section
    def extract_nom_prenom(section_name):
        pattern = rf"{section_name}.*?Nom\s*:\s*(.*?)(?:\n|$).*?Pr[eé]nom\s*:\s*(.*?)(?:\n|$)"
        m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if m:
            return m.group(1).strip(), m.group(2).strip()
        return "", ""

    donateur_nom, donateur_prenom = extract_nom_prenom("DONATEUR")
    beneficiaire_nom, beneficiaire_prenom = extract_nom_prenom("BÉNÉFICIAIRE")

    # Biens donnés
    biens_match = re.search(r"Bien\s*\(s\)\s*:\s*(.*)", text)
    biens = biens_match.group(1).strip() if biens_match else ""

    return [
        numero,
        beneficiaire_nom,
        beneficiaire_prenom,
        donateur_nom,
        donateur_prenom,
        biens
    ]



def parse_demande_aide(content):
    content = content.replace("Ã©", "é").replace("Ã¨", "è").replace("Ã", "à").replace("â€™", "'")

    nom = re.search(r"Nom:\s*(.*)", content)
    prenom = re.search(r"Pr[ée]nom:\s*(.*)", content)
    tel = re.search(r"T[ée]l[ée]phone:\s*(.*)", content)
    message = re.search(r"Message:\s*(.*)", content)

    return [
        nom.group(1).strip() if nom else "",
        prenom.group(1).strip() if prenom else "",
        tel.group(1).strip() if tel else "",
        message.group(1).strip() if message else ""
    ]

def parse_inscription_benevole(content):
    content = (
        content.replace("Ã©", "é")
               .replace("Ã¨", "è")
               .replace("Ã", "à")
               .replace("â€™", "'")
               .replace("\r", "")
               .strip()
    )

    nom = re.search(r"Nom:\s*(.*)", content)
    prenom = re.search(r"Pr[ée]nom:\s*(.*)", content)
    tel = re.search(r"T[ée]l[ée]phone:\s*(.*)", content)
    groupe = re.search(r"Groupe sanguin:\s*(.*)", content)
    aides = re.search(r"Aides proposées:\s*(.*?)(?:Autre:|$)", content, re.DOTALL)
    autre = re.search(r"Autre:\s*(.*)", content)

    return [
        nom.group(1).strip() if nom else "",
        prenom.group(1).strip() if prenom else "",
        tel.group(1).strip() if tel else "",
        groupe.group(1).strip() if groupe else "",
        aides.group(1).strip().replace("\n", " ") if aides else "",
        autre.group(1).strip() if autre else ""
    ]

def generate_csv(data):
    filenames = []

    for subject, rows in data.items():
        filename = f"{subject.replace(' ', '_').replace('–', '-')}.csv"
        filenames.append(filename)

        with open(filename, "w", newline="", encoding="utf-8") as f:
            # Ajouter delimiter=';' et quoting pour bien séparer les colonnes
            writer = csv.writer(f, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)

            if subject == "Nouvelle inscription bénévole":
                writer.writerow(["Date", "Nom", "Prénom", "Téléphone", "Groupe sanguin", "Aides proposées", "Autre"])

            elif subject == "Nouvelle demande d'aide":
                writer.writerow(["Date", "Nom", "Prénom", "Téléphone", "Message"])

            elif subject == "Bon de Don – Association AUBE Ait Bouyahia":
                writer.writerow(["Date", "Numéro bon", "Nom bénéficiaire", "Prénom bénéficiaire", "Nom donateur", "Prénom donateur", "Biens"])

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
            msg.add_attachment(f.read(), maintype="text", subtype="csv", filename=file)

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASSWORD)
        smtp.send_message(msg)

def main():
    data = read_sent_emails()
    files = generate_csv(data)
    send_email_with_csv(files)

if __name__ == "__main__":
    main()
