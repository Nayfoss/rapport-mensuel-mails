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
    "Bon de Don Association AUBE Ait Bouyahia",
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

    results = {
        "Nouvelle demande d'aide": [],
        "Nouvelle inscription bénévole": [],
        "Bon de Don – Association AUBE Ait Bouyahia": []
    }

    status, messages = mail.search(None, "ALL")
    mail_ids = messages[0].split()

    for mail_id in mail_ids:
        status, msg_data = mail.fetch(mail_id, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        subject = decode_subject(msg["Subject"])
        date = msg["Date"]

        treated = False  # 🔐 sécurité

        # Récupérer le texte
        content = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    content += part.get_payload(decode=True).decode("utf-8", errors="ignore")
        else:
            content = msg.get_payload(decode=True).decode("utf-8", errors="ignore")

        if "Nouvelle demande d'aide" in subject:
            row = parse_demande_aide(content)
            results["Nouvelle demande d'aide"].append([date] + row)
            treated = True

        elif "Nouvelle inscription bénévole" in subject:
            row = parse_inscription_benevole(content)
            results["Nouvelle inscription bénévole"].append([date] + row)
            treated = True

        elif "Bon de Don – Association AUBE Ait Bouyahia" in subject:
            pdf_path = None
            for part in msg.walk():
                if part.get_content_type() == "application/pdf":
                    pdf_path = f"bon_{mail_id.decode()}.pdf"
                    with open(pdf_path, "wb") as f:
                        f.write(part.get_payload(decode=True))

            if pdf_path:
                row = parse_bon_de_don_pdf(pdf_path)
                results["Bon de Don – Association AUBE Ait Bouyahia"].append([date] + row)
                os.remove(pdf_path)
                treated = True

        # ✅ supprimer UNIQUEMENT si traité
        if treated:
            mail.store(mail_id, '+FLAGS', '\\Deleted')

    mail.expunge()
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


from datetime import datetime, timedelta

def get_previous_month():
    today = datetime.today()
    first_day_this_month = today.replace(day=1)
    last_day_previous_month = first_day_this_month - timedelta(days=1)

    year = last_day_previous_month.year
    month = last_day_previous_month.month

    return year, month

REPORT_RECEIVER = "aubeaitbouyahia09@gmail.com" 
def send_email_with_csv(files):
    year, month = get_previous_month()

    msg = EmailMessage()
    msg["Subject"] = f"Rapport mensuel {year}_{month:02d}"
    msg["From"] = EMAIL_USER
    msg["To"] = REPORT_RECEIVER

    msg.set_content(
        f"Bonjour,\n\n"
        f"Veuillez trouver ci-joint le rapport du mois {month:02d}/{year}.\n\n"
        f"Cordialement."
    )

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
