import os
import time
import smtplib
import openpyxl
from dotenv import load_dotenv
from site_1 import site_1_scrap
from site_2 import site_2_scrap
from site_3 import site_3_scrap
from site_4 import site_4_scrap
from site_5 import site_5_scrap
from excel_file import excel_file
from email.mime.text import MIMEText
from donnees_de_comparaison import donnees_de_comparaison
from email.mime.multipart import MIMEMultipart

load_dotenv()


def send_email(subject, content) -> None:
    sender_email = os.getenv("SENDER_EMAIL")
    sender_password = os.getenv("SENDER_PASSWORD")
    recipient_email = os.getenv("EMAIL_RECIPIENT")
    subject = f"Scraping"

    message = MIMEMultipart()
    message.attach(MIMEText(content, 'plain'))
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = recipient_email

    outlook_smtp_server = "smtp.office365.com"
    outlook_smtp_port = 587

    with smtplib.SMTP_SSL(outlook_smtp_server, outlook_smtp_port) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, message.as_string())


chemin_fichier_excel = os.getenv("EXCEL_FILE")
classeur = openpyxl.load_workbook(chemin_fichier_excel)
liste_sheets_excel = classeur.sheetnames

for i in liste_sheets_excel:

    code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(
        chemin_fichier_excel, i)

    try:
        site_1_scrap(code, df, profil, marques, saisons)
        time.sleep(5)
    except Exception as ex:
        error_message = str(ex)
        print("Erreur scrap web scrap web site 1", error_message, code)
        send_email("Erreur scrap web site 1", error_message)

for i in liste_sheets_excel:

    code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)

    try:
        site_2_scrap(df, code_article)
        time.sleep(5)
    except Exception as ex:
        error_message = str(ex)
        print("Erreur scrap web site 2", error_message, code)
        send_email("Erreur scrap web site 2", error_message)

for i in liste_sheets_excel:

    code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)
    try:
        site_3_scrap(df, marques, gencode)
        time.sleep(5)
    except Exception as ex:
        error_message = str(ex)
        print("Erreur scrap web gettygo_scrap", error_message, code)
        send_email("Erreur scrap web gettygo_scrap", error_message)

for i in liste_sheets_excel:

    code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)
    try:
        site_4_scrap(df, marques, gencode)
        time.sleep(5)
    except Exception as ex:
        error_message = str(ex)
        print("Erreur scrap web site 4", error_message, code)
        send_email("Erreur scrap web site 4", error_message)

for i in liste_sheets_excel:

    code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)
    try:
        site_5_scrap(code, df, marques, saisons, profil)
        time.sleep(5)
    except Exception as ex:
        error_message = str(ex)
        print("Erreur scrap web site 5", error_message, code)
        send_email("Erreur scrap web site 5", error_message)

try:
    donnees_de_comparaison()
except Exception as ex:
    error_message = str(ex)
    print("Erreur scrap web donnees_chrono", error_message, code)
    send_email("Erreur scrap web donnees_chrono", error_message)
