import os
import time
import pyodbc
import openpyxl
import pandas as pd
from datetime import datetime
from selenium import webdriver
from dotenv import load_dotenv
from excel_file import excel_file
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

load_dotenv()


def login(driver):
    driver.find_element(
        By.XPATH, '//*[@id="alzura-cookie-consent"]/div/div/div/div[2]/a[1]').click()
    time.sleep(2)

    login = driver.find_element(By.NAME, "userid")
    login.send_keys(os.getenv('TYRE24_LOG'))
    time.sleep(1)

    password = driver.find_element(By.NAME, "password")
    password.send_keys(os.getenv('TRYE24_PASS'))
    time.sleep(1)

    password.send_keys(Keys.ENTER)
    time.sleep(10)


def extract_article_info(article, liste_article, marque):

    try:
        date_du_jour = datetime.now().strftime("%d-%m-%Y")
        prix = article.find_element(
            By.XPATH, 'div[5]/div[2]/div/div/span')
        saison = article.find_element(
            By.XPATH, 'div[4]/div[2]/div')
        saison_clean = saison.text.split(' ', 1)
        prix = prix.text.replace('€', '').replace(',', '.')

        article_info = {
            'gencode': liste_article,
            'Prix': round(float(prix) / 1.2, 2),
            'Saison': saison_clean[1],
            'Date': date_du_jour,
            'Site': 'Tyre24'
        }

        return article_info

    except Exception as ex:
        print(ex)
        return None


def save_donnees_sql(data, df_base):

    str_connection_string = (
        "DRIVER={SQL Server};"
        f"SERVER={os.getenv('HOST')};"
        f"DATABASE={os.getenv('DATABASE')};"
        f"UID={os.getenv('ID')};"
        f"PWD={os.getenv('PSWD')};"
    )

    data = [item for item in data if item is not None]

    df = pd.DataFrame(data)
    df.replace('été', 'Eté', inplace=True)

    try:
        correspondance_code_article = dict(
            zip(df_base['gencode'], df_base['Code_article']))
        df['Code_article'] = df['gencode'].map(correspondance_code_article)
        correspondance_marque = dict(
            zip(df_base['gencode'], df_base['Marque']))
        df['Marque'] = df['gencode'].map(correspondance_marque)
        correspondance_profil = dict(
            zip(df_base['gencode'], df_base['Profil']))
        df['Profil'] = df['gencode'].map(correspondance_profil)
        correspondance_profil = dict(
            zip(df_base['gencode'], df_base['Code']))
        df['Code'] = df['gencode'].map(correspondance_profil)
        df.drop('gencode', axis=1, inplace=True)
        try:
            conn = pyodbc.connect(str_connection_string)
            conn.timeout = 0
            cursor = conn.cursor()
            cursor.execute("set transaction isolation level read uncommitted")

            for index, row in df.iterrows():
                cursor.execute("""
                        INSERT INTO price_tracking (Code, Marque, Profil, Prix,	Saison,	Date, Site, Code_article)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                                """, row['Code'], row['Marque'], row['Profil'], row['Prix'], row['Saison'],
                                    row['Date'], row['Site'], row['Code_article'])

            conn.commit()
            cursor.close()
            conn.close()
        except Exception as ex:
            pass
    except Exception as ex:
        print(ex)
        pass


def site_4_scrap(df, marques, gencode):

    driver = webdriver.Firefox()
    driver.get(os.getenv("URL_SITE_4"))
    time.sleep(2)

    login(driver)

    data = []
    i = 0
    for liste_article in gencode:
        driver.find_element(
            By.XPATH, '//*[@id="megaMenu"]/ul/li[1]/a').click()
        time.sleep(5)
        search_article = driver.find_element(
            By.XPATH, '//*[@id="vs1__combobox"]/div[1]/input')
        search_article.send_keys(liste_article)
        time.sleep(1)
        search_article.send_keys(Keys.ENTER)
        time.sleep(7)
        articles = driver.find_elements(
            By.CLASS_NAME, 'item')
        for article in articles:
            article_info = extract_article_info(article, liste_article, marques[i])
            data.append(article_info)

        i += 1

    save_donnees_sql(data, df)

    driver.quit()


if __name__ == "__main__":

    chemin_fichier_excel = os.getenv("EXCEL_FILE")
    classeur = openpyxl.load_workbook(chemin_fichier_excel)
    liste_sheets_excel = classeur.sheetnames

    for i in liste_sheets_excel:
        code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)
    
        site_4_scrap(df, marques, gencode)
