import os
import time
import pyodbc
import openpyxl
import pandas as pd
from excel_file import excel_file
from requete import requete
from datetime import datetime
from selenium import webdriver
from dotenv import load_dotenv
from sqlalchemy import create_engine
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

load_dotenv()


def login(driver):
    login_elem = driver.find_element(By.NAME, "login")
    login_elem.send_keys(os.getenv('site_2_log'))

    password_elem = driver.find_element(By.NAME, "password")
    password_elem.send_keys(os.getenv('site_2_pass'))
    password_elem.send_keys(Keys.ENTER)


def search_article(driver, liste_article):
    while True:
        try:
            search_article = driver.find_element(By.ID, "code_article_input")
            search_article.clear()
            search_article.send_keys(liste_article)
            time.sleep(2)
            driver.find_element(By.ID, 'envoie_formulaire').click()
            time.sleep(15)
            driver.find_element(By.XPATH, '//*[@id="loader"]/div/p/a').click()
            time.sleep(2)
            driver.refresh()
        except Exception as ex:
            break


def extract_article_info(article, liste_article):
    date_du_jour = datetime.now().strftime("%d-%m-%Y")
    try:
        marque = article.find_element(
            By.XPATH, 'td[9]/img').get_attribute('title')
        prix = article.find_element(By.XPATH, 'td[16]/span/b')
        saison = article.find_element(
            By.XPATH, 'td[29]/b/img').get_attribute('title')
        code_article = article.find_element(By.XPATH, 'td[10]')
        article_info = {
            'Marque': marque.capitalize(),
            'Code_article': code_article.text,
            'Prix': prix.text,
            'Saison': saison,
            'Date': date_du_jour,
            'Site': 'Districash'
        }
        if liste_article != article_info["Code_article"]:
            return None
        else:
            return article_info
    except Exception as ex:
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
    df.replace("été / hiver", '4 Saisons', inplace=True)
    df.replace("été", 'Eté', inplace=True)
    try:
        correspondance_profil = dict(
            zip(df_base['Code_article'], df_base['Profil']))
        correspondance_code = dict(
            zip(df_base['Code_article'], df_base['Code']))
        df['Profil'] = df['Code_article'].map(correspondance_profil)
        df['Code'] = df['Code_article'].map(correspondance_code)
        df.replace("", pd.NA, inplace=True)
        df.dropna(inplace=True)
        try:
            conn = pyodbc.connect(str_connection_string)
            conn.timeout = 0
            cursor = conn.cursor()
            cursor.execute("set transaction isolation level read uncommitted")

            for index, row in df.iterrows():
                cursor.execute("""
                                INSERT INTO price_tracking (Code, Marque, Profil, Prix,	Saison,	Date, Site, Code_article )
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                                """, row['Code'], row['Marque'], row['Profil'], row['Prix'], row['Saison'],
                                    row['Date'], row['Site'], row['Code_article'])

            conn.commit()
            cursor.close()
            conn.close()
        except Exception as ex:
            pass
    except Exception as ex:
        pass


def site_2_scrap(df, code_article):
    driver = webdriver.Firefox('')
    driver.get(os.getenv('URL_SITE_2'))

    time.sleep(2)
    login(driver)
    time.sleep(10)

    data = []

    for liste_article in code_article:
        search_article(driver, liste_article)
        articles = driver.find_elements(
            By.XPATH, "//*[@id='TABLEAU_RECHERCHE_ARTICLE']/tbody/tr")

        for article in articles:
            article_info = extract_article_info(article, liste_article)

            if article_info:
                data.append(article_info)
            else:
                pass
        
    save_donnees_sql(data, df)

    driver.quit()


if __name__ == "__main__":
    chemin_fichier_excel = os.getenv("EXCEL_FILE")
    classeur = openpyxl.load_workbook(chemin_fichier_excel)
    liste_sheets_excel = classeur.sheetnames

    for i in liste_sheets_excel:
        code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)
            
        site_2_scrap(df, code_article)
