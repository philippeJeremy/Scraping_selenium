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
    username_input = driver.find_element(By.ID, "nx-login-form-username")
    username_input.send_keys(os.getenv('SITE_3_LOG'))
    password_input = driver.find_element(By.ID, "nx-login-form-password")
    password_input.send_keys(os.getenv('SITE_3_PASS'))
    password_input.send_keys(Keys.ENTER)


def search_article(driver, liste_article):
    try:
        time.sleep(1)
        reset = driver.find_element(
            By.XPATH, '//*[@id="nx-main-background-container"]/div/div/div/section/div/nx-tyre-quick-search/div/nx-tyre-quick-options/div/div[2]/div[1]/div[3]/nx-button[2]/div/button/span')
        reset.click()
        search_article = driver.find_element(
                    By.ID, "nx-tyre-quick-options-matchcode")
        search_article.send_keys(liste_article)
        time.sleep(2)
        search_article.send_keys(Keys.ENTER)

        time.sleep(20)
    except Exception as ex:
        print(ex)


def extract_article_info(article, liste_article, marques):
    try:
        date_du_jour = datetime.now().strftime("%d-%m-%Y")
        marque = article.find_element(
            By.XPATH, 'div[1]/div/div[1]/div/div[2]/div[1]/b')
        prix = article.find_element(
            By.XPATH, 'div[3]/div/div[2]/div/div[2]/div[1]/div[2]/div')
        saison = article.find_element(
            By.XPATH, 'div[1]/div/div[1]/div/div[2]/div[3]')

        article_info = {
            'gencode': liste_article,
            'Marque': marque.text.capitalize(),
            'Prix': prix.text.replace('€', '').replace(',', '.'),
            'Saison': saison.text.replace('VL', '').replace(' ', ''),
            'Date': date_du_jour,
            'Site': os.getenv('SITE_3'),
        }
        if marques != article_info["Marque"]:
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
    df.replace('ÉTÉ', 'Eté', inplace=True)
    df.replace('4SAISONS', '4 Saisons', inplace=True)

    try:
        correspondance_code_article = dict(
            zip(df_base['gencode'], df_base['Code_article']))
        df['Code_article'] = df['gencode'].map(correspondance_code_article)
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
        pass


def site_3_scrap(df, marques, gencode):

    driver = webdriver.Firefox()
    driver.get(os.getenv('SITE_3_URL'))
    time.sleep(2)

    login(driver)
    time.sleep(15)

    data = []
    i = 0
    for liste_article in gencode:
        try:
            search_article(driver, liste_article)
            articles = driver.find_elements(
                By.CLASS_NAME, "row.nx-table-body.nx-table-body-no-alternating-lines")

            if articles != None:

                for article in articles:
                    try:
                        article_info = extract_article_info(
                            article, liste_article, marques[i])
                        data.append(article_info)
                    except Exception as ex:
                        print(ex)

                        pass
            i += 1

        except Exception as ex:
            print(ex)
            break

    save_donnees_sql(data, df)

    driver.quit()


if __name__ == "__main__":

    chemin_fichier_excel = os.getenv("EXCEL_FILE")
    classeur = openpyxl.load_workbook(chemin_fichier_excel)
    liste_sheets_excel = classeur.sheetnames

    for i in liste_sheets_excel:
        code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)
    
        site_3_scrap(df, marques, gencode)
