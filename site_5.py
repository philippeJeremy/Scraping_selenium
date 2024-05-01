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
from selenium.webdriver.support.select import Select

load_dotenv()


def select_marque(driver, marques):
    driver.find_element(
        By.XPATH, '//*[@id="home-selector"]/form/div[2]/div[8]/div[1]').click()
    time.sleep(2)

    for i in marques:
        try:
            driver.find_element(By.XPATH, f'//label[text()="{i}"]').click()
            time.sleep(2)
        except Exception as ex:
            pass

    driver.find_element(
        By.XPATH, '//*[@id="selector-marque"]/div[4]/span').click()
    time.sleep(2)


def search_article(driver, liste_article, saisons):
    
    if saisons == 'été':
        saison = Select(driver.find_element(By.XPATH, '//select[@name="s"]'))
        saison.select_by_value('S')
        time.sleep(2)
    if saisons == '4 saisons':
        saison = Select(driver.find_element(By.XPATH, '//select[@name="s"]'))
        saison.select_by_value('G')
        time.sleep(2)
    largeur = Select(driver.find_element(By.XPATH, '//select[@name="l"]'))
    largeur.select_by_value(f'{liste_article[0:3]}')
    time.sleep(2)
    hauteur = Select(driver.find_element(By.XPATH, '//select[@name="h"]'))
    hauteur.select_by_value(f'{liste_article[3:5]}')
    time.sleep(2)
    diametre = Select(driver.find_element(By.XPATH, '//select[@name="d"]'))
    diametre.select_by_value(f'{liste_article[5:7]}')
    time.sleep(2)
    driver.find_element(
        By.CLASS_NAME, 'upper.btn.btn-danger.btn-block.btn-large').click()
    time.sleep(5)
    if len(liste_article) == 10:
        charge = Select(driver.find_element(By.XPATH, '//select[@name="c"]'))
        charge.select_by_value(f'{liste_article[8:10]}')
        time.sleep(2)
    else:
        charge = Select(driver.find_element(By.XPATH, '//select[@name="c"]'))
        charge.select_by_value(f'{liste_article[8:11]}')
        time.sleep(2)
    vitesse = Select(driver.find_element(By.XPATH, '//select[@name="v"]'))
    vitesse.select_by_value(f'{liste_article[7:8]}')
    time.sleep(2)
    driver.find_element(
        By.CLASS_NAME, 'upper.btn.btn-danger.btn-block.btn-large').click()


def extract_article(artilce, liste_article):
    date_du_jour = datetime.now().strftime("%d-%m-%Y")
    code = artilce.find_element(By.XPATH, 'td[3]/a/small')
    clean_code = code.text.split('\n')
    split_code = clean_code[0].split(' ')
    digits = ""
    letters = ""

    for char in split_code[2]:
        if char.isdigit():
            digits += char
        elif char.isalpha():
            letters += char

    code_clean = split_code[0].replace(
        '/', '') + split_code[1] + letters + digits

    marque = artilce.find_element(By.XPATH, 'td[3]/a/span[1]')
    designation = artilce.find_element(By.XPATH, 'td[3]/a/span[2]')
    profil = designation.text.replace(marque.text + ' ', '')
    prix = artilce.find_element(By.XPATH, 'td[6]/div/span[1]')
    saison = artilce.find_element(By.XPATH, 'td[5]')

    article_info = {
        'Code': code_clean.replace('R', ''),
        'Marque': marque.text,
        'Profil': profil.upper(),
        'Prix': round(float(prix.text.replace(',', '.').replace('€', '')) / 1.2, 2),
        'Saison': saison.text.replace('Tourisme ', ''),
        'Date': date_du_jour,
        'Site': os.getenv("SITE_5")
    }

    if liste_article != article_info["Code"]:
        return None
    else:
        return article_info


def save_donnees_sql(data, df_base, profil):
    
    str_connection_string = (
        "DRIVER={SQL Server};"
        f"SERVER={os.getenv('HOST')};"
        f"DATABASE={os.getenv('DATABASE')};"
        f"UID={os.getenv('ID')};"
        f"PWD={os.getenv('PSWD')};"
                            )
    
    data = [item for item in data if item is not None]
    df = pd.DataFrame(data)
    df.replace("", pd.NA, inplace=True)
    df.replace("été", "Eté", inplace=True)
    df.dropna(inplace=True)
    df = df[df['Profil'].isin(profil)]
    try:
        correspondance_profil_code_article = dict(
            zip(df_base['Profil'], df_base['Code_article']))
        df['Code_article'] = df['Profil'].map(correspondance_profil_code_article)
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


def site_5_scrap(code, df, marques, saisons, profil):

    # connexion au site web
    driver = webdriver.Firefox()
    driver.get(os.getenv('SITE_5'))
    time.sleep(5)

    data = []

    driver.find_element(By.XPATH, '/html/body/div[8]/div[1]/a').click()
    time.sleep(5)

    select_marque(driver, marques)

    for liste_article in code:
        search_article(driver, liste_article, saisons)
        articles = driver.find_elements(By.CLASS_NAME, 'tr-to-product')
        for artilce in articles:
            article_info = extract_article(artilce, liste_article)
            data.append(article_info)

    save_donnees_sql(data, df, profil)

    driver.quit()


if __name__ == "__main__":

    chemin_fichier_excel = os.getenv("EXCEL_FILE")
    classeur = openpyxl.load_workbook(chemin_fichier_excel)
    liste_sheets_excel = classeur.sheetnames

    for i in liste_sheets_excel:
        code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(
                chemin_fichier_excel, i)
    
        site_5_scrap(code, df, marques, saisons, profil)
