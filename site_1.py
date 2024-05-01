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
from selenium.webdriver.support.select import Select

load_dotenv()


def login(driver: webdriver) -> None:
    """
    Fonction qui permet de se connecter
    :param driver: webdriver
    """
    login = driver.find_element(By.ID, "login_form_customer_code")
    login.send_keys(os.getenv("SITE_1_LOG"))
    time.sleep(2)
    password = driver.find_element(By.ID, "login_form_password")
    password.send_keys(os.getenv("SITE_1_PASS"))
    time.sleep(2)
    password.send_keys(Keys.ENTER)
    time.sleep(15)


def select_saison(driver: webdriver, saison: list) -> None:
    """
    Fonction qui permet de selectionné les saisons souhaitées
    :param driver: webdriver
    :param saison: listes des saisons à selectionner
    """
    saison_ete = driver.find_element(By.ID, "customCheck2")
    if 'été' in saison:
        if saison_ete.is_selected():
            pass
        else:
            driver.execute_script("arguments[0].click();", saison_ete)
            time.sleep(1)
    else:
        if saison_ete.is_selected():
            driver.execute_script("arguments[0].click();", saison_ete)
            time.sleep(1)
    saison_4saison = driver.find_element(By.ID, "customCheck4")
    time.sleep(1)
    if '4 saisons' in saison:
        if saison_4saison.is_selected():
            pass
        else:
            driver.execute_script(
                "arguments[0].click();", saison_4saison)
            time.sleep(1)
    else:
        if saison_4saison.is_selected():
            driver.execute_script(
                "arguments[0].click();", saison_4saison)
            time.sleep(1)
    saison_hiver = driver.find_element(By.ID, "customCheck3")
    time.sleep(1)
    if 'hiver' in saison:
        if saison_hiver.is_selected():
            pass
        else:
            driver.execute_script(
                "arguments[0].click();", saison_hiver)
            time.sleep(1)
    else:
        if saison_hiver.is_selected():
            driver.execute_script(
                "arguments[0].click();", saison_hiver)
            time.sleep(1)


def select_marque(driver: webdriver, marque: list) -> None:
    """
    Fonction permettant de selectionner les marques de la liste
    :param driver: webdriver
    :param marque: liste des marques à rechercher
    """
    try:
        time.sleep(1)
        select_marque = driver.find_element(
            By.CLASS_NAME, 'multiselect__input')
        select_marque.send_keys(marque)
        time.sleep(1)
        select_marque.send_keys(Keys.ENTER)
        time.sleep(1)
        select_marque.send_keys(Keys.ESCAPE)
    except Exception as ex:
        pass


def search_article(driver: webdriver, article: str) -> None:
    """
    Fonction qui permet de rechercher l'article
    :param driver: webdriver
    :param article: article à chercher
    """
    try:
        time.sleep(1)
        search_article = driver.find_element(By.ID, "validationServer01")
        driver.execute_script("arguments[0].value = '';", search_article)
        search_article.clear()
        search_article.send_keys(
            f"{article[0:3] + '/' + f'{article[3:5]}' + 'R' + f'{article[5:7]}'}")
        time.sleep(1)
        if len(article) <= 10:
            charge = Select(driver.find_element(By.ID, "validationServer02"))
            charge.select_by_value(f'{article[8:10]}')
        if len(article) > 10:
            charge = Select(driver.find_element(By.ID, "validationServer02"))
            charge.select_by_value(f'{article[8:11]}')
        time.sleep(1)
        vitesse = Select(driver.find_element(By.ID, "validationServer03"))
        vitesse.select_by_value(f'{article[7]}')
        time.sleep(1)
        search_article.send_keys(Keys.ENTER)
        time.sleep(30)
    except Exception as ex:
        pass


def extract_article_info(article: webdriver, liste_article: list) -> dict:
    """
    Fonction qui extrait les informations des articles selectionné
    :param article: webdriver contenant les articles selectionner
    :param liste_article: listes des articles recherchés pour verification
    return : dictionnaire des informations
    """
    date_du_jour = datetime.now().strftime("%d-%m-%Y")
    # Code
    ind_C = article.find_element(
        By.XPATH, 'div[2]/div/div/div[2]/div/div/ul/li[1]')
    tab_c = ind_C.text.split()
    ind_V = article.find_element(
        By.XPATH, 'div[2]/div/div/div[2]/div/div/ul/li[2]')
    tab_v = ind_V.text.split()
    code = article.find_element(
        By.XPATH, 'div/div/div[1]/div/div[2]/div/div[1]/span')
    text_code = code.text.replace("/", "").replace("-", "").replace(" ", "")
    # Marque
    marque = article.find_element(
        By.XPATH, 'div/div/div[1]/div/div[2]/div/div[1]/p/strong')
    marque_clean = marque.text.replace(' ', '')
    # Prix
    prix = article.find_element(
        By.XPATH, 'div[1]/div/div[4]/div/div[2]/div/div[2]/span[2]')
    # Saison
    saison = article.find_element(
        By.XPATH, 'div[2]/div/div/div[2]/div/div/ul/li[3]')
    saison_tab = saison.text.split()
    if saison_tab[-2] == '4':
        texte_saison = saison_tab[-2] + ' ' + saison_tab[-1]
    else:
        texte_saison = saison_tab[-1]

    profil = article.find_element(
        By.XPATH, 'div/div/div[1]/div/div[2]/div/div[1]/p')
    profil_clean = profil.text.replace(marque_clean + ' ', '')

    article_info = {
        'Code': text_code[0:7] + tab_v[-1] + tab_c[-1],
        'Marque': marque_clean,
        'Profil': profil_clean.upper(),
        'Prix': prix.text.replace('€', ''),
        'Saison': texte_saison,
        'Date': date_du_jour,
        'Site': os.getenv("SITE_1")
    }

    if liste_article != article_info["Code"]:
        return None
    else:
        return article_info


def save_donnees_sql(data: dict, df_base: pd.DataFrame, profil: list) -> None:
    """
    Fonction qui sauvegarde les données dans une base SQL
    :param data: dictionnaire contenant les infos récolté
    :param df_base: dataframe contenant les infos à récolter
    :param profil: profil des pneus recherché
    """
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
                               row['Date'], row['Site'], row['Code_article'], )

            conn.commit()
            cursor.close()
            conn.close()
        except Exception as ex:
            print(ex)
            pass
    except Exception as ex:
        print(ex)
        pass


def site_1_scrap(code: list, df: pd.DataFrame, profil: list, marques: list, saisons: list) -> None:
    """
    Fonction principal qui permet de lancer le scraping
    :param code: Code des articles rechercher
    :param df: DataFrame contenant les infos pour la rechercher
    :param profil: profil de pneu recherché
    :param marques: marques des pneus rechercher
    :param saisons: saisons rechercher
    """
    driver = webdriver.Firefox('')
    driver.get(os.getenv("SITE_1_LOG"))
    time.sleep(2)

    login(driver)

    try:
        driver.find_element(By.XPATH, '/html/body/div[5]/div/div/a').click()
    except Exception as ex:
        pass

    data = []

    for marque in marques:
        select_marque(driver, marque)

    for liste_article in code:
        select_saison(driver, saisons)
        time.sleep(1)
        search_article(driver, liste_article)
        driver.find_element(
            By.XPATH, '/html/body/main/div/div/div[2]/div/div[2]/div[1]/div/div[2]/form/div[3]/label').click()
        time.sleep(2)
        articles = driver.find_elements(
            By.CLASS_NAME, 'row.mx-0.border-top.border-gray.is-tertiary')

        for article in articles:
            article_info = extract_article_info(article, liste_article)
            if article_info:
                data.append(article_info)

    save_donnees_sql(data, df, profil)

    driver.quit()


if __name__ == "__main__":

    chemin_fichier_excel = os.getenv("EXCEL_FILE")
    classeur = openpyxl.load_workbook(chemin_fichier_excel)
    liste_sheets_excel = classeur.sheetnames

    for i in liste_sheets_excel:
        code, profil, marque_abrege, marques, gencode, saisons, df, code_article = excel_file(chemin_fichier_excel, i)

        site_1_scrap(code, df, profil, marques, saisons)