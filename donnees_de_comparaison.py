import pandas as pd
import openpyxl
import pyodbc
import os
from datetime import datetime
from dotenv import load_dotenv
from requete import requete

load_dotenv()

def donnees_de_comparaison():
    chemin_fichier_excel = os.getenv('EXCEL_FILE')
    classeur = openpyxl.load_workbook(chemin_fichier_excel)
    liste_sheets_excel = classeur.sheetnames

    df_final = pd.DataFrame()

    for i in liste_sheets_excel:
        taille = pd.read_excel(chemin_fichier_excel, sheet_name=i)
        code_article = taille['Code_article'].unique().tolist()
        df = requete(code_article)
        df['Code'] = taille['Code']
        df_final = pd.concat([df_final, df], ignore_index=True)

    df_final['Date'] = datetime.now().strftime("%d-%m-%Y")
    df_final['Site'] = 'Chrono'
    df = df_final[['Code', 'Marque', 'Profil', 'Prix', 'Saison', 'Date', 'Site', 'Code_article']]

    HOST = os.getenv('HOSTSIMON')
    ID = os.getenv('IDSIMON')
    DATABASE = os.getenv('DATABASESIMON')
    PSWD = os.getenv('PSWDSIMON')
        
    str_connection_string = (
        "DRIVER={SQL Server};"
        f"SERVER={HOST};"
        f"DATABASE={DATABASE};"
        f"UID={ID};"
        f"PWD={PSWD};")
        
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


if __name__ == "__main__":
    donnees_chrono()
