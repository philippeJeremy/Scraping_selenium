import os
import pyodbc
import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine

load_dotenv()


def requete(code_article) -> pd.DataFrame:
    HOST = os.getenv('HOST')
    ID = os.getenv('ID')
    DATABASE = os.getenv('DATABASE')
    PSWD = os.getenv('PSWD')

    connection_string = f"mssql+pyodbc://{ID}:{PSWD}@{HOST}/{DATABASE}?driver=ODBC+Driver+17+for+SQL+Server"

    engine = create_engine(connection_string)
    codes_in_clause = ', '.join(f"'{code}'" for code in code_article)
    
    sql_prix = f"""
            SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;
            SELECT 
                RTRIM(a.c_art) AS Code_article, RTRIM(a.profil) AS Profil, RTRIM(a.c_marque) AS Abrev_marque, 
                RTRIM(m.nom_marque) AS Marque, RTRIM(t.prix_ht) AS Prix, RTRIM(a.gencode) AS gencode,
                CASE 
                    WHEN c_sfam_art IN ('4H', 'CH', 'TH') THEN 'Hiver'
                    WHEN c_sfam_art IN ('TE', 'CE', '44') THEN 'Eté' 
                    WHEN c_sfam_art IN ('4TS', 'CTS', 'TTS') THEN '4 Saisons' ELSE ' ' 
                    END AS Saison,  
                lib_art AS Libelle
            FROM 
                wp_chrono..article a, wp_chrono..marque m, wp_chrono..lig_tarif t   
            WHERE 
                a.c_marque = m.c_marque AND a.c_art IN ({codes_in_clause}) 
                AND t.c_art = a.c_art AND t.c_tarif = 'PLOMB' 
        """  
    df = pd.read_sql_query(sql_prix, engine)
    
    resultat = df.loc[df['Marque'] == 'Hankook', 'Libelle'].tolist()
    
    if resultat:
        hank = resultat[0].split(' ')
        df['Profil'] = df.apply(lambda row: row['Profil'] + ' ' + hank[-1] if row['Marque'] == 'Hankook' else row['Profil'], axis=1)
    
    engine.dispose()
    
    return df