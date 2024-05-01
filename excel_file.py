import pandas as pd
from requete import requete
from pandas import DataFrame
from typing import Tuple, Any


def excel_file(adress: str, i) -> tuple[Any, Any, Any, Any, Any, str, DataFrame, Any]:

    taille = pd.read_excel(adress, sheet_name=i)
    code_article = taille['Code_article'].unique().tolist()
    df = requete(code_article)
    df['Code'] = taille['Code']
    code = df['Code'].unique().tolist()
    profil = df['Profil'].unique().tolist()
    marque_abrege = df['Abrev_marque'].unique().tolist()
    marques = df['Marque'].unique().tolist()
    gencode = df['gencode'].unique().tolist()
    saisons = ""
    if i[0:3] == 'ÉTÉ':
        saisons = 'été'
    if i[0:2] == 'TS':
        saisons = '4 saisons'

    return code, profil, marque_abrege, marques, gencode, saisons, df, code_article