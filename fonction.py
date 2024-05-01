# def save_donnees_sql_local(data, marques):
#     HOST = os.getenv('HOSTLOCAL')
#     PORT = os.getenv('PORT')
#     ID = os.getenv('IDLOCAL')
#     DATABASE = os.getenv('DATABASE')
#     PSWD = os.getenv('PSWDLOCAL')
#
#     # data = [item for item in data if item is not None]
#
#     df = pd.DataFrame(data)
#     # df.replace("", pd.NA, inplace=True)
#     # df.dropna(inplace=True)
#     # print(df.head(5))
#     df = df[df['Marque'].isin(marques)]
#     df = df[df['Saison'] != 'hiver']
#     df.replace("été / hiver", '4 Saisons', inplace=True)
#     df.replace("été", 'Eté', inplace=True)
#
#     connection_string = f"postgresql+psycopg2://{ID}:{PSWD}@{HOST}:{PORT}/{DATABASE}"
#     engine = create_engine(connection_string)
#
#     df.to_sql('scrap_pneu', engine, if_exists='append', index=False)
#
# def save_donnees_excel(data, marques):
#     df = pd.DataFrame(data)
#     df.replace("", pd.NA, inplace=True)
#     df.dropna(inplace=True)
#     df = df[df['saison'] != 'hiver']
#     df.replace("été / hiver", '4 Saisons', inplace=True)
#     df.replace("été", 'Eté', inplace=True)
#
#     workbook = load_workbook('Scrap_site_web.xlsx')
#     with pd.ExcelWriter('Scrap_site_web.xlsx', if_sheet_exists="overlay", mode='a', engine='openpyxl') as writer:
#         for marque in marques:
#
#             df_marque = df[df['Marque'] == marque]
#
#             df_marque.to_excel(writer, sheet_name=marque,
#                                startrow=workbook[marque].max_row, index=False, header=False)