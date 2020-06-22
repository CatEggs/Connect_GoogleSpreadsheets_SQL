import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import pyodbc
import config

### Connect to Google spreadsheet ###
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_cred.json', scope)
client = gspread.authorize(creds)
sheet = client.open('Master Deficiency List').sheet1

### Connect to SQL ###
connection = pyodbc.connect(
        r'DRIVER={SQL Server Native Client 11.0};'
        r'SERVER=' + config.server + ';'
        r'DATABASE=' + config.database + ';'
        r'UID=' + config.username + ';'
        r'PWD=' + config.password
        )
cursor = connection.cursor()

### Get List of Missing Discrepancies ###
result = []
header = ['DeficiencyId', 'TypeName', 'EntityId', 'EntityTypeName', 'Note', 'StageName', 'CreatedByName']
disc_list = sheet.col_values(4)
qmark = "?"
for rows in range(len(disc_list)-1):
    qmark += ",?"
sql = (
""" select DeficiencyId, TypeName, EntityId, EntityTypeName, Note, StageName, CreatedByName
    from Deficiencies
    where Note not in ({0})""".format(qmark))

for row in cursor.execute(sql, disc_list).fetchall():
    result.append(dict(zip(header, row)))

cursor.commit()

### Export List into Excel Spreadsheet ###
pd.DataFrame.from_dict(result).to_excel("missing_deficiencies.xlsx", index = False)