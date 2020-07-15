from oauth2client.service_account import ServiceAccountCredentials
from email_results import email
from datetime import date
import pandas as pd
import gspread
import pyodbc
import config
import time
import os

def get_deficiencies(filename):
    logg_file = open(filename, "a+")
    
### Connect to Google spreadsheet ###
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name('client_cred.json', scope)
        client = gspread.authorize(creds)
        sheet = client.open('Master Deficiency List').sheet1
        logg_file.write('INFO - Connection to Google was successful.\n')

    except Exception as e:
        logg_file.write('CRITICAL - Connection to Google failed. Error: {0}.\n'.format(str(e)))

### Connect to SQL ###
    try:
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
        logg_file.write('INFO - Deficiency file created successfully.\n')
    except Exception as e:
        logg_file.write('CRITICAL - Deficiency file created with errors. Error: {0}\n'.format(str(e)))


def main():
    today = date.today()
    logg_file = "log_file\\logfile-{0}.txt".format(str(today))
    get_deficiencies(logg_file)
    
### Send attachments via email ###
    email()

main()

