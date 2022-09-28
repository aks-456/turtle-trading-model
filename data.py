from openpyxl import Workbook
from openpyxl import load_workbook
import firebase_admin
from firebase_admin import credentials, firestore

cred = credentials.Certificate("insert_credentials_certificate_file_path_here")
firebase_admin.initialize_app(cred)

db = firestore.client()

#Read excel file
wb = load_workbook(filename = 'insert_ticker_filename')
tickers = wb['Ticker'] #Header of column on A1

#Read ticker values in column A, row 2 onwards
i = 2
while(tickers['A' + str(i)].value is not None):
    row_val = tickers['A' + str(i)].value #Ticker values are on A2 all the way to An
    
    data = {
        u'Name': row_val #set ticker value in name category
    }

# Add a new ticker in collection tickers, with id "name_of_ticker"
    db.collection(u'tickers').document(row_val).set(data)
    i += 1