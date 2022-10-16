from flask import Flask
from flask import request
from flask import abort, redirect, url_for
import yfinance as yf
import firebase_admin
from firebase_admin import credentials, firestore
import xlsxwriter

cred = credentials.Certificate("turtle-trading-model-firebase-adminsdk-4rv25-ce9ed9fa72.json")
firebase_admin.initialize_app(cred)

db = firestore.client()

#Create new file report.xlsx
workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

#Color the text green in the cell
cell_format_green = workbook.add_format()
cell_format_green.set_font_color('green')

#Color the text red in the cell
cell_format_red = workbook.add_format()
cell_format_red.set_font_color('red')

#Write headings to file
worksheet.write('A1', 'Ticker')
worksheet.write('B1', '20-Day High')
worksheet.write('C1', '20-Day Low')
worksheet.write('D1', '55-Day High')
worksheet.write('E1', '55-Day Low')

#Get file from collection 'tickers'
docs = db.collection(u'tickers').stream()

i = 2 #Track row number in file

#Iterate through each document in the collection
for doc in docs:
    ticker = doc.to_dict()['Name']

    print(ticker) #Tracking which ticker is being uploaded

    increase_count = 0
    #Catch any errors with the API 
    try:
        # Get data
        data_20 = yf.download(tickers=str(ticker), period="20d", interval="1d")
        data_55 = yf.download(tickers=str(ticker), period="55d", interval="1d")

        #Write to excel doc
        td_high_20 = data_20['High'].iloc[-1]
        td_high_55 = data_55['High'].iloc[-1]
        
        #Check if today's price is a 20 day high
        if(data_20['High'].max() == td_high_20):
            worksheet.write('A' + str(i), str(ticker)) #Write ticker name in cell
            worksheet.write('B' + str(i), '20-day-high', cell_format_green) #Write '20-day-high' in cell with color green
            increase_count = 1 #Tracks if the current document was written to the file
        #Check if today's price is a 20 day low
        if(data_20['High'].min() == td_high_20):
            worksheet.write('A' + str(i), str(ticker))
            worksheet.write('C' + str(i), '20-day-low', cell_format_red)
            increase_count = 1
        #Check if today's price is a 55 day high
        if(data_55['High'].max() == td_high_55):
            worksheet.write('A' + str(i), str(ticker))
            worksheet.write('D' + str(i), '55-day-high', cell_format_green)
            increase_count = 1
        #Check if today's price is a 55 day low
        if(data_55['High'].min() == td_high_55):
            worksheet.write('A' + str(i), str(ticker))
            worksheet.write('E' + str(i), '55-day-low', cell_format_red)
            increase_count = 1
    except:
        continue
    
    #If the current document was written to the file, increase i by 1 to move to the next row
    if(increase_count == 1):
        i += 1

workbook.close()
