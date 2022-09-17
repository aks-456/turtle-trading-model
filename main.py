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
    
workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

cell_format_green = workbook.add_format()
cell_format_green.set_font_color('green')

cell_format_red = workbook.add_format()
cell_format_red.set_font_color('red')


worksheet.write('A1', 'Ticker')
worksheet.write('B1', '20-Day High')
worksheet.write('C1', '20-Day Low')
worksheet.write('D1', '55-Day High')
worksheet.write('E1', '55-Day Low')


docs = db.collection(u'tickers').stream()
i = 2
for doc in docs:
    ticker = doc.to_dict()['Name']

    # Get the data
    data_20 = yf.download(tickers=str(ticker), period="20d", interval="1d")
    data_55 = yf.download(tickers=str(ticker), period="55d", interval="1d")

    #Write to excel doc
    td_high_20 = data_20['High'].iloc[-1]
    td_high_55 = data_55['High'].iloc[-1]

    if(data_20['High'].max() == td_high_20):
        worksheet.write('A' + str(i), str(ticker))
        worksheet.write('B' + str(i), '20-day-high', cell_format_green)
    if(data_20['High'].min() == td_high_20):
        worksheet.write('A' + str(i), str(ticker))
        worksheet.write('C' + str(i), '20-day-low', cell_format_red)

    if(data_55['High'].max() == td_high_55):
        worksheet.write('A' + str(i), str(ticker))
        worksheet.write('D' + str(i), '55-day-high', cell_format_green)

    if(data_55['High'].min() == td_high_55):
        worksheet.write('A' + str(i), str(ticker))
        worksheet.write('E' + str(i), '55-day-low', cell_format_red)
        

    i += 1

workbook.close()
