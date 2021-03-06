from yahooquery import Ticker
import pandas as pd
from datetime import datetime, timedelta, date
from openpyxl import Workbook, load_workbook
import time
import bs4

startTime = time.time()

def job(x,y):
    yplus1 = str(y.date() +timedelta(1))
    yminus1 = str(y.date() -timedelta(1))
    y = y.date()
    df15 = Ticker(x).history(interval='15m', start=str(y), end=yplus1)
    df1d = Ticker(x).history(interval='1d', start=str(y), end=yplus1)
    openPrice = df1d.iloc[0]['open']
    prevClose = Ticker(x).history(interval='1d', start=yminus1, end=str(y)).iloc[0]['close']

    data =   {'Ticker':Ticker(x).price[x]['symbol'],
              'Sector':Ticker(x).asset_profile[x]['sector'],
              'Country':Ticker(x).summary_profile[x]['country'],
              'Index':Ticker(x).price[x]['exchangeName'],
              'Date':str(y),
              'Closed G/R EOD':'Green' if (Ticker(x).history(interval='1d', start=str(y), end=str(yplus1)).iloc[0]['close'] - openPrice)*100 > 0 else 'Red',
              '% at Open': ((openPrice - prevClose)/prevClose)*100,
              '% at Close':((df1d.iloc[0]['close'] - openPrice)/openPrice)*100,
              'High/Low %': ((df1d.iloc[0]['close'] - openPrice)/openPrice)*100,
              'Pattern': 'Fall' if (((df1d.iloc[0]['close'] - openPrice)/openPrice)*100 < -5) else 'Go' if (((df1d.iloc[0]['close'] - openPrice)/openPrice)*100 > 5 or ((df1d.iloc[0]['close'] - openPrice)/openPrice)*100 < 5) else 'Hold',
              'Prev. Close':prevClose,
              'Prev. High': round(Ticker(x).history(interval='1d', start = yminus1, end=str(y)).iloc[0]['high'],2),
              'Open':openPrice,
              'High':df1d.iloc[0]['high'],
              'Low':df1d.iloc[0]['low'],
              'Close':df1d.iloc[0]['close'],

            'First 5 min High': round(Ticker(x).history(interval='5m', start=str(y), end=yplus1).iloc[0]['high'], 2),
            'First 10 min High': round(Ticker(x).history(interval='5m', start=str(y), end=yplus1).iloc[1]['high'], 2),
            '1st 15 min High': round(df15.iloc[1]['high'], 2),
            '2nd 15 min High': round(df15.iloc[2]['high'], 2),
            '3rd 15 min High': round(df15.iloc[3]['high'], 2),
            '4th 15 min High': round(df15.iloc[4]['high'], 2),
            '5th 15 min High': round(df15.iloc[5]['high'], 2),
            '6th 15 min High': round(df15.iloc[6]['high'], 2),
            '7th 15 min High': round(df15.iloc[7]['high'], 2),
            '8th 15 min High': round(df15.iloc[8]['high'], 2),
            '9th 15 min High': round(df15.iloc[9]['high'], 2),
            '10th 15 min High': round(df15.iloc[10]['high'], 2),
            '11th 15 min High': round(df15.iloc[11]['high'], 2),
            '12th 15 min High': round(df15.iloc[12]['high'], 2),
            '13th 15 min High': round(df15.iloc[13]['high'], 2),
            '14th 15 min High': round(df15.iloc[14]['high'], 2),
            '15th 15 min High': round(df15.iloc[15]['high'], 2),
            '16th 15 min High': round(df15.iloc[16]['high'], 2),
            '17th 15 min High': round(df15.iloc[17]['high'], 2),
            '18th 15 min High': round(df15.iloc[18]['high'], 2),
            '19th 15 min High': round(df15.iloc[19]['high'], 2),
            '20th 15 min High': round(df15.iloc[20]['high'], 2),
            '21st 15 min High': round(df15.iloc[21]['high'], 2),
            '22nd 15 min High': round(df15.iloc[22]['high'], 2),
            '23rd 15 min High': round(df15.iloc[23]['high'], 2),
            '24th 15 min High': round(df15.iloc[24]['high'], 2),
            '25th 15 min High': round(df15.iloc[25]['high'], 2),

            '5 min %':round(((Ticker(x).history(interval='5m', start=str(y), end=yplus1).iloc[0]['close']-openPrice)/openPrice)*100,2),
            '10 min %': round(((Ticker(x).history(interval='5m', start=str(y), end=yplus1).iloc[1]['close'] - openPrice) / openPrice) * 100, 2),
            '1st 15 min %': round(((df15.iloc[1]['close'] -openPrice) /openPrice) * 100, 2),
            '2nd 15 min %': round(((df15.iloc[2]['close'] -openPrice) /openPrice) * 100, 2),
            '3rd 15 min %': round(((df15.iloc[3]['close'] -openPrice) /openPrice) * 100, 2),
            '4th 15 min %': round(((df15.iloc[4]['close'] -openPrice) /openPrice) * 100, 2),
            '5th 15 min %': round(((df15.iloc[5]['close'] -openPrice) /openPrice) * 100, 2),
            '6th 15 min %': round(((df15.iloc[6]['close'] -openPrice) /openPrice) * 100, 2),
            '7th 15 min %': round(((df15.iloc[7]['close'] -openPrice) /openPrice) * 100, 2),
            '8th 15 min %': round(((df15.iloc[8]['close'] -openPrice) /openPrice) * 100, 2),
            '9th 15 min %': round(((df15.iloc[9]['close'] -openPrice) /openPrice) * 100, 2),
            '10th 15 min %': round(((df15.iloc[10]['close'] -openPrice) /openPrice) * 100, 2),
            '11th 15 min %': round(((df15.iloc[11]['close'] -openPrice) /openPrice) * 100, 2),
            '12th 15 min %': round(((df15.iloc[12]['close'] -openPrice) /openPrice) * 100, 2),
            '13th 15 min %': round(((df15.iloc[13]['close'] -openPrice) /openPrice) * 100, 2),
            '14th 15 min %': round(((df15.iloc[14]['close'] -openPrice) /openPrice) * 100, 2),
            '15th 15 min %': round(((df15.iloc[15]['close'] -openPrice) /openPrice) * 100, 2),
            '16th 15 min %': round(((df15.iloc[16]['close'] -openPrice) /openPrice) * 100, 2),
            '17th 15 min %': round(((df15.iloc[17]['close'] -openPrice) /openPrice) * 100, 2),
            '18th 15 min %': round(((df15.iloc[18]['close'] -openPrice) /openPrice) * 100, 2),
            '19th 15 min %': round(((df15.iloc[19]['close'] -openPrice) /openPrice) * 100, 2),
            '20th 15 min %': round(((df15.iloc[20]['close'] -openPrice) /openPrice) * 100, 2),
            '21st 15 min %': round(((df15.iloc[21]['close'] -openPrice) /openPrice) * 100, 2),
            '22nd 15 min %': round(((df15.iloc[22]['close'] -openPrice) /openPrice) * 100, 2),
            '23rd 15 min %': round(((df15.iloc[23]['close'] -openPrice) /openPrice) * 100, 2),
            '24th 15 min %': round(((df15.iloc[24]['close'] -openPrice) /openPrice) * 100, 2),
            '25th 15 min %': round(((df15.iloc[25]['close'] -openPrice) /openPrice) * 100, 2),

              'Morning Low':round(Ticker(x).history(interval='60m', start=str(y), end=yplus1).iloc[1]['low'], 2),
              'Market Cap':Ticker(x).key_stats[x]['sharesOutstanding'] * openPrice,
              'Volume':df1d.iloc[0]['volume'],
              'Float':Ticker(x).key_stats[x]['floatShares'],
              'Float vs Volume':Ticker(x).summary_detail[x]['volume']/Ticker(x).key_stats[x]['floatShares']}

    print(data)
    df = pd.DataFrame(data, index=[0])
    print(df)
    writer = pd.ExcelWriter('historicalGappers.xlsx', engine='openpyxl')
    writer.book = load_workbook('historicalGappers.xlsx')
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    reader = pd.read_excel('historicalGappers.xlsx')
    df.to_excel(writer, index=False, header=False, startrow=len(reader) + 1, sheet_name='main')
    writer.close()

list1 = ['KXIN','HOFV','FRGI']
for x in list1:
    try:
        job(x,datetime(2020, 11, 5))
    except:
        continue