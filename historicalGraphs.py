from yahooquery import Ticker
import mplfinance as mpf
import pandas as pd
from datetime import datetime, timedelta

df = pd.read_excel('historicalGappers.xlsx')
mkt = ''

def job(x):
    datedate = datetime.strptime(df['Date'][x], '%Y-%m-%d').date()
    openPrice = Ticker(df['Ticker'][x]).history(interval='1d', start=str(df['Date'][x]), end=str(datedate + timedelta(1))).iloc[0]['open']
    mktDay = Ticker(df['Ticker'][x]).key_stats[df['Ticker'][x]]['sharesOutstanding'] * openPrice
    tick = Ticker(df['Ticker'][x]).history(interval='5m', start=str(df['Date'][x]), end=str(datedate + timedelta(1)))
    tick.index.name = 'Date'

    if mktDay <= 10000000:
        mkt = '0-10 Mkt Cap'
    elif mktDay <= 50000000:
        mkt = '10-50 Mkt Cap'
    elif mktDay <= 100000000:
        mkt = '50-100 Mkt Cap'
    elif mktDay <= 200000000:
        mkt = '100-200 Mkt Cap'
    elif mktDay <= 400000000:
        mkt = '200-400 Mkt Cap'
    else:
        mkt = '>400 Mkt Cap'

    name = '$' + str(df['Ticker'][x]) + '; '+ df['Date'][x] + '; '+ df['Closed G/R EOD'][x] + "; Gap " + str(round(df['% at Open'][x],2)) + '%; ' + mkt
    mpf.plot(tick, type='candle', volume=True, title= name, style='yahoo',hlines=dict(hlines=[tick['open'][0]], colors=['g'], linestyle='-.'), savefig = 'graphs/'+name +'.png')


for x in range(269):
    try:
        job(x)
        print(x)
    except:
        pass

