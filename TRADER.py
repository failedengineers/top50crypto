from pycoingecko import CoinGeckoAPI
import pandas as pd

import xlwings as xw4
import schedule
import time

from openpyxl import load_workbook,Workbook
from datetime import datetime

file='new.xlsx'

import os



3#file_path = "data.xlsx"

cg = CoinGeckoAPI()
def report(file):
    capsorted=[]
    i=0
    t=[]
    change=[]

    coins50 = cg.get_coins_markets(vs_currency='usd', order='market_cap_desc', per_page=50, page=1, sparkline=False)
    data=[['name','coin','currprice','marketcap','totalvol','pricechg24h','pricechg24h%']]
    
    
    try:
        for coin in coins50:
            datas=data.append([coin['name'],coin['symbol'],coin['current_price'],coin['market_cap'],coin['total_volume'],coin['price_change_24h'],coin['price_change_percentage_24h']])
            
            new=data[1:]
        for j in new:
            capsorted.append(j[0])
            i=i+1
            if i==5:
                break
            else:
                pass
        for k in new:
            total=t.append(k[2])
        totals=sum(t)
        avg=round(totals/len(t),2)
        for m in new:
            c=change.append(m[6])
        maxm=max(change)
        minm=min(change)

        dic={'top5crypto'.title():capsorted,'AVG':avg,"highest24h":maxm,'minm24h':minm}
        print(dic)
        if not os.path.exists(file):
            df = pd.DataFrame(data[1:], columns=data[0])
            df.to_excel(file, index=False, engine='openpyxl')
            print(f"Created new file: {file}")
        
        wb = load_workbook(file)
        sheet = wb.active  

        for g in new:
            sheet.append(g)

        wb.save(file)
        print(f" Data appended successfully! at this time {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    except Exception as e:
        print('SORRY NO DATA FOUND',e)
schedule.every(5).minutes.do(report, file)

print("Scheduler started...")

while True:
    schedule.run_pending()
    time.sleep(2)


report(file)


