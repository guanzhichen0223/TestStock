import pandas as pd
import requests
import xlwings as xw
import numpy as np
import datetime
import time


def price(date):
    url = 'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date=' + \
        str(date).split(' ')[0].replace('-', '') + '&type=ALL'
    res = requests.get(url)
    data = res.text
    i = 0

    # 把爬下來的資料整理乾淨
    cleaned_data = []
    for da in data.split('\n'):
        if len(da.split('","')) == 16 and da.split('","')[0][0] != '=':
            cleaned_data.append([ele.replace('",\r', '').replace('"', '')
                                for ele in da.split('","')])

    # 輸出成表格並呈現到excel上
    df = pd.DataFrame(cleaned_data, columns=cleaned_data[0])
    df = df.set_index('證券代號')[1:]
    return xw.view(df)


data = {}
n_days = 5
date = datetime.datetime.now()
fail_count = 0
allow_continuous_fail_count = 5
while len(data) < n_days:

    print('parsing', date)
    # 使用 crawPrice 爬資料
    try:
        # 抓資料
        data[date.date()] = price(date)
        print('success!')
        fail_count = 0
    except:
        # 假日爬不到
        print('fail! check the date is holiday')
        fail_count += 1
        if fail_count == allow_continuous_fail_count:
            raise
            break

    # 減一天
    date -= datetime.timedelta(days=1)
    time.sleep(5)
