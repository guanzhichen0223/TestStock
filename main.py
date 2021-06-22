import pandas as pd
import requests
import xlwings as xw


datastr = input('請輸入日期:')


def price():
    url = 'https://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date=' + \
        datastr + '&type=ALL'
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
    xw.view(df)
