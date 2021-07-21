## -*- coding: utf-8 -*-

# initalize by load libraries, and execl file
import pandas as pd
import numpy as np
import xlwings as xw
import os, datetime
import requests
import datetime as dt

# set path and file name, read into sheet, convert to dataframe
path = 'C:\\Users\\weiping\\Documents\\eclincloud2020\\training\\Python_dev\\code\\streamlit\\my_test\\器械模拟20210623 - 副本.xls'
wb = xw.Book(path)
df = pd.DataFrame()

# read from sheet to concat total dataframe with all sites

for sheet_num in range(0, len(wb.sheets)-4):


    sheet = wb.sheets[sheet_num]

    table_range = sheet.range('A1').options(expand='table',empty = 'NA',numbers = int)

    df_read  = pd.DataFrame(table_range.value[1:], columns=table_range.value[0],)

    new_column = ['中心名称'] + df_read.columns.tolist() + ['中心名称1']

#     df_read = df_read.reindex(columns=new_column, fill_value= [sheet.name,sheet.name.strip()[2:]'])
    df_read = df_read.reindex(columns=new_column, fill_value= sheet.name)
    df_read['中心名称1'] = sheet.name.replace(" ","")[2:]

    df = pd.concat([df, df_read])

# —————————————————————————————————————— status——————————————————————————————————————
status= []
for s in df['使用日期'].astype(str):
    if len(s) == 0:
        status.append('可用')
    else:
        status.append('完成')
df['状态'] = status


# convert time to datatime format and create data_range
df['效期'] = pd.to_datetime(df['效期']).dt.date
df['接收日期'] = pd.to_datetime(df['接收日期']).dt.date
df['使用日期'] = pd.to_datetime(df['使用日期']).dt.date


# df['效期'] = df['效期'].apply(lambda x : x.date())
# df['接收日期'] = df['接收日期'].apply(lambda x : x.date())
# df['使用日期'] = df['使用日期'].apply(lambda x : x.date())
df['date_range']=(df['效期'] - datetime.datetime.now().date()).apply(lambda x : x.days)

# df['效期'] = pd.to_datetime(df['效期'])
# df['接收日期'] = pd.to_datetime(df['接收日期'])
# df['使用日期'] = pd.to_datetime(df['使用日期'])

# df['效期'] = pd.to_datetime(df['效期'])
# df['接收日期'] = pd.to_datetime(df['接收日期'])
# df['使用日期'] = pd.to_datetime(df['使用日期'])
# df['效期'] = df['效期'].apply(lambda x : x.date())
# df['接收日期'] = df['接收日期'].apply(lambda x : x.date())
# df['使用日期'] = df['使用日期'].apply(lambda x : x.date())
# df['date_range']=(df['效期'] - datetime.datetime.now().date()).apply(lambda x : x.days)

# —————————————————————————————————————— convert device and model——————————————————————————————————————
item_name = []
item_model = [ ]
for item in df["规格型号"].str.split("-"):
    if len(item) == 5:
        item_name.append( "-".join(item[0:3]))
        item_model.append(item[3])
    else:
        item_name.append( "-".join(item[0:2]))
        item_model.append(item[2])

df['model_name'] = item_name
df['model_model'] = item_model

df_columns = ['中心名称', '序号', '器械名称', '规格型号',  'model_name', 'model_model','序列号', '批号', '效期', '接收日期',
       '使用日期', '状态','受试者编号', '中心名称1','date_range']
df = df[df_columns]


# create site name without number, write into 'Whole' sheet
df_geo = pd.DataFrame(df.groupby('中心名称1').size(),columns=['num'])
df_geo['date_range_mean'] = df.groupby('中心名称1')['date_range'].mean().astype(int)

# reindex with 中心名称，序号
df['序号'] = df['序号'].astype(int)
df = df.sort_values(by=['中心名称','序号'],ascending=[True,True])
df.set_index(['中心名称', '序号'],inplace=True)

sheet1 = wb.sheets('Whole')   # create Working sheet
sheet1.range('A1').options(index=True,numbers = str).value = df # write to Working sheet, number must be int
sheet1.autofit()                                                                       # auto fit Working sheet
wb.save()

import requests
import json

def getUrl(*address):
    '''
    调用地图API获取待查询地址专属url
    最高查询次数30w/天，最大并发量160/秒
    '''
    ak = 'biWEYo2PUIpAFDojunotOUBy1agBYUGl'
    if len(address) < 1:
        return None
    else:
        for add in address:
            url = 'http://api.map.baidu.com/geocoding/v3/?address={inputAddress}&output=json&ak={myAk}'.format(inputAddress=add,myAk=ak)
            yield url


def getPosition(url):
    '''返回经纬度信息'''
    res = requests.get(url)
    json_data = json.loads(res.text)

    if json_data['status'] == 0:
        lat = json_data['result']['location']['lat'] #纬度
        lng = json_data['result']['location']['lng'] #经度
    else:
        print("Error output!")
        return json_data['status']
    return lat,lng



if __name__ == "__main__":
#     address = ['上海市嘉定区科贸路8号','苏州市虎丘区青花路9号','沈阳辉山农业高新技术开发区宏业街73号','广东省佛山市魁奇一路9号','四川省成都市锦江区中纱帽街8号']
    geo_lat = []
    geo_lng = []
    for add in df_geo.index.tolist():
        add_url = list(getUrl(add))[0]
#         print(add_url)
        try:
            lat,lng = getPosition(add_url)
            geo_lat.append(lat)
            geo_lng.append(lng)
#             print("运单地址：{0}|经度:{1}|纬度:{2}.".format(add,lng,lat))
        except Error as e:
            print(e)

df_geo['lon'] = geo_lng
df_geo['lat'] = geo_lat
sheet1 = wb.sheets('Geo')   # create Geo sheet
sheet1.range('A1').options(index=True,numbers = str).value = df_geo # write to Working sheet, number must be int
sheet1.autofit()                                                                       # auto fit Working sheet
wb.save()

df = df.join(df_geo, on ='中心名称1')
sheet1 = wb.sheets('Whole_geo')   # create Working sheet
#sheet1.range('E1').api.EntireColumn.NumberFormat = '@'
sheet1.range('A1').options(index=True,numbers = str).value = df # write to Working sheet, number must be int
sheet1.autofit()                                                                       # auto fit Working sheet
wb.save()

#———————————————————————————————— read and write to handle NA ————————————————————————————————————————————-
sheet_name = 'Whole_geo'
excel_file = 'C:\\Users\\weiping\\Documents\\eclincloud2020\\training\\Python_dev\\code\\streamlit\\my_test\\器械模拟20210623 - 副本.xls'
df_whole_geo = pd.read_excel(excel_file, sheet_name=sheet_name)
df_whole_geo.fillna(0,inplace=True)

df_whole_geo['序号'] = df_whole_geo['序号'].astype(int)
df_whole_geo = df_whole_geo.sort_values(by=['中心名称','序号'],ascending=[True,True])
df_whole_geo.set_index(['中心名称', '序号'],inplace=True)



sheet1 = wb.sheets('Whole_geo')   # create Working sheet
sheet1.range('A1').options(index=True,numbers = str).value = df_whole_geo # write to Working sheet, number must be int
sheet1.range('A1').expand('table').color = 224,238,224
sheet1.range('A1').expand('table').api.Borders(9).Weight = 2


xw.Range('A1').color = (255,255,255)
sheet1.autofit()


                                                               # auto fit Working sheet
wb.save()