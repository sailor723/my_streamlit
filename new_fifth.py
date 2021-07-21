import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import xlwings as xw
from PIL import Image
import os, datetime
import pydeck as pdk
import datetime as dt
from urllib.error import URLError

st.set_page_config(page_title='器械调度管理2021.1')

st.header('为全球的医生和患者提供安全创新的心血管医疗器械')

image = Image.open('C:\\Users\\weiping\\Documents\\eclincloud2020\\Customers\\元心\\images\\logo.png')
st.sidebar.image(image,
        caption ='Design by ByotyK/ECC',
        width=200)


excel_file = 'C:\\Users\\weiping\\Documents\\eclincloud2020\\training\\Python_dev\\code\\streamlit\\my_test\\器械模拟20210623 - 副本.xls'
wb = xw.Book(excel_file)

#  read from sheet Working————————————————————————————————————————————————————————————————————————————————————————————————————————————

sheetMain = wb.sheets('Whole_geo')
sheetMainTableRange = sheetMain.range('A1').options(expand='table',empty = 'NA')

df = pd.DataFrame(sheetMainTableRange.value[1:], columns=sheetMainTableRange.value[0])       # read df_post if data exists

df = df.reset_index().set_index('序号',drop=False)
# df['效期'] = pd.to_datetime(df['效期']).dt.date
# df['接收日期'] = pd.to_datetime(df['接收日期']).dt.date
# df['使用日期'] = pd.to_datetime(df['使用日期']).dt.date



# initialize Carrier, or read from sheet Carrier if data exists——————————————————————————————————————————————————————————————————————————————

sheetCarrier = wb.sheets('Carrier')
sheetCarrierTableRange = sheetCarrier.range('A1').options(expand='table',empty = 'NA',numbers = int)
try:
    st.write('start')
    df_post = pd.DataFrame(sheetCarrierTableRange.value[1:], columns=sheetCarrierTableRange.value[0])  # read df_post if data exists
    st.write('after')
    df_post = df_post[df_post.columns.drop('NA')]
    df_post['发送日期'] = pd.to_datetime(df_post['发送日期'])
    df_post[df_post['接收日期'] != 'NA']['接收日期'] = pd.to_datetime(df_post[df_post['接收日期'] != 'NA']['接收日期']).dt.date
    st.write('end')
    st.dataframe(df_post)

except:
    df_post = pd.DataFrame(columns=['运单号','序号','始发库','目标中心','发送日期','接收日期','s_lon','s_lat','r_lon','r_lat'])     # initial sheetCarrier if on data
    sheetCarrier.range('A1').options(numbers = str,transpose=False,index=True,dates=dt.date).value = df_post
    sheetCarrier.autofit()
    st.write('hello')



def load_data_geo():

        sheet_name = "Geo"
        # excel_file = '器械模拟20210623 - 副本.xls'
        df = pd.read_excel(excel_file,
            sheet_name=sheet_name,
            keep_default_na = False
            )
        return df
df_geo = load_data_geo()

st.text('df')
st.dataframe(df)
st.text('df_post')
st.dataframe(df_post)
st.text('df_geo')
st.dataframe(df_geo)


#————————————————Geo Chart ————————————————————————————————————————————————

# df_geo['date_range_mean'] = df[mask].groupby('中心名称1')['date_range'].mean().astype(int)
# df_geo.fillna(0, inplace=True)
GREEN_RGB = [0, 255, 0, 40]
RED_RGB = [240, 100, 0, 40]

try:
    ALL_LAYERS = {
        "Device Number": pdk.Layer(
            "HexagonLayer",
            data=df,
            get_position="[lon, lat]",
            # get_position=["lon", "lat"],
            radius=2000,
            elevation_scale=400,
            elevation_range=[0, 1000],
            pickable=True,
            extruded=True,
        ),
        "Date_Range": pdk.Layer(
            "ScatterplotLayer",
            data=df_geo,
            get_position=["lon", "lat"],
            get_color=[200, 30, 0, 160],
            get_radius="[date_range_mean]",
            pickable=True,
            radius_scale=200,
        ),
        "Carrier_Route": pdk.Layer(
            "ArcLayer",

            data=df_post[df_post['接收日期'] == 'NA'],
            get_source_position=["s_lon", "s_lat"],
            get_target_position=["r_lon", "r_lat"],
            auto_highlight=True,
            width_scale=0.01,
            get_width="outbound",
            width_min_pixels=10,
            width_max_pixels=60,
            get_source_color=RED_RGB,
            get_target_color=GREEN_RGB,
            pickable=True,
        ),

    }
    st.sidebar.markdown('### Map Layers')
    selected_layers = [
        layer for layer_name, layer in ALL_LAYERS.items()
        if st.sidebar.checkbox(layer_name, True)]
    if selected_layers:
        st.pydeck_chart(pdk.Deck(
            map_style="mapbox://styles/mapbox/light-v9",
            initial_view_state={"latitude": 33.2164668532657,
                                "longitude": 110.472886,
                                "zoom": 4,
                                "pitch": 0,
                                },
            layers=selected_layers,
        ))
    else:
        st.error("Please choose at least one layer above.")
except URLError as e:
    st.error("""
        **This demo requires internet access.**

        Connection error: %s
    """ % e.reason)