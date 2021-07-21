
import streamlit as st
import pydeck as pdk
import xlwings as xw
import os, datetime
# To make things easier later, we're also importing numpy and pandas for
# working with sample data.
import numpy as np
import pandas as pd



path = 'C:\\Users\\weiping\\Documents\\eclincloud2020\\Customers\\元心\\器械模拟20210623 - 副本.xls'
wb = xw.Book(path)
df = pd.DataFrame()

sheet = wb.sheets['Geo']

table_range = sheet.range('A1').options(expand='table',empty = 'NA',numbers = int)

df_read  = pd.DataFrame(table_range.value[1:], columns=table_range.value[0])

# st.map(df_read)



st.pydeck_chart(pdk.Deck(
  map_style='mapbox://styles/mapbox/light-v9',
  initial_view_state=pdk.ViewState(
    latitude=31.21646685,
    longitude=121.472886,
    zoom=11,
    pitch=50,
     ),
    layers=[
         pdk.Layer(
            'HexagonLayer',
            data=df_read,
            get_position='[lon, lat]',
            get_elevation=40,
            radius=200,
            elevation_scale=1,
            elevation_range=[0, 10],
            pickable=True,
            extruded=True,
         ),
         pdk.Layer(
             'ScatterplotLayer',
             data=df_read,
             get_position='[lon, lat]',
             get_color='[200, 30, 0, 160]',
             get_radius=200,
         ),
     ],
 ))