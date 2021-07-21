from numpy.lib.arraysetops import unique
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from PIL import Image
import os, datetime
import pydeck as pdk
from urllib.error import URLError
import altair as alt

st.set_page_config(page_title='器械转运2021')
st.header('为全球的医生和患者提供安全创新的心血管医疗器械')
st.subheader('器械中心分布')

excel_file = 'C:\\Users\\weiping\\Documents\\eclincloud2020\\training\\Python_dev\\code\\streamlit\\my_test\\器械模拟20210623 - 副本.xls'

image = Image.open('C:\\Users\\weiping\\Documents\\eclincloud2020\\Customers\\元心\\images\\logo.png')
st.sidebar.image(image,
        caption ='Design by ByotyK/ECC',
        width=200)

def load_data_whole():

        sheet_name = 'Whole_geo'
        # excel_file = '器械模拟20210623 - 副本.xls'
        df = pd.read_excel(excel_file,
            sheet_name=sheet_name,
            )
        return df


def load_data_geo():

        sheet_name = "Geo"
        # excel_file = '器械模拟20210623 - 副本.xls'
        df = pd.read_excel(excel_file,
            sheet_name=sheet_name,
            keep_default_na = False
            )
        return df


# def load_data_post():

#         sheet_name = "Carrier"
#         # excel_file = '器械模拟20210623 - 副本.xls'
#         df = pd.read_excel(excel_file,
#             sheet_name=sheet_name,
#             keep_default_na = False
#             )
#         return df

df_geo = load_data_geo()
df = load_data_whole()
df = df.reset_index().set_index('序号',drop=False)
df['效期'] = pd.to_datetime(df['效期'])
df['接收日期'] = pd.to_datetime(df['接收日期'])
df['使用日期'] = pd.to_datetime(df['使用日期'])
# df_post = load_data_post()
# st.dataframe(df_post)



#———————————————————————————— Whole_geo View ——————————————————————————————————


# st.dataframe(df)

# ———————————————————————————— streamlit selection ——————————————————————————————————

site = df.site_name.unique().tolist()
dates = df['date_range'].unique().tolist()
device_model = df['model_name'].unique().tolist()
device_spec = df['model_spec'].unique().tolist()


with st.sidebar.form("my_form"):

    date_selection = st.sidebar.slider('date range to expire',   ## display date range selection
                                min_value = min(dates),
                                max_value = max(dates),
                                value = (min(dates), max(dates)))


    site_selection = st.sidebar.multiselect( 'site_name',              ## display site selection
                                    site,
                                    default= site)


    model_selection = st.sidebar.multiselect( '型号',              ## display site selection
                                    device_model,
                                    default= device_model)

    spec_selection = st.sidebar.multiselect( '规格',              ## display site selection
                                    device_spec,
                                    default= device_spec)

    submitted = st.form_submit_button("Submit")


# ————————————— fileter dataframe based on selection————————————————————————————————————————————

mask = (df['date_range'].between(*date_selection)) & (df['site_name'].isin(site_selection)& 
            (df['model_name'].isin(model_selection)) & (df['model_spec'].isin(spec_selection)))
number_of_result = df[mask].shape[0]
st.sidebar.markdown(f'*Avaiable Result: {number_of_result}*')

#———————————————Group dataframe after selection————————————————————————————————————————————————
df_fillter = df[mask].groupby(by=['model_name','model_spec']).count()[['site_name']]
df_fillter = df_fillter.rename(columns ={'site_name':'器械数量'})
df_fillter = df_fillter.reset_index()

GREEN_RGB = [0, 255, 0, 40]
RED_RGB = [240, 100, 0, 40]

#————————————————Geo Chart ————————————————————————————————————————————————

# df_geo['date_range_mean'] = df[mask].groupby('中心名称1')['date_range'].mean().astype(int)
# df_geo.fillna(0, inplace=True)

# try:
#     ALL_LAYERS = {
#         "Device Number": pdk.Layer(
#             "HexagonLayer",
#             data=df,
#             get_position=["lon", "lat"],
#             radius=2000,
#             elevation_scale=400,
#             elevation_range=[0, 1000],
#             pickable=True,
#             extruded=True,
#         ),
#         "Date_Range": pdk.Layer(
#             "ScatterplotLayer",
#             data=df_geo,
#             get_position=["lon", "lat"],
#             get_color=[200, 30, 0, 160],
#             get_radius="[date_range_mean]",
#             pickable=True,
#             radius_scale=200,
#         ),
#         "Carrier_Route": pdk.Layer(
#             "ArcLayer",

#             data=df_post[df_post['接收日期'] == 'NA'],
#             get_source_position=["s_lon", "s_lat"],
#             get_target_position=["r_lon", "r_lat"],
#             auto_highlight=True,
#             width_scale=0.01,
#             get_width="outbound",
#             width_min_pixels=10,
#             width_max_pixels=60,
#             get_source_color=RED_RGB,
#             get_target_color=GREEN_RGB,
#             pickable=True,
#     )
#     }
#     st.sidebar.markdown('### Map Layers')
#     selected_layers = [
#         layer for layer_name, layer in ALL_LAYERS.items()
#         if st.sidebar.checkbox(layer_name, True)]
#     if selected_layers:
#         st.pydeck_chart(pdk.Deck(
#             map_style="mapbox://styles/mapbox/light-v9",
#             initial_view_state={"latitude": 31.2164668532657,
#                                 "longitude": 121.472886,
#                                 "zoom": 6,
#                                 "pitch": 50,
#                                 },
#             layers=selected_layers,

#             # tooltip={

#             #     'html': '<b>器械数量:</b> {elevationValue} <b>时间:</b> {[date_range_mean]} ',
#             #     'style': {
#             #         'color': 'white'
#             #     }
#             #     }

#         ))
#     else:
#         st.error("Please choose at least one layer above.")
# except URLError as e:
#     st.error("""
#         **This demo requires internet access.**

#         Connection error: %s
#     """ % e.reason)

# st.map(df_geo)

#———————————————Ploying the area and bar Chart——————————————————————————————————————————————————————————————————

c = alt.Chart(df_geo).mark_circle().encode(
     x='site_name', y='date_range_mean', size='site_dev_num', color='site_dev_num')

st.altair_chart(c, use_container_width=True)


st.area_chart(df_geo[['site_name', 'site_dev_num']])


st.bar_chart(df_geo[['site_dev_num','date_range_mean']])



#———————————————Plot Bar Chart——————————————————————————————————————————————————————————————————


st.dataframe(df_fillter)
bar_chart = px.bar(df_fillter,
                  x = 'model_name',
                  y = '器械数量',
                #   text = 'model',
                  color_discrete_sequence = ['#F63366'] * len(df_fillter),
                  template = 'plotly_white')
st.plotly_chart(bar_chart,use_container_width=True)

st.write(len(df_fillter))
with st.beta_expander("选择后的数据"):
    st.dataframe(df_fillter)

#_______________________________________________________________________

df_x = df.sort_values(['model_spec','model_name'],ascending=True)
fig = px.histogram(df_x, x="model_name",color='model_spec')
st.plotly_chart(fig)

df_x = df.sort_values(['model_spec','model_name'],ascending=True)
fig = px.histogram(df_x, x="model_name",color='model_spec')
st.plotly_chart(fig)

#_____________________________________________________________________________
import plotly.io as pio

df_3 = df[['model_name','model_spec','使用日期']]
df_3 = df_3[df_3['使用日期'] != '1970-01-01']

df_x = df_3[(df_3['使用日期'] <= '2019-12-31' ) & (df_3['使用日期'] >= '2019-01-01' )].sort_values(['model_spec','model_name'],ascending=True)
fig = px.histogram(df_x, x="model_name",color='model_spec')
st.plotly_chart(fig)


#——————————————————————————————————————————————————————————————————————————————————————————
df_4 = pd.DataFrame(df_3.groupby(['model_name']).size(),columns=['num_device']).reset_index().transpose()
st.dataframe(df_4)

df_4 = pd.DataFrame(df_3.groupby(['model_name','model_spec']).size(),columns=['num_device']).reset_index().transpose()
st.dataframe(df_4)


df_4 = pd.DataFrame(df_3.groupby(['model_spec']).size(),columns=['num_device']).reset_index().transpose()
st.dataframe(df_4)
#———————————————Ploying the Pie Chart——————————————————————————————————————————————————————————————————

col4, col5 = st.beta_columns([3,2])
with col4:
    df_site = pd.DataFrame(df[mask].groupby('site_name').size(),columns=['num'])
    pie_chart = px.pie(df_site,
                    title = '研究中心库存器械分析',
                    values = 'num',
                    names = df_site.index)
    st.plotly_chart(pie_chart)

with col5:
    df_status = pd.DataFrame(df.groupby('状态').size(),columns=['status'])
    pie_chart1 = px.pie(df_status,
                    title = '状态分析',
                    values = 'status',
                    names = df_status.index)
    st.plotly_chart(pie_chart1)



# ———————————————————— heatmap for site, device, number ————————————————————————————————————————————————————————

df_num = pd.DataFrame(df[mask].groupby(['site_name','model_name']).size(),columns=['num_device']).reset_index()

c1 = alt.Chart(df_num).mark_rect().encode(
    x='site_name:O',
    y='model_name:O',
    color='num_device:Q'
)
st.altair_chart(c1, use_container_width=True)


df_num = pd.DataFrame(df[mask].groupby(['site_name','model_spec']).size(),columns=['num_device']).reset_index()

c1 = alt.Chart(df_num).mark_rect().encode(
    x='site_name:O',
    y='model_spec:O',
    color='num_device:Q'
)
st.altair_chart(c1, use_container_width=True)

# ———————————————————— heatmap for site, device, date range——————————————————————————————————————————————————————

df_range = pd.DataFrame(df[mask]['date_range'].groupby([df['site_name'],df['规格型号']]).mean()).reset_index()
c2 = alt.Chart(df_range).mark_rect().encode(
    x='site_name:O',
    y='规格型号:O',
    color='date_range:Q'
)
st.altair_chart(c2, use_container_width=True)

#_________________________________________ something real __________________________________________
df['model_spec'].astype(int)
table = pd.pivot_table(df[df['状态'] != '完成'], values=['date_range'], index=['model_name', 'model_spec','状态'],
                    columns=['site_name'], aggfunc={'date_range':  np.min}).fillna('-')

st.table(table)