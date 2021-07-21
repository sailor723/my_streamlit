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

# initialize Carrier, or read from sheet Carrier if data exists——————————————————————————————————————————————————————————————————————————————

sheetCarrier = wb.sheets('Carrier')
sheetCarrierTableRange = sheetCarrier.range('A1').options(expand='table',empty = 'NA',numbers = int)
try:
    df_post = pd.DataFrame(sheetCarrierTableRange.value[1:], columns=sheetCarrierTableRange.value[0])  # read df_post if data exists

    df_post = df_post[df_post.columns.drop('NA')]
    df_post['发送日期'] = pd.to_datetime(df_post['发送日期'])
    df_post[df_post['接收日期'] != 'NA']['接收日期'] = pd.to_datetime(df_post[df_post['接收日期'] != 'NA']['接收日期']).dt.date

except:
    df_post = pd.DataFrame(columns=['运单号','序号','始发库','目标中心','发送日期','接收日期','s_lon','s_lat','r_lon','r_lat'])     # initial sheetCarrier if on data
    sheetCarrier.range('A1').options(numbers = str,transpose=False,index=True,dates=dt.date).value = df_post
    sheetCarrier.autofit()



def load_data_geo():

        sheet_name = "Geo"
        # excel_file = '器械模拟20210623 - 副本.xls'
        df = pd.read_excel(excel_file,
            sheet_name=sheet_name,
            keep_default_na = False
            )
        return df
df_geo = load_data_geo()
#————————————————Geo Chart ————————————————————————————————————————————————

# df_geo['date_range_mean'] = df[mask].groupby('中心名称1')['date_range'].mean().astype(int)
# df_geo.fillna(0, inplace=True)
Source_RGB = [0, 255, 0, 40]

Target_RGB = [240, 100, 0, 40]

try:
    ALL_LAYERS = {
        "Device Number": pdk.Layer(
            "HexagonLayer",
            data=df,
            get_position=["lon", "lat"],
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
            get_source_color=Source_RGB,
            get_target_color=Target_RGB,
            pickable=True,
            )
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
                                "zoom": 3.5,
                                "pitch": 0,
                                },
            layers=selected_layers,

            tooltip={

                'html': '<b>中心名称:</b> {site_name} 器械数量:{elevationValue} <b>  \
                         <br />效期时间:</b> {date_range_mean} <br />始发中心:</b> {始发库} <br />接收中心:</b> {目标中心}  ',

                'style': {
                    'color': 'white'
                }
                }

        ))
    else:
        st.error("Please choose at least one layer above.")
except URLError as e:
    st.error("""
        **This demo requires internet access.**

        Connection error: %s
    """ % e.reason)

# st.map(df_geo)


# generate post_slips, need input device_list, post_id, post_data, target_site ——————————————————————————————————————————————————————————————————
# parametor: device_id_list, post_id, post_data, target_site, surgery_date)

def device_sent (device_id_list, post_id, post_date,target_site):
    global df, df_post

    if post_id not in df_post['运单号'].tolist():

        for device_id in device_id_list:
                print('device_id_list', device_id_list)

                if device_id in df.index:

                    if (target_site != df.loc[device_id,'site_name']) & (target_site in df.site_name.tolist()):


                        print('device_id', device_id)

                        if df.loc[device_id, '状态'] == '可用':

                            series = pd.Series({"运单号":post_id,"序号":device_id,"始发库":df.loc[device_id,
                                                'site_name'],"目标中心":target_site,"发送日期":post_date,"接收日期":'NA',
                                               's_lon': df.loc[device_id]['lon'],
                                               's_lat': df.loc[device_id]['lat'],
                                               'r_lon': df[df.site_name == target_site]['lon'].unique()[0],
                                               'r_lat': df[df.site_name == target_site]['lat'].unique()[0],
                                                })


                            df_post = df_post.append(series,ignore_index=True)
                            df.loc[device_id, 'lon'] = df[df.site_name == target_site]['lon'].unique()[0]
                            update_excel(device_id, column_name = 'lon', sheet_name = sheetMain)#  update 状态

                            df.loc[device_id, 'lat'] = df[df.site_name == target_site]['lat'].unique()[0]
                            update_excel(device_id, column_name = 'lat', sheet_name = sheetMain)#  update 状态

                            df.loc[device_id, 'site_name'] = target_site
                            # df.loc[device_id, '中心名称1'] = target_site[3:]
                            df.loc[device_id,'状态'] = 'in carrier'
                            update_excel(device_id, column_name = '状态', sheet_name = sheetMain)#  update 状态
                            update_excel(device_id, column_name = 'site_name', sheet_name = sheetMain)#  update 状态
                            update_excel(device_id, column_name = '中心名称1', sheet_name = sheetMain)#  update 状态

                            row = str(len(df_post) + 1)            # find the last row to be add in the sheet

                            range1 = 'A' + row                      # find starting cell to be add in the sheet

                            sheetCarrier.range(range1).options(numbers = str,transpose=True,index=False).value = df_post.loc[len(df_post) - 1]
                            sheetCarrier.autofit()                   # feel sheet along row by transposoe = True

                            print('成功发送运单。设备号%s, 运单号%s'%(device_id, post_id))

                        else:
                            print('%s设备状态错误，请确认: %s' %(device_id, df.loc[device_id,'状态']))
                            return ('4')
                    else:
                        print('中心名称错误，请确认: %s' %target_site)
                        return ('3')
                else:
                    print('设备号错误，请确认:%s' %(device_id))
                    return ('2')


    else:
        print('运单号重复，请检查 %s' %(post_id))
        return ('1')

    return('0')


# receive and update status, will check device_id, post_id before update——————————————————————————————————————————————————————————————————————————————
# input post_id, post_data, target_site
# receive and update status, will check device_id, post_id before update
# input post_id, post_data, target_site
def device_receive ( post_id,target_site, receive_date):

    global df, df_post

    if post_id in df_post['运单号'].tolist():

        for i in df_post[df_post['运单号'] == post_id].index.tolist():

            device_id = df_post.loc[i]['序号']

            if device_id in df.index:

                if target_site == df.loc[device_id,"site_name"]:

                    if df.loc[device_id,'状态'] == 'in carrier':

                        df.loc[device_id,'状态'] = '可用'
                        update_excel(device_id, column_name = '状态', sheet_name = sheetMain)      #  update 状态
                        print('成功接收运单。设备号:%s, 运单号%s:，状态更新完毕:%s' %(device_id,post_id,df.loc[device_id,'状态']))

                        df_post.loc[i, '接收日期'] = receive_date

                        col = df_post.columns.tolist().index('接收日期')

                        col = chr(ord('A') + col + 1)

                        row = str(df_post[df_post['序号'] == device_id].index.tolist()[0] + 2)    # to find rows in the sheetCarrier

                        range1 = col + row                           # find cell of receive date in the sheetCarrier

                        sheetCarrier.range(range1).options(index=False).options(numbers = str).value = df_post.loc[i, '接收日期']

                        sheetCarrier.autofit()

                    else:

                        print('%s状态错误，请确认: %s' %(device_id, df.loc[device_id,'状态']))

                else:
                    print("%s中心名称错误，请确认:" %(target_site))
                    return('2')

            else:
                print('%s设备号错误，请确认:' %(device_id))
                return('1')

            df_post.loc[i, '接收日期'] = 'receive_date'

    else:
        print('运单号错误，请确认:')
        return('4')
    return('0')

# update status ,input df, device_id, column_name, sheet——————————————————————————————————————————————————————————————————————————————————————————————————————————————

def update_excel (device_id,column_name,sheet_name):
    global df, df_post

    if device_id in df.index:

        if column_name in df.keys():
            col = np.where(df.keys()== column_name)[0][0]
            col = chr(col + 64)
            print('col:', col)
            row = str(df.loc[device_id,'index'] + 2 )
            range1 = 'A' + row + ':' + col + row

            if df.loc[device_id, '状态'] == 'in plan':

                sheet_name.range(range1).color = 209,238,238
                sheet_name.range(range1).api.Borders(9).Weight = 2


            elif df.loc[device_id, '状态'] == 'in carrier':

                sheet_name.range(range1).color = 255,240,245
                sheet_name.range(range1).api.Borders(9).Weight = 2

            elif df.loc[device_id, '状态'] == '日期告警':

                sheet_name.range(range1).color = 255,248,220
                sheet_name.range(range1).api.Borders(9).Weight = 2

            else:
                sheet_name.range(range1).color = 224,238,224
                sheet_name.range(range1).api.Borders(9).Weight = 2



            sheet_name.range(col+row).options(index=False).options(numbers = str).value = df.loc[device_id, column_name]
            sheetMain.autofit()
            return('0')

        else:
            print('列名错误，请确:', column_name)
            return('1')

    else:
        print('设备号错误，请确认: ', device_id)
        return('1')

#———————————————————————————————————————— surgery plan   ————————————————————————————————————————————————————————————

def surgery_plan (device_id,target_site,patient_id,target_date,sheet_name):
    global df, df_post

    if device_id in df.index:


        if target_site == df.loc[device_id,"site_name"]:
            if df.loc[device_id,'状态'] == '可用':
                df.loc[device_id,'状态'] = 'in plan'                               # updatee 状态
                update_excel(device_id, column_name = '状态', sheet_name = sheet_name )     # update 状态
                df.loc[device_id,'受试者编号'] = patient_id                            # update 筛选号
                update_excel(device_id, column_name = '受试者编号', sheet_name = sheet_name)    # update 筛选号
                df.loc[device_id,'使用日期'] = target_date                         # update 手术日期
                update_excel(device_id, column_name = '使用日期', sheet_name = sheet_name)  # update 手术日期
                print('成功制定手术计划。设备号:%s, 手术日期%s:，状态更新完毕:%s' %(device_id,df.loc[device_id,'使用日期'],df.loc[device_id,'状态']))
                return('0')
            else:
                print('设备不可用:%s %s' %(device_id,df.loc[device_id, '状态']))
                return('3')
        else:
            print('中心名称错误，请确认 %s  %s :' %(device_id,df.loc[device_id, 'site_name']))
            return('2')
    else:
        print('设备号错误，请确认: ', device_id)
        return('1')


#———————————————————————————————————————— surgery complete  ————————————————————————————————————————————————————————————

def surgery_complete (device_id,target_site):
    global df, df_post
    if device_id in df.index:
        if target_site == df.loc[device_id,"site_name"]:
            if df.loc[device_id,'状态'] == 'in plan':
                df.loc[device_id,'状态'] = '完成'
                update_excel(device_id, column_name = '状态', sheet_name = sheetMain)
                update_excel(device_id, column_name = '受试者编号', sheet_name = sheetMain)
                update_excel(device_id, column_name = '使用日期', sheet_name = sheetMain)
                print('成功完成手术，设备号:%s, 手术日期%s:，状态更新完毕:%s' %(device_id,df.loc[device_id,'使用日期'],df.loc[device_id,'状态']))
                return('0')

            else:
                print('设备没有计划，请确认:', df.loc[device_id, '状态'])
                return('1')
        else:
            print('中心名称错误，请确认:', target_site)
            return('1')
    else:
        print('设备号错误，请确认: ', device_id)
        return('1')


#————————————————————————————————————————logistic plan UI   device send   ————————————————————————————————————————————————————————————


sent1, sent2, rec = st.beta_columns(3)

with sent1:


    st.subheader("器械发送调度")

    device_selection = st.selectbox(
    "请选择器械名称",
    df['器械名称'].unique().tolist())


    list1 = df[df['器械名称'] == device_selection]['model_name'].unique().tolist()
    list1.sort()

    model_selection = st.selectbox("请选择器械型号", list1)

    list1 = df[(df['器械名称'] == device_selection) & (df['model_name'] == model_selection)]['model_spec'].unique().tolist()
    list1.sort()

    spec_selection = st.selectbox( "请选择器械规格", list1)

    df_selection = df[(df['器械名称'] == device_selection) & (df['model_name'] == model_selection) & (df['model_spec'] == spec_selection)]

    original_site = st.selectbox(
    "请选择发货中心",
    df_selection.site_name.unique().tolist())

    target_site = st.selectbox(
    "请选择收货中心",
    df.site_name.unique().tolist())

with sent2:
    st.subheader("器械发送运单")
    if original_site != target_site:

        device_list = st.multiselect( "请选择设备",
                            df_selection[(df.site_name == original_site) & (df_selection['状态'] == '可用') ] ['序号'].tolist())

        carrier_number = st.text_input('Please input carrier number')


        carrier_date = st.date_input(
            "请输入发货日期",
            datetime.date.today())

        if st.button('请确认以上信息'):
            st.write('谢谢')
            if device_sent (device_list, carrier_number, carrier_date, target_site) == '0':
                st.write('system update completed!')
            else:
                st.write('请重新选择')
                st.markdown('Streamlit is **_really_ cool**.')
    else:
        # st.text('收发中心不能是同一个。请重新选择')
        st.markdown('Streamlit is **_really_ cool**.')

#————————————————————————————————————————logistic plan UI  device receive   ————————————————————————————————————————————————————————————

with rec:
    st.subheader("器械接收管理")
    receive_number = st.selectbox(
    "请选择运单号",
    df_post['运单号'].unique().tolist())

    st.text("device_list: %s" %(df_post[df_post['运单号'] == receive_number]['序号'].tolist()))


    st.text("始发中心: %s" %(df_post[df_post['运单号'] == receive_number]['始发库'].unique()))
    st.text("接收中心: %s" %(df_post[df_post['运单号'] == receive_number]['目标中心'].unique()))


    receive_date = st.date_input(
        "请输入收货日期",
         datetime.date.today())

    st.write('收货日期:', receive_date)

    if st.button('请确认收货信息'):
        st.write('谢谢')
        if device_receive(post_id=receive_number,target_site=df_post[df_post['运单号'] == receive_number]['目标中心'].unique(), receive_date = receive_date) == '0':
            st.write('系统收货完成！')
    else:
        st.write('请重新选择')

df_post10 = df_post.sort_values(by=['序号', '运单号'], axis=0, ascending=[True,True])

with st.beta_expander("请见详细运单信息"):
    st.table(df_post10[['运单号','序号','始发库','目标中心','发送日期','接收日期']])

# ————————————————————————————————————————surgery UI surgery plan   ————————————————————————————————————————————————————————————


col4, col5, col6 = st.beta_columns([2, 2, 2])

with col4:
    st.subheader("手术计划")
    surgery_site = st.selectbox(
    "请选择研究中心",
    df.site_name.unique().tolist())


    surgery_device = st.selectbox(
        "请选择手术器械序号",
         df[(df.site_name == surgery_site) & (df['状态'] == '可用')] ['序号'].tolist())


    patient_id = st.text_input('受试者筛选号：')
    st.write('受试者筛选号', patient_id)

    surgery_date = st.date_input(
        "请输入手术日期",
        datetime.date.today() + datetime.timedelta(days=30))

    st.write('手术日期:', surgery_date)

    if st.button('Please confirm3'):
        st.write('谢谢')
        if surgery_plan (surgery_device,surgery_site,patient_id,surgery_date,sheetMain) == '0':
            st.write('system update completed!')
        else:
            st.write('手术计划设定失败，请重新计划')

with col5:
        st.image(".\\carrier.jfif")


#————————————————————————————————————————surgery UI  surgery complete   ————————————————————————————————————————————————————————————

with col6:
    st.subheader("完成手术")
    surgery_site = st.selectbox(
    "请选择研究中心",
    df[df['状态'] == 'in plan'].site_name.unique().tolist(),
    key = 'surgey_site_complete')

    surgery_device = st.selectbox(
        "请选择手术器械序号",
         df[(df.site_name == surgery_site) & (df['状态'] == 'in plan')] ['序号'].tolist(),
         key='surgery_complete')

    surgery_date = st.date_input(
        "请输入手术日期",
        (datetime.date.today() + datetime.timedelta(days=30)),
        key = 'surgery_complete_date')


    st.text("受试者筛选号: %s " %(df[df['序号'] == surgery_device ] ['受试者编号'].tolist()))


    if st.button('请确认手术完成信息'):

        st.write('谢谢')

        if surgery_complete (surgery_device,surgery_site) == '0':
            st.write('系统处理完成！')
    else:
        st.write('请重新选择')



df_post10 = df_post.sort_values(by=['序号', '运单号'], axis=0, ascending=[True,True])

with st.beta_expander("计划和完成的器械"):
    st.table( df[(df['状态'] == 'in plan') | (df['状态'] == '完成')][['site_name','器械名称','受试者编号','使用日期','状态']].sort_values('序号'))

with st.beta_expander("完整的数据"):
    st.dataframe(df)