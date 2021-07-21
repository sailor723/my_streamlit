# from streamlit.my_test.new_second import load_data
import streamlit as st
import pandas as pd
import pydeck as pdk

from urllib.error import URLError

sheet_name = 'Whole_geo'
excel_file = 'C:\\Users\\weiping\\Documents\\eclincloud2020\\training\\Python_dev\\code\\streamlit\\my_test\\器械模拟20210623 - 副本.xls'

df_whole_geo = pd.read_excel(excel_file, sheet_name=sheet_name,keep_default_na = False)


sheet_name = 'Geo'
df_geo = pd.read_excel(excel_file, sheet_name=sheet_name)



try:
    ALL_LAYERS = {
        "Device Number": pdk.Layer(
            "HexagonLayer",
            data=df_whole_geo,
            get_position=["lon", "lat"],
            radius=4000,
            elevation_scale=400,
            elevation_range=[0, 1000],
            extruded=True,
        ),
        "Date_Range": pdk.Layer(
            "ScatterplotLayer",
            data=df_geo,
            get_position=["lon", "lat"],
            get_color=[200, 30, 0, 160],
            get_radius="date_range_mean",
            radius_scale=90,
        ),
        # "ColumnLayer": pdk.Layer(
        #     "ColumnLayer",
        #     data=df_geo,
        #     get_position=["lon", "lat"],
        #     get_elevation="date_range_mean",
        #     elevation_scale=500,
        #     radius=1200,
        #     get_fill_color=[200, 30, 0, 160],
        #     # get_fill_color=["mrt_distance * 10", "mrt_distance", "mrt_distance * 10", 140],
        #     pickable=True,
        #     auto_highlight=True,
        # )
        # "Bart Stop Names": pdk.Layer(
        #     "TextLayer",
        #     data=from_data_file("bart_stop_stats.json"),
        #     get_position=["lon", "lat"],
        #     get_text="name",
        #     get_color=[0, 0, 0, 200],
        #     get_size=15,
        #     get_alignment_baseline="'bottom'",
        # ),
        # "Outbound Flow": pdk.Layer(
        #     "ArcLayer",
        #     data=from_data_file("bart_path_stats.json"),
        #     get_source_position=["lon", "lat"],
        #     get_target_position=["lon2", "lat2"],
        #     get_source_color=[200, 30, 0, 160],
        #     get_target_color=[200, 30, 0, 160],
        #     auto_highlight=True,
        #     width_scale=0.0001,
        #     get_width="outbound",
        #     width_min_pixels=3,
        #     width_max_pixels=30,
        # ),
    }
    st.sidebar.markdown('### Map Layers')
    selected_layers = [
        layer for layer_name, layer in ALL_LAYERS.items()
        if st.sidebar.checkbox(layer_name, True)]
    if selected_layers:
        st.pydeck_chart(pdk.Deck(
            map_style="mapbox://styles/mapbox/light-v9",
            initial_view_state={"latitude": 31.2164668532657,
                                "longitude": 121.472886,
                                "zoom": 6,
                                "pitch": 50,
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


date_selection = st.sidebar.slider('date range to expire',   ## display date range selection
                            min_value = min(dates),
                            max_value = max(dates),
                            value = (min(dates), max(dates)))