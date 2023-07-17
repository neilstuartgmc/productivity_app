import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Sales Dashboard", page_icon=":bar_chart:", layout="wide")

@st.cache_data
def get_data_from_excel():

    df = pd.read_excel(io=r'PE Completions Raw.xlsx',
                        engine='openpyxl',
                        sheet_name='PE',
                        usecols='A:AZ')


    return df

df = get_data_from_excel()

# ---- SIDEBAR ----
st.sidebar.header("Please Filter Here:")
year = st.sidebar.multiselect(
    "Select the Year:",
    options=df["Year"].unique(),
    default=df["Year"].unique()
)

month = st.sidebar.multiselect(
    "Select the Month:",
    options=df["Month"].unique(),
    default=df["Month"].unique(),
)

week = st.sidebar.multiselect(
    "Select the Week:",
    options=df["Week"].unique(),
    default=[41],
)

cell = st.sidebar.multiselect(
    "Select the Cell:",
    options=df["Cell"].unique(),
    default=df["Cell"].unique()
)

resource = st.sidebar.multiselect(
    "Select the Resource:",
    options=df["Resource"].unique(),
    default=df["Resource"].unique()
)

outcome = st.sidebar.multiselect(
    "Select the Outcome:",
    options=df["Outcome"].unique(),
    default=df["Outcome"].unique()
)

df_selection = df.query(
    "Year == @year & Month ==@month & Week ==@week & Cell == @cell & Resource == @resource & Outcome == @outcome"
)

@st.experimental_memo
def convert_df(df):
   return df.to_csv(index=False).encode('utf-8')

csv = convert_df(df_selection)

# TOP KPI's
sales_value = int(df_selection["PR Total Cost"].sum())
total_wo = int(df_selection["PR Total Cost"].count())
average_sale_by_wo = round(df_selection["PR Total Cost"].mean(), 2)

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader("Total Sales:")
    st.subheader(f" € {sales_value:,}")
with middle_column:
    st.subheader("Total WO Complete:")
    st.subheader(f"{total_wo:,}")
with right_column:
    st.subheader("Average Sales Per WO:")
    st.subheader(f" € {average_sale_by_wo}")

st.dataframe(df_selection)

st.download_button(
   "Press to Download",
   csv,
   "selected_file.csv",
   "text/csv",
   key='download-csv'
)

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
