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
df["Weekending"] = pd.to_datetime(df["Weekending"])
df["Weekending1"] = df["Weekending"].dt.date

# ---- SIDEBAR ----
st.sidebar.header("Please Filter Here:")
year = st.sidebar.multiselect(
    "Select the Year:",
    options=sorted(df["Year"].unique()),
    default=sorted(df["Year"].unique())
)

month = st.sidebar.multiselect(
    "Select the Month:",
    options=sorted(df["Month"].unique()),
    default=sorted(df["Month"].unique()),
)

week = st.sidebar.multiselect(
    "Select the Week:",
    options=sorted(df["Weekending1"].unique()),
    default=sorted(df["Weekending1"].unique()),
)

cell = st.sidebar.multiselect(
    "Select the Cell:",
    options=sorted(df["Cell"].unique()),
    default=sorted(df["Cell"].unique())
)

resource = st.sidebar.multiselect(
    "Select the Resource:",
    options=sorted(df["Resource"].unique()),
    default=sorted(df["Resource"].unique())
)

outcome = st.sidebar.multiselect(
    "Select the Outcome:",
    options=sorted(df["Outcome"].unique()),
    default=sorted(df["Outcome"].unique())
)

df_selection = df.query(
    "Year == @year & Month ==@month & Weekending1 ==@week & Cell == @cell & Resource == @resource & Outcome == @outcome"
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
