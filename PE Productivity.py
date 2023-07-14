import pandas as pd
import plotly.express as px
import streamlit as st
import base64
import io

st.set_page_config(page_title="Sales Dashboard", page_icon=":bar_chart:", layout="wide")

@st.cache_data
def get_data_from_excel():

    df = pd.read_excel(io='PE Completions Raw.xlsx',
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
    "Year == @year & Month ==@month & Cell == @cell & Resource == @resource & Outcome == @outcome"
)
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

df_filtered_sidebar = df[
    df["Year"].isin(year) &
    df["Month"].isin(month) &
    df["Cell"].isin(cell) &
    df["Resource"].isin(resource) &
    df["Outcome"].isin(outcome)
]




st.markdown("### Detailed Data View")
st.dataframe(df_filtered_sidebar[['Work Order No.','Resolution Type','Job Plan Code (WO)','Cell','Resource','Week','Program Code','All Address Field','PR Total Cost']], use_container_width=True)


# Convert the filtered DataFrame into a XLSX file in memory
to_excel = io.BytesIO()
df_filtered_sidebar.to_excel(to_excel, index=False, header=True)
to_excel.seek(0)

# Encode XLSX file in base64
b64 = base64.b64encode(to_excel.read()).decode()

# Generate download link
href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="follow_up_data.xlsx">Download file in XLSX</a>'
st.markdown(href, unsafe_allow_html=True)


# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
