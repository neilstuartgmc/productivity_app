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

# Dynamic month filter
year = st.sidebar.multiselect(
    "Year:",
    options=df["Year"].unique(),
    default=[]

)

# Filter the DataFrame with selected years
year_selec = df[df["Year"].isin(year)]

# unique years
weekending_selec = year_selec["Weekending1"].unique()
cell_selec = year_selec["Cell"].unique()
resource_selec = year_selec["Resource"].unique()
outcome_selec = year_selec["Outcome"].unique()



if year:
    weekending = st.sidebar.multiselect(
        "Weekending:",
        options=sorted(weekending_selec),
        default=[]
    )
else:
    weekending= []


if weekending:
    cell = st.sidebar.multiselect(
        "Cell:",
        options=sorted(df[df["Weekending1"].isin(weekending)]["Cell"].unique()),
        default=sorted(df[df["Weekending1"].isin(weekending)]["Cell"].unique())

)
else:
    cell = []

if cell:
    resource = st.sidebar.multiselect(
        "Resource:",
        options=sorted(df[df["Cell"].isin(cell)]["Resource"].unique()),
        default=sorted(df[df["Cell"].isin(cell)]["Resource"].unique())
)
else:
    resource = []

if resource:
    outcome = st.sidebar.multiselect(
        "Outcome:",
        options=sorted(df[df["Resource"].isin(resource)]["Outcome"].unique()),
        default=sorted(df[df["Resource"].isin(resource)]["Outcome"].unique())
)
else:
    outcome = []

df_filtered_sidebar = df[
    df["Year"].isin(year) &
    df["Weekending1"].isin(weekending) &
    df["Cell"].isin(cell) &
    df["Resource"].isin(resource) &
    df["Outcome"].isin(outcome)
]

df_filtered_sidebar = df_filtered_sidebar.sort_values(by=["Year","Resource"], ascending=False)


@st.experimental_memo
def convert_df(df):
    return df.to_csv(index=False).encode('utf-8')


csv = convert_df(df_filtered_sidebar)

# TOP KPI's
sales_value = round(df_filtered_sidebar["PR Total Cost"].sum(),2)
total_wo = int(df_filtered_sidebar["PR Total Cost"].count())
average_sale_by_wo = round(df_filtered_sidebar["PR Total Cost"].mean(), 2)

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

st.dataframe(df_filtered_sidebar)

st.download_button(
    "Press to Download",
    csv,
    "selected_file.csv",
    "text/csv",
    key='download-csv'
)

sales_by_cell = (
    df_filtered_sidebar.groupby(by=["Resource"]).count()[["Outcome"]]
)

fig_product_sales = px.bar(
    sales_by_cell,
    x="Outcome",
    y=sales_by_cell.index,
    orientation="h",
    title="<b>Sales by Resource</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_cell),
    template="plotly_white",
)
fig_product_sales.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

sales_value_by_cell = (
    df_filtered_sidebar.groupby(by=["Resource"]).count()[["PR Total Cost"]]
)

fig_sales_value_by_cell = px.bar(
    sales_value_by_cell,
    x="PR Total Cost",
    y=sales_by_cell.index,
    orientation="h",
    title="<b>Sales by Resource</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_value_by_cell),
    template="plotly_white",
)
fig_sales_value_by_cell.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)
left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_sales_value_by_cell, use_container_width=True)
right_column.plotly_chart(fig_product_sales, use_container_width=True)
# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                header {visibility: hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)
