import streamlit as st
import pandas as pd
import calendar
import altair as alt


# Load Excel files
def load_data():
    file_2023 = 'Sales2023.xlsx'
    file_2024 = 'Sales2024.xlsx'

    data_frames = []
    for file in [file_2023, file_2024]:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            sheet_data = xls.parse(sheet_name)
            data_frames.append(sheet_data)

    combined_data = pd.concat(data_frames)
    return combined_data


# Preprocess data
def preprocess_data(data):
    data['Month'] = pd.to_datetime(data['Month'], errors='coerce')
    data = data.dropna(subset=['Month'])
    data = data.sort_values(by='Month')
    return data


# Load and preprocess data
data = load_data()
data = preprocess_data(data)

# Streamlit application
st.title("Vehicle Sales Dashboard")

# Sidebar for user input
make_options = sorted(data['Make'].unique())
selected_makes = st.sidebar.multiselect("Select Vehicle Makes:", make_options, default=make_options[:1])

model_options = sorted(data[data['Make'].isin(selected_makes)]['Model'].unique())
selected_models = st.sidebar.multiselect("Select Vehicle Models:", model_options, default=model_options[:1])

st.sidebar.subheader("Total Sales Analysis")
selected_make = st.sidebar.selectbox("Select a Make for Total Sales Analysis:", make_options)

st.sidebar.subheader("Select Month and Year")
unique_years = data['Month'].dt.year.unique()
selected_year = st.sidebar.selectbox("Select Year:", sorted(unique_years))
unique_months = data[data['Month'].dt.year == selected_year]['Month'].dt.month_name().unique()
selected_month = st.sidebar.selectbox("Select Month:", sorted(unique_months, key=lambda x: pd.to_datetime(x, format='%B')))

filtered_data = data[(data['Make'].isin(selected_makes)) & (data['Model'].isin(selected_models))]

# Sales trends plot
st.subheader(f"Sales Trends for Selected Makes and Models")
if not filtered_data.empty:
    sales_trends = alt.Chart(filtered_data).mark_line(point=True).encode(
        x=alt.X('Month:T', title="Month"),
        y=alt.Y('Net Sales:Q', title="Net Sales"),
        color=alt.Color('Model:N', title="Model"),
        tooltip=['Make', 'Model', 'Net Sales', 'Month']
    ).properties(
        width=800,
        height=400,
        title="Sales Trends for Selected Makes and Models"
    )
    st.altair_chart(sales_trends, use_container_width=True)

# Total sales per model for a make
make_data = data[data['Make'] == selected_make]
grouped_data = make_data.groupby(['Model', 'Month']).sum().reset_index()

if not grouped_data.empty:
    total_sales_chart = alt.Chart(grouped_data).mark_line(point=True).encode(
        x=alt.X('Month:T', title="Month"),
        y=alt.Y('Net Sales:Q', title="Total Sales"),
        color=alt.Color('Model:N', title="Model"),
        tooltip=['Model', 'Net Sales', 'Month']
    ).properties(
        width=800,
        height=400,
        title=f"Total Sales Per Month for Each Model of {selected_make}"
    )
    st.altair_chart(total_sales_chart, use_container_width=True)

# Filtered data table
st.subheader("Filtered Data")
filtered_data['Month'] = filtered_data['Month'].dt.strftime('%B %Y')
st.dataframe(filtered_data[["Make", "Model", "Net Sales", "Month"]])

# Best-selling models bar chart
selected_month_data = data[(data['Month'].dt.year == selected_year) & (data['Month'].dt.month_name() == selected_month)]

st.subheader(f"Best-Selling Models in {selected_month} {selected_year}")
if not selected_month_data.empty:
    best_selling_models = (selected_month_data.groupby('Model')['Net Sales']
                           .sum()
                           .sort_values(ascending=False)
                           .head(30)
                           .reset_index())

    best_selling_chart = alt.Chart(best_selling_models).mark_bar().encode(
        x=alt.X('Net Sales:Q', title="Total Sales"),
        y=alt.Y('Model:N', title="Model", sort='-x'),
        tooltip=['Model', 'Net Sales']
    ).properties(
        width=800,
        height=600,
        title=f"Top 30 Best-Selling Models in {selected_month} {selected_year}"
    )
    st.altair_chart(best_selling_chart, use_container_width=True)
else:
    st.write("No data available for the selected month and year.")

# Total sales by year for a make
st.sidebar.subheader("Total Sales of Models by Year")
selected_year_for_make = st.sidebar.selectbox("Select Year for Make Analysis:", sorted(unique_years))
selected_make_for_year = st.sidebar.selectbox("Select Make for Year Analysis:", make_options)

data_for_year_make = data[(data['Month'].dt.year == selected_year_for_make) & (data['Make'] == selected_make_for_year)]

if not data_for_year_make.empty:
    total_sales_models = (data_for_year_make.groupby('Model')['Net Sales']
                          .sum()
                          .sort_values(ascending=False)
                          .reset_index())

    total_sales_chart = alt.Chart(total_sales_models).mark_bar().encode(
        x=alt.X('Net Sales:Q', title="Total Sales"),
        y=alt.Y('Model:N', title="Model", sort='-x'),
        tooltip=['Model', 'Net Sales']
    ).properties(
        width=800,
        height=600,
        title=f"Total Sales of Models for {selected_make_for_year} in {selected_year_for_make}"
    )
    st.altair_chart(total_sales_chart, use_container_width=True)
else:
    st.write(f"No data available for {selected_make_for_year} in {selected_year_for_make}.")

# Monthly sales of a model
st.sidebar.subheader("Monthly Sales of a Model")

# Ensure 'Make' and 'Model' are properly formatted as strings
data['Make'] = data['Make'].astype(str)
data['Model'] = data['Model'].astype(str)

# Create concatenated make and model column
data['Make_Model'] = data['Make'] + ' ' + data['Model']

# Dropdown for Make_Model selection
make_model_options = sorted(data['Make_Model'].unique())
selected_make_model = st.sidebar.selectbox("Select Make & Model for Monthly Sales Analysis", make_model_options)

if selected_make_model:
    # Filter data for the selected Make_Model
    make_model_data = data[data['Make_Model'] == selected_make_model]

    if not make_model_data.empty:
        # Group data by year and month
        monthly_sales = make_model_data.groupby([make_model_data['Month'].dt.year.rename('Year'),
                                                 make_model_data['Month'].dt.month.rename('Month')])['Net Sales'].sum().reset_index()

        # Add full month names for better x-axis labels
        monthly_sales['Month_Name'] = monthly_sales['Month'].apply(lambda x: calendar.month_name[x])

        # Plot monthly sales
        monthly_sales_chart = alt.Chart(monthly_sales).mark_bar().encode(
            x=alt.X('Month_Name:N', sort=list(calendar.month_name[1:]), title="Month"),
            y=alt.Y('Net Sales:Q', title="Net Sales"),
            color=alt.Color('Year:N', title="Year"),
            tooltip=['Year', 'Month_Name', 'Net Sales']
        ).properties(
            width=800,
            height=400,
            title=f"Monthly Sales of {selected_make_model} (2023 & 2024)"
        )
        st.altair_chart(monthly_sales_chart, use_container_width=True)
    else:
        st.write(f"No sales data available for {selected_make_model}.")
