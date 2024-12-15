import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import calendar


# Load Excel files
def load_data():
    file_2023 = 'Sales2023.xlsx'
    file_2024 = 'Sales2024.xlsx'

    # Combine data from all sheets in both files
    data_frames = []
    for file in [file_2023, file_2024]:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            sheet_data = xls.parse(sheet_name)
            data_frames.append(sheet_data)

    # Concatenate all data
    combined_data = pd.concat(data_frames)
    return combined_data

# Preprocess data
def preprocess_data(data):
    data['Month'] = pd.to_datetime(data['Month'], errors='coerce')  # Handle invalid dates
    data = data.dropna(subset=['Month'])  # Drop rows with NaT values
    data = data.sort_values(by='Month')
    return data

# Load and preprocess data
data = load_data()
data = preprocess_data(data)

# Streamlit application
st.title("Vehicle Sales Dashboard")

# Sidebar for user input
make_options = sorted(data['Make'].unique())  # Sort alphabetically
selected_makes = st.sidebar.multiselect("Select Vehicle Makes:", make_options, default=make_options[:1])

# Filter model options based on selected makes
model_options = sorted(data[data['Make'].isin(selected_makes)]['Model'].unique())  # Sort alphabetically
selected_models = st.sidebar.multiselect("Select Vehicle Models:", model_options, default=model_options[:1])

# Additional functionality: Total sales per month for a selected make
st.sidebar.subheader("Total Sales Analysis")
selected_make = st.sidebar.selectbox("Select a Make for Total Sales Analysis:", make_options)
make_data = data[data['Make'] == selected_make]

# Selector for month and year
st.sidebar.subheader("Select Month and Year")
unique_years = data['Month'].dt.year.unique()
selected_year = st.sidebar.selectbox("Select Year:", sorted(unique_years))
unique_months = data[data['Month'].dt.year == selected_year]['Month'].dt.month_name().unique()
selected_month = st.sidebar.selectbox("Select Month:", sorted(unique_months, key=lambda x: pd.to_datetime(x, format='%B')))

# Filter data based on user selection
filtered_data = data[(data['Make'].isin(selected_makes)) & (data['Model'].isin(selected_models))]

# Plot sales trend for selected makes and models
st.subheader(f"Sales Trends for Selected Makes and Models")
fig1, ax1 = plt.subplots(figsize=(12, 6))  # Adjust figure size to make the x-axis wider

for make in selected_makes:
    for model in selected_models:
        model_data = filtered_data[(filtered_data['Make'] == make) & (filtered_data['Model'] == model)]
        if not model_data.empty:
            months = model_data['Month'].to_numpy()  # Convert to numpy array
            sales = model_data['Net Sales'].to_numpy()  # Convert to numpy array
            ax1.plot(months, sales, marker='o', linestyle='-', label=f"{make} {model}")

ax1.set_xlabel("Month")
ax1.set_ylabel("Net Sales")
ax1.set_title("Sales Trends for Selected Makes and Models")
ax1.legend()
plt.xticks(rotation=45)
plt.tight_layout()  # Ensure everything fits within the figure
st.pyplot(fig1)

# Group data by model and month to calculate total sales per month
grouped_data = make_data.groupby(['Model', 'Month']).sum().reset_index()

fig2, ax2 = plt.subplots(figsize=(12, 6))  # Adjust figure size to make the x-axis wider

models = grouped_data['Model'].unique()
for model in models:
    model_data = grouped_data[grouped_data['Model'] == model]
    months = model_data['Month'].to_numpy()  # Convert to numpy array
    sales = model_data['Net Sales'].to_numpy()  # Convert to numpy array
    ax2.plot(months, sales, marker='o', linestyle='-', label=model)

ax2.set_xlabel("Month")
ax2.set_ylabel("Total Sales")
ax2.set_title(f"Total Sales Per Month for Each Model of {selected_make}")
ax2.legend()
plt.xticks(rotation=45)
plt.tight_layout()  # Ensure everything fits within the figure
st.pyplot(fig2)

# Display filtered data table
st.subheader("Filtered Data")
filtered_data['Month'] = filtered_data['Month'].dt.strftime('%B %Y')  # Format month as "January 2023"
st.dataframe(filtered_data[["Make", "Model", "Net Sales", "Month"]])

# Filter data based on month and year
selected_month_data = data[(data['Month'].dt.year == selected_year) & (data['Month'].dt.month_name() == selected_month)]

# Display best-selling models for selected month and year as horizontal bar chart
st.subheader(f"Best-Selling Models in {selected_month} {selected_year}")
if not selected_month_data.empty:
    best_selling_models = (selected_month_data.groupby('Model')['Net Sales']
                           .sum()
                           .sort_values(ascending=False)
                           .head(30)  # Limit to top 30 models
                           .reset_index())
    best_selling_models = best_selling_models[::-1]  # Reverse the order for descending display

    fig3, ax3 = plt.subplots(figsize=(10, 6))
    ax3.barh(best_selling_models['Model'].astype(str), best_selling_models['Net Sales'], color='skyblue')
    ax3.set_xlabel("Total Sales")
    ax3.set_ylabel("Model")
    ax3.set_title(f"Top 30 Best-Selling Models in {selected_month} {selected_year}")
    plt.tight_layout()
    st.pyplot(fig3)
else:
    st.write("No data available for the selected month and year.")

# Selector for year and make, and display total sales for models
st.sidebar.subheader("Total Sales of Models by Year")
selected_year_for_make = st.sidebar.selectbox("Select Year for Make Analysis:", sorted(unique_years))
selected_make_for_year = st.sidebar.selectbox("Select Make for Year Analysis:", make_options)

data_for_year_make = data[(data['Month'].dt.year == selected_year_for_make) & (data['Make'] == selected_make_for_year)]
if not data_for_year_make.empty:
    total_sales_models = (data_for_year_make.groupby('Model')['Net Sales']
                          .sum()
                          .sort_values(ascending=False)
                          .reset_index())

    st.subheader(f"Total Sales of Models for {selected_make_for_year} in {selected_year_for_make}")
    fig4, ax4 = plt.subplots(figsize=(10, 6))
    ax4.barh(total_sales_models['Model'].astype(str), total_sales_models['Net Sales'], color='lightcoral')
    ax4.set_xlabel("Total Sales")
    ax4.set_ylabel("Model")
    ax4.set_title(f"Total Sales of Models for {selected_make_for_year} in {selected_year_for_make}")
    plt.tight_layout()
    st.pyplot(fig4)
else:
    st.write(f"No data available for {selected_make_for_year} in {selected_year_for_make}.")

# Ensure 'Model' is properly formatted as strings
data['Model'] = data['Model'].astype(str)

# Create concatenated make and model column
data['Make_Model'] = data['Make'] + ' ' + data['Model']

# Selector for model and display monthly sales for 2023 and 2024
st.sidebar.subheader("Monthly Sales of a Model")

# Specify file paths
file_2023 = "Sales2023.xlsx"  # Replace with the correct file path
file_2024 = "Sales2024.xlsx"  # Replace with the correct file path
try:
    excel_2023 = pd.ExcelFile(file_2023)
    excel_2024 = pd.ExcelFile(file_2024)

    # Combine data from all sheets in 2023
    data_2023 = []
    for sheet in excel_2023.sheet_names:
        df = excel_2023.parse(sheet)
        df['Month'] = pd.to_datetime(sheet, format='%B', errors='coerce')  # Convert sheet name to datetime
        data_2023.append(df)
    df_2023 = pd.concat(data_2023, ignore_index=True)

    # Combine data from all sheets in 2024
    data_2024 = []
    for sheet in excel_2024.sheet_names:
        df = excel_2024.parse(sheet)
        df['Month'] = pd.to_datetime(sheet, format='%B', errors='coerce')  # Convert sheet name to datetime
        data_2024.append(df)
    df_2024 = pd.concat(data_2024, ignore_index=True)

    # Combine 'Make' and 'Model' into a single column for both datasets
    df_2023['Make'] = df_2023['Make'].astype(str)
    df_2023['Model'] = df_2023['Model'].astype(str)
    df_2024['Make'] = df_2024['Make'].astype(str)
    df_2024['Model'] = df_2024['Model'].astype(str)

    df_2023['Make_Model'] = df_2023['Make'] + ' ' + df_2023['Model']
    df_2024['Make_Model'] = df_2024['Make'] + ' ' + df_2024['Model']

    # Aggregate Net Sales by Month and Make_Model
    net_sales_2023 = df_2023.groupby([df_2023['Month'].dt.month, 'Make_Model'])['Net Sales'].sum().unstack()
    net_sales_2024 = df_2024.groupby([df_2024['Month'].dt.month, 'Make_Model'])['Net Sales'].sum().unstack()

    # Ensure both datasets have the same index for plotting
    combined_index = net_sales_2023.index.union(net_sales_2024.index)
    net_sales_2023 = net_sales_2023.reindex(combined_index, fill_value=0)
    net_sales_2024 = net_sales_2024.reindex(combined_index, fill_value=0)

    # Create a sorted list of Make_Model options for the dropdown
    make_model_options = sorted(set(net_sales_2023.columns).union(net_sales_2024.columns))

    # Dropdown for Make_Model selection
    selected_make_model = st.sidebar.selectbox("Select Make & Model", make_model_options)

    if selected_make_model:
        # Prepare data for the selected Make_Model
        net_sales_2023_selected = net_sales_2023.get(selected_make_model, pd.Series(0, index=combined_index))
        net_sales_2024_selected = net_sales_2024.get(selected_make_model, pd.Series(0, index=combined_index))
        
        # Combine the data for plotting
        df_plot = pd.DataFrame({
            '2023': net_sales_2023_selected,
            '2024': net_sales_2024_selected
        }).reindex(range(1, 13), fill_value=0)  # Ensure all months are included
        
        # Month names for x-axis labels
        month_names = [calendar.month_name[m] for m in range(1, 13)]

        # Plot the data
        fig, ax = plt.subplots(figsize=(10, 6))
        df_plot.plot(kind='bar', ax=ax)
        ax.set_title(f"Monthly Net Sales for {selected_make_model} (2023 vs 2024)", fontsize=16)
        ax.set_xlabel("Month", fontsize=14)
        ax.set_ylabel("Net Sales", fontsize=14)
        ax.set_xticks(range(12))
        ax.set_xticklabels(month_names, rotation=45)
        ax.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()

        # Display the plot
        st.pyplot(fig)

except FileNotFoundError:
    st.error("One or both Excel files could not be found. Please check the file paths.")