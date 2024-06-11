import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

# Streamlit application
st.title("Monthly Opportunity Analysis")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the data from the uploaded file
    data = pd.read_excel(uploaded_file)

    # Convert the 'Close Date' column to datetime format
    data['Close Date'] = pd.to_datetime(data['Close Date'])

    # Extract month and year for columns
    data['Month'] = data['Close Date'].dt.strftime('%Y-%m')

    # Create a pivot table with the count of 'Opportunity Name' for each month
    count_pivot_table = data.pivot_table(index='Opportunity Name', columns='Month', values='Contact Name', aggfunc='count', fill_value=0)

    # Display the pivot table
    st.subheader("Monthly Opportunity Counts")
    st.dataframe(count_pivot_table)

    # Optionally, display the pivot table as a downloadable CSV
    st.download_button(
        label="Download data as CSV",
        data=count_pivot_table.to_csv().encode('utf-8'),
        file_name='monthly_opportunity_counts.csv',
        mime='text/csv',
    )

    # Filter the data for only 'Won' milestones
    won_data = data[data['Milestone'] == 'Won']

    # Group by 'Month' and count cumulative wins
    won_data['YearMonth'] = won_data['Close Date'].dt.to_period('M')
    cumulative_wins = won_data.groupby('YearMonth').size().cumsum().reset_index(name='Cumulative Wins')

    # Plot the cumulative wins by month
    st.subheader("Cumulative Wins by Month")
    plt.figure(figsize=(12, 6))
    plt.plot(cumulative_wins['YearMonth'].astype(str), cumulative_wins['Cumulative Wins'], marker='o')
    plt.title('Cumulative Wins by Month')
    plt.xlabel('Month')
    plt.ylabel('Cumulative Wins')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()

    # Display the plot in Streamlit
    st.pyplot(plt)

    # Convert USD values to GBP using the exchange rate
    exchange_rate = 1.27
    data['Value in GBP'] = data.apply(lambda row: row['Estimated Value'] if row['Currency'] == 'GBP' else row['Estimated Value'] / exchange_rate, axis=1)

    # Pivot table with months as columns and opportunity names as rows for GBP values
    pivot_table_with_contact = data.pivot_table(index=['Opportunity Name', 'Contact Name'], columns='Month', values='Value in GBP', aggfunc='sum', fill_value=0)

    # Display the pivot table with opportunity values in GBP
    st.subheader("Monthly Opportunity Values in GBP")
    st.dataframe(pivot_table_with_contact)

    # Optionally, display the pivot table as a downloadable CSV
    st.download_button(
        label="Download opportunity values as CSV",
        data=pivot_table_with_contact.to_csv().encode('utf-8'),
        file_name='monthly_opportunity_values.csv',
        mime='text/csv',
    )

    # Sum the values by month for the chart
    monthly_values = data.groupby('Month')['Value in GBP'].sum().reset_index()

    # Plot the monthly opportunity values in GBP
    st.subheader("Monthly Opportunity Values in GBP")
    plt.figure(figsize=(12, 6))
    plt.bar(monthly_values['Month'], monthly_values['Value in GBP'])
    plt.title('Monthly Opportunity Values in GBP')
    plt.xlabel('Month')
    plt.ylabel('Total Value in GBP')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()

    # Display the plot in Streamlit
    st.pyplot(plt)

    # Calculate cumulative values
    monthly_values['Cumulative Value'] = monthly_values['Value in GBP'].cumsum()

    # Plot the cumulative value of opportunities in GBP
    st.subheader("Cumulative Opportunity Value in GBP")
    plt.figure(figsize=(12, 6))
    plt.plot(monthly_values['Month'], monthly_values['Cumulative Value'], marker='o')
    plt.title('Cumulative Opportunity Value in GBP')
    plt.xlabel('Month')
    plt.ylabel('Cumulative Value in GBP')
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.tight_layout()

    # Display the plot in Streamlit
    st.pyplot(plt)

else:
    st.write("Please upload an Excel file to proceed.")
