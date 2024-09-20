import pandas as pd
import streamlit as st

# Streamlit application
st.set_page_config(page_title="Pipeline Review Data Processing", layout="wide")
st.title("📊 Pipeline Review Data Processing")

# Upload section
st.header("Upload Excel Files")
input_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

# User input for starting column index
starting_column = st.number_input("Enter the starting column index for sales data:", min_value=0)

if input_files:
    all_final_df = pd.DataFrame()  # Initialize an empty DataFrame to concatenate results
    download_links = []  # Initialize list to store download links

    for input_file in input_files:
        # Read the Excel file
        df = pd.read_excel(input_file, sheet_name="1) Base Business", header=3)

        # Select relevant columns based on user input
        df = df.iloc[:, list(range(10)) + list(range(starting_column, starting_column + 8))]

        # Rename the columns as needed
        df = df.rename(columns={"2024 NTS $.6": "Estimated_Organic_Gains", "2024 NTS $.7": "Estimated_Losses"})

        # Filter for rows with valid Estimated Organic Gains
        organic_gains_filtered = df[
            df["Estimated_Organic_Gains"].notna() & 
            (df["Estimated_Organic_Gains"] != 0) & 
            (df["Estimated_Organic_Gains"] != "-")
        ]

        # Create new DataFrame for Organic Gains
        new_data = pd.DataFrame(columns=[
            'SalesOrgCode', 'DivCode', 'ProfitCenter', 'CustomerCode', 
            'SalesDistrictCode', 'SalesRepName', 'CurrencyCode', 
            'SalesYear', 'Estimated_Organic_Gains', 'Estimated_Losses', 
            'Comments'
        ])
        
        new_data['SalesOrgCode'] = ['1030'] * len(organic_gains_filtered)
        new_data['DivCode'] = ['02'] * len(organic_gains_filtered)
        new_data['ProfitCenter'] = organic_gains_filtered['Profit Center'].values
        new_data['CustomerCode'] = organic_gains_filtered['Customer #'].values
        new_data['SalesDistrictCode'] = organic_gains_filtered['Sales District #'].values
        new_data['SalesRepName'] = organic_gains_filtered['Sales Rep'].values
        new_data['CurrencyCode'] = ['USD'] * len(organic_gains_filtered)
        new_data['SalesYear'] = ['2024'] * len(organic_gains_filtered)
        new_data['Estimated_Organic_Gains'] = organic_gains_filtered['Estimated_Organic_Gains'].values
        new_data['Estimated_Losses'] = organic_gains_filtered['Estimated_Losses'].values
        new_data['Comments'] = organic_gains_filtered['Comments (For September Review)'].values

        # Filter for rows with valid Estimated Losses
        loss_filtered = df[
            df["Estimated_Losses"].notna() & 
            (df["Estimated_Losses"] != 0) & 
            (df["Estimated_Losses"] != "-")
        ]

        # Create new DataFrame for Losses
        new_data1 = pd.DataFrame(columns=[
            'SalesOrgCode', 'DivCode', 'ProfitCenter', 'CustomerCode', 
            'SalesDistrictCode', 'SalesRepName', 'CurrencyCode', 
            'SalesYear', 'Estimated_Organic_Gains', 'Estimated_Losses', 
            'Comments'
        ])
        
        new_data1['SalesOrgCode'] = ['1030'] * len(loss_filtered)
        new_data1['DivCode'] = ['02'] * len(loss_filtered)
        new_data1['ProfitCenter'] = loss_filtered['Profit Center'].values
        new_data1['CustomerCode'] = loss_filtered['Customer #'].values
        new_data1['SalesDistrictCode'] = loss_filtered['Sales District #'].values
        new_data1['SalesRepName'] = loss_filtered['Sales Rep'].values
        new_data1['CurrencyCode'] = ['USD'] * len(loss_filtered)
        new_data1['SalesYear'] = ['2024'] * len(loss_filtered)
        new_data1['Estimated_Organic_Gains'] = loss_filtered['Estimated_Organic_Gains'].values
        new_data1['Estimated_Losses'] = loss_filtered['Estimated_Losses'].values
        new_data1['Comments'] = loss_filtered['Comments (For September Review)'].values

        # Concatenate the results
        final_df = pd.concat([new_data, new_data1], ignore_index=True)
        all_final_df = pd.concat([all_final_df, final_df], ignore_index=True)

        # Save each file's final_df to a CSV file
        output_filename = f"final_data_{input_file.name}.csv"
        final_df.to_csv(output_filename, index=False)
        download_links.append(output_filename)

    # Display download links for the user
    st.header("Download Final Data Files")
    for link in download_links:
        st.markdown(f"[Download {link}]({link})")

# End of Streamlit application
