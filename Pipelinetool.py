import pandas as pd
import streamlit as st
import warnings

# Suppress warnings
warnings.filterwarnings("ignore")

# Function to process the pipeline data
def process_pipeline_data(input_file):
    # Read the Excel file and specify the sheet and header row
    df = pd.read_excel(input_file, sheet_name="1) Base Business", header=3)

    # Select specific columns from the DataFrame (A:J and AJ:AL)
    df = df.iloc[:, list(range(10)) + [52,53,54,55,56,57,58,59]]

    # Rename columns for easier access
    df=df.rename(columns={"2024 NTS $.9": "Estimated_Organic_Gains", "2024 NTS $.10": "Estimated_Losses"})
    # Filter for rows with valid Estimated Organic Gains
    organic_gains_filtered = df[
        df["Estimated_Organic_Gains"].notna() & 
        (df["Estimated_Organic_Gains"] != 0) & 
        (df["Estimated_Organic_Gains"] != "-")
    ]

    # Prepare a DataFrame for Organic Gains
    column_names = [
        'SalesOrgCode', 'DivCode', 'ProfitCenterCode', 'CustomerCode',
        'SalesDistrictCode', 'SalesRepName', 'CurrencyCode', 
        'SalesYear', 'Estimated_Organic_Gains', 'Estimated_Losses', 'Comments'
    ]
    new_data = pd.DataFrame(columns=column_names)
    new_data['SalesOrgCode'] = ['1030'] * len(organic_gains_filtered)
    new_data['DivCode'] = ['02'] * len(organic_gains_filtered)
    new_data['ProfitCenterCode'] = organic_gains_filtered['Profit Center #'].values
    new_data['CustomerCode'] = organic_gains_filtered['Customer #'].values
    new_data['SalesDistrictCode'] = organic_gains_filtered['Sales District #'].values
    new_data['SalesRepName'] = organic_gains_filtered['Sales Rep'].values
    new_data['CurrencyCode'] = ['USD'] * len(organic_gains_filtered)
    new_data['SalesYear'] = ['2024'] * len(organic_gains_filtered)
    new_data['Estimated_Organic_Gains'] = organic_gains_filtered['Estimated_Organic_Gains'].values
    new_data['Estimated_Losses'] = organic_gains_filtered['Estimated_Losses'].values
    new_data['Comments'] = organic_gains_filtered['Comments (For September Review).1'].values

    # Filter for rows with valid Estimated Losses
    loss_filtered = df[
        df["Estimated_Losses"].notna() & 
        (df["Estimated_Losses"] != 0) & 
        (df["Estimated_Losses"] != "-")
    ]

    # Prepare a DataFrame for Losses
    new_data1 = pd.DataFrame(columns=column_names)
    new_data1['SalesOrgCode'] = ['1030'] * len(loss_filtered)
    new_data1['DivCode'] = ['02'] * len(loss_filtered)
    new_data1['ProfitCenterCode'] = loss_filtered['Profit Center #'].values
    new_data1['CustomerCode'] = loss_filtered['Customer #'].values
    new_data1['SalesDistrictCode'] = loss_filtered['Sales District #'].values
    new_data1['SalesRepName'] = loss_filtered['Sales Rep'].values
    new_data1['CurrencyCode'] = ['USD'] * len(loss_filtered)
    new_data1['SalesYear'] = ['2024'] * len(loss_filtered)
    new_data1['Estimated_Organic_Gains'] = loss_filtered['Estimated_Organic_Gains'].values
    new_data1['Estimated_Losses'] = loss_filtered['Estimated_Losses'].values
    new_data1['Comments'] = loss_filtered['Comments (For September Review)'].values

    # Concatenate the two DataFrames
    final_df = pd.concat([new_data, new_data1], ignore_index=True)
    final_df= final_df.drop_duplicates()

    return final_df

# To run this application, use the command:
# streamlit run your_script.py


# Streamlit application
st.set_page_config(page_title="Pipeline Review Data Processing", layout="wide")
st.title("📊 Pipeline Review Data Processing")

# Upload section
st.header("Upload Excel Files")
input_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

# Process button
if st.button("Process Data"):
    if input_files:
        all_final_df = pd.DataFrame()  # Initialize an empty DataFrame
        download_links = []  # Initialize list to store download links

        for input_file in input_files:
            final_df = process_pipeline_data(input_file)
            all_final_df = pd.concat([all_final_df, final_df], ignore_index=True)  # Concatenate results
            
            # Save each file's final_df to a CSV file
            output_file = f"{input_file.name}_output.csv"
            final_df.to_csv(output_file, index=False)

            # Add the download link for the current file's final_df
            download_links.append((input_file.name, output_file))

        # Save the final DataFrame to a CSV file
        overall_output_file = "final_output.csv"
        all_final_df.to_csv(overall_output_file, index=False)

        st.success("Data processed successfully!")

        # Provide a download link for each processed file
        st.header("Download Processed Files")
        for original_file_name, output_file in download_links:
            with open(output_file, "rb") as file:
                st.download_button(label=f"Download {original_file_name} Output CSV", data=file, file_name=output_file)

        # Provide a download link for the overall output
        with open(overall_output_file, "rb") as file:
            st.download_button(label="Download Overall Output CSV", data=file, file_name=overall_output_file)
    else:
        st.error("Please upload at least one Excel file.")

