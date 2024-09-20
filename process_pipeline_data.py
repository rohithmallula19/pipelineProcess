def process_pipeline_data(input_file):
    # Read the Excel file and specify the sheet and header row
    df = pd.read_excel(input_file, sheet_name="1) Base Business", header=3)

    # Select specific columns from the DataFrame (A:J and AJ:AL)
    df = df.iloc[:, list(range(10)) + [34, 35, 36, 37, 38, 39, 40, 41]]

    # Rename columns for easier access
    df = df.rename(columns={
        "2024 NTS $.6": "Estimated_Organic_Gains",
        "2024 NTS $.7": "Estimated_Losses"
    })

    # Filter for rows with valid Estimated Organic Gains
    organic_gains_filtered = df[
        df["Estimated_Organic_Gains"].notna() & 
        (df["Estimated_Organic_Gains"] != 0) & 
        (df["Estimated_Organic_Gains"] != "-")
    ]

    # Prepare a DataFrame for Organic Gains
    column_names = [
        'SalesOrgCode', 'DivCode', 'ProfitCenter', 'CustomerCode',
        'SalesDistrictCode', 'SalesRepName', 'CurrencyCode', 
        'SalesYear', 'Estimated_Organic_Gains', 'Estimated_Losses', 'Comments'
    ]
    new_data = pd.DataFrame(columns=column_names)
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

    # Prepare a DataFrame for Losses
    new_data1 = pd.DataFrame(columns=column_names)
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

    # Concatenate the two DataFrames
    final_df = pd.concat([new_data, new_data1], ignore_index=True)

    return final_df
