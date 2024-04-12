import pandas as pd

# Define the filenames for input Excel sheet and output Excel file
input_excel_file = r'C:\Users\ffb-9\Desktop\BIZ.xlsx'
list2_excel_file = r'C:\Users\ffb-9\Documents\List2.xlsx'  # Change the path accordingly
output_excel_file = r'C:\Users\ffb-9\Desktop\output_data.xlsx'  # Change the output file path accordingly

# List of sheet names to extract data from
sheet_names_to_extract = [ "Week 42","Week 41","Week 40","Week 39","Week 38", "Week 37", "Week 36", "Week 35", "Week 34", "Week 33", "Week 32", "Week 31", "Week 30", "Week 29", "Week 28", "Week 27"]

# Read the PRODUCT_ID information from List2.xlsx
list2_data = pd.read_excel(list2_excel_file)

# Convert 'PRODUCT_ID' column to string in List2 data
list2_data['PRODUCT_ID'] = list2_data['PRODUCT_ID'].astype(str)

# Initialize an empty DataFrame to store the combined data
combined_data = pd.DataFrame()

# Loop through the specified sheets in the input data
for sheet_name in sheet_names_to_extract:
    # Read data from the current sheet in BIZ.xlsx
    df = pd.read_excel(input_excel_file, sheet_name=sheet_name)
    
    # Convert 'PRODUCT_ID' column to string in BIZ data
    df['PRODUCT_ID'] = df['PRODUCT_ID'].astype(str)
    
    # Merge the data with the PRODUCT_ID information using a common column (e.g., 'PRODUCT_ID')
    merged_data = pd.merge(df, list2_data, on='PRODUCT_ID', how='inner')
    
    # Sum the values for each PRODUCT_ID
    grouped_data = merged_data.groupby(['PRODUCT_ID']).sum().reset_index()
    
    # Add a space character after each PRODUCT_ID
    grouped_data['PRODUCT_ID'] = grouped_data['PRODUCT_ID'] + ' '
    
    # Add a 'Sheet' column to indicate which week the data is from
    grouped_data['Sheet'] = sheet_name
    
    # Format 'Inventory Spin L7D' and 'Markdown' columns as percentages
    grouped_data['Inventory Spin L7D'] = grouped_data['Inventory Spin L7D'].apply(lambda x: f"{x:.2%}")
    grouped_data['Markdown'] = grouped_data['Markdown'].apply(lambda x: f"{x:.2%}")
    
    # Append the grouped data to the combined_data DataFrame
    combined_data = pd.concat([combined_data, grouped_data], ignore_index=True)  # Concatenate the data

# Sort the combined data by 'PRODUCT_ID' and 'Sheet' in descending order
combined_data.sort_values(by=['PRODUCT_ID', 'Sheet'], ascending=[True, False], inplace=True)

# Write the combined data to an output Excel file
combined_data.to_excel(output_excel_file, index=False)

print("Data successfully combined and saved to", output_excel_file)
