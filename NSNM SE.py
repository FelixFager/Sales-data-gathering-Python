import pandas as pd


input_excel_file = r'C:\Users\ffb-9\Desktop\BIZ.xlsx'
list2_excel_file = r'C:\Users\ffb-9\Documents\List2.xlsx' 
output_excel_file = r'C:\Users\ffb-9\Desktop\output_data.xlsx'  


sheet_names_to_extract = [ "Week 42","Week 41","Week 40","Week 39","Week 38", "Week 37", "Week 36", "Week 35", "Week 34", "Week 33", "Week 32", "Week 31", "Week 30", "Week 29", "Week 28", "Week 27"]


list2_data = pd.read_excel(list2_excel_file)


list2_data['PRODUCT_ID'] = list2_data['PRODUCT_ID'].astype(str)


combined_data = pd.DataFrame()


for sheet_name in sheet_names_to_extract:

    df = pd.read_excel(input_excel_file, sheet_name=sheet_name)
    
    
    df['PRODUCT_ID'] = df['PRODUCT_ID'].astype(str)
    
  
    merged_data = pd.merge(df, list2_data, on='PRODUCT_ID', how='inner')
    

    grouped_data = merged_data.groupby(['PRODUCT_ID']).sum().reset_index()
    
    grouped_data['PRODUCT_ID'] = grouped_data['PRODUCT_ID'] + ' '
    

    grouped_data['Sheet'] = sheet_name
    
 
    grouped_data['Inventory Spin L7D'] = grouped_data['Inventory Spin L7D'].apply(lambda x: f"{x:.2%}")
    grouped_data['Markdown'] = grouped_data['Markdown'].apply(lambda x: f"{x:.2%}")
    

    combined_data = pd.concat([combined_data, grouped_data], ignore_index=True)  


combined_data.sort_values(by=['PRODUCT_ID', 'Sheet'], ascending=[True, False], inplace=True)


combined_data.to_excel(output_excel_file, index=False)

print("Data successfully combined and saved to", output_excel_file)
