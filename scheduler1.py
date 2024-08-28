# import pandas as pd
# from openpyxl import load_workbook

# def modify_excel():
#     # Step 1: Read the Excel file from the 'tco' sheet
#     input_file_path = './err_tco.xlsx'
#     df = pd.read_excel(input_file_path, sheet_name='Sheet2')
    
#     # Step 2: Sort by column A
#     df.sort_values(by=df.columns[0], inplace=True)
    
#     # Step 3: Remove columns J and K
#     df.drop(columns=[df.columns[9], df.columns[10]], inplace=True)
    
#     # Step 4: Insert new column after column E
#     df.insert(5, 'NewColumn', df.iloc[:, 4] / 100)
    
#     # Step 5: Add formula in F column =E/100
#     df['NewColumn'] = df.iloc[:, 4] / 100
    
#     # Step 6: Copy entire F column and paste special values in same column F
#     df['NewColumn'] = df['NewColumn'].values
    
#     # Step 7: Insert new row in 1st line
#     new_header = ['Name', 'AGID', 'AGID NAME', 'TABLE NAME', 'TABLE COUNT', 'NewColumn', 'APPLICATION NAME', 'APPLICATION DESCRIPTION', 'MNEMONIC', 'COST CODE', 'UNIQUE IDENTIFIER', 'COST CODE2']
#     if len(new_header) == len(df.columns):
#         df.loc[-1] = new_header  # adding a row
#         df.index = df.index + 1  # shifting index
#         df.sort_index(inplace=True)
#     else:
#         print("Error: The number of columns in new_header does not match the number of columns in the DataFrame.")
#         return
    
#     # Step 8: Add each column heading
#     df.columns = new_header
    
#     # Step 9: Add formula in column K =CONCATENATE(D2,G2,H2)
#     df['UNIQUE IDENTIFIER'] = df.apply(lambda row: f"{row['TABLE NAME']}{row['APPLICATION NAME']}{row['APPLICATION DESCRIPTION']}", axis=1)
    
#     # Step 10: Copy column K and paste special values in same column K
#     df['UNIQUE IDENTIFIER'] = df['UNIQUE IDENTIFIER'].values
    
#     # Step 11: Add formula in column L =LEFT(J2,8)
#     df['COST CODE2'] = df['COST CODE'].str[:8]
    
#     # Step 12: Copy column L and paste special values in same column L
#     df['COST CODE2'] = df['COST CODE2'].values
    
#     # Step 13: Filter column J for "BEA MMD"
#     df_filtered = df[df['COST CODE'] == 'BEA MMD']
    
#     # Step 14: Add sum of entire E column to row number 20
#     df.loc[19, 'TABLE COUNT'] = df['TABLE COUNT'].sum()
    
#     # Save the modified file to a different location
#     output_file_path = './err_tco_out.xlsx'
#     df.to_excel(output_file_path, index=False)
    
#     # Load the workbook to perform additional operations
#     wb = load_workbook(output_file_path)
#     ws = wb['Sheet2']
    
#     # Step 15-35: Perform additional operations using openpyxl
#     # (These steps are complex and require detailed implementation using openpyxl)
#     # Example for step 15:
#     cost_sheet = wb.create_sheet('cost')
#     for row in ws.iter_rows(min_row=20, max_row=20, min_col=3, max_col=25, values_only=True):
#         cost_sheet.append(row)
    
#     # Save the workbook
#     wb.save(output_file_path)

# # Call the function to execute the modifications
# modify_excel()