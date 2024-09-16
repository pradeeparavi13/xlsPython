import os
import pandas as pd
from openpyxl import load_workbook
import numpy as np

def modify_excel():

    # Define the file path
    output_file_path = './err_tco_out.xlsx'

    # Check if the file exists
    if os.path.exists(output_file_path):
        # Delete the file
        os.remove(output_file_path)
        print(f"File {output_file_path} has been deleted.")
    else:
        print(f"File {output_file_path} does not exist.")

    # Define the file path
    output_file_path1 = './err_tco_out.xlsx'

    # Check if the file exists
    if os.path.exists(output_file_path1):
        # Delete the file
        os.remove(output_file_path1)
        print(f"File {output_file_path} has been deleted.")
    else:
        print(f"File {output_file_path} does not exist.")

    # Step 1: Read the Excel file from the 'tco' sheet
    input_file_path = './err_tco.xlsx'
    df = pd.read_excel(input_file_path, sheet_name='Sheet2')
    
    # Step 2: Sort by column A (assuming the first column is 'A')
    df.sort_values(by=df.columns[0], inplace=True)
    
    # Step 3: Remove columns J and K (which are at index 9 and 10 respectively)
    df.drop(columns=[df.columns[9], df.columns[10]], inplace=True)

    # Step 4: Insert a new row in the first line
    new_header = ['Name', 'AGID', 'AGID NAME', 'TABLE NAME', 'TABLE COUNT', 'APPLICATION NAME', 'APPLICATION DESCRIPTION', 'MNEMONIC', 'COST CODE2']
    if len(new_header) == len(df.columns):
        df.loc[-1] = new_header  # adding a row
        df.index = df.index + 1  # shifting index
        df.sort_index(inplace=True)
    else: 
        print("Error: The number of columns in new_header does not match the number of columns in the DataFrame.")
        return
    
    # Step 5: Update the dataframe with the new headers
    df.columns = new_header

    # Step 6: Delete the second row (index 0)
    df = df.drop(0)  # Modify in-place or assign the result

    
    # Step 7: Insert a new column after column E 
    df.insert(5, "NEW COUNT", df.iloc[:, 4])  
    
    # Step 8: Copy entire F column and paste special values in the same column
    df['NEW COUNT'] = df['NEW COUNT'].values

    # Step 9: 'UNIQUE IDENTIFIER', 'COST CODE'
    df.insert(10, 'UNIQUE IDENTIFIER', np.nan)
    df.insert(11, 'COST CODE', np.nan)

    # num_columns = df.shape[1]

    # print(f"Number of columns in the DataFrame: {num_columns}")

    
    
    # # Step 10: Add formula in column K
    df['UNIQUE IDENTIFIER'] = df.apply(lambda row: f"{row['TABLE NAME']}{row['APPLICATION NAME']}{row['APPLICATION DESCRIPTION']}", axis=1)
    
    # # Step 11: Copy column K and paste special values in the same column
    df['UNIQUE IDENTIFIER'] = df['UNIQUE IDENTIFIER'].values
    
    # # Step 12: Add formula in column L (COST CODE)
    df['COST CODE'] = df['COST CODE2'].str[:8]
    
    # # Step 13: Copy column L and paste special values in the same column
    df['COST CODE'] = df['COST CODE'].values
    
    # # Step 14: Filter column J (COST CODE) for "BEA MMD"
    df_filtered = df[df['COST CODE2'] == 'BEA MMED (Money Management Division)']

    
    # # Step 15: Add sum of the entire E column (TABLE COUNT) to its bottom
    total_sum = df_filtered['NEW COUNT'].sum()
    # df.loc[df.shape[0]] = ['', '', '', '', total_sum] + [''] * (df.shape[1] - 5)

    print("total_sum ==",total_sum)


    df = df[~(df['COST CODE2'] == 'BEA MMED (Money Management Division)')]
    count = len(df)
    print("rows ===>>",len(df))


    # Step 16: Read the second Excel file from the 'aca' sheet
    input_file_path = './mmdcost.xlsx'
    df1 = pd.read_excel(input_file_path, sheet_name='aca')

    # Print column names to check for typos
    print(df1.columns)


    df1['acaPercentageallocation'] = (df1['acaPercentageallocation'] / 100) * total_sum / 2

    df1.dropna(subset=['acaPercentageallocation'], inplace=True)

    print(df1['acaPercentageallocation'])
    print(df1['acaMu'])



    # Create a new DataFrame with the values from df1['acaPercentageallocation']
    # new_df = pd.DataFrame(df1['acaPercentageallocation'].values.reshape(-1, 1), columns=['NEW COUNT'])
    new_df = pd.DataFrame({'NEW COUNT': df1['acaPercentageallocation'], 'COST CODE': df1['acaMu'], 'UNIQUE IDENTIFIER': df1['acaMu']})


    # Insert the new DataFrame into df after the nth row
    df = pd.concat([df[:count], new_df, df[count:]], ignore_index=True)

    df.loc[count:, 'UNIQUE IDENTIFIER'] = "BEA MMD " + df.loc[count:, 'UNIQUE IDENTIFIER'].astype(str)
    
    df['COST CODE2'] = df['COST CODE2'].fillna('blank')
    df['MNEMONIC'] = df['MNEMONIC'].fillna('MMD')

    pivot_table = df.pivot_table(
        index=['UNIQUE IDENTIFIER', 'COST CODE', 'COST CODE2', 'MNEMONIC'],
    values='NEW COUNT',
    aggfunc='sum',
    fill_value=0
    )

    print(pivot_table)

    pivot_table = pivot_table.reset_index()


    pivot_table.to_excel('./pivot_table.xlsx', engine='openpyxl')

    print("Updated values have been saved to", output_file_path)

    # Print the DataFrame to verify the changes
    # print(colVlaue)
    
    # Save the modified file to a different location
    output_file_path = './err_tco_out.xlsx'
    df.to_excel(output_file_path, index=False)


    # Step last: Load the workbook to perform additional operations (if necessary)
    wb = load_workbook(output_file_path)
    ws = wb.active
    

    
    # If needed, you can perform more operations using openpyxl here
    
    # Save the final workbook
    wb.save(output_file_path)

# Call the function to execute the modifications
modify_excel()
