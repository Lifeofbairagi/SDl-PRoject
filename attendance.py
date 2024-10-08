import pandas as pd
import glob
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


path = "./*.xlsx" 
all_files = glob.glob(path)

#empty list bana di hai storage ke liye
df_list = []


for file in all_files:
    try:
        # Load each Excel file without header to inspect the first few rows and then print the first 5 rows to check fromating
        df = pd.read_excel(file, header=None)

        print(f"Processing file: {file}")
        print(f"Shape: {df.shape}")
        print(df.head(5))  

        if df.shape[0] < 2 or df.shape[1] < 4:  
            print(f"Skipping empty or improperly formatted file: {file}")
            continue

      
        subject_name = df.iloc[0, 0] 

        # Set the second row which contains subject name as header
        df.columns = df.iloc[1]  # Set the second row as header here 
        df = df.drop(index=[0, 1]) 

        # Reset index so that the deletion of the first two rows does not affect the indexing 
        df.reset_index(drop=True, inplace=True)

        
        df.columns = df.columns.str.strip()  # Remove any spaces in front or back in colum name

        if 'Enrollment No.' not in df.columns or 'Name' not in df.columns:
            print(f"Skipping file due to missing 'Enrollment No.' or 'Name' columns: {file}")
            continue

        # converting all values to numeric to make sure nothing illigitimate is in the attencace count
        df['Total Theory'] = pd.to_numeric(df['Total Theory'], errors='coerce')
        df['Attended'] = pd.to_numeric(df['Attended'], errors='coerce')

  
        if 'Lab' in df.columns and 'Lab Attended' in df.columns:
            # Select columns for lab and theory seperately
            df['Lab'] = pd.to_numeric(df['Lab'], errors='coerce')
            df['Lab Attended'] = pd.to_numeric(df['Lab Attended'], errors='coerce')
            df_theory = df[['Enrollment No.', 'Name', 'Total Theory', 'Attended']].copy()
            df_lab = df[['Enrollment No.', 'Name', 'Lab', 'Lab Attended']].copy()
        else:
         
            df_theory = df[['Enrollment No.', 'Name', 'Total Theory', 'Attended']].copy()
            df_lab = None  

        
        df_theory['Theory Percentage'] = (df_theory['Attended'] * 100 / df_theory['Total Theory']).fillna(0).round(2)

        # Set index for theory DataFrame
        df_theory.set_index(['Enrollment No.', 'Name'], inplace=True)
        df_theory.columns = pd.MultiIndex.from_product([[subject_name], df_theory.columns])

       
        df_list.append(df_theory)

        # If lab data exists take it similary as theory 
        if df_lab is not None:
            df_lab['Lab Percentage'] = (df_lab['Lab Attended'] * 100 / df_lab['Lab']).fillna(0).round(2)
            df_lab.set_index(['Enrollment No.', 'Name'], inplace=True)
            df_lab.columns = pd.MultiIndex.from_product([[subject_name], df_lab.columns])
            df_list.append(df_lab)

        print(f"DataFrame for {subject_name} (Theory):")
        print(df_theory.head())  

        if df_lab is not None:
            print(f"DataFrame for {subject_name} (Lab):")
            print(df_lab.head())  

    except Exception as e:
        print(f"Error processing file {file}: {e}")

# Concatinate all DataFrames in the list into a single DataFrame
if df_list:
    merged_df = pd.concat(df_list, axis=1)  

   
    merged_df['Total Classes'] = merged_df.xs('Total Theory', axis=1, level=1).sum(axis=1, skipna=True) + \
                                  merged_df.xs('Lab', axis=1, level=1).sum(axis=1, skipna=True)
    merged_df['Total Attended'] = merged_df.xs('Attended', axis=1, level=1).sum(axis=1, skipna=True) + \
                                   merged_df.xs('Lab Attended', axis=1, level=1).sum(axis=1, skipna=True)

    
    total_percentage = (merged_df['Total Attended'] * 100 / merged_df['Total Classes']).replace([float('inf'), -float('inf')], 0).round(0).fillna(0)

    # Adding the Total Percentage column to the DataFrame
    merged_df['Total Percentage'] = total_percentage

    # Print the merged DataFrame for verification looking for any formating mistakes
    print("Merged DataFrame with Total Columns and Percentages:")
    print(merged_df.head())  

    # creating final output dataframe 
    output_file = "FinalAttendace.xlsx"
    merged_df.to_excel(output_file, sheet_name='Summary', index=True)  # Include index