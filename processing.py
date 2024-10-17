# def transformation(keka_path, biometric_path):
#     import pandas as pd
#     import re

#     # Load the Excel file into a Pandas DataFrame
#     keka_df = pd.read_excel(keka_path, header=2, skipfooter=26, engine='openpyxl')
#     keka_df.drop(columns=['06-Feb-2024'], inplace=True)
    
#     # This regex matches columns with a pattern like 'dd-mmm-yyyy', e.g., '04-feb-2024'
#     date_pattern = re.compile(r'\d{2}-[a-zA-Z]{3}-\d{4}')

#     # read biometric file
#     file = pd.read_excel(biometric_path, header=9)   # header will start from the 9th row and rest it will skip
#     file = file.dropna(how='all')
#     file = file.dropna(axis=1, how='all')

#     # Creating a new dataframe by selecting the values from Unnamed: 2 or Unnamed: 3 (whichever has a string instead of NaN)
#     biometrics_df = pd.DataFrame({
#         'Employee Name': file['Unnamed: 2'].combine_first(file['Unnamed: 3']),
#     })
#     biometrics_df = biometrics_df.dropna(how='all')
#     biometrics_df = biometrics_df.dropna(axis=1, how='all')
#     # print(biometrics_df.to_string())
#     def mark_presence(value):
#         if isinstance(value, str) and re.match(r'\d{2}:\d{2}', value):  # Check for time format
#             return 'P'
#         elif value == 'A':
#             return 'A'
#         else:
#             return value  # Keep other values as they are (like "WO-II" or NaN)

#     # Apply the function to all columns dynamically (starting from column 7 to last)
#     # df = file.copy()
#     for col in file.columns[5:]:
#         biometrics_df[col] = file[col].apply(mark_presence)

#     biometrics_df.to_excel('biometric_optimized.xlsx')

#     #Extract day from 'dd-mmm-yyyy' columns in keka
#     date_columns = [col for col in keka_df.columns if date_pattern.match(col)]

#     # # Create a mapping of date columns to day numbers
#     day_map = {col: col.split('-')[0].lstrip('0') for col in date_columns}

#     # # Normalize employee names in both DataFrames to remove leading/trailing spaces and make them lowercase
#     keka_df['Employee Name'] = keka_df['Employee Name'].str.strip().str.upper()
#     biometrics_df['Employee Name'] = biometrics_df['Employee Name'].str.strip().str.upper()

#     # # Set 'Employee Name' as index for fast lookup in both DataFrames
#     keka_df.set_index('Employee Name', inplace=True)
#     biometrics_df.set_index('Employee Name', inplace=True)

#     # # Optimize by iterating only over rows present in both DataFrames
#     common_employees = keka_df.index.intersection(biometrics_df.index)

#     # # Update attendance based on biometric data
#     for employee in common_employees:
#         for keka_date, day in day_map.items():
#             if keka_df.at[employee, keka_date] == 'A' and biometrics_df.at[employee, day] == 'P':
#                 # Update keka_df from 'A' to 'P'
#                 keka_df.at[employee, keka_date] = 'P'
#                 keka_df.at[employee,'Absent Days'] = keka_df.at[employee,'Absent Days']-1
#                 keka_df.at[employee,'Present Days'] = keka_df.at[employee,'Present Days']+1

#     # # Reset index if needed and print or save the updated DataFrame
#     keka_df.reset_index(inplace=True)
#     consolidated_path = 'consolidated_attendance_optimized.xlsx'
#     # Optionally save to an Excel or CSV file
#     keka_df.to_excel(consolidated_path, index=False)
#     return consolidated_path
import pandas as pd
import re

def transformation(keka_path, indore_biometric_path, raipur_biometric_path):
    # Load the Excel file into a Pandas DataFrame
    keka_df = pd.read_excel(keka_path, header=2, skipfooter=26, engine='openpyxl')
    # get merged biometric dataframe with indore and raipur record
    biometrics_df = preprocess_and_merge_biometric_data(indore_biometric_path, raipur_biometric_path)
    
    # # Normalize employee names in both DataFrames to remove leading/trailing spaces and make them lowercase
    keka_df['Employee Name'] = keka_df['Employee Name'].str.strip().str.upper()
    biometrics_df['Employee Name'] = biometrics_df['Employee Name'].str.strip().str.upper()
    
    biometrics_df.to_excel('biometric_optimized.xlsx', index=False)
    
    # Check for duplicate employee names
    if keka_df['Employee Name'].duplicated().any():
        duplicated = keka_df[keka_df['Employee Name'].duplicated(keep=False)]
        print("Duplicate Employee Names found in keka_df:")
        print(duplicated[['Employee Name']])
        return
    if biometrics_df['Employee Name'].duplicated().any():
        duplicated = biometrics_df[biometrics_df['Employee Name'].duplicated(keep=False)]
        print("Duplicate Employee Names found in biometrics_df:")
        print(duplicated[['Employee Name']])
        return
   
    # # Set 'Employee Name' as index for fast lookup in both DataFrames
    keka_df.set_index('Employee Name', inplace=True)
    biometrics_df.set_index('Employee Name', inplace=True)
    print(len(biometrics_df))
    # Ensure there are common employees to process
    common_employees = keka_df.index.intersection(biometrics_df.index)
    print(f"Number of common employees: {len(common_employees)}")
    
    if len(common_employees) == 0:
        print("No common employees found between Keka and Biometric data.")
        return
    
    # identify date columns in keka DataFrame
    # This regex matches columns with a pattern like 'dd-mmm-yyyy', e.g., '04-feb-2024'
    date_pattern = re.compile(r'\d{2}-[a-zA-Z]{3}-\d{4}')
    #Extract day from 'dd-mmm-yyyy' columns in keka
    date_columns = [col for col in keka_df.columns if date_pattern.match(col)]

    if not date_columns:
        print("No date columns found in Keka DataFrame.")
        return

    # Create a mapping from Keka date columns to Biometric day numbers
    # Assuming Biometric day columns are numbered as strings without leading zeros
    day_map = {col: str(int(col.split('-')[0])) for col in date_columns}
    
    # Verify that all day numbers exist in Biometric DataFrame columns
    missing_days = set(day_map.values()) - set(biometrics_df.columns.astype(str))
    if missing_days:
        print(f"The following day columns are missing in Biometric DataFrame: {missing_days}")
        # Optionally, remove these dates from processing
        for col, day in list(day_map.items()):
            if day in missing_days:
                print(f"Removing date column '{col}' due to missing day '{day}' in Biometric data.")
                del day_map[col]

    if not day_map:
        print("No valid date mappings found between Keka and Biometric data.")
        return    
    
    # # Update attendance based on biometric data
    for employee in common_employees:
        for keka_date, day in day_map.items():
            keka_val = keka_df.at[employee, keka_date] if pd.notna(keka_df.at[employee, keka_date]) else None
            biometric_val = biometrics_df.at[employee, day] if pd.notna(biometrics_df.at[employee, day]) else None
            if keka_val == 'A' and biometric_val == 'P':
            # if keka_df.at[employee, keka_date] == 'A' and biometrics_df.at[employee, day] == 'P':
                # Update keka_df from 'A' to 'P'
                # print(employee +', '+keka_date +', '+day)
                keka_df.at[employee, keka_date] = 'P'
                keka_df.at[employee,'Absent Days'] -= 1
                keka_df.at[employee,'Present Days'] += 1

    # # Reset index if needed and print or save the updated DataFrame
    keka_df.reset_index(inplace=True)

    try:
        output_path = 'consolidated_attendance_optimized.xlsx'
        keka_df.to_excel(output_path, index=False)
        print(f"Transformation complete. Saved to '{output_path}'.")
        return output_path
    except PermissionError:
        print(f"Permission denied: Could not save to '{output_path}'. Ensure the file is closed before running the script.")
    

def obtimized_biometric(biometrics):
    biometrics.drop(columns=['Unnamed: 0','Unnamed: 1'], inplace=True)
    biometrics.rename(columns={'Unnamed: 2': 'Employee Name'}, inplace=True)

    biometrics.dropna(how='all', inplace=True)
    biometrics.dropna(axis=1, how='all', inplace=True)
    biometrics.dropna(subset=['Employee Name'],inplace=True)
    # Replace 'MIS' with 'P' (Present)
    biometrics.replace('MIS', 'P', inplace=True)

    return biometrics



def preprocess_and_merge_biometric_data(indore_path, raipur_path, indore_header=8, raipur_header=5):
    """
    This function preprocesses two Excel files for biometric attendance data (Indore and Raipur),
    standardizes them, and merges them into a single DataFrame.

    Parameters:
    indore_path (str): Path to the Indore biometric Excel file.
    raipur_path (str): Path to the Raipur biometric Excel file.
    indore_header (int): Header row index for the Indore file (default is 8).
    raipur_header (int): Header row index for the Raipur file (default is 5).

    Returns:
    pd.DataFrame: The merged DataFrame containing preprocessed data from both files.
    """

    # Load the Indore file
    indore = pd.read_excel(indore_path, header=indore_header)
    indore.drop(columns=['Unnamed: 0', 'Unnamed: 1'], inplace=True)
    indore.rename(columns={'Unnamed: 2': 'Employee Name'}, inplace=True)
    indore = indore.dropna(axis=1, how='all')  # Drop columns that are completely empty
    indore = indore[indore['Employee Name'].notna()]  # Remove rows where 'Employee Name' is NaN

    # Clean columns: Drop columns that do not match the date pattern
    for column in indore.columns:
        if column != 'Employee Name' and not re.match(r'\d', column):
            indore.drop(columns=column, inplace=True)

    # Replace 'MIS' with 'P' and fill NaNs with 'A'
    indore.replace('MIS', 'P', inplace=True)
    indore.fillna('A', inplace=True)

    # Load the Raipur file
    raipur = pd.read_excel(raipur_path, header=raipur_header)
    raipur = raipur.dropna(how='all')  # Drop rows that are completely empty
    raipur = raipur.dropna(axis=1, how='all')  # Drop columns that are completely empty

    # Remove unwanted columns
    raipur.drop(columns=['Employee ID', 'Department'], inplace=True)

    # Clean columns: Drop columns that do not match the date pattern
    for column in raipur.columns:
        if column != 'First Name' and not re.match(r'\d', column):
            raipur.drop(columns=column, inplace=True)

    # Standardize the column names to match the Indore DataFrame
    raipur.columns = indore.columns

    # Replace 'HD' with 'P' and fill NaNs with 'A'
    raipur.replace('HD', 'P', inplace=True)
    raipur.fillna('A', inplace=True)

    # Merge the two DataFrames
    union_df = pd.concat([indore, raipur], ignore_index=True)

    return union_df


# # Example usage:
# merged_df = preprocess_and_merge_biometric_data("sept_master_biometric.xlsx", "raipur_biometric.xlsx")
# print(merged_df)


# function calling
transformation(r'C:\Users\DELL\Downloads\test_keka.xlsx',r'C:\Users\DELL\Downloads\sept_master_biometric.xlsx', r'C:\Users\DELL\Downloads\raipur_biometric.xlsx')