def transformation(keka_path, biometric_path):
    import pandas as pd
    import re

    # Load the Excel file into a Pandas DataFrame
    keka_df = pd.read_excel(keka_path, header=2, skipfooter=26, engine='openpyxl')
    keka_df.drop(columns=['06-Feb-2024'], inplace=True)
    
    # This regex matches columns with a pattern like 'dd-mmm-yyyy', e.g., '04-feb-2024'
    date_pattern = re.compile(r'\d{2}-[a-zA-Z]{3}-\d{4}')

    # read biometric file
    file = pd.read_excel(biometric_path, header=9)   # header will start from the 9th row and rest it will skip
    file = file.dropna(how='all')
    file = file.dropna(axis=1, how='all')

    # Creating a new dataframe by selecting the values from Unnamed: 2 or Unnamed: 3 (whichever has a string instead of NaN)
    biometrics_df = pd.DataFrame({
        'Employee Name': file['Unnamed: 2'].combine_first(file['Unnamed: 3']),
    })
    biometrics_df = biometrics_df.dropna(how='all')
    biometrics_df = biometrics_df.dropna(axis=1, how='all')
    # print(biometrics_df.to_string())
    def mark_presence(value):
        if isinstance(value, str) and re.match(r'\d{2}:\d{2}', value):  # Check for time format
            return 'P'
        elif value == 'A':
            return 'A'
        else:
            return value  # Keep other values as they are (like "WO-II" or NaN)

    # Apply the function to all columns dynamically (starting from column 7 to last)
    # df = file.copy()
    for col in file.columns[5:]:
        biometrics_df[col] = file[col].apply(mark_presence)

    biometrics_df.to_excel('biometric_optimized.xlsx')

    #Extract day from 'dd-mmm-yyyy' columns in keka
    date_columns = [col for col in keka_df.columns if date_pattern.match(col)]

    # # Create a mapping of date columns to day numbers
    day_map = {col: col.split('-')[0].lstrip('0') for col in date_columns}

    # # Normalize employee names in both DataFrames to remove leading/trailing spaces and make them lowercase
    keka_df['Employee Name'] = keka_df['Employee Name'].str.strip().str.upper()
    biometrics_df['Employee Name'] = biometrics_df['Employee Name'].str.strip().str.upper()

    # # Set 'Employee Name' as index for fast lookup in both DataFrames
    keka_df.set_index('Employee Name', inplace=True)
    biometrics_df.set_index('Employee Name', inplace=True)

    # # Optimize by iterating only over rows present in both DataFrames
    common_employees = keka_df.index.intersection(biometrics_df.index)

    # # Update attendance based on biometric data
    for employee in common_employees:
        for keka_date, day in day_map.items():
            if keka_df.at[employee, keka_date] == 'A' and biometrics_df.at[employee, day] == 'P':
                # Update keka_df from 'A' to 'P'
                keka_df.at[employee, keka_date] = 'P'
                keka_df.at[employee,'Absent Days'] = keka_df.at[employee,'Absent Days']-1
                keka_df.at[employee,'Present Days'] = keka_df.at[employee,'Present Days']+1

    # # Reset index if needed and print or save the updated DataFrame
    keka_df.reset_index(inplace=True)
    consolidated_path = 'consolidated_attendance_optimized.xlsx'
    # Optionally save to an Excel or CSV file
    keka_df.to_excel(consolidated_path, index=False)
    return consolidated_path
