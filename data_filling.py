import pandas as pd

# Load the Excel file
file_path = 't1.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Iterate over the rows of the dataframe
for index, row in df.iterrows():
    col_L_value = str(row.iloc[11])  # 12th column (index 11)
    col_O_value = str(row.iloc[13])  # 14th column (index 13)
    
    # Check if '150x224' is in column 'L' and column 'O' is blank
    if ' 270x270' in col_L_value and pd.isna(row.iloc[13]):
        df.at[index, df.columns[13]] = 'size=Double'
    
    # Check if 'Single' is in column 'O' and column 'L' is blank
    elif 'Double' in col_O_value and pd.isna(row.iloc[11]):
        df.at[index, df.columns[11]] = 'Dimension:  270x270 CM'

# Save the updated Excel file
df.to_excel('t2.xlsx', index=False)
