import pandas as pd

# Load the Excel file
file_path = 't2.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Drop rows with any NaN (blank) values
df_cleaned = df.dropna()

# Save the updated Excel file without rows that had blank elements
df_cleaned.to_excel('cleaned_excel_file.xlsx', index=False)

print(f"Original rows: {df.shape[0]}, Rows after cleaning: {df_cleaned.shape[0]}")
