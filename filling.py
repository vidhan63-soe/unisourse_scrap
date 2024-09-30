import pandas as pd

# Load the Excel file
file_path = 'f1.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    # Check if there are any missing (NaN) values in the row
    if row.isna().any():
        # Identify the columns that are NOT missing in the current row
        non_missing_columns = row.dropna().index.tolist()

        # Find rows where the values in the non-missing columns are the same
        similar_rows = df.dropna().loc[(df[non_missing_columns] == row[non_missing_columns]).all(axis=1)]

        # If we find similar rows, use the first one to fill in the missing values
        if not similar_rows.empty:
            for col in df.columns:
                if pd.isna(row[col]):  # If the value in the current column is NaN, fill it
                    df.at[index, col] = similar_rows.iloc[0][col]

# Save the updated Excel file with filled values
df.to_excel('f2.xlsx', index=False)

print("Missing values filled based on similar entries.")
