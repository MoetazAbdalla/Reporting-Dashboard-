import pandas as pd

# Load the Excel files
file1 = 'application-list-NEW.xlsx'
file2 = 'application-list.xlsx'

# Read the entire DataFrames
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# Find IDs present in df1 but not in df2
missing_ids = df1[~df1['ID'].isin(df2['ID'])]

# Save the resulting DataFrame to a new Excel file
missing_ids.to_excel('missing_ids.xlsx', index=False)

print("Missing IDs and their corresponding rows have been saved to 'missing_ids.xlsx'.")
