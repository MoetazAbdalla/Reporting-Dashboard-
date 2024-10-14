import pandas as pd

# Load the Excel file
file_path = 'find-regions.xlsx'  # Update with your file path
df = pd.read_excel(file_path)

# Assuming your columns are named 'Region' and 'Country'
# Group by 'Region', aggregate unique countries into a list
grouped = df.groupby('Region')['Nationality'].apply(lambda x: list(set(x))).reset_index()

# Save the result to a new CSV file
grouped.to_csv('regions_with_unique_countries.csv', index=False)

print("Regions and their unique corresponding countries have been saved to 'regions_with_unique_countries.csv'.")
