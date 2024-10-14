import pandas as pd
import re

# Load the Excel file
input_file = "application-list-NEW.xlsx"  # Replace with your file path

# Try loading the file with the correct header row (adjust the header parameter if needed)
df = pd.read_excel(input_file, header=0)  # Change header=0 if the headers are not in the first row

# Print the first few rows to inspect the structure
print(df.head())

# Strip any leading/trailing whitespace in column names
df.columns = df.columns.str.strip()

# Print column names for debugging
print("Column names after stripping:", df.columns)

# Assuming 'Created By' is the column you're interested in
agency_column = 'Created By'

# Check if the 'Created By' column exists
if agency_column not in df.columns:
    print(f"Column '{agency_column}' not found!")
else:
    # Count the occurrences of each agency
    agency_counts = df[agency_column].value_counts()

    # Create an ExcelWriter object to write multiple sheets to the same file
    output_file = "agencies_data.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for agency_name, count in agency_counts.items():
            # Sanitize the agency name to make it a valid sheet name (max 31 characters, no special chars)
            sanitized_agency_name = re.sub(r'[\\/*?:"<>|]', "", agency_name)[:31]

            # Filter rows for the current agency
            agency_df = df[df[agency_column] == agency_name]

            # Write the data to a new sheet in the same file
            agency_df.to_excel(writer, sheet_name=sanitized_agency_name, index=False)
            print(f"Saved data for {agency_name} to the '{sanitized_agency_name}' sheet.")

    print(f"All agency data saved to {output_file}")
