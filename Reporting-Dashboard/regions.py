import pandas as pd

# Load the Excel file into a DataFrame
file_path = 'Nationalities.xlsx'  # Replace with the actual file path
df = pd.read_excel(file_path)

# Define the regions and their corresponding countries for mapping
regions = {
    'Central & Eastern Europe': ["Czech Republic", "Slovakia", "Cyprus", "Kosovo", "Bosnia and Herzegovina", "Hungary",
                                 "Poland", "Estonia", "Romania", "Montenegro", "Albania", "Greece", "Serbia",
                                 "Bulgaria"],
    'Central Asia (CIS)': ["Belarus", "Tajikistan", "Azerbaijan", "Moldova", "Armenia", "Kyrgyzstan", "Georgia",
                           "Ukraine", "Kazakhstan", "Uzbekistan"],

    'East, Southeast Asia & Pacific': ["Thailand", "Cambodia", "Taiwan", "Malaysia", "Myanmar", "Hong Kong", "Japan",
                                       "China", "Mongolia", "South Korea", "Philippines", "American Samoa",
                                       "Australia"],
    'Indonesia': ["Indonesia"],

    'Iran': ["Iran"],

    'Latin America & The Caribbean': ["Saint Kitts and Nevis", "Brazil", "Saint Lucia", "Mexico", "Belize", "Colombia",
                                      "Jamaica", "East Timor", "Venezuela", "Antigua and Barbuda", "Dominica"],

    'MENA': ["Libya", "Israel", "Qatar", "Tunisia", "Syria", "Algeria", "Yemen", "Jordan", "Lebanon", "Saudi Arabia",
             "Egypt", "Iraq", "Palestine", "Morocco"],

    'Middle East': ["Qatar", "Israel", "Syria", "Yemen", "Kuwait", "Jordan", "Lebanon", "Saudi Arabia", "Iraq", "Oman",
                    "United Arab Emirates", "Palestine"],

    'North Africa': ["Libya", "Tunisia", "Liechtenstein", "Algeria", "Bahrain", "Egypt", "Morocco"],
    'North America': ["Canada", "United States"],

    'Northern & Western Europe': ["Netherlands", "Luxembourg", "Austria", "Spain", "Norway", "Sweden", "Belgium",
                                  "Germany", "Italy", "Finland", "Ireland", "Switzerland", "Denmark", "Portugal",
                                  "Former Yugoslav Republic of Macedonia", "United Kingdom", "France"],

    'Russia': ["Russia"],

    'South Asia': ["Maldives", "Sri Lanka", "Nepal", "Pakistan", "Afghanistan", "Bangladesh", "India"],

    'Sub-Saharan Africa': ["Eritrea", "Central African Republic", "Bhutan", "Congo", "Mauritius", "South Africa",
                           "Guinea-Bissau", "Niger", "Burundi", "Rwanda", "Sudan", "Djibouti", "Uganda", "Nigeria",
                           "Mali", "Ghana", "Kenya", "Zimbabwe", "Gabon", "Burkina Faso", "Mauritania", "Senegal",
                           "Guinea", "Chad", "Comoros", "Botswana", "Malawi", "Benin", "Lesotho", "Mozambique",
                           "Zambia", "Sierra Leone", "Tanzania", "Cameroon", "Liberia", "South Sudan",
                           "Equatorial Guinea", "Angola", "Somalia", "Ethiopia", "Ivory Coast", "Swaziland",
                           "Democratic Republic of the Congo", "The Gambia"],

    'Türkiye': ["Turkish Republic of Northern Cyprus", "Turkey"],
    'Turkmenistan': ["Turkmenistan"],


}


# Create a function to map the country to the respective region
def find_region(country):
    for region, countries in regions.items():
        if country in countries:
            return region
    return 'Unknown'  # Return 'Unknown' if the country is not found in any region


# Apply the function to the DataFrame to create a new 'Region' column
df['Region'] = df['Nationality'].apply(find_region)

# Save the updated DataFrame to a new Excel file
output_file_path = 'regions.xlsx'  # Replace with your desired output file path
df.to_excel(output_file_path, index=False)

print("The region mapping has been successfully added to the Excel file.")
