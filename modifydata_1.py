import pandas as pd
import re

# Load the Excel file
input_file = "CBP_Ruling.xlsx"
output_file = "new_CBP_Ruling.xlsx"

# Read the Excel file
df = pd.read_excel(input_file)

# Assuming the third column is the one with the text to process
# Adjust the column index if necessary
third_column = df.iloc[:, 2]

# Initialize lists to store the extracted data
categories = []
tariff_nos = []
texts = []

# Process each row in the third column
for text in third_column:
    category_match = re.search(r"CATEGORY:\s+(.*)", text)
    tariff_no_match = re.search(r"TARIFF NO.:\s+([\d.]+)", text)

    category = category_match.group(1) if category_match else None
    tariff_no = tariff_no_match.group(1) if tariff_no_match else None

    categories.append(category)
    tariff_nos.append(tariff_no)
    texts.append(text)

# Create a new DataFrame with the extracted data
new_df = pd.DataFrame({
    "CATEGORY": categories,
    "TARIFF NO": tariff_nos,
    "text": texts
})

# Save the new DataFrame to a new Excel file
new_df.to_excel(output_file, index=False)

print(f"Processed data has been saved to {output_file}")
