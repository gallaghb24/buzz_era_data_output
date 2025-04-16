import pandas as pd
import numpy as np

# Read the Excel file
df = pd.read_excel('Parts Export.xlsx')

# Print column names
print("Column names in the Excel file:")
print(df.columns.tolist())

# Convert DataFrame to dictionary
PARTS_DATA_RAW = {}
for _, row in df.iterrows():
    # Convert row to dictionary, keeping all columns
    row_dict = row.to_dict()
    # Replace nan values with None for proper string representation
    row_dict = {k: (None if pd.isna(v) else v) for k, v in row_dict.items()}
    # Use the Part URN as the key
    key = str(int(row_dict['Part URN']))  # Convert to int first to remove decimal places
    PARTS_DATA_RAW[key] = row_dict

print("\nFirst row raw data:")
print(df.iloc[0])

print("\nFirst entry in PARTS_DATA_RAW:")
first_key = list(PARTS_DATA_RAW.keys())[0]
print(f"Key: {first_key}")
print("Values:")
for k, v in PARTS_DATA_RAW[first_key].items():
    print(f"  {k}: {v}")

# Create a string representation of the dictionary
dict_str = "PARTS_DATA_RAW = " + str(PARTS_DATA_RAW)

# Write to parts_data.py
with open('parts_data.py', 'w') as f:
    f.write(dict_str)

print(f"\nCreated PARTS_DATA_RAW dictionary with {len(PARTS_DATA_RAW)} entries")
print("Dictionary has been saved to 'parts_data.py'") 