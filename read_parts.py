import pandas as pd
import json

# Read the Excel file
df = pd.read_excel('Parts Export.xlsx')

# Display column information
print("\nColumn names in the Excel file:")
print(df.columns.tolist())

print("\nFirst few rows of data:")
print(df.head())

print("\nSample row with all columns:")
print(df.iloc[0])

# Create a structured dictionary
parts_dict = {}
for _, row in df.iterrows():
    part_id = str(row.iloc[0])  # Convert ID to string
    description = row.iloc[1]    # Description is in second column
    
    # Extract size from description (assuming format like "1000x2400")
    size = ""
    if "x" in description:
        size_parts = description.split("x")
        if len(size_parts) >= 2:
            height = size_parts[0].strip()
            width = size_parts[1].split()[0].strip()
            size = f"{height}x{width}"
    
    parts_dict[part_id] = {
        "description": description,
        "size": size,
        "pagination": "No of Pages",  # Placeholder as this column wasn't visible in sample
        "material": "Material",       # Placeholder as this column wasn't visible in sample
        "finishing": "Production Finishing Notes"  # Placeholder as this column wasn't visible in sample
    }

# Save to a Python file
with open('parts_data.py', 'w') as f:
    f.write('PARTS_DATA = {\n')
    for part_id, data in parts_dict.items():
        f.write(f"    '{part_id}': {{\n")
        f.write(f"        'description': '{data['description']}',\n")
        f.write(f"        'size': '{data['size']}',\n")
        f.write(f"        'pagination': '{data['pagination']}',\n")
        f.write(f"        'material': '{data['material']}',\n")
        f.write(f"        'finishing': '{data['finishing']}'\n")
        f.write('    },\n')
    f.write('}\n')

print(f"Created parts_data.py with {len(parts_dict)} items") 