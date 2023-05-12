import pandas as pd
import numpy as np


# Chichi Qi code for examining samples in intersecting orders
sheet_name = 'Sheet2'

# Read the Excel file
df = pd.read_excel('Scenario 1.xlsx', sheet_name='Sheet2')

# Create an empty dictionary to store the item counts
item_counts = {}

# Loop through each column
for column in df.columns:
    # Loop through each item in the column
    for item in df[column]:
        # Check if the item already exists in the dictionary
        if item in item_counts:
            # Increment the count of the item
            item_counts[item] += 1
        else:
            # Add the item to the dictionary with a count of 1
            item_counts[item] = 1

del item_counts[np.nan]

# Print the items that appear more than twice
print("Items that appear more than twice:")
for item, count in item_counts.items():
    if count > 2:
        print(item)

# Get the item that appears the most
most_common_item = max(item_counts, key=item_counts.get)
print("The most common item is:", most_common_item)
