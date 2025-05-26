"""
This file includes runner code for Automated PPT Generation
for Male and Female Athletes.
"""

# Import required modules
import pandas as pd
import sys

sys.path.append("Classes")
import ppt_generator

# Get PPT filter data ready
PPT_DATA = {
    "Male": {
        "name": "Male",
    },
    "Female": {
        "name": "Female",
    },
}

# Load data
complete_data = pd.read_excel("processed_data.xlsx")
personal_data = pd.read_excel("personal_data.xlsx")


# Instantiate PPT Generator Class
runner = ppt_generator.PowerPointGenerator(
    data=complete_data, personal_data=personal_data
)

# Loop through each Key as in each PPT we wish to create

for key, item in PPT_DATA.items():
    print(f"PPT Generating for: {key}")
    print(item)

    runner._handle_general_flow(key=key, name=item["name"])

    print("\n\n")

print("🙂 Done: Generating PPTs")
