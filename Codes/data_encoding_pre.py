"""
This script was written by Pejman Ebrahimi (pejman.ebrahimi@uni.li)

Purpose:
--------
This script reads an Excel file containing Likert scale responses for pre-interaction survey.
It applies different string-to-numeric mappings to different column ranges
for analysis.

Column mappings:
----------------
- Columns C to G (ordinal scale from 1 to 7):
    "Strongly Disagree" → 1
    "Disagree"          → 2
    "Somewhat Disagree" → 3
    "Neutral"           → 4
    "Somewhat Agree"    → 5
    "Agree"             → 6
    "Strongly Agree"    → 7

- Columns H to M (balanced scale from -3 to 3):
    "Strongly Disagree" → -3
    "Disagree"          → -2
    "Somewhat Disagree" → -1
    "Neutral"           → 0
    "Somewhat Agree"    → 1
    "Agree"             → 2
    "Strongly Agree"    → 3
"""

import pandas as pd
import os

# Load the Excel file
file_path = os.path.join("Excel_Data", "Pre-Interaction.xlsx")
df = pd.read_excel(file_path)

# Define column groups (0-based index: A=0, B=1, ..., C=2, ..., M=12)
columns_scale_1 = df.columns[2:7]  # Columns C to G
columns_scale_2 = df.columns[7:13] # Columns H to M

# Define mappings
scale_1_map = {
    "Strongly Disagree": 1,
    "Disagree": 2,
    "Somewhat Disagree": 3,
    "Neutral": 4,
    "Somewhat Agree": 5,
    "Agree": 6,
    "Strongly Agree": 7
}

scale_2_map = {
    "Strongly Disagree": -3,
    "Disagree": -2,
    "Somewhat Disagree": -1,
    "Neutral": 0,
    "Somewhat Agree": 1,
    "Agree": 2,
    "Strongly Agree": 3
}

# Apply the mappings to respective columns
df[columns_scale_1] = df[columns_scale_1].replace(scale_1_map)
df[columns_scale_2] = df[columns_scale_2].replace(scale_2_map)

# Save the converted DataFrame to a new file
output_path = os.path.join("Excel_Data", "Pre-Interaction_Converted.xlsx")
df.to_excel(output_path, index=False)

print("Conversion complete. File saved to:", output_path)
