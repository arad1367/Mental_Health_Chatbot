"""
Code written by Pejman Ebrahimi (pejman.ebrahimi@uni.li or pejman.ebrahimi77@gmail.com)

This script converts specified columns in the Excel file 'Post-Interaction.xlsx' 
from string responses to numeric scores according to the following schemes:

Columns C to G follow:
    Strongly Agree → 3
    Agree → 2
    Somewhat Agree → 1
    Neutral → 0
    Somewhat Disagree → -1
    Disagree → -2
    Strongly Disagree → -3

Columns H and K follow:
    High level of Satisfaction
    Strongly Agree → 3
    Agree → 2
    Somewhat Agree → 1
    Neutral → 0
    Somewhat Disagree → -1
    Disagree → -2
    Strongly Disagree → -3

Columns I, J, L, M, N follow:
    Reverse Scoring for negative Questions
    Strongly Agree → -3
    Agree → -2
    Somewhat Agree → -1
    Neutral → 0
    Somewhat Disagree → 1
    Disagree → 2
    Strongly Disagree → 3

Columns S and T follow:
    Reverse Scoring for negative Questions (same as above)
    Strongly Agree → -3
    Agree → -2
    Somewhat Agree → -1
    Neutral → 0
    Somewhat Disagree → 1
    Disagree → 2
    Strongly Disagree → 3

Columns U, V, O, P, Q, R follow:
    Scoring on Outcomes (same as C-G and H,K)
    Strongly Agree → 3
    Agree → 2
    Somewhat Agree → 1
    Neutral → 0
    Somewhat Disagree → -1
    Disagree → -2
    Strongly Disagree → -3
"""

import pandas as pd

# Load data
file_path = "Excel_Data/Post-Interaction.xlsx"
df = pd.read_excel(file_path)

# Define mappings
standard_score = {
    "Strongly Agree": 3,
    "Agree": 2,
    "Somewhat Agree": 1,
    "Neutral": 0,
    "Somewhat Disagree": -1,
    "Disagree": -2,
    "Strongly Disagree": -3
}

reverse_score = {
    "Strongly Agree": -3,
    "Agree": -2,
    "Somewhat Agree": -1,
    "Neutral": 0,
    "Somewhat Disagree": 1,
    "Disagree": 2,
    "Strongly Disagree": 3
}

# Columns grouped by scoring scheme (using Excel letter columns)
# Note: pandas uses zero-based numeric indices, but it's easier to just use column names or indices as strings here

# Convert Excel columns to zero-based indices to help:
# A=0, B=1, C=2, ... etc.

# C to G => columns 2 to 6 (inclusive)
cols_standard_1 = df.columns[2:7].tolist()  # C-G

# H and K => columns 7 and 10
cols_standard_2 = [df.columns[7], df.columns[10]]

# I, J, L, M, N => columns 8, 9, 11, 12, 13
cols_reverse_1 = [df.columns[i] for i in [8, 9, 11, 12, 13]]

# S and T => columns 18, 19
cols_reverse_2 = [df.columns[18], df.columns[19]]

# U, V, O, P, Q, R => columns 20, 21, 14, 15, 16, 17
cols_standard_3 = [df.columns[i] for i in [20, 21, 14, 15, 16, 17]]

# Apply mappings
for col in cols_standard_1 + cols_standard_2 + cols_standard_3:
    df[col] = df[col].map(standard_score)

for col in cols_reverse_1 + cols_reverse_2:
    df[col] = df[col].map(reverse_score)

# Save to a new Excel file or overwrite
output_path = "Excel_Data/Post-Interaction_Converted.xlsx"
df.to_excel(output_path, index=False)
