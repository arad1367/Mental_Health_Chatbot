# Author: Pejman Ebrahimi
# Emails: pejman.ebrahimi77@gmail.com | pejman.ebrahimi@uni.li

"""
This script merges two Excel files: Pre-Interaction and Post-Interaction.
Both files contain a column representing the participant's Prolific ID:
- 'Participant Prolific ID Pre' in the pre-interaction file
- 'Participant Prolific ID Post' in the post-interaction file

These two columns are used as the key to merge the files. 
"""

import pandas as pd
import os

# Define folder and file paths
data_folder = "Excel_Data"
pre_file = os.path.join(data_folder, "Pre-Interaction_Final.xlsx")
post_file = os.path.join(data_folder, "Post-Interaction_Final.xlsx")
output_file = os.path.join(data_folder, "Final_data.xlsx")

# Load the Excel files
pre_df = pd.read_excel(pre_file)
post_df = pd.read_excel(post_file)

# Rename columns for consistency (optional but helpful)
pre_df = pre_df.rename(columns={"Participant Prolific ID Pre": "Prolific_ID"})
post_df = post_df.rename(columns={"Participant Prolific ID Post": "Prolific_ID"})

# Merge the two dataframes on the common Prolific_ID
merged_df = pd.merge(pre_df, post_df, on="Prolific_ID", how="inner")

# Save the final merged data
merged_df.to_excel(output_file, index=False)

print("Merge complete. Final file saved as:", output_file)
