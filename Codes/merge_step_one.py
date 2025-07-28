# This file was created by Pejman Ebrahimi (23.07.2025)
"""
We need to merge our previous 50 data entries with the newly 139 collected data.
After execution, we will have one 'pre-interaction' and one 'post-interaction' Excel file (189 data).
"""
# merge_excel_files.py

import pandas as pd
import os

def merge_excel_files(file1, file2, output_file):
    """
    Merges two Excel files with the same column names and saves the result to a new file.

    Parameters:
    - file1: Path to the first Excel file
    - file2: Path to the second Excel file
    - output_file: Path for the merged output Excel file
    """
    # Read the Excel files
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Merge the data
    merged_df = pd.concat([df1, df2], ignore_index=True)

    # Save to a new Excel file
    merged_df.to_excel(output_file, index=False)
    print(f"Merged file saved as: {output_file}")

if __name__ == "__main__":
    # Define the folder and filenames
    folder = "Excel_Data"
    file1 = os.path.join(folder, "Post-Interaction50.xlsx")
    file2 = os.path.join(folder, "Post-Interaction140.xlsx")
    output_file = os.path.join(folder, "Post-Interaction.xlsx")

    merge_excel_files(file1, file2, output_file)
