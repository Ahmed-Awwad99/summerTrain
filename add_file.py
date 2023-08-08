import os
import pandas as pd

folder_path = 'data'

# get all xls files
xls_files = [f for f in os.listdir(folder_path) if f.endswith('.xls')]

dfs = []

for file in xls_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path)
    
    # remove unuseful rows
    title_row_index = df.index[df.iloc[:, 0].str.contains("Title", na=False)].min()
    
    # Set the header
    df.columns = df.iloc[title_row_index]
    
    # Remove the rows above the header row
    df = df.iloc[title_row_index + 1:]
    
    dfs.append(df)

# Concatenate the list of DataFrames into one DataFrame
combined_df = pd.concat(dfs, ignore_index=True)

# Reset the index of the combined DataFrame and rename the index column
combined_df.reset_index(drop=True, inplace=True)
combined_df.rename_axis('index', axis=1, inplace=True)

# Remove columns that contain only zeros
combined_df = combined_df.loc[:, (combined_df != 0).any()]

columns_to_remove = [
 'Editors',
 'Book Editors',
 'Source Title',
 'Volume',
 'Issue',
 'Part Number',
 'Supplement',
 'Special Issue',
 'Beginning Page',
 'Ending Page',
 'Article Number',
 'Conference Title',
 'Conference Date'
]

output_file_path = 'output_combined.xlsx'  # Change the file extension to .xlsx
combined_df.to_excel(output_file_path, index=False)

print("Combined DataFrame is saved to:", output_file_path)

# Remove unwanted columns
combined_df.drop(columns=columns_to_remove, inplace=True)

output_dataset_path = 'dataset.xlsx'  # Change the file extension to .xlsx
combined_df.to_excel(output_dataset_path, index=False)

print("Combined dataset is saved to:", output_dataset_path)
