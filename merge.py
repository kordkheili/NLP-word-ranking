import pandas as pd
import numpy as np
from nltk.stem import PorterStemmer

# # Function to extract the root of a word (basic example)
# def get_root(word):
#     # This is a simple example; you can modify it based on your needs
#     # Here, we just return the word itself for demonstration
#     # You can implement more complex logic or use libraries like nltk or spacy
#     return word.split()[0]  # Example: take the first part of the word

stemmer = PorterStemmer()
# Read the Excel file into a DataFrame
df = pd.read_excel('input-all.xlsx', sheet_name='all2')

# Convert the 'Word' column to string data type
df['Word'] = df['Word'].astype(str).str.lower()

# Create a new column 'Root' and apply the get_root function
df['Root'] = df['Word'].apply(stemmer.stem)

# Sort the DataFrame by the 'Word' column
df = df.sort_values('Word')

# Reset the index of the DataFrame
df = df.reset_index(drop=True)

# Create a list to store the starting row index for each unique 'Word' value
start_rows = [0]
for i in range(1, len(df)):
    if df.loc[i, 'Word'] != df.loc[i-1, 'Word']:
        start_rows.append(i)

# Merge the rows based on the 'Word' column
for i in range(len(start_rows)):
    start_row = start_rows[i]
    if i == len(start_rows) - 1:
        end_row = len(df) - 1
    else:
        end_row = start_rows[i+1] - 1
    
    if end_row == start_row:
        df.at[start_row, 'Synonyms'] = str(df.loc[start_row, 'Synonyms']) if pd.notna(df.loc[start_row, 'Synonyms']) else ''
        df.at[start_row, 'Meaning'] = str(df.loc[start_row, 'Meaning']) if pd.notna(df.loc[start_row, 'Meaning']) else ''
        df.at[start_row, 'Persian'] = str(df.loc[start_row, 'Persian']) if pd.notna(df.loc[start_row, 'Persian']) else ''
        df.at[start_row, 'Example'] = str(df.loc[start_row, 'Example']) if pd.notna(df.loc[start_row, 'Example']) else ''
    else:
        df.at[start_row, 'Synonyms'] = '\n'.join(map(str, df.loc[start_row:end_row, 'Synonyms'].fillna('')))
        df.at[start_row, 'Meaning'] = '\n'.join(map(str, df.loc[start_row:end_row, 'Meaning'].fillna('')))
        df.at[start_row, 'Persian'] = '\n'.join(map(str, df.loc[start_row:end_row, 'Persian'].fillna('')))
        df.at[start_row, 'Example'] = '\n'.join(map(str, df.loc[start_row:end_row, 'Example'].fillna('')))
        df = df.drop(range(start_row+1, end_row+1))

# Write the merged DataFrame back to an Excel file
df.to_excel('output-all.xlsx', index=False)