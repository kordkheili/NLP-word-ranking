import pandas as pd
import spacy
from collections import Counter

# Step 1: Load the spaCy English model
nlp = spacy.load('en_core_web_trf')

# Step 2: Read the Excel file
df = pd.read_excel('input.xlsx', sheet_name='Sheet1')

# Step 3: Process the words in the DataFrame using spaCy
text = ' '.join(df['Word'])
doc = nlp(text)

# Step 4: Count the frequency of each word, excluding stop words and punctuation
word_freq = Counter([token.text.lower() for token in doc if not token.is_stop and not token.is_punct])

# Step 5: Add a new column to the DataFrame with the frequencies
df['Frequency'] = df['Word'].apply(lambda x: word_freq[x.lower()] if x.lower() in word_freq else 0)

# Step 6: Sort the DataFrame by the Frequency column in descending order
sorted_df = df.sort_values(by='Frequency', ascending=False)

# Step 7: Write the sorted DataFrame to a new Excel file
sorted_df.to_excel('output.xlsx', index=False)

print("Word frequencies have been sorted and saved to 'output.xlsx'.")