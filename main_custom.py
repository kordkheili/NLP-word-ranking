import pandas as pd
from collections import Counter
from nltk.stem import PorterStemmer
import nltk

# Download the NLTK data files (only need to run this once)
# nltk.download('punkt')

# Initialize the stemmer
stemmer = PorterStemmer()

# Load the Excel file containing the word list
df_words = pd.read_excel('input-custom1.xlsx',sheet_name='Sheet1')
df_words['Word'] = df_words['Word'].astype(str).str.lower()

# Read the text file and count word frequencies
with open('dataset/dataset.txt', 'r', encoding='utf-8') as file:
    text = file.read()

# Tokenize the text into words
word_list = nltk.word_tokenize(text)

# Stem the words to their root forms
stemmed_words = [stemmer.stem(word) for word in word_list]

# Count frequencies of the stemmed words
word_counts = Counter(word_list)

# Create a DataFrame from the word counts
df_counts = pd.DataFrame(word_counts.items(), columns=['Stemmed_Word', 'Freq'])

# Stem the words in the Excel file for matching
df_words['Stemmed_Word'] = df_words['Word'].apply(stemmer.stem)

# Merge the two DataFrames on the stemmed word column
df_merged = pd.merge(df_words, df_counts, on='Stemmed_Word', how='left')

# Sort by frequency in descending order
df_sorted = df_merged.sort_values(by='Freq', ascending=False)

# Print the sorted DataFrame
print(df_sorted[['Word', 'Freq']])

# Optionally, save the sorted DataFrame to a new Excel file
df_sorted.to_excel('output-custom1.xlsx', index=False)