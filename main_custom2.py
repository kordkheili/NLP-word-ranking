import pandas as pd
from collections import Counter
from nltk.stem import PorterStemmer
import nltk

# Download the NLTK data files (only need to run this once)
# nltk.download('punkt')

# Initialize the stemmer
stemmer = PorterStemmer()

# Load the Excel file containing the word list
df_words = pd.read_excel('input-custom2.xlsx',sheet_name='Sheet18')
df_words['Word'] = df_words['Word'].astype(str).str.lower()

# Read the text file and count word frequencies
with open('dataset/dataset.txt', 'r', encoding='utf-8') as file:
    text = file.read()

# Tokenize the text into words
word_list = nltk.word_tokenize(text)

# Count frequencies of the original words
raw_word_counts = Counter(word.lower() for word in word_list)

# Create a DataFrame from the raw word counts
df_raw_counts = pd.DataFrame(raw_word_counts.items(), columns=['Word', 'raw_rank'])

# Stem the words in the Excel file for matching
df_words['Stemmed_Word'] = df_words['Word'].apply(stemmer.stem)

# Merge the two DataFrames on the original word column to get raw_rank
df_merged = pd.merge(df_words, df_raw_counts, on='Word', how='left')

# Count frequencies of the stemmed words
stemmed_word_counts = Counter(stemmer.stem(word) for word in word_list)
df_stemmed_counts = pd.DataFrame(stemmed_word_counts.items(), columns=['Stemmed_Word', 'Freq'])

# Merge the stemmed counts with the existing merged DataFrame
df_merged = pd.merge(df_merged, df_stemmed_counts, on='Stemmed_Word', how='left')

# Sort by frequency in descending order
df_sorted = df_merged.sort_values(by='Freq', ascending=False)

# Print the sorted DataFrame
print(df_sorted[['Word', 'Freq', 'raw_rank']])

# Optionally, save the sorted DataFrame to a new Excel file
df_sorted.to_excel('output-custom2.xlsx', index=False)