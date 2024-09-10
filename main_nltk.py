import pandas as pd
import nltk
from nltk import FreqDist
from nltk.corpus import brown, gutenberg, webtext, twitter_samples, movie_reviews

# Ensure that the Brown corpus is downloaded
# nltk.download('brown')
# nltk.download('gutenberg')
# nltk.download('webtext')
# nltk.download('twitter_samples')
# nltk.download('movie_reviews')




# Step 1: Read the Excel file
df = pd.read_excel('input.xlsx', sheet_name='Sheet1')

# Step 2: Create a frequency distribution of words in the Brown corpus
brown_freqs = FreqDist([w.lower() for w in brown.words()])
gutenberg_freqs = FreqDist([w.lower() for w in gutenberg.words()])
webtext_freqs = FreqDist([w.lower() for w in webtext.words()])
twitter_freqs = FreqDist([w.lower() for w in twitter_samples.strings()])
movie_reviews_freqs = FreqDist([w.lower() for w in movie_reviews.words()])


# Step 3: Combine frequencies from all corpora
combined_freqs = brown_freqs.copy()
for freq in [gutenberg_freqs, webtext_freqs, twitter_freqs, movie_reviews_freqs]:
    for word, count in freq.items():
        combined_freqs[word] += count

# Step 3: Add a new column to the DataFrame with the frequencies
df['Frequency'] = df['Word'].apply(lambda x: combined_freqs[x.lower()] if x.lower() in combined_freqs else 0)

# Step 4: Sort the DataFrame by the Frequency column in descending order
sorted_df = df.sort_values(by='Frequency', ascending=False)

# Step 5: Write the sorted DataFrame to a new Excel file
sorted_df.to_excel('output.xlsx', index=False)