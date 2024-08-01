import requests
from bs4 import BeautifulSoup
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import re

nltk.download('punkt')

output_structure_path = 'Output Data Structure.xlsx'
output_df = pd.read_excel(output_structure_path)

# Function to extract article title and text from URL
def extract_article_text(url):
    try:
        # Send a GET request to the URL
        response = requests.get(url)
        response.raise_for_status()

        # Parse HTML content using BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract article title
        title = soup.title.text.strip()

        # Extract article text
        article_text = ''
        article_body = soup.find('div', {'class':'td-post-content tagdiv-type'})
        
        if article_body:
            for paragraph in article_body.find_all(['p', 'h2', 'h3', 'h5', 'li']):
                article_text += paragraph.text + '\n'
        else:
            article_body = soup.find_all('div', class_= 'tdb-block-inner td-fix-index')
            if article_body:
                for div in article_body:
                    if div.find('p'):
                        for paragraph in div.find_all(['p', 'h2', 'h3', 'h5', 'li']):
                            article_text += paragraph.text + '\n'
                        break

        return title, article_text
    except Exception as e:
        print(f"Error extracting data from {url}: {e}")
        return None, None
    
# Load stop words from multiple files
def load_stop_words():
    stop_words = set()
    stop_words_files = [
        'StopWords_Auditor.txt',
        'StopWords_Currencies.txt',
        'StopWords_DatesandNumbers.txt',
        'StopWords_Generic.txt',
        'StopWords_GenericLong.txt',
        'StopWords_Geographic.txt',
        'StopWords_Names.txt'
    ]

    for file in stop_words_files:
        with open(file, 'r', encoding='latin-1') as stop_file:
            stop_words.update(stop_file.read().split())

    return stop_words

# Function to perform text analysis and compute variables
def perform_text_analysis(text, stop_words):
    # Tokenize the text into words and sentences
    words = word_tokenize(text)
    sentences = sent_tokenize(text)

    # Clean the text by removing stop words and punctuations
    cleaned_words = [word.lower() for word in words if word.isalnum() and word.lower() not in stop_words]

    # Positive and negative words lists (replace with actual paths)
    positive_words = set(open('positive-words.txt').read().split())
    negative_words = set(open('negative-words.txt').read().split())

    # Extract derived variables
    positive_score = sum(1 for word in cleaned_words if word in positive_words)
    negative_score = sum(1 for word in cleaned_words if word in negative_words)
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)

    # Analysis of readability
    average_sentence_length = len(words) / len(sentences)
    complex_words_count = sum(1 for word in cleaned_words if len(word) > 2)
    percentage_complex_words = complex_words_count / len(cleaned_words)
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
    avg_words_per_sentence = len(words) / len(sentences)

    # Syllable count per word
    def count_syllables(word):
        vowels = "aeiouy"
        count = 0
        word = word.lower()
        if word[0] in vowels:
            count += 1
        for index in range(1, len(word)):
            if word[index] in vowels and word[index - 1] not in vowels:
                count += 1
        if word.endswith("e"):
            count -= 1
        if word.endswith("le") and len(word) > 2 and word[-3] not in vowels:
            count += 1
        return max(1, count)

    syllables_per_word = sum(count_syllables(word) for word in cleaned_words) / len(cleaned_words)

    # Personal pronouns count
    personal_pronouns_count = sum(1 for word in cleaned_words if word.lower() in ["i", "we", "my", "ours", "us"])

    # Average word length
    avg_word_length = sum(len(word) for word in cleaned_words) / len(cleaned_words)

    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': average_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': complex_words_count,
        'WORD COUNT': len(cleaned_words),
        'SYLLABLE PER WORD': syllables_per_word,
        'PERSONAL PRONOUNS': personal_pronouns_count,
        'AVG WORD LENGTH': avg_word_length
    }

# Read URLs from Excel file
excel_file_path = 'Input.xlsx'
df = pd.read_excel(excel_file_path)
# Iterate through each row in the DataFrame
for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    # Extract article title and text
    title, article_text = extract_article_text(url)

    if title and article_text:
    # Save the extracted data in a text file
        output_file_path = f"{url_id}.txt"
        with open(output_file_path, 'w', encoding='utf-8') as file:
            file.write(f"{title}\n\n{article_text}")

        print(f"Extracted data from {url} and saved to {output_file_path}")
    else:
        print(f"Failed to extract data from {url}")

# Load stop words
stop_words = load_stop_words()

for index, row in output_df.iterrows():
    url_id = row['URL_ID']
    file_path = f"{url_id}.txt"

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            article_text = file.read()

        # Perform text analysis
        analysis_result = perform_text_analysis(article_text, stop_words)

        # Update the output DataFrame with the analysis result
        for key, value in analysis_result.items():
            output_df.at[index, key] = value

        print(f"Text analysis for {file_path} completed.")
    except FileNotFoundError:
        print(f"File {file_path} not found.")

# Save the updated DataFrame to a new Excel file
output_df.to_excel('Output Result.xlsx', index=False)
print("Analysis results saved to 'Output Result.xlsx'.")
