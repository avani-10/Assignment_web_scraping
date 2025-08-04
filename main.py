import os
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

try:
    nltk.data.find('tokenizers/punkt')
    nltk.data.find('corpora/stopwords')
except nltk.downloader.DownloadError:
    logging.info("Downloading NLTK data (punkt, stopwords)...")
    nltk.download('punkt', quiet=True)
    nltk.download('stopwords', quiet=True)

INPUT_FILE = 'Input.xlsx'
OUTPUT_FILE = 'Output.xlsx'
STOPWORDS_DIR = 'StopWords'
MASTER_DICT_DIR = 'MasterDictionary'
ARTICLES_DIR = 'extracted_articles'

if not os.path.exists(ARTICLES_DIR):
    os.makedirs(ARTICLES_DIR)

def extract_article_text(url, url_id):
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
    except requests.RequestException as e:
        logging.error(f"Failed to fetch URL {url}: {e}")
        return False

    soup = BeautifulSoup(response.content, 'html.parser')

    title_element = soup.find('h1', class_='tdb-title-text')
    title = title_element.get_text().strip() if title_element else ""

    content_element = soup.find('div', class_='td-post-content')
    if not content_element:
        content_element = soup.find('div', class_=re.compile(r'tdb_single_content'))
    
    if not content_element:
        logging.warning(f"Could not find article content for URL_ID {url_id} ({url})")
        return False
        
    article_text = content_element.get_text(separator='\n').strip()

    full_text = f"{title}\n\n{article_text}"

    file_path = os.path.join(ARTICLES_DIR, f"{url_id}.txt")
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(full_text)
    
    logging.info(f"Successfully extracted and saved article {url_id}")
    return True

def load_stop_words(stopwords_dir):
    stop_words = set()
    for filename in os.listdir(stopwords_dir):
        with open(os.path.join(stopwords_dir, filename), 'r', encoding='ISO-8859-1') as f:
            stop_words.update(line.strip().lower() for line in f)
    return stop_words

def load_sentiment_words(master_dict_dir):
    with open(os.path.join(master_dict_dir, 'positive-words.txt'), 'r', encoding='ISO-8859-1') as f:
        positive_words = set(f.read().splitlines())
    with open(os.path.join(master_dict_dir, 'negative-words.txt'), 'r', encoding='ISO-8859-1') as f:
        negative_words = set(f.read().splitlines())
    return positive_words, negative_words

def syllable_count(word):
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word.endswith(('es', 'ed')):
        pass
    elif word.endswith('e'):
        word = word[:-1]
        
    for index in range(len(word)):
        if word[index] in vowels and (index == 0 or word[index-1] not in vowels):
            count += 1
            
    return max(1, count)

def analyze_text(text, stop_words, positive_words, negative_words):
    words = word_tokenize(text.lower())
    cleaned_words = [word for word in words if word.isalnum() and word not in stop_words]
    word_count = len(cleaned_words)

    if word_count == 0:
        return {key: 0 for key in ['POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE', 
                                   'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
                                   'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
                                   'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']}

    positive_score = sum(1 for word in cleaned_words if word in positive_words)
    negative_score = sum(1 for word in cleaned_words if word in negative_words)
    
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 1e-6)
    subjectivity_score = (positive_score + negative_score) / (word_count + 1e-6)

    sentences = sent_tokenize(text)
    sentence_count = len(sentences)
    avg_sentence_length = word_count / sentence_count if sentence_count > 0 else 0
    
    complex_word_count = sum(1 for word in cleaned_words if syllable_count(word) > 2)
    percentage_complex_words = complex_word_count / word_count if word_count > 0 else 0
    
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    total_syllables = sum(syllable_count(word) for word in cleaned_words)
    syllable_per_word = total_syllables / word_count if word_count > 0 else 0
    
    pronoun_pattern = r'\b(I|we|my|ours|us)\b'
    pronouns = re.findall(pronoun_pattern, text, re.IGNORECASE)
    personal_pronouns_count = len([p for p in pronouns if p.lower() != 'us' or (p == 'us' and p in text.split())])

    total_chars = sum(len(word) for word in cleaned_words)
    avg_word_length = total_chars / word_count if word_count > 0 else 0

    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_sentence_length,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': personal_pronouns_count,
        'AVG WORD LENGTH': avg_word_length,
    }

def main():
    logging.info("Loading input data and dictionaries...")
    try:
        df_input = pd.read_excel(INPUT_FILE)
    except FileNotFoundError:
        logging.error(f"Input file not found: {INPUT_FILE}. Please place it in the script's directory.")
        return

    stop_words = load_stop_words(STOPWORDS_DIR)
    positive_words, negative_words = load_sentiment_words(MASTER_DICT_DIR)
    
    results = []

    for index, row in df_input.iterrows():
        url_id = row['URL_ID']
        url = row['URL']
        
        logging.info(f"Processing URL_ID: {url_id} ({index + 1}/{len(df_input)})")
        
        if extract_article_text(url, url_id):
            article_path = os.path.join(ARTICLES_DIR, f"{url_id}.txt")
            with open(article_path, 'r', encoding='utf-8') as f:
                text = f.read()
            
            analysis_scores = analyze_text(text, stop_words, positive_words, negative_words)
            result_row = {**row.to_dict(), **analysis_scores}
            results.append(result_row)
        else:
            logging.warning(f"Skipping analysis for URL_ID {url_id} due to extraction failure.")
            empty_scores = {key: None for key in ['POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE', 
                                   'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
                                   'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
                                   'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH']}
            result_row = {**row.to_dict(), **empty_scores}
            results.append(result_row)

    logging.info("Analysis complete. Saving output file...")
    df_output = pd.DataFrame(results)

    output_columns = list(df_input.columns) + [
        'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
        'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
        'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
        'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
    ]
    df_output = df_output[output_columns]
    
    df_output.to_excel(OUTPUT_FILE, index=False, engine='openpyxl')
    
    logging.info(f"Successfully created output file: {OUTPUT_FILE}")

if __name__ == '__main__':
    main()