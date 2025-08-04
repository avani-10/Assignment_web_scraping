Blackcoffer Assignment: Data Extraction & Text Analysis
👤 Author: Avani Aggarwal
Submission for: Data Extraction and NLP Test Assignment


🎯 Objective
This project extracts article content from URLs given in Input.xlsx, processes each article using NLP techniques, and generates an output Excel file containing various textual analysis metrics.
All steps are implemented in Python using libraries like BeautifulSoup, NLTK, and pandas.

Script role:
Reads URLs from Input.xlsx.
Extracts the article title and body only — skipping headers, footers, ads, etc.
Cleans and processes the text using NLP.
Computes 13 linguistic variables, as listed in the provided Text Analysis.docx.
Outputs the results in Output.xlsx, matching the format in “Output Data Structure.xlsx”.

📦 Files in Submission
File Name	      Purpose
main.py	          Main script to extract, analyze, and generate results.
Input.xlsx	      Input file with URL IDs and article links.
Output.xlsx	      Final auto generated output file with all variables and input columns.
requirements.txt  Python dependencies required to run the script.
StopWords/	      Folder with various stopword files used for cleaning text.
MasterDictionary/ Contains positive-words.txt and negative-words.txt.
extracted_articles/Folder auto-created to store downloaded articles as text files.

🔧 How to Run the Script
1. Clone/Download and Navigate
Make sure all files are in the same directory:

main.py
Input.xlsx
requirements.txt
StopWords/
MasterDictionary/

2. Install Dependencies
Use pip to install required packages:

pip install -r requirements.txt


3. Run the Script

python main.py 


The script will:
Create a folder called extracted_articles
Fetch each article
Analyze text and compute all metrics
Save the final result in Output.xlsx

📊 Output Variables
Each row in Output.xlsx will contain:

All original input columns from Input.xlsx

The following analysis metrics:

POSITIVE SCORE

NEGATIVE SCORE

POLARITY SCORE

SUBJECTIVITY SCORE

AVG SENTENCE LENGTH

PERCENTAGE OF COMPLEX WORDS

FOG INDEX

AVG NUMBER OF WORDS PER SENTENCE

COMPLEX WORD COUNT

WORD COUNT

SYLLABLE PER WORD

PERSONAL PRONOUNS

AVG WORD LENGTH


⚙️ Methodology Summary


Read input URLs from Input.xlsx.

Scrape articles using requests and BeautifulSoup, extracting only the title and main text.

Save each article as a .txt file named with its URL_ID.

Load stopwords from multiple files and load sentiment word lists (positive & negative).

Clean and tokenize text using nltk: remove punctuation, stopwords, and non-alphanumeric tokens.

Calculate sentiment metrics like Positive Score, Negative Score, Polarity, and Subjectivity.

Compute readability metrics like Avg Sentence Length, Fog Index, Syllables per Word, etc.

Handle missing/invalid articles by skipping them and assigning None for analysis fields.

Combine results with input data and save final output to Output.xlsx.