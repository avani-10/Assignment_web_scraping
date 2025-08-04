# 🧠 Blackcoffer NLP Assignment – Web Scraping & Text Analysis

This project is a solution to Blackcoffer’s data extraction and NLP assignment. It automates the process of scraping article content from URLs, cleaning the text, and performing NLP-based sentiment and readability analysis to generate structured insights.

---

## 🚀 Objective

- Extract **article title and main content** from URLs listed in `Input.xlsx`
- Ignore headers, footers, ads, and irrelevant sections
- Perform NLP analysis as per `Text Analysis.docx`
- Output the result in the format of `Output Data Structure.xlsx`

---

## 💡 My Approach

1. **Web Scraping**
   - Used `requests` + `BeautifulSoup` to fetch and parse HTML.
   - Extracted only the article `<title>` and `<p>` text body.
   - Saved each article as a `.txt` file named after `URL_ID`.

2. **Text Cleaning**
   - Removed extra whitespace, HTML tags, and non-content elements.
   - Normalized all text to lowercase.

3. **NLP Analysis**
   - Calculated all metrics: sentiment scores, word/sentence averages, fog index, personal pronouns, syllables, etc.
   - Used `nltk`, `textstat`, and `regex` for analysis.

4. **Final Output**
   - Created an Excel output (`output.xlsx`) using the same structure as `Output Data Structure.xlsx`.

---

## 🔧 Project Setup

### 📁 Prerequisites
Make sure you have Python 3.x installed.

### 📦 Install Dependencies
```bash
pip install -r requirements.txt
