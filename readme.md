Text Analysis and Sentiment Analysis
This project explores techniques for analyzing textual data from websites and extracting sentiment information.

Features:

Scrapes website content using URLs provided in an Excel file.
Performs sentiment analysis by classifying the text as positive, negative, or neutral.
Calculates various metrics related to readability, complexity, and personal pronoun usage.
Outputs the results to a CSV file.
Requirements:

Python 3.x
Libraries:
requests
BeautifulSoup4
openpyxl
nltk
csv
os
re
Installation:

Ensure you have Python 3 installed. You can check by running python --version in your terminal. If you don't have it, download the installer from https://www.python.org/downloads/.

Clone or download the project files.

Navigate to the project directory using the cd command in your terminal.

Install the required libraries by running the following command:

Bash
pip install -r requirements.txt


This will download and install all the necessary libraries from the Python Package Index (PyPI).

Usage:

Make sure you have an Excel file named Input.xlsx containing URLs in the specified sheet and column (see code for details).

Ensure you have separate text files containing positive and negative words (e.g., positive-words.txt, negative-words.txt).

Run the main script using the following command:

Bash
python main.py


This will initiate the scraping, analysis, and output the results to a CSV file named Output.csv.

Explanation:

The code utilizes the following libraries for specific tasks:

requests: Fetches website content.
BeautifulSoup: Parses the HTML content of the downloaded web pages.
openpyxl: Reads data from the Excel file containing URLs for analysis.
nltk: Provides functionalities for stop word removal (currently not used in sentiment analysis).
csv: Creates and writes data to a CSV file.
os: Used for checking directory paths.
re: Used for regular expressions (counting personal pronouns).
Customization:

You can modify the script (main.py) to:
Change the Excel file name, sheet name, and URL column.
Adjust the HTML element identification logic for scraping body text based on website structures.
Implement different sentiment analysis algorithms (currently basic word counting is used).
Add additional metrics for text analysis as needed.

<h3>What does this program do</h3>
Sentiment Analysis:

Cleaning: Stop words are removed to focus on meaningful content.
Dictionary Creation: Positive and negative words are identified from a master dictionary, excluding stop words.
Score Calculation: Positive, negative, polarity, and subjectivity scores are derived based on word occurrence in the dictionaries.

Readability Analysis:

Gunning Fog Index: This formula considers sentence length and complex word percentage to assess difficulty.

Word and Sentence Analysis:

Average words per sentence is calculated.
Complex word count (words with more than two syllables) is determined.
Total word count is obtained after cleaning.

Syllable Count:

Vowels are counted in each word to estimate syllable count, with some exception handling.

Personal Pronouns:

Regular expressions are used to identify and count specific pronouns ("I," "we," etc.), excluding the country name "US."

Average Word Length:

This is calculated by dividing the total character count by the total word count.