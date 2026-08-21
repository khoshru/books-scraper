# Books Scraper

A Python scraper that collects book data from books.toscrape.com
and exports it to an Excel file.

## What it does

- Fetches book listings across multiple pages
- Extracts title, price, and image URL for each book
- Saves everything into a single .xlsx file

## Stack

- requests — HTTP requests
- BeautifulSoup — HTML parsing
- openpyxl — Excel output

## Usage

python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 scraper.py

Output: books.xlsx
