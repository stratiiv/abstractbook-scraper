# PDF to Excel Data Extraction Script

This Python script is designed to extract data from a PDF file containing abstracts of articles. It parses the content of the PDF, extracts relevant information such as session names, article titles, author names and affiliations, and presentation abstracts, and then populates an Excel spreadsheet with the extracted data.

## Purpose
The script is specifically tailored to parse the PDF file named "Abstracts from the 5th World Psoriasis & Psoriatic Arthritis Conference 2018". It handles different patterns found within the PDF content.

## Usage
1. Clone the repository:
```bash
git clone https://github.com/stratiiv/abstractbook-scraper.git
```
2. Install the required dependencies using Pipenv:
```bash
pipenv install
```
3. Activate the virtual environment:
```bash
pipenv shell
```
4. Make sure the PDF file `abstractbook.pdf` is in the same directory as this script.
5. Run the script:
```bash
python scraper.py
```

5. The script will parse the PDF, extract the relevant data, and generate an Excel file named `output.xlsx` containing the extracted information.

Please note that the script includes manual handling for specific articles with predefined patterns. Make sure to customize the script if you encounter similar cases in other PDF files.

## PDF Source
`Abstracts from the 5th World Psoriasis & Psoriatic Arthritis Conference 2018`
