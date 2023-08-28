import re
from PyPDF2 import PdfReader
from openpyxl import Workbook

FLAGS = re.DOTALL | re.MULTILINE

reader = PdfReader('abstractbook.pdf')
warning_tag = r'\1\n---TO FILL---\n---to fill---\nIntroduction: ---to fill---\2'
article_pattern = r'P\d{3}.*?(?=P\d{3})'
standard_parse_pattern = (
                    r'(P\d{3})([/\\A-Z()\s’;:,.\xad0-9-–*@\“”?Α&^=%#!+]+$)'
                    r'(.*?)(?=Introduction|Background|Objective)(.*)'
)
non_standard_parse_pattern = (
    r'(P\d{3})([/\\A-Z()\s’;:,.\xad0-9-–*@\“”?Α&^=%#!+]+$)'
)

workbook = Workbook()
worksheet = workbook.active
worksheet.title = 'Abstracts'
worksheet.append(['Session Name', 'Title', 'Names and Affiliation',
                  'Presentation Abstract'])

pdf_text = ""
for page in reader.pages[6:64]:  # 64th page is last
    page_text = page.extract_text()
    pdf_text += page_text

# articles with "PXXX" in text have to be manually filled
pdf_text = re.sub(r'(P035).*?(P036)', warning_tag,
                  pdf_text, flags=FLAGS)
pdf_text = re.sub(r'(P052).*?(P053)', warning_tag,
                  pdf_text, flags=FLAGS)
pdf_text = re.sub(r'(P061).*?(P062)', warning_tag,
                  pdf_text, flags=FLAGS)

articles = re.findall(article_pattern, pdf_text, FLAGS)

# include last article
last_article = re.search(r'P155.*', pdf_text, FLAGS)
articles.append(last_article[0])

# remove empty strings
articles = [article for article in articles if article.strip()]

# parsing articles
for article in articles:
    header = re.search(standard_parse_pattern, article, FLAGS)
    if header:
        session_id = header[1]
        title = header[2]
        names_and_affilation = header[3]
        content = header[4]
    else:
        header = re.match(non_standard_parse_pattern, article, FLAGS)
        session_id = header[1]
        title = header[2]
        names_and_affilation = "---to fill---"
        content = "Introduction: ---to fill---"
    worksheet.append([session_id, title, names_and_affilation, content])

workbook.save("output.xlsx")
