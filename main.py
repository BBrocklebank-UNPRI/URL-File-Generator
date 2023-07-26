from openpyxl import load_workbook
import requests

data_file = 'data/URLs_without_files.xlsx'

# Provide PDF host server with credentials (prevent 406 error)
headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; PPC Mac OS X 10_8_7 rv:5.0; en-US) AppleWebKit/533.31.5 (KHTML, like Gecko) Version/4.0 Safari/533.31.5',
}

# Load the entire workbook.
wb = load_workbook(data_file, data_only=True)

# Load specific worksheet.
ws = wb['0kbUrls']

# Store URLs list
urls = []

# Append column 3 values to urls list
for row in ws:
    url = row[3].value
    urls.append(url)

# name pdfs correctly
# save pdfs to folder
# loop through urls

# Download PDF
downloadDoc = 'https://aipmanagement.dk/wp-content/uploads/2023/06/AIP_ESG-report_2022.pdf'
r = requests.get(downloadDoc, allow_redirects=True, headers=headers)

open('test_download1.pdf', 'wb').write(r.content)
#open('test_download2.pdf', 'wb').close()

