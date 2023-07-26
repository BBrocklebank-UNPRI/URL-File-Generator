from openpyxl import load_workbook
import requests, re

data_file = 'data/URLs_without_files.xlsx'

# Credential to prevnt 406 error
headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; PPC Mac OS X 10_8_7 rv:5.0; en-US) AppleWebKit/533.31.5 (KHTML, like Gecko) Version/4.0 Safari/533.31.5',
}

# Load the entire workbook.
wb = load_workbook(data_file, data_only=True)

# Load specific worksheet.
ws = wb['0kbUrls']

# Store URLs
urls = []

# Filter values, add to list
for row in ws:
    url = row[3].value
    testList = ['https:', 'www.', '.pdf']
    if any(testString in url for testString in testList):
        urls.append(url)

print(urls)

# save pdfs to folder
# loop through urls

def getFilename_fromCd(cd, url):
    """
    #Get filename from content-disposition
    """
    if not cd:
        if url.find('/'):
            fname = url.rsplit('/', 1)[1]
            return fname
    fname = re.findall('filename=(.+)', cd)
    if len(fname) == 0:
         if url.find('/'):
            fname = url.rsplit('/', 1)[1]
            return fname
    return fname[0]

# Download PDF
##for url in urls:

downloadUrl = "https://www.bancobpi.pt/contentservice/getContent?documentName=PR_UCMS02081560"
r = requests.get(downloadUrl, allow_redirects=True, headers=headers)
filename = getFilename_fromCd(r.headers.get('content-disposition'), downloadUrl)
open(filename, 'wb').write(r.content)
#open('test_download2.pdf', 'wb').close() """

