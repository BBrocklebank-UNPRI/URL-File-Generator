import requests, re, os
from openpyxl import load_workbook
from requests.exceptions import RequestException

# Define file to scan
data_file = 'data/URLs_without_files.xlsx'

# Credential to prevnt 406 error
headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh; PPC Mac OS X 10_8_7 rv:5.0; en-US) AppleWebKit/533.31.5 (KHTML, like Gecko) Version/4.0 Safari/533.31.5',
}

# Load the entire workbook
wb = load_workbook(data_file, data_only=True)

# Load specific worksheet
ws = wb['0kbUrls']

# Store URLs
urls = []

# Filter values, add to list
for row in ws:
    url = row[3].value
    testList = ['https:', 'www.', '.pdf', '.com', '.co.uk']
    if any(testString in url for testString in testList):
        urls.append(url)



def nameFile (url):
    """Validate file names using URL"""

    fname = url.rsplit('/', 1)[1]

    if len(fname) == 0:
        fname = url

    if ".pdf" not in fname :
            fname = f"{fname}.html"
            print(f"Split: {fname}")
            #convert to PDF

    invalid = '<>:"/\|?* '
    for char in invalid:
        fname = fname.replace(char, '')

    if len(fname) > 256:
        fname = fname[:250]
    
    return fname

def download():
    """Validate URLs/Download"""

    for url in urls:

        if "https://" not in url :
            url = f"https://{url}"

        try:
            with requests.get(url, allow_redirects=True, headers=headers) as r:
            
                fname = nameFile(r.url)

                f = open(fname, 'wb')
                f.write(r.content)
                f.close


        except RequestException as e:
            print(e)

download()