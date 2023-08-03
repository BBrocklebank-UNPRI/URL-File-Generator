import requests, re, os, logging
from openpyxl import load_workbook
from requests.exceptions import RequestException
from convert import PdfGenerator
from urllib.parse import urljoin

logging.basicConfig(level=logging.WARNING, format='%(message)s')

# Define Excel file to scan
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
urls = set()


def fetchURLs (urls):
    """
    Loop Excel column, add to set
    """
    for row in ws:
        url = row[3].value
        testList = ['https:', 'www.', '.pdf', '.com', '.co.uk']
        if any(testString in url for testString in testList):
            urls.add(url)

    downloadPdf(urls)


def nameFile (url):
    """
    Filter invalid filename characters
    """
    
    fname = url

    invalid = '<>:"/\|?* '
    for char in invalid:
        fname = fname.replace(char, '')

    if len(fname) > 256:
        fname = fname[:250]

    return fname


def downloadPdf(urls):
    """
    Loop URLs, check content type, save
    """

    pdf_content_types = {'application/pdf'}
    html_content_types = {'text/html'}

    for url in urls:

        full_url = urljoin('https://', url)

        try:
            with requests.get(full_url, allow_redirects=True, headers=headers) as r:
                contentType = r.headers.get('content-type')
            
                fname = nameFile(r.url)

                if contentType in pdf_content_types:
                    fname = f"{fname}.pdf"
                    with open(fname, 'wb', buffering=8192) as f:
                        f.write(r.content)

                elif contentType in html_content_types:
                    pdf_file = PdfGenerator([url]).main()
                    fname = f"{fname}.pdf"
                    with open(fname, "wb", buffering=8192) as outfile:
                        outfile.write(pdf_file[0].getbuffer())

            logging.warning(f"Organisation Name: URL: {full_url} Result: Success")

        except requests.exceptions.RequestException as e:
            logging.warning(f" Organisation Name: URL: {full_url} Result: Failed")

fetchURLs(urls)