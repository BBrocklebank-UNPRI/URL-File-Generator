import requests, re, os, convert, logging
from openpyxl import load_workbook
from requests.exceptions import RequestException
from convert import PdfGenerator

logging.basicConfig(level=logging.WARNING)

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
    #pdf_file = []

    #fname = url.rsplit('/', 1)[1]
    fname = url
    #if len(fname) == 0:
       #fname = url

    #if ".pdf" not in url :
            #fname = f"{fname}.pdf"
            #Convert to PDF
            #pdf_file = PdfGenerator([url]).main()

    invalid = '<>:"/\|?* '
    for char in invalid:
        fname = fname.replace(char, '')

    if len(fname) > 256:
        fname = fname[:250]

    #if not pdf_file :
        #return fname, None
    
    #else :
        #return fname, pdf_file[0]
    return fname


def downloadPdf(urls):
    """
    Loop URLs, check content type, save
    """

    for url in urls:
        print(url)
        if "https://" not in url :
            url = f"https://{url}"

        try:
            with requests.get(url, allow_redirects=True, headers=headers) as r:
                contentType = r.headers.get('content-type')
                print(contentType)
            
                fname = nameFile(r.url)

                if 'application/pdf' in contentType:
                    f = open(f"{fname}.pdf", 'wb')
                    f.write(r.content)
                    f.close

                elif 'text/html' in contentType:
                    pdf_file = PdfGenerator([url]).main()
                    with open(f"{fname}.pdf", "wb") as outfile:
                        outfile.write(pdf_file[0].getbuffer())

            logging.warning(f"Organisation Name: URL: {url} Result: Success")

        except RequestException as e:
            logging.warning(f" Organisation Name: URL: {url} Result: Failed")

fetchURLs(urls)