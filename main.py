import requests, re, os, convert
from openpyxl import load_workbook
from requests.exceptions import RequestException
from convert import PdfGenerator

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

def fetchURLs (urls):
    """
    Loop Excel file, add to list
    """
    for row in ws:
        url = row[3].value
        #testList = ['https:', 'www.', '.pdf', '.com', '.co.uk']
        #if any(testString in url for testString in testList):
        urls.append(url)
    urls = list( dict.fromkeys(urls) )
    print(urls)
    downloadPdf(urls)


def nameFile (url):
    """
    Validate file names/Format HTML
    """
    pdf_file = []

    print(f"First URL: {url}")
    fname = url.rsplit('/', 1)[1]
    print(f"Split Name: {fname}")

    if len(fname) == 0:
        fname = url
        print(f"Zero Name: {fname}")

    if ".pdf" not in url :
            fname = f"{fname}.pdf"
            #print(f"PDF Extension Added: {fname}")
            #Convert to PDF
            pdf_file = PdfGenerator([url]).main()

    invalid = '<>:"/\|?* '
    for char in invalid:
        fname = fname.replace(char, '')

    if len(fname) > 256:
        fname = fname[:250]

    if not pdf_file :
        return fname, None
    
    else :
        return fname, pdf_file[0]


def downloadPdf(urls):
    """
    Loop through Urls/Save PDFs
    """

    for url in urls:

        if "https://" not in url :
            url = f"https://{url}"

        try:
            with requests.get(url, allow_redirects=True, headers=headers) as r:
            
                fname_PDF = nameFile(r.url)
                print(f"fname_PDF:{fname_PDF}")

                if fname_PDF[1] == None : #logic error, check for second value
                    print(f"No converted PDF:{fname_PDF}")
                    f = open(fname_PDF[0], 'wb')
                    f.write(r.content)
                    f.close
                
                else :
                    print(f"Converted PDF present:{fname_PDF}")
                    with open(fname_PDF[0], "wb") as outfile:
                        outfile.write(fname_PDF[1].getbuffer())


        except RequestException as e:
            print(e)

fetchURLs(urls)