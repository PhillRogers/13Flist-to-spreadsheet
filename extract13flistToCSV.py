# Purpoose:
"""	Parse a specific PDF to a CSV spreadsheet. """
# Usage:
#   Download the latest PDF from the official web site.
#   Run while in the folder where the desired PDF(s) are.
# Notes:
#   https://www.sec.gov/divisions/investment/13flists.htm -> https://www.sec.gov/files/investment/13flist2021q2.pdf
# History:
#	2025-09-29    Phill
#	Fresh minimal version to convert only. Needs manual download. Does no notifications.
#   output to CSV only.

import csv
import glob
import os.path
import sys
import pdfplumber

try:
    pdf_path = sorted(glob.glob('13flist????q?.pdf'), key=os.path.getmtime, reverse=True)[0]
    csv_path = os.path.splitext(pdf_path)[0] +'.csv'
except:
    print('Failure to find any files like "13flist????q?.pdf" in the current folder.')
    sys.exit(-1)

output_columns = ["CUSIP","CUSIP text","Issuer Name","Issuer Description","Status"]
with pdfplumber.open(pdf_path) as pdf, open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
    writer = csv.DictWriter(csv_file, quoting=csv.QUOTE_ALL, fieldnames=output_columns)
    writer.writeheader()
    for page in pdf.pages[2:]:
        pn = page.page_number
        if (pn % 10) == 0: print(f'Page: {pn}')
        box = page.crop((63,130,536,714), strict=False, relative=False) # l,t,r,b pts
        table = box.extract_text(keep_blank_chars=True)
        if table:
            for row in table.splitlines():
                if 'Total Count: ' in row: continue
                col_A = col_B = col_C = col_D = col_E = ''
                try:
                    col_A = row[0:11].replace(' ','').strip()
                    col_B = row[0:13].strip()
                    col_C = row[14:42].strip()
                    col_D = row[43:59].strip()
                    col_E = row[60:].strip()
                    if col_A == '' and col_B == '' and col_C == '' and col_D == '' and col_E == '': continue
                    writer.writerow({
                        "CUSIP": col_A,
                        "CUSIP text": col_B,
                        "Issuer Name": col_C,
                        "Issuer Description": col_D,
                        "Status": col_E
                    })
                except IndexError:
                    print('IndexError')
        page.close()

#FIN
