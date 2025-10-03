# Purpoose:
"""	Parse SEC 13flist PDF to spreadsheet format. """
# Usage:
#   Download the latest PDF from the official web site and save into the same folder as this.
#   Run this.
# Notes:
#   https://www.sec.gov/divisions/investment/13flists.htm -> https://www.sec.gov/files/investment/13flist2021q2.pdf
# History:
#   https://github.com/PhillRogers/13Flist-to-spreadsheet
#	2025-09-29    Phill
#	Fresh minimal version to convert only. Needs manual download. Does no notifications.
#	2025-10-03    Phill
#   Assume downloaded PDF is in same folder as this script. If not then launch browser to get one.
#   Add XLSX file output.  Tested with auto-py-to-exe  Check with pylint = 10.0
# License:
#   SEC-13Flist-ToSpreadsheet - Parse SEC 13flist PDF to spreadsheet format.
#   Copyright (C) 2025  Phill Rogers
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#   PRogers at Enhance.Group

import csv
import glob
import os.path
import sys
import webbrowser
import pdfplumber
import openpyxl

download_page = 'https://www.sec.gov/divisions/investment/13flists'

# __file__ does not work with auto-py-to-exe
if sys.argv[0]: prog_fpn = os.path.abspath(sys.argv[0])
else: prog_fpn = os.getcwd()+os.sep+'PyCLI'
obeydir,prog_fn = os.path.split(prog_fpn)
os.chdir(obeydir)

try:
    pdf_path = sorted(glob.glob('13flist????q?.pdf'), key=os.path.getmtime, reverse=True)[0]
    csv_path = os.path.splitext(pdf_path)[0] +'.csv'
    xlsx_path = os.path.splitext(pdf_path)[0] +'.xlsx'
except:
    print('Failure to find any files like "13flist????q?.pdf" in the current folder.')
    webbrowser.open(download_page)
    print(f'Please retry after downloading the PDF and saving it into this folder:\n\t"{obeydir}"')
    _ = input('Press RETURN to finish.')
    sys.exit(-1)

print(f'Reading from {pdf_path} ')
output_columns = ["CUSIP","CUSIP text","Issuer Name","Issuer Description","Status"]
grand_table = [output_columns]
with pdfplumber.open(pdf_path) as pdf, open(csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
    print(f'Parsing {len(pdf.pages)} pages ... ')
    writer = csv.DictWriter(csv_file, quoting=csv.QUOTE_ALL, fieldnames=output_columns)
    writer.writeheader()
    for page in pdf.pages[2:]:
        pn = page.page_number
        if (pn % 10) == 0: print(f'Page: {pn}')
        box = page.crop((63,130,536,714), strict=False, relative=False) # l,t,r,b pts
        page_table = box.extract_text(keep_blank_chars=True)
        if page_table:
            for row in page_table.splitlines():
                if 'Total Count: ' in row: continue
                col_A = col_B = col_C = col_D = col_E = ''
                try:
                    col_A = row[0:11].replace(' ','').strip()
                    col_B = row[0:13].strip()
                    col_C = row[14:42].strip()
                    col_D = row[43:59].strip()
                    col_E = row[60:].strip()
                    if col_A == '' and col_B == '' and col_C == '' and col_D == '' and col_E == '': continue
                    grand_table.append([col_A,col_B,col_C,col_D,col_E])
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

if len(grand_table) >1:
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        for row in grand_table:
            ws.append(row)
        wb.save(xlsx_path)
    except Exception as e:
        print(f'Failure to write XLSX file.\n\t{e}')

_ = input('Press RETURN to finish.')
#FIN
