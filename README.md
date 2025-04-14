#  Amazon Invoice PDF Extractor with Table Parsing and OCR

This Python script automates the extraction of data from **Amazon invoice PDFs**, including invoice details and product table rows for a specific type of Amazon bill. 
The extracted data is saved into a structured **Excel file**.

---

##  Features

-  Extracts key invoice metadata (Order No., Invoice Date, Customer Info, etc.)
-  Extracts tabular product data (SKU, Qty, Amount, etc.)
-  Supports batch processing of multiple PDFs
-  Saves all extracted info into a single **Excel spreadsheet**

---

##  Requirements

- Python 3.x
- Install the packages: pip install -r requirements.txt
- Tesseract OCR must be installed separately:
  [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)
  Add the Tesseract executable path to your environment variables.

Set:
pdf_folder = r"your/invoice/folder/path"
excel_path = r"your/output/excel/file.xlsx"
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

To run the script:
- Clone the repo.
- python pdf_to_excel.py
