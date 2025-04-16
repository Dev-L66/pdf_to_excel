#  Amazon Invoice PDF Extractor(Text, Tables, Images)

This Python script automates the extraction of data from **Amazon invoice PDFs**, including invoice details and product table rows for a specific type of Amazon bill. 
The extracted data is saved into a structured **Excel file**.

---

##  Features

-  Extracts key invoice metadata (Order No., Invoice Date, Customer Info, etc.)
-  Extracts tabular product data (SKU, Qty, Amount, etc.)
-  Extracts barcodes
-  Supports batch processing of multiple PDFs
-  Saves all extracted info into a single **Excel spreadsheet**

---

##  Requirements

- Python 3.x
- Install the packages: pip install -r requirements.txt


Set the paths:
- pdf_folder = r"your/invoice/folder/path"
- excel_path = r"your/output/excel/file.xlsx"


To run the script:
- Clone the repo.
- Run the command: python pdf_to_excel.py
