import re
import pdfplumber
import pandas as pd
from datetime import datetime
from pyzbar.pyzbar import decode
from PIL import Image
import os
import glob

# Define the folder containing PDFs
pdf_folder = r"C:\Users\sasas\OneDrive\Desktop\New folder"

# Get all PDF file paths from the folder
pdf_paths = glob.glob(os.path.join(pdf_folder, "*.pdf"))

#Define your excel path
excel_path = r"C:\Users\sasas\OneDrive\Desktop\output.xlsx"

combined_data = []



#Format date
def reformat_date(date_str):
    try:
       
        date_obj = pd.to_datetime(date_str, format="%d.%m.%Y")

        # Convert the datetime object to an Excel date (numeric representation)
        excel_date = (date_obj - pd.Timestamp('1899-12-30')).days
        return excel_date
    except ValueError:
        return ""

def extract_invoice_details(text):
    # Extract invoice details using regex
    product_sku = re.search(r"F\/[A-Z]+\/[A-Z]+\/\d+(?:\.\d+)?\/[A-Z0-9]+\/\d+[A-Z]+", text)
    sub_order_no = re.search(r"Order Number[:\s]*([0-9\-]+)", text, re.IGNORECASE)
    invoice_date = re.search(r'Invoice Date\s*:\s*(\d{2}\.\d{2}\.\d{4})', text, re.IGNORECASE)
    payment_date = re.search(r"Payment Date[:\s]*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", text, re.IGNORECASE)
    order_date = re.search(r"Order Date[:\s]*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", text, re.IGNORECASE)
    customer_name = re.search(r"cottonode\s+([A-Za-z\s]+)", text, re.IGNORECASE)
    customer_address = re.search(r"Shipping Address\s*:\s*\n[^\n]+\n[^\n]+\n([^\n]+(?:\n[^\n]+)*)(?=\nIN\b)", text, re.IGNORECASE)
    cus_state = re.search(r"Place of delivery:\s*([A-Za-z\s]+)(?=\n|$)", text, re.IGNORECASE)
    cus_city = re.search(r"Shipping\s+Address\s*:\s*(?:[\s\S]+?)\n([A-Za-z\s]+)\s*,\s*[A-Za-z\s]+,\s*\d{6}(?=\s*IN)", text, re.IGNORECASE)
    cus_pincode = re.search(r"Shipping\s+Address\s*:\s*[\s\S]+?(\d{6})(?=\s*IN)", text, re.IGNORECASE)
    reseller_name = re.search(r"cottonode\s+([A-Za-z\s]+)", text, re.IGNORECASE)
    reseller_state = re.search(r"Place of delivery:\s*([A-Za-z\s]+)(?=\n|$)", text, re.IGNORECASE)
    reseller_city = re.search(r"Shipping\s+Address\s*:\s*(?:[\s\S]+?)\n([A-Za-z\s]+)\s*,\s*[A-Za-z\s]+,\s*\d{6}(?=\s*IN)", text, re.IGNORECASE)
    reseller_pincode = re.search(r"Shipping\s+Address\s*:\s*[\s\S]+?(\d{6})(?=\s*IN)", text, re.IGNORECASE)
    group_code = re.search(r"F\/PLZ\/RYN\/\d", text, re.IGNORECASE)
    color_code =re.search(r"F\/[A-Z]+\/[A-Z]+\/\d+(?:\.\d+)?\/([A-Z0-9]+)", text)
    # style_code = re.search(r"F\/[A-Z]+\/[A-Z]+\/\d+(?:\.\d+)?\/[A-Z0-9]+", text)
    # size = re.search(r"\/(\d+[A-Z]+)\s*\)", text, re.IGNORECASE)
    
    

    formatted_invoice_date = reformat_date(invoice_date.group(1)) if invoice_date else " "
    formatted_order_date = reformat_date(order_date.group(1)) if order_date else " "
    return {
        "Product SKU": product_sku.group(0) if product_sku else " ",
        "Sub Order No.": sub_order_no.group(1) if sub_order_no else " ",
        "Company": "Amazon",
        "Delivery Partner": "Easy Ship",
        "Invoice Date": formatted_invoice_date,
        "Order Status": "",
        "Payment Date": payment_date.group(1) if payment_date else " ",
        "Cost": "",
        "Delivery Charges": "",
        "Payout Amount": "",
        "Profit": "",
        "Payout Done?": "",
        "NOTE": "",
        "Order Date": formatted_order_date,
        "Pickup Date": "",
        "Return Type": "",
        "Return/Exchange Issue": "",
        "Return/Exchange/Rating Comment": "",
        "Rating": "",
        "Cashback": "",
        "Customer Name": customer_name.group(1) if customer_name else " ",
        "Customer Address": customer_address.group(1) if customer_address else " ",
        "Cus. State": cus_state.group(1) if cus_state else " ",
        "Cus. City": cus_city.group(1) if cus_city else " ",
        "Cus. pincode": cus_pincode.group(1) if cus_pincode else " ",
        "Reseller Name": reseller_name.group(1) if reseller_name else " ",
        "Reseller State": reseller_state.group(1) if reseller_state else " ",
        "Reseller City": reseller_city.group(1) if reseller_city else " ",
        "Reseller pincode": reseller_pincode.group(1) if reseller_pincode else " ",
        "Order Number": sub_order_no.group(1) if sub_order_no else " ",
        "Group Code": group_code.group(0) if group_code else " ",
        # "Style code": style_code.group(0) if style_code else " ",
        "style Code": f"{group_code.group(0) if group_code else " "}/{color_code.group(1) if color_code else " "}",
        "Color Code": color_code.group(1) if color_code else " ",
        # "Size": size.group(1) if size else " ",
        "size": ""
    }

#Extract SINo. and Qty from table
def extract_si_no_and_qty_from_table(pdf_path):
    extracted_table_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            print(f"Processing page {i + 1}...")
            
            # Extract text from the page
            text = page.extract_text()
            
            # Extract Order Number from the text
            order_number_match = re.search(r"Order Number[:\s]*([0-9\-]+)", text, re.IGNORECASE)
            order_number = order_number_match.group(1) if order_number_match else None
            
            # Extract tables from the page
            tables = page.extract_tables(table_settings={
                "vertical_strategy": "lines", 
                "horizontal_strategy": "lines", 
                "snap_tolerance": 3,
                "intersection_tolerance": 3
            })
            
            if tables:
                for table in tables:
                    # Convert the table into a DataFrame
                    df = pd.DataFrame(table[1:], columns=table[0])
                    
                    # Clean headers (remove newlines and extra spaces)
                    df.columns = df.columns.str.replace('\n', ' ').str.strip().str.replace(r'\s+', ' ', regex=True)
                    
                    # Print the cleaned DataFrame for debugging
                    print("Cleaned DataFrame:")
                    print(df)
                    
                    # Print available columns for debugging
                    print("Available columns:", df.columns.tolist())
                    
                    # Define possible column names for Sl. No and Qty
                    sl_no_columns = ["Sl.", "Sl. No", "No"]
                    qty_columns = ["Qty", "Quantity"]
                    amount_columns = ['Total Amount']
                    
                    # Find the Sl. No and Qty columns
                    sl_no_column = next((col for col in sl_no_columns if col in df.columns), None)
                    qty_column = next((col for col in qty_columns if col in df.columns), None)
                    amount_column = next((col for col in amount_columns if col in df.columns), None)
                    print(sl_no_column, amount_column)
                    
                    if sl_no_column and qty_column:
                        # Filter rows where Sl. No and Qty are numeric
                        df_filtered = df[
                            df[sl_no_column].astype(str).str.isnumeric() & 
                            df[qty_column].astype(str).str.isnumeric()
                        ]
                        
                        # Extract Sl. No and Qty values
                        for _, row in df_filtered.iterrows():
                            extracted_table_data.append({
                                "SI No": row[sl_no_column],
                                "Qty": row[qty_column],
                                "Order Number": order_number,
                                "Total Amount": row[amount_column] if amount_column else "0"
                                  # Add Order Number to table data
                            })
                    else:
                        print("Required columns (Sl. No and Qty) not found in this table.")
            else:
                print("No tables found on this page.")
    
    return extracted_table_data

def extract_text_from_pdf(pdf_path):
    extracted_invoice_data = []
    with pdfplumber.open(pdf_path) as pdf:
          for page in pdf.pages:
            text = page.extract_text()
            print(text)

            # Extract invoice details
            invoice_details = extract_invoice_details(text)
            if invoice_details["Order Number"] != " ":
                extracted_invoice_data.append(invoice_details)
    
    return extracted_invoice_data


for pdf_path in pdf_paths:
    print(f"Processing PDF: {pdf_path}")
    
    # Extract invoice details and table data
    invoice_data = extract_text_from_pdf(pdf_path)
    table_data = extract_si_no_and_qty_from_table(pdf_path)

    # Dictionary to track the last SI No used for each Order Number
    order_si_no_tracker = {}

    for invoice in invoice_data:
        order_number = invoice["Order Number"]
        # Find all table entries that match the current invoice's order number
        matching_table_entries = [table for table in table_data if table.get("Order Number") == order_number]

        if not matching_table_entries:
            combined_data.append({**invoice, "Qty": "", "Invoice Amount": "", "Sub Order No.": order_number + "_"})
        else:
            for table in matching_table_entries:
                if order_number in order_si_no_tracker:
                    order_si_no_tracker[order_number] += 1
                else:
                    order_si_no_tracker[order_number] = 1
                order_no_si_no = f"{order_number}_{order_si_no_tracker[order_number]}"
                combined_data.append({**invoice, "Qty": int(table["Qty"]), "Invoice Amount": float(table["Total Amount"].replace('₹', '').strip()), "Sub Order No.": order_no_si_no})

# Save combined data after processing all PDFs
if combined_data:
    df = pd.DataFrame(combined_data)
    # Reorder columns
    columns = df.columns.tolist()
    columns.remove("Sub Order No.")
    columns.insert(1, "Sub Order No.")
    
    df = df[columns]
    column_to_move = df.pop('Qty')
    df.insert(4, 'Qty', column_to_move)

    column_to_move = df.pop('Invoice Amount')
    df.insert(9, 'Invoice Amount', column_to_move)

    df.insert(22, 'Shipment Id/AWB', None)
    df.insert(32, 'Exchanged  AWB(In another sheet)', None)
    df.insert(33, 'Return Partner', None)
    df.insert(34, 'Return Id/AWB', None)
    
    
    df.to_excel(excel_path, index=False)

    print(f"Data extracted and saved to Excel successfully: {excel_path}")
else:
    print("No valid data extracted.")
