from pyzbar.pyzbar import decode
from PIL import Image
import pandas as pd
import numpy as np
import pdfplumber
import logging
import glob
import re
import io
import os
import cv2

logging.getLogger("pdfminer").setLevel(logging.ERROR)


# Define the folder containing PDFs
pdf_folder = r"C:\Users\sasas\OneDrive\Desktop\New folder"

# Get all PDF file paths from the folder
pdf_paths = glob.glob(os.path.join(pdf_folder, "*.pdf"))

#Define your excel path
excel_path = r"C:\Users\sasas\OneDrive\Desktop\output.xlsx"

ordered_list = []


# Format date to excel date
def reformat_date(date_str):
    try:      
       date_obj =  pd.to_datetime(date_str, format="%d.%m.%Y")
       return date_obj.date()
    except ValueError:
        return ""

# Extract invoice details
def extract_invoice_details(text):
 try:
    order_no = re.search(r"Order Number" + r"[:\s]*([0-9\-]+)", text, re.IGNORECASE)
    invoice_date = re.search(r'Invoice Date\s*:\s*(\d{2}\.\d{2}\.\d{4})', text, re.IGNORECASE)
    order_date = re.search(r"Order Date[:\s]*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", text, re.IGNORECASE)
    shipmentId_awb = re.search(r'(?:AWB|SWB)[#\s:]*(\d{9,20})', text, re.IGNORECASE)
    customer_name = re.search(r"Shipping Address\s*:\s*([^\n]+)", text, re.IGNORECASE)
    customer_address = re.search(r"Shipping Address\s*:\s*(?:.*\n){2}((?:.*\n)*?[A-Z ]+,\s*[A-Z ]+,\s*\d{6})", text, re.IGNORECASE)
    customer_state = re.search(r"Place of delivery:\s*([A-Za-z\s]+)(?=\n|$)", text, re.IGNORECASE)
    customer_city = re.search(r"Billing Address\s*:.*?\n(?:.*\n){2,}([A-Z ]+),\s*[A-Z ]+,\s*\d{6}", text, re.IGNORECASE)
    customer_pincode = re.search(r"Shipping\s+Address\s*:\s*[\s\S]+?(\d{6})(?=\s*IN)", text, re.IGNORECASE)
    group_code = re.search(r"(?:F)?\/[A-Z]+\/[A-Z]+\/\d+", text, re.IGNORECASE)
    color_code = re.search(r"(?:F)?\/[A-Z]+\/[A-Z]+\/\d+(?:\.\d+)?\/([A-Z0-9]+)", text, re.IGNORECASE)

    #Format date to excel date
    formatted_invoice_date = reformat_date(invoice_date.group(1)) if invoice_date else " "
    formatted_order_date = reformat_date(order_date.group(1)) if order_date else " "
   

    return{
        "Order Number": order_no.group(1) if order_no else "",
        "Company": "Amazon",
        "Delivery Partner": "Easy Ship",
        "Invoice Date": formatted_invoice_date,
        "Order Status": " ",
        "Payment": " ",
        "Cost": " ",
        "Delivery Charges": " ",
        "Payout Amount": " ",
        "Profit": " ",
        "Payout Done?": " ",
        "NOTE":" ",
        "Order Date": formatted_order_date,
        "Pickup Date": " ",
        "Return Type": " ",
        "Return/Exchange Issue": " ",
        "Return/Exchange/Rating Comment": " ",
        "Rating": " ",
        "Cashback": " ",
        "ShipmentId/AWB": shipmentId_awb.group(1) if shipmentId_awb else " ",
        "Customer Name": customer_name.group(1) if customer_name else " ",
        "Customer Address": customer_address.group(1) if customer_address else " ",
        "Customer State": customer_state.group(1) if customer_state else " ",
        "Customer City": customer_city.group(1) if customer_city else " ",
        "Customer Pincode": customer_pincode.group(1) if customer_pincode else " ",
        "Reseller Name": customer_name.group(1) if customer_name else " ",
        "Reseller Address": customer_address.group(1) if customer_address else " ",
        "Reseller State": customer_state.group(1) if customer_state else " ",
        "Reseller City": customer_city.group(1) if customer_city else " ",
        "Reseller Pincode": customer_pincode.group(1) if customer_pincode else " ",
        "Exchanged AW(In another sheet)": " ",
        "Return Partner": " ",
        "Return Id/AWB": " ",
        "Group Code": group_code.group(0) if group_code else " ",
        "Style Code": f"{group_code.group(0) if group_code else " "}/{color_code.group(1) if color_code else " "}",
        "Color Code": color_code.group(1) if color_code else " ",
        "Size": " ",
        "Contact": " "
        
    }
 except Exception as e:
    print(f"Some error occurred in the regex {e}")

def extract_pdf(pdf_paths):
     
      for pdf_path in pdf_paths:
       try: 
        with pdfplumber.open(pdf_path) as pdf:
         for page in pdf.pages:
            # Extract text
            print("Extracting text....")  
            text = page.extract_text() 
            left = page.crop((0, 0, 0.5 * page.width, 0.9 * page.height))
            right = page.crop((0.5 * page.width, 0, page.width, page.height))
            r_text = right.extract_text()
            # convert barcodes
            print("Processing images....")  
            pil_image = page.to_image(resolution = 500).original
            open_cv_image = cv2.cvtColor(np.array(pil_image),cv2.COLOR_RGB2BGR)
                    
            if not text:
                # Extract barcodes
                print("Reading barcodes...")
                barcodes = decode(open_cv_image)
                for barcode in barcodes:
                    barcode_data = barcode.data.decode('utf-8')
                
            else:
                details = extract_invoice_details(text)
                r_details = extract_invoice_details(r_text) 
                
            

            # Extract tables
            print("Extracing tables....")
            print("Processing...")
            tables = page.extract_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy":"lines", "snap_tolerance": 5, "intersection_tolerance": 5})
            
            for table in tables:
                
                df = pd.DataFrame(table[1:], columns=table[0])
                pd.set_option('display.max_colwidth', None)
                pd.set_option('display.expand_frame_repr', False)  

                
                if "Sl.\nNo" in df.columns and "Description" in df.columns and "Qty" in df.columns and details:
                    df["Sl.\nNo"]= df["Sl.\nNo"].astype(str)
                    valid_rows= df[df["Sl.\nNo"].str.isnumeric()]
                    
                    for _, row in valid_rows.iterrows():
                        description = row["Description"]
                        sku_match = re.search(r"\(\s*([A-Z0-9/.\-]+)\s*\)", description)
                        product_sku = sku_match.group(1) if sku_match else " "
                        slno = row["Sl.\nNo"]
                        qty = float(row["Qty"])
                        total_amount = float(row["Total\nAmount"].replace('₹', '').strip())
                        ordered_list.append((product_sku, f"{details["Order Number"]}_{slno}",f"{details["Company"]}", details["Delivery Partner"], qty, details["Invoice Date"], details["Order Status"], details["Payment"],details["Cost"], total_amount, details["Delivery Charges"],  details["Payout Amount"],details["Profit"],details["Payout Done?"],details["NOTE"],details["Order Date"],details["Pickup Date"],details["Return Type"],details["Return/Exchange Issue"], details["Return/Exchange/Rating Comment"],details["Rating"],details["Cashback"], barcode_data, r_details["Customer Name"], r_details["Customer Address"], details["Customer State"], details["Customer City"], details["Customer Pincode"], r_details["Reseller Name"], r_details["Reseller Address"], details["Reseller State"], details["Reseller City"], details["Reseller Pincode"],"","","", details["Order Number"],details["Group Code"], details["Style Code"],details["Color Code"],details["Size"],details["Contact"]))
       except FileNotFoundError:
        print(f"File not found: {pdf_path}") 
       except Exception as e:
        print(f"Error processing file {os.path.basename(pdf_path)}:{e}")          
            
extract_pdf(pdf_paths) 
#ordered_list  
# for item in ordered_list:
    # print(item)

try:
 df = pd.DataFrame(ordered_list, columns=["Product SKU","Sub Order No.", "Company", "Delivery Partner", "Qty", "Invoice Date", "Order Status", "Payment", "Cost", "Total Amount","Delivery Charges",  "Payout Amount","Profit","Payout Done?","NOTE","Order Date","Pickup Date","Return Type","Return/Exchange Issue", "Return/Exchange/Rating Comment","Rating","Cashback","ShipmentId/AWB", "Customer Name", "Customer Address", "Customer State", "Customer City", "Customer Pincode","Reseller Name", "Reseller Address", "Reseller State", "Reseller City", "Reseller Pincode","Exchanged  AWB(In another sheet)","Return Partner", "Return Id/AWB", "Order No.", "Group Code", "Style Code","Color Code","Size","Contact"]) 
 if not df.empty:
    try: 
     df.to_excel(excel_path,index = False, engine='openpyxl' )
     print(f"Excel file saved successfully at: {excel_path}")
    except PermissionError as e:
     print(f"Please, close the Excel file: {e}")
    except FileNotFoundError as e:
     print(f"File path not found.{e}") 
 else:
    print(f"No data Found. {excel_path}")

except Exception as e:
    print(f"Unexpected error while exporting to Excel: {e}")
