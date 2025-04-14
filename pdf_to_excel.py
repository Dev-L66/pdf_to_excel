import logging
import pdfplumber
import re
import pandas as pd
from PIL import Image
import pytesseract
import io

logging.getLogger("pdfminer").setLevel(logging.ERROR)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

pdf_paths = [r"C:\Users\sasas\OneDrive\Desktop\New folder\103328959845020116.pdf",r"C:\Users\sasas\OneDrive\Desktop\New folder\102505269325020105.pdf",r"C:\Users\sasas\OneDrive\Desktop\New folder\102798760075020109.pdf"]
excel_path = r"C:\Users\sasas\OneDrive\Desktop\output.xlsx"

order = []

def reformat_date(date_str):
    try:      
       date_obj =  pd.to_datetime(date_str, format="%d.%m.%Y")
       return date_obj.date()
    except ValueError:
        return ""




def extract_invoice_details(text):
    order_no = re.search(r"Order Number" + r"[:\s]*([0-9\-]+)", text, re.IGNORECASE)
    invoice_date = re.search(r'Invoice Date\s*:\s*(\d{2}\.\d{2}\.\d{4})', text, re.IGNORECASE)
    order_date = re.search(r"Order Date[:\s]*([0-9]{2}\.[0-9]{2}\.[0-9]{4})", text, re.IGNORECASE)
    shipmentId_awb = re.search(r'AWB[#:\s]*([0-9]{9,20})', text, re.IGNORECASE)
    customer_name = re.search(r'Ship To:.*\n\s*([a-zA-Z][^\n]+)', text, re.IGNORECASE)
    customer_address = re.search(r'Ship To:.*?\n+(?:.*\n+)?(?:.*\n+)?([^\n]*?(?:\n(?!\s*Invoice|\s*Total|\s*Amount|\s*AWB|\s*Order ID|\s*Shipped|\s*Landmark|\s*NDL).+)+)', text, re.IGNORECASE)
    customer_state = re.search(r"Place of delivery:\s*([A-Za-z\s]+)(?=\n|$)", text, re.IGNORECASE)
    customer_city = re.search(r"Shipping\s+Address\s*:.*?\n(?:.*\n)*?([A-Za-z]+),\s*[A-Z\s]+,\s*\d{6}\s*\nIN", text, re.IGNORECASE)



    






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
        "Order Date": formatted_order_date,
        "Pickup Date": " ",
        "Return Type": " ",
        "Return/Exchange Issue": " ",
        "Return/Exchange/Rating Comment": " ",
        "Rating": " ",
        "Cashback": " ",
        "ShipmentId/AWB": shipmentId_awb.group(1) if shipmentId_awb else " ",
        "Customer Name": customer_name.group(1) if customer_name else " ",
        "Customer Address": customer_address.group(1)  if customer_address else " ",
        "Customer State": customer_state.group(1) if customer_state else " ",
        "Customer City": customer_city.group(1) if customer_city else " "

        
        
    }

for pdf_path in pdf_paths:
 with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text()          
        if text:
            details = extract_invoice_details(text)
        else:
            page_image = page.to_image(resolution = 300)
            pil_image = page_image.original
            text = pytesseract.image_to_string(pil_image)
            # print(text)
            detail = extract_invoice_details(text)
        

        
        tables = page.extract_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy":"lines", "snap_tolerance": 5, "intersection_tolerance": 5})
        
        for table in tables:
             
             df = pd.DataFrame(table[1:], columns=table[0])
             pd.set_option('display.max_colwidth', None)
             pd.set_option('display.expand_frame_repr', False)  


            #  print(df)
            
             if "Sl.\nNo" in df.columns and "Description" in df.columns and "Qty" in df.columns and details:
                df["Sl.\nNo"]= df["Sl.\nNo"].astype(str)
                valid_rows= df[df["Sl.\nNo"].str.isnumeric()]
                
                for _, row in valid_rows.iterrows():
                    description = row["Description"]
                    sku_match = re.search(r"\(\s*([A-Z0-9/.\-]+)\s*\)", description)
                    product_sku = sku_match.group(1) if sku_match else " "
                    slno = row["Sl.\nNo"]
                    qty = row["Qty"]
                    total_amount = row["Total\nAmount"]
                    order.append((product_sku, f"{details["Order Number"]}_{slno}",f"{details["Company"]}", details["Delivery Partner"], qty, details["Invoice Date"], details["Order Status"], details["Payment"],details["Cost"], total_amount, details["Delivery Charges"],  details["Payout Amount"],details["Profit"],details["Payout Done?"],details["Order Date"],details["Pickup Date"],details["Return Type"],details["Return/Exchange Issue"], details["Return/Exchange/Rating Comment"],details["Rating"],details["Cashback"], detail["ShipmentId/AWB"], detail["Customer Name"], detail["Customer Address"], details["Customer State"], details["Customer City"]))
                
            
             
                 
              
for item in order:
    print(item)

df = pd.DataFrame(order, columns=["Product SKU","Sub Order No.", "Company", "Delivery Partner", "Qty", "Invoice Date", "Order Status", "Payment", "Cost", "Total Amount","Delivery Charges",  "Payout Amount","Profit","Payout Done?","Order Date","Pickup Date","Return Type","Return/Exchange Issue", "Return/Exchange/Rating Comment","Rating","Cashback","ShipmentId/AWB", "Customer Name", "Customer Address", "Customer State", "Customer City"])
df.to_excel(excel_path,index = False )
