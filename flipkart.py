import pdfplumber
import re
import csv
from num2words import num2words

pdf_files = ['flip1.pdf']

fieldnames = [
    "Invoice Number", "Invoice Details", "Invoice Date", "Order Number", "Order Date",
    "PAN No", "GST Registration Number", "Sold By", "Billing Address", "Shipping Address",
    "Description", "Unit Price", "Qty", "Net Amount", "Tax Rate", "Tax Type", "Tax Amount",
    "Total Amount", "Amount in Words"
]

# Open CSV file to write multiple rows
with open("flip.csv", "w", newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()

    for loc in pdf_files:
        with pdfplumber.open(loc) as pdf:
            page = pdf.pages[0]
            flipkart_text = page.extract_text()

        data = {key: "Not found" for key in fieldnames}
        # print(flipkart_text)

        # Extract fields
        invoice_number = re.search(r'Invoice Number\s*#\s*([A-Z0-9]+)', flipkart_text)
        invoice_date = re.search(r'Invoice Date\s*:\s*(\d{2}-\d{2}-\d{4})', flipkart_text)
        order_number = re.search(r'Order ID:.*?\n([A-Z0-9]+)', flipkart_text)
        order_date = re.search(r'Order Date:\s*(\d{2}-\d{2}-\d{4})', flipkart_text)
        pan_no = re.search(r'PAN:\s*([A-Z0-9]+)', flipkart_text)
        gst_no = re.search(r'GSTIN\s*-\s*([A-Z0-9]+)', flipkart_text)
        sold_by = re.search(r'Sold By:\s*(.*?),', flipkart_text)
        billing_address = re.search(r'Bill To\s*(.*?)Order Date:', flipkart_text, re.DOTALL)
        shipping_address = re.search(r'Ship To\s*(.*?)Invoice Date:', flipkart_text, re.DOTALL)

        # Extract item details
        item = re.search(
            r'Product Title\s+(.*?)\s+(\d+)\s+([\d.,]+)\s+[-]?[\d.,]+\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)',
            flipkart_text, re.DOTALL
        )
        grand_total = re.search(r'Grand Total\s*₹\s*([\d.,]+)', flipkart_text)

        # Assign extracted values to data dictionary
        data["Invoice Number"] = invoice_number.group(1) if invoice_number else "Not found"
        data["Invoice Date"] = invoice_date.group(1) if invoice_date else "Not found"
        data["Order Number"] = order_number.group(1) if order_number else "Not found"
        data["Order Date"] = order_date.group(1) if order_date else "Not found"
        data["PAN No"] = pan_no.group(1) if pan_no else "Not found"
        data["GST Registration Number"] = gst_no.group(1) if gst_no else "Not found"
        data["Sold By"] = sold_by.group(1).strip() if sold_by else "Not found"
        data["Billing Address"] = billing_address.group(1).strip().replace("\n", ", ") if billing_address else "Not found"
        data["Shipping Address"] = shipping_address.group(1).strip().replace("\n", ", ") if shipping_address else "Not found"

        if item:
            data["Description"] = item.group(1).strip().replace("\n", " ")
            data["Qty"] = item.group(2)
            data["Unit Price"] = item.group(3)
            data["Net Amount"] = item.group(4)
            data["Tax Amount"] = item.group(5)
            data["Total Amount"] = grand_total.group(1) if grand_total else item.group(6)
            data["Tax Rate"] = "18%"  # Flipkart fixed rate
            data["Tax Type"] = "IGST"
        else:
            data["Description"] = data["Qty"] = data["Unit Price"] = data["Net Amount"] = "Not found"

        # data["Amount in Words"] = "Not available in Flipkart invoice"
        amount_str = data["Total Amount"].replace(",", "")
        amount_number = float(amount_str)
        data["Amount in Words"] = num2words(amount_number, lang='en_IN').replace(" and", "").capitalize() + " only"
        data["Invoice Details"] = "Flipkart Tax Invoice"

        writer.writerow(data)

print(" All Flipkart invoice data saved to flip.csv")
