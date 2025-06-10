import pdfplumber
import re
from word2number import w2n
import csv
# Invoice number
# invoice details
# invoice date
# order number
# order date
# pan no
# gst registration number
# total amount
# amount in words
# Plcae of supply
# place of delivery


loc='invoice.pdf'

with pdfplumber.open(loc) as pdf:
    page=pdf.pages[0]
    text = page.extract_text()
    # for line in text.split('\n'):
    #     print(line)
    # print(text)


invoice_number = re.search(r'Invoice Number :\s*([A-Z0-9-]+)', text) 
print(f"Invoice Number: {invoice_number.group(1) if invoice_number else 'Not found'}")

invoice_details = re.search(r'Invoice Details :\s*(.*?)\s+Invoice Date', text)
print(f"Invoice Details: {invoice_details.group(1) if invoice_details else 'Not found'}")

invoice_date = re.search(r'Invoice Date :\s*(\b\d{2}\.\d{2}\.\d{4}\b)', text)
print(f"Invoice Date: {invoice_date.group(1) if invoice_date else 'Not found'}")

order_number = re.search(r'Order Number:\s*([A-Z0-9-]+)', text)
print(f"Order Number: {order_number.group(1) if order_number else 'Not found'}")

order_date = re.search(r'Order Date:\s*(\b\d{2}\.\d{2}\.\d{4}\b)', text)
print(f"Order Date: {order_date.group(1) if order_date else 'Not found'}")

PAN_No= re.search(r'PAN No:\s*([A-Z0-9-]+)', text)
print(f"PAN No: {PAN_No.group(1) if PAN_No else 'Not found'}")

gst_registration_number = re.search(r'GST Registration No:\s*([A-Z0-9-]+)', text)
print(f"GST Registration Number: {gst_registration_number.group(1) if gst_registration_number else 'Not found'}")
#--------------


amount_in_words = re.search(r'Amount in Words:\s*(.+?)\s+only', text, re.IGNORECASE)
print(f"Amount in Words: {amount_in_words.group(1) if amount_in_words else 'Not found'}")




total_amount = re.search(r'Total Amount:\s*([A-Z0-9\s,.-]+)', text)
if total_amount == None:
    amount_words = amount_in_words.group(1)
    amount_number = w2n.word_to_num(amount_words)
print(f"Total Amount: {amount_number}")


# #--------------------
# place_of_supply = re.search(r'Place of Supply:\s*([A-Za-z\s,.-]+)', text)       #x
# print(f"Place of Supply: {place_of_supply.group(1) if place_of_supply else 'Not found'}")

# place_of_delivery = re.search(r'Place of Delivery:\s*([A-Za-z\s,.-]+)', text)       #x
# print(f"Place of Delivery: {place_of_delivery.group(1) if place_of_delivery else 'Not found'}")

# #----------------------

data = {
    "Invoice Number": invoice_number.group(1) if invoice_number else "Not found",
    "Invoice Details": invoice_details.group(1) if invoice_details else "Not found",
    "Invoice Date": invoice_date.group(1) if invoice_date else "Not found",
    "Order Number": order_number.group(1) if order_number else "Not found",
    "Order Date": order_date.group(1) if order_date else "Not found",
    "PAN No": PAN_No.group(1) if PAN_No else "Not found",
    "GST Registration Number": gst_registration_number.group(1) if gst_registration_number else "Not found",
    "Amount in Words": amount_in_words.group(1) if amount_in_words else "Not found",
    "Total Amount": amount_number
}
print(data)
csv_file = "invoice_data.csv"
with open(csv_file, mode='w', newline='') as file:
    writer = csv.DictWriter(file, fieldnames=data.keys())
    writer.writeheader()
    writer.writerow(data)

print(f"Data saved to {csv_file}")