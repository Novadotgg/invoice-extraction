import pdfplumber
import re
import csv
from word2number import w2n

pdf_files = ['invoice.pdf', 'invoice1.pdf', 'invoice2.pdf', 'invoice3.pdf']
csv_file = "invoice_data.csv"

fieldnames = [
    "Invoice Number", "Invoice Details", "Invoice Date", "Order Number", "Order Date",
    "PAN No", "GST Registration Number", "Sold By", "Billing Address", "Shipping Address",
    "Description", "Unit Price", "Qty", "Net Amount", "Tax Rate", "Tax Type", "Tax Amount",
    "Total Amount", "Amount in Words"
]

with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.DictWriter(file, fieldnames=fieldnames, quoting=csv.QUOTE_ALL)
    writer.writeheader()

    for loc in pdf_files:
        try:
            with pdfplumber.open(loc) as pdf:
                page = pdf.pages[0]
                text = page.extract_text()

            # Normalize text
            # text = text.replace('\r', '\n')
            print(text)

            # Extract fields
            invoice_number = re.search(r'Invoice Number\s*:\s*([A-Z0-9-]+)', text)
            invoice_details = re.search(r'Invoice Details\s*:\s*(.+)', text)
            invoice_date = re.search(r'Invoice Date\s*:\s*(\d{2}\.\d{2}\.\d{4})', text)
            order_number = re.search(r'Order Number\s*:\s*([0-9-]+)', text)
            order_date = re.search(r'Order Date\s*:\s*(\d{2}\.\d{2}\.\d{4})', text)
            pan_no = re.search(r'PAN No\s*:\s*([A-Z0-9]+)', text)
            gst_no = re.search(r'GST Registration No\s*:\s*([A-Z0-9]+)', text)

            sold_by = re.search(r'Sold By\s*:\s*(.*?)Billing Address', text, re.DOTALL)
            billing_address = re.search(r'Billing Address\s*:\s*(.*?)\n(?:Shipping Address|PAN No|GST Registration No)', text, re.DOTALL)
            shipping_address = re.search(r'Shipping Address\s*:\s*(.*?)\n(?:Dynamic QR Code|Place of supply)', text, re.DOTALL)
            total_amount= re.search(r'TOTAL:\s*₹[\d.,]+\s*₹([\d.,]+)', text)
            amount_words = re.search(r'Amount in Words\s*:\s*(.+?)\s+only', text, re.IGNORECASE)




            # Extract item block
            # Extract item fields individually
            description_match = re.search(
                r'\d+\s+(INOVERA.*?Laptop)',  # adjust this pattern to be more general if needed
                text, re.DOTALL
            )
            unit_price_match = re.search(r'Unit Price\s*₹([\d.,]+)', text)
            qty_match = re.search(r'\s+(\d+)\s+₹[\d.,]+\s+\d+%\s+[A-Z]+\s+₹[\d.,]+\s+₹[\d.,]+', text)
            # net_amount_match = re.search(r'Net Amount\s*₹([\d.,]+)', text)
            net_amount_match = re.search(r'₹([\d.,]+)\s+\d+\s+₹([\d.,]+)\s+\d+% IGST ₹[\d.,]+\s+₹[\d.,]+', text)
            if net_amount_match:
                net_amount = net_amount_match.group(2)

            tax_rate_match = re.search(r'(\d+%)\s+[A-Z]+\s+₹[\d.,]+', text)
            tax_type_match = re.search(r'\d+%\s+([A-Z]+)\s+₹[\d.,]+', text)
            tax_amount_match = re.search(r'\d+%\s+[A-Z]+\s+₹([\d.,]+)', text)
            total_amount_match = re.search(r'TOTAL:\s*₹[\d.,]+\s*₹([\d.,]+)', text)

            # Assign values
            description = description_match.group(1).strip().replace('\n', ' ') if description_match else "Not found"
            # unit_price = unit_price_match.group(1) if unit_price_match else "Not found"
            qty = qty_match.group(1) if qty_match else "Not found"
            net_amount = net_amount_match.group(1) if net_amount_match else "Not found"
            unit_price=net_amount_match.group(1) if net_amount_match else "Not found"
            tax_rate = tax_rate_match.group(1) if tax_rate_match else "Not found"
            tax_type = tax_type_match.group(1) if tax_type_match else "Not found"
            tax_amount = tax_amount_match.group(1) if tax_amount_match else "Not found"
            total_amount = total_amount_match.group(1) if total_amount_match else "Not found"

            # Convert total amount to float or fallback to amount in words
            try:
                total_amount_num = float(total_amount.replace(",", ""))
            except:
                try:
                    total_amount_num = w2n.word_to_num(amount_words.group(1)) if amount_words else "Not found"
                except:
                    total_amount_num = "Conversion failed"


            # Construct dictionary for CSV row
            data = {
                "Invoice Number": invoice_number.group(1) if invoice_number else "Not found",
                "Invoice Details": invoice_details.group(1).strip() if invoice_details else "Not found",
                "Invoice Date": invoice_date.group(1) if invoice_date else "Not found",
                "Order Number": order_number.group(1) if order_number else "Not found",
                "Order Date": order_date.group(1) if order_date else "Not found",
                "PAN No": pan_no.group(1) if pan_no else "Not found",
                "GST Registration Number": gst_no.group(1) if gst_no else "Not found",
                "Sold By": sold_by.group(1).strip().replace('\n', ', ') if sold_by else "Not found",
                "Billing Address": billing_address.group(1).strip().replace('\n', ', ') if billing_address else "Not found",
                "Shipping Address": shipping_address.group(1).strip().replace('\n', ', ') if shipping_address else "Not found",
                "Description": description,
                "Unit Price": unit_price,
                "Qty": qty,
                "Net Amount": net_amount,
                "Tax Rate": tax_rate,
                "Tax Type": tax_type,
                "Tax Amount": tax_amount,
                "Total Amount": total_amount_num,
                "Amount in Words": (amount_words.group(1).strip() + " only") if amount_words else "Not found"
            }

            print(f"Processed {loc}")
            writer.writerow(data)

        except Exception as e:
            print(f" Error processing {loc}: {e}")

print(f"\nAll data saved to {csv_file}")
