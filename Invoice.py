import csv
import sys
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
import sys
import os
import glob

# -------------------------
# Config
# -------------------------

if len(sys.argv) > 1:
    csv_files = [sys.argv[1]]
else:
    csv_files = glob.glob("*.csv")

if not csv_files:
    print("No CSV files found.")
    sys.exit(1)

for CSV_FILE in csv_files:
    print(f"Processing {CSV_FILE}...")

    NZ_TAX_CODE = "S15"
    NZ_JOB = "Online NZ"

    INT_TAX_CODE = "N-T"
    INT_JOB = "Online International"

    GST_RATE = 0.15

    # CSV_FILE = sys.argv[1] if len(sys.argv) > 1 else "shopifyexport.csv"

    CUSTOMER_NAME = "Kolorex.com(Online sales)"
    ADDRESS1 = "Kolorex.com(Online sales)"
    ADDRESS2 = "142 Collingwood Street"
    ADDRESS3 = "Nelson 7010"
    ADDRESS4 = "New Zealand"

    customerpo = ""

    # -------------------------
    # Totals
    # -------------------------
    shipping_nz = 0
    shipping_int = 0
    fees_nz = 0
    shipping_nz_gst = 0
    fees_int = 0

    rows = []
    payout_date_str = None
    my_list = []

    # -------------------------
    # MYOB header structure
    # -------------------------
    fieldnames = [
        "Co./Last Name","First Name","Addr 1 - Line 1","Addr 1 - Line 2","Addr 1 - Line 3","Addr 1 - Line 4",
        "Inclusive","Invoice No.","Date","Customer PO","Ship Via","Delivery Status",
        "Item Number","Quantity","Description","Price","Discount","Total","Job","Comment",
        "Journal Memo","Salesperson Last Name","Salesperson First Name","Shipping Date","Referral Source",
        "GST Code","GST Amount","Freight Amount","Freight GST Code","Freight GST Amount",
        "Sale Status","Currency Code","Exchange Rate",
        "Terms - Payment is Due"," - Discount Days"," - Balance Due Days",
        " - % Discount"," - % Monthly Charge",
        "Amount Paid","Payment Method","Payment Notes","Name on Card","Card Number",
        "Authorisation Code","Cheque Number","Category","Location ID","Card ID","Record ID"
    ]

    # -------------------------
    # Read Shopify CSV (pass 1)
    # -------------------------
    with open(CSV_FILE, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)

        for row in reader:
            product = row["SKU"].strip()
            # chargeback = row["Dispute"].strip()
            
            
            if product.lower().startswith("total") or not product:
                continue
            # if chargeback.lower().startswith("chargeback") or not chargeback:
                # continue
            my_list.append(list(row.values())[2])

            cleaned_numbers = []
            for item in my_list:
                if not item:
                    continue

                value = item.replace("#", "").strip()

                if value == "":
                    continue

                if value.isdigit():
                    cleaned_numbers.append(int(value))

            if cleaned_numbers:
                lowest = min(cleaned_numbers)
                highest = max(cleaned_numbers)
            else:
                lowest = highest = None

            customerpo = str(lowest) + "-" + str(highest)

            # Capture NZ shipping total
            taxes = float(row["Total taxes"].replace("$","").replace(",",""))
            shipping = float(row["Total shipping"].replace("$","").replace(",",""))

            if taxes > 0:
                # Calculate GST for this shipping line
                shippinggst = round(shipping * GST_RATE, 2)
                
                # Add GST to shipping for NZ
                shipping_incl_gst = round(shipping + shippinggst, 2)
                
                # Accumulate NZ shipping total
                shipping_nz += shipping
                
                # Accumulate NZ shipping GST separately
                shipping_nz_gst += shippinggst
            else:
                shipping_int += shipping

            # Round totals at the end
            shipping_nz = round(shipping_nz, 2)
            shipping_nz_gst = round(shipping_nz_gst, 2)

    # -------------------------
    # Read Shopify CSV (pass 2)
    # -------------------------
    with open(CSV_FILE, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)

        for row in reader:

            # Capture payout date
            if payout_date_str is None:
                raw_date = list(row.values())[0]

                try:
                    payout_date = datetime.strptime(raw_date, "%Y-%m-%d")
                except:
                    payout_date = datetime.strptime(raw_date, "%d/%m/%Y")

                payout_date_str = payout_date.strftime("%Y-%m-%d")

            product = row["SKU"].strip()
            # chargeback = row["Dispute"].strip()
            
            
            if product.lower().startswith("total") or not product:
                continue
            # if chargeback.lower().startswith("chargeback") or not chargeback:
                # continue
                
            if product == "GCCB30TR":
                product = "GCCBTR30"

            qty = int(row["Net quantity"])

            net_sales = float(row["Total net sales"].replace("$","").replace(",",""))
            taxes = float(row["Total taxes"].replace("$","").replace(",",""))
            fees = float(row["SUM Allocated payout fee"].replace("$","").replace(",",""))

            # Determine region
            if taxes > 0:
                job = NZ_JOB
                gst_code = NZ_TAX_CODE

                gst_amount = round(net_sales * GST_RATE, 2)
                total = round(net_sales, 2)

                #print(gst_amount)
                #print(total)
                #print(net_sales)

                fees_nz += fees

            else:
                job = INT_JOB
                gst_code = INT_TAX_CODE
                gst_amount = 0
                total = round(net_sales, 2)

                fees_int += fees

            if qty <= 0 or product.startswith("[Shipping]"):
                continue

            # FIX: Use Decimal for precise division, then re-derive total from price * qty
            # so that price * qty always equals the displayed total (no 1-cent variance).
            price = float(
                (Decimal(str(total)) / Decimal(str(qty))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            )
            total = round(price * qty, 2)  # re-derive total so it's always consistent with price

            rows.append({
                "Co./Last Name": CUSTOMER_NAME,
                "First Name": "",
                "Addr 1 - Line 1": ADDRESS1,
                "Addr 1 - Line 2": ADDRESS2,
                "Addr 1 - Line 3": ADDRESS3,
                "Addr 1 - Line 4": ADDRESS4,
                "Inclusive": "X",
                "Date": payout_date_str,
                "Customer PO": customerpo,
                "Ship Via": "NZ Post",
                "Delivery Status": "P",
                "Item Number": product,
                "Quantity": qty,
                "Description": "",
                "Price": f"${price:.2f}",
                "Total": f"${total:.2f}",
                "Job": job,
                "Journal Memo": f"Sale; {CUSTOMER_NAME}",
                "Salesperson Last Name": "Osborne",
                "Salesperson First Name": "Suzanne",
                "GST Code": gst_code,
                "GST Amount": f"${gst_amount:.2f}",
                "Freight Amount": shipping_nz,
                "Freight GST Code": "S15",
                "Freight GST Amount": shipping_nz_gst,
                "Sale Status": "I",
                "Terms - Payment is Due": "5",
                " - Discount Days": "1",
                " - Balance Due Days": "20",
                " - % Discount": "0",
                " - % Monthly Charge": "0",
                "Amount Paid": "$0.00",
                "Category": "New Zealand",
                "Location ID": "FHR-ONL"
            })

    # -------------------------
    # Cost line functions
    # -------------------------
    def add_cost_line(name, amount, job, gst_applicable):

        if amount <= 0:
            return

        qty = -1

        if gst_applicable:
            gst_code = NZ_TAX_CODE
            gst_amount = amount * 3 / 23
            net_amount = amount - gst_amount
            #print(net_amount)# strip GST
            # gst_amount = net_amount * GST_RATE
        else:
            gst_code = INT_TAX_CODE
            net_amount = amount
            gst_amount = 0

        price = net_amount   # unit price for MYOB import (negative if needed)
        total = -net_amount   # total excluding GST
        gst_amount = -gst_amount
        
        rows.append({
            "Co./Last Name": CUSTOMER_NAME,
            "Addr 1 - Line 1": ADDRESS1,
            "Addr 1 - Line 2": ADDRESS2,
            "Addr 1 - Line 3": ADDRESS3,
            "Addr 1 - Line 4": ADDRESS4,
            "Inclusive": "X",
            "Date": payout_date_str,
            "Customer PO": customerpo,
            "Ship Via": "NZ Post",
            "Delivery Status": "P",
            "Item Number": name,
            "Quantity": qty,
            "Description": name,
            "Price": f"${price:.2f}",
            "Total": f"${total:.2f}",
            "Job": job,
            "Journal Memo": f"Sale; {CUSTOMER_NAME}",
            "Salesperson Last Name": "Osborne",
            "Salesperson First Name": "Suzanne",
            "GST Code": gst_code,
            "GST Amount": f"${gst_amount:.2f}",
            "Freight Amount": shipping_nz,
            "Freight GST Code": "S15",
            "Freight GST Amount": shipping_nz_gst,
            "Sale Status": "I",
            "Terms - Payment is Due": "5",
            " - Discount Days": "1",
            " - Balance Due Days": "20",
            " - % Discount": "0",
            " - % Monthly Charge": "0",
            "Amount Paid": "$0.00",
            "Category": "New Zealand",
            "Location ID": "FHR-ONL"
        })


    def add_cost_line_shipping(name, desc, amount, job, gst_applicable):

        if amount <= 0:
            return

        qty = 1
        price = amount

        if gst_applicable:
            gst_code = NZ_TAX_CODE
            gst_amount = (amount * GST_RATE)
        else:
            gst_code = INT_TAX_CODE
            gst_amount = 0

        rows.append({
            "Co./Last Name": CUSTOMER_NAME,
            "Addr 1 - Line 1": ADDRESS1,
            "Addr 1 - Line 2": ADDRESS2,
            "Addr 1 - Line 3": ADDRESS3,
            "Addr 1 - Line 4": ADDRESS4,
            "Inclusive": "X",
            "Date": payout_date_str,
            "Customer PO": customerpo,
            "Ship Via": "NZ Post",
            "Delivery Status": "P",
            "Item Number": name,
            "Quantity": qty,
            "Description": desc,
            "Price": f"${price:.2f}",
            "Total": f"${price:.2f}",
            "Job": job,
            "Journal Memo": f"Sale; {CUSTOMER_NAME}",
            "Salesperson Last Name": "Osborne",
            "Salesperson First Name": "Suzanne",
            "GST Code": gst_code,
            "GST Amount": f"${gst_amount:.2f}",
            "Freight Amount": shipping_nz,
            "Freight GST Code": "S15",
            "Freight GST Amount": shipping_nz_gst,
            "Sale Status": "I",
            "Terms - Payment is Due": "5",
            " - Discount Days": "1",
            " - Balance Due Days": "20",
            " - % Discount": "0",
            " - % Monthly Charge": "0",
            "Amount Paid": "$0.00",
            "Category": "New Zealand",
            "Location ID": "FHR-ONL"
        })

    # -------------------------
    # Add lines
    # -------------------------
    # add_cost_line_shipping("Freight","Freight GST", shipping_nz, NZ_JOB, True)
    add_cost_line_shipping("Freight","Freight GST E", shipping_int, INT_JOB, False)

    add_cost_line("SF", fees_nz, NZ_JOB, True)
    add_cost_line("SF", fees_int, INT_JOB, False)

    # -------------------------
    # Write CSV
    # -------------------------
    OUTPUT_CSV = f"myob_import_{payout_date_str}.txt"

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:

        raw_writer = csv.writer(f)
        raw_writer.writerow(["{}"])

        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for r in rows:
            writer.writerow(r)

#print("MYOB import file created:", OUTPUT_CSV)