import pandas as pd
from faker import Faker
import random
import datetime

fake = Faker("en_US")
Faker.seed(606)
random.seed(606)

OUTPUT_FILE = "telecom_data.xlsx"

df_invoices = pd.read_excel(OUTPUT_FILE, sheet_name="Invoices")
df_invoices = df_invoices.drop_duplicates(subset="Invoice_ID")

payment_methods = ["Cash", "Credit Card", "Mobile Wallet", "Bank Transfer"]

def to_date(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return datetime.datetime.strptime(str(value), "%Y-%m-%d").date()

payments = []
payment_id = 1

for inv in df_invoices.itertuples():
    status = inv.Invoice_Status
    if status not in ("Paid", "Partially Paid"):
        continue

    billing_month = to_date(inv.Billing_Month)
    due_date = to_date(inv.Due_Date)
    window_end = due_date if due_date else billing_month + datetime.timedelta(days=30)
    window_end = window_end + datetime.timedelta(days=15)

    if status == "Paid":
        paid_amount = inv.Amount
    else:
        paid_amount = round(inv.Amount - inv.Outstanding_Balance, 2)

    if paid_amount <= 0:
        continue

    split_payment = status == "Paid" and random.random() < 0.08

    if split_payment:
        first_share = round(paid_amount * random.uniform(0.3, 0.6), 2)
        second_share = round(paid_amount - first_share, 2)
        shares = [first_share, second_share]
    else:
        shares = [paid_amount]

    for share in shares:
        pay_date = fake.date_between(start_date=billing_month, end_date=window_end)
        method = random.choice(payment_methods)
        if random.random() < 0.04:
            method = None

        payments.append({
            "Payment_ID": payment_id,
            "Invoice_ID": inv.Invoice_ID,
            "Customer_ID": inv.Customer_ID,
            "Payment_Date": pay_date,
            "Amount_Paid": share,
            "Payment_Method": method,
            "Payment_Status": "Completed"
        })
        payment_id += 1

    if random.random() < 0.015:
        fail_date = fake.date_between(start_date=billing_month, end_date=window_end)
        payments.append({
            "Payment_ID": payment_id,
            "Invoice_ID": inv.Invoice_ID,
            "Customer_ID": inv.Customer_ID,
            "Payment_Date": fail_date,
            "Amount_Paid": inv.Amount,
            "Payment_Method": random.choice(payment_methods),
            "Payment_Status": "Failed"
        })
        payment_id += 1

df_payments = pd.DataFrame(payments)

num_duplicates = max(1, int(len(df_payments) * 0.01))
duplicate_rows = df_payments.sample(n=num_duplicates, random_state=606)
df_payments = pd.concat([df_payments, duplicate_rows], ignore_index=True)
df_payments = df_payments.sample(frac=1, random_state=606).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_payments.to_excel(writer, sheet_name="Payments", index=False)

print(f"Payments sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_payments)} (including {num_duplicates} intentional duplicates)")
print(f"Payment_Status breakdown:\n{df_payments['Payment_Status'].value_counts()}")
print(f"Nulls in Payment_Method: {df_payments['Payment_Method'].isna().sum()}")

completed = df_payments[df_payments["Payment_Status"] == "Completed"]
reconciliation = completed.groupby("Invoice_ID")["Amount_Paid"].sum().reset_index()
check = df_invoices.merge(reconciliation, on="Invoice_ID", how="left")
check["Amount_Paid"] = check["Amount_Paid"].fillna(0)
check["Expected_Paid"] = check["Amount"] - check["Outstanding_Balance"]
check["Diff"] = (check["Amount_Paid"] - check["Expected_Paid"]).abs()
mismatches = check[check["Diff"] > 0.01]
print(f"Reconciliation mismatches: {len(mismatches)}")