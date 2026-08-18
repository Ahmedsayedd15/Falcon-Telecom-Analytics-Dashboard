import pandas as pd
from faker import Faker
import random
import datetime
from dateutil.relativedelta import relativedelta

fake = Faker("en_US")
Faker.seed(505)
random.seed(505)

OUTPUT_FILE = "telecom_data.xlsx"

df_subscriptions = pd.read_excel(OUTPUT_FILE, sheet_name="Subscriptions")
df_plans = pd.read_excel(OUTPUT_FILE, sheet_name="Plans")

df_subscriptions = df_subscriptions.drop_duplicates(subset="Subscription_ID")
plans_lookup = df_plans.set_index("Plan_ID").to_dict(orient="index")

def to_date(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return datetime.datetime.strptime(str(value), "%Y-%m-%d").date()

def month_starts(start, end):
    current = start.replace(day=1)
    end = end.replace(day=1)
    months = []
    while current <= end:
        months.append(current)
        current = current + relativedelta(months=1)
    return months

TODAY = datetime.date.today()

invoices = []
invoice_id = 1

for sub in df_subscriptions.itertuples():
    start = to_date(sub.Start_Date)
    end = to_date(sub.End_Date)
    if end is None:
        end = TODAY
    if start is None or start > end:
        continue

    plan = plans_lookup.get(sub.Plan_ID)
    if plan is None:
        continue

    amount = plan["Monthly_Price"]

    for month in month_starts(start, end):
        billing_date = month
        due_date = billing_date + datetime.timedelta(days=15)

        status_roll = random.random()
        if status_roll < 0.75:
            status = "Paid"
        elif status_roll < 0.83:
            status = "Partially Paid"
        else:
            status = "Unpaid"

        if status != "Paid" and due_date < TODAY:
            status = "Overdue"

        if status == "Paid":
            outstanding = 0.0
        elif status == "Partially Paid":
            outstanding = round(amount * random.uniform(0.2, 0.7), 2)
        else:
            outstanding = amount

        if random.random() < 0.02:
            due_date = None

        invoices.append({
            "Invoice_ID": invoice_id,
            "Subscription_ID": sub.Subscription_ID,
            "Customer_ID": sub.Customer_ID,
            "Billing_Month": billing_date,
            "Amount": amount,
            "Due_Date": due_date,
            "Invoice_Status": status,
            "Outstanding_Balance": outstanding
        })
        invoice_id += 1

df_invoices = pd.DataFrame(invoices)

num_duplicates = max(1, int(len(df_invoices) * 0.01))
duplicate_rows = df_invoices.sample(n=num_duplicates, random_state=505)
df_invoices = pd.concat([df_invoices, duplicate_rows], ignore_index=True)
df_invoices = df_invoices.sample(frac=1, random_state=505).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_invoices.to_excel(writer, sheet_name="Invoices", index=False)

print(f"Invoices sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_invoices)} (including {num_duplicates} intentional duplicates)")
print(f"Status breakdown:\n{df_invoices['Invoice_Status'].value_counts()}")
print(f"Nulls in Due_Date: {df_invoices['Due_Date'].isna().sum()}")