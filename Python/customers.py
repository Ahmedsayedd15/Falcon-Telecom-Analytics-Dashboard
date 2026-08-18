import pandas as pd
from faker import Faker
import random

fake = Faker("en_US")
Faker.seed(202)
random.seed(202)

OUTPUT_FILE = "telecom_data.xlsx"

df_regions = pd.read_excel(OUTPUT_FILE, sheet_name="Regions")
df_branches = pd.read_excel(OUTPUT_FILE, sheet_name="Branches")

churn_reasons = [
    "High Price",
    "Poor Network Coverage",
    "Better Competitor Offer",
    "Poor Customer Service",
    "Relocated",
    "No Longer Needed",
    "Billing Issues"
]

genders = ["Male", "Female"]

NUM_CUSTOMERS = 3000
customers = []

for i in range(1, NUM_CUSTOMERS + 1):
    branch = df_branches.sample(1).iloc[0]
    region_id = branch["Region_ID"]

    reg_date = fake.date_between(start_date="-5y", end_date="-1M")

    status = random.choices(["Active", "Churned"], weights=[0.82, 0.18])[0]

    churn_date = None
    churn_reason = None
    if status == "Churned":
        churn_date = fake.date_between(start_date=reg_date, end_date="today")
        churn_reason = random.choice(churn_reasons)

    name = fake.name()
    if random.random() < 0.12:
        space_choice = random.choice(["leading", "trailing", "double"])
        if space_choice == "leading":
            name = "  " + name
        elif space_choice == "trailing":
            name = name + "  "
        else:
            name = name.replace(" ", "  ")

    email = fake.email()
    if random.random() < 0.05:
        email = None

    phone = fake.numerify("01#########")

    customers.append({
        "Customer_ID": i,
        "Customer_Name": name,
        "Gender": random.choice(genders),
        "Age": random.randint(18, 70),
        "Phone": phone,
        "Email": email,
        "Region_ID": region_id,
        "Branch_ID": branch["Branch_ID"],
        "Registration_Date": reg_date,
        "Status": status,
        "Churn_Date": churn_date,
        "Churn_Reason": churn_reason
    })

df_customers = pd.DataFrame(customers)

num_duplicates = max(1, int(len(df_customers) * 0.01))
duplicate_rows = df_customers.sample(n=num_duplicates, random_state=202)
df_customers = pd.concat([df_customers, duplicate_rows], ignore_index=True)
df_customers = df_customers.sample(frac=1, random_state=202).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_customers.to_excel(writer, sheet_name="Customers", index=False)

print(f"Customers sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_customers)} (including {num_duplicates} intentional duplicates)")
print(f"Nulls in Email: {df_customers['Email'].isna().sum()}")
print(f"Status breakdown:\n{df_customers['Status'].value_counts()}")