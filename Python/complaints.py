import pandas as pd
from faker import Faker
import random
import datetime

fake = Faker("en_US")
Faker.seed(707)
random.seed(707)

OUTPUT_FILE = "telecom_data.xlsx"

df_customers = pd.read_excel(OUTPUT_FILE, sheet_name="Customers")
df_employees = pd.read_excel(OUTPUT_FILE, sheet_name="Employees")

df_customers = df_customers.drop_duplicates(subset="Customer_ID")
df_support = df_employees[df_employees["Role"].str.strip().str.upper() == "CUSTOMER SUPPORT"]

branch_to_agents = {}
for branch_id, group in df_support.groupby("Branch_ID"):
    branch_to_agents[branch_id] = group["Employee_ID"].tolist()

all_support_ids = df_support["Employee_ID"].tolist()

def pick_agent(branch_id):
    agents = branch_to_agents.get(branch_id)
    if agents:
        return random.choice(agents)
    return random.choice(all_support_ids)

complaint_types = [
    "Billing Issue",
    "Network Coverage",
    "Data Overcharge",
    "Poor Customer Service",
    "SIM Issue",
    "Slow Internet Speed",
    "Call Drops",
    "Wrong Plan Activation"
]

statuses = ["Resolved", "Closed", "In Progress", "Open"]
status_weights = [0.55, 0.20, 0.15, 0.10]

def to_date(value):
    if pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return datetime.datetime.strptime(str(value), "%Y-%m-%d").date()

TODAY = datetime.date.today()

complaints = []
complaint_id = 1

for cust in df_customers.itertuples():
    if random.random() >= 0.40:
        continue

    reg_date = to_date(cust.Registration_Date)
    end_date = to_date(cust.Churn_Date) if cust.Status == "Churned" else TODAY
    if reg_date is None or reg_date >= end_date:
        continue

    num_complaints = random.choices([1, 2, 3], weights=[0.6, 0.3, 0.1])[0]

    for _ in range(num_complaints):
        complaint_date = fake.date_between(start_date=reg_date, end_date=end_date)
        status = random.choices(statuses, weights=status_weights)[0]
        agent_id = pick_agent(cust.Branch_ID)

        resolution_hours = None
        satisfaction = None
        if status in ("Resolved", "Closed"):
            resolution_hours = round(random.uniform(1, 120), 1)
            if random.random() >= 0.15:
                satisfaction = random.randint(1, 5)

        complaints.append({
            "Complaint_ID": complaint_id,
            "Customer_ID": cust.Customer_ID,
            "Assigned_Employee_ID": agent_id,
            "Complaint_Type": random.choice(complaint_types),
            "Complaint_Date": complaint_date,
            "Status": status,
            "Resolution_Time_Hours": resolution_hours,
            "Satisfaction": satisfaction
        })
        complaint_id += 1

df_complaints = pd.DataFrame(complaints)

num_duplicates = max(1, int(len(df_complaints) * 0.015))
duplicate_rows = df_complaints.sample(n=num_duplicates, random_state=707)
df_complaints = pd.concat([df_complaints, duplicate_rows], ignore_index=True)
df_complaints = df_complaints.sample(frac=1, random_state=707).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_complaints.to_excel(writer, sheet_name="Complaints", index=False)

print(f"Complaints sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_complaints)} (including {num_duplicates} intentional duplicates)")
print(f"Status breakdown:\n{df_complaints['Status'].value_counts()}")
print(f"Nulls in Satisfaction: {df_complaints['Satisfaction'].isna().sum()}")
print(f"Nulls in Resolution_Time_Hours: {df_complaints['Resolution_Time_Hours'].isna().sum()}")