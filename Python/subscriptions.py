import pandas as pd
from faker import Faker
import random
import datetime

fake = Faker("en_US")
Faker.seed(303)
random.seed(303)

OUTPUT_FILE = "telecom_data.xlsx"

df_customers = pd.read_excel(OUTPUT_FILE, sheet_name="Customers")
df_plans = pd.read_excel(OUTPUT_FILE, sheet_name="Plans")
df_employees = pd.read_excel(OUTPUT_FILE, sheet_name="Employees")

df_customers = df_customers.drop_duplicates(subset="Customer_ID")
df_sales_agents = df_employees[df_employees["Role"].str.strip().str.upper() == "SALES AGENT"]

branch_to_agents = {}
for branch_id, group in df_sales_agents.groupby("Branch_ID"):
    branch_to_agents[branch_id] = group["Employee_ID"].tolist()

all_agent_ids = df_sales_agents["Employee_ID"].tolist()

def pick_agent(branch_id):
    agents = branch_to_agents.get(branch_id)
    if agents:
        return random.choice(agents)
    return random.choice(all_agent_ids)

subscriptions = []
subscription_id = 1

for cust in df_customers.itertuples():
    reg_date = cust.Registration_Date
    if isinstance(reg_date, str):
        reg_date = datetime.datetime.strptime(reg_date, "%Y-%m-%d").date()
    elif isinstance(reg_date, pd.Timestamp):
        reg_date = reg_date.date()

    is_churned = cust.Status == "Churned"
    churn_date = cust.Churn_Date
    if isinstance(churn_date, pd.Timestamp):
        churn_date = churn_date.date()

    first_plan = df_plans.sample(1).iloc[0]
    agent_id = pick_agent(cust.Branch_ID)

    has_change = random.random() < 0.25 and not is_churned
    has_change_before_churn = random.random() < 0.15 and is_churned

    if is_churned:
        end_date_final = churn_date
        final_status = "Cancelled"
    else:
        end_date_final = None
        final_status = "Active"

    if has_change or has_change_before_churn:
        if is_churned:
            max_change_date = churn_date
        else:
            max_change_date = datetime.date.today()

        if max_change_date > reg_date:
            change_date = fake.date_between(start_date=reg_date, end_date=max_change_date)
        else:
            change_date = reg_date

        subscriptions.append({
            "Subscription_ID": subscription_id,
            "Customer_ID": cust.Customer_ID,
            "Plan_ID": first_plan["Plan_ID"],
            "Employee_ID": agent_id,
            "Start_Date": reg_date,
            "End_Date": change_date,
            "Status": "Changed",
            "Change_Type": "New"
        })
        subscription_id += 1

        second_plan = df_plans.sample(1).iloc[0]
        while second_plan["Plan_ID"] == first_plan["Plan_ID"]:
            second_plan = df_plans.sample(1).iloc[0]

        change_type = "Upgrade" if second_plan["Monthly_Price"] > first_plan["Monthly_Price"] else "Downgrade"

        subscriptions.append({
            "Subscription_ID": subscription_id,
            "Customer_ID": cust.Customer_ID,
            "Plan_ID": second_plan["Plan_ID"],
            "Employee_ID": agent_id,
            "Start_Date": change_date,
            "End_Date": end_date_final,
            "Status": final_status,
            "Change_Type": change_type
        })
        subscription_id += 1

    else:
        subscriptions.append({
            "Subscription_ID": subscription_id,
            "Customer_ID": cust.Customer_ID,
            "Plan_ID": first_plan["Plan_ID"],
            "Employee_ID": agent_id,
            "Start_Date": reg_date,
            "End_Date": end_date_final,
            "Status": final_status,
            "Change_Type": "New"
        })
        subscription_id += 1

df_subscriptions = pd.DataFrame(subscriptions)

num_duplicates = max(1, int(len(df_subscriptions) * 0.01))
duplicate_rows = df_subscriptions.sample(n=num_duplicates, random_state=303)
df_subscriptions = pd.concat([df_subscriptions, duplicate_rows], ignore_index=True)
df_subscriptions = df_subscriptions.sample(frac=1, random_state=303).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_subscriptions.to_excel(writer, sheet_name="Subscriptions", index=False)

print(f"Subscriptions sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_subscriptions)} (including {num_duplicates} intentional duplicates)")
print(f"Status breakdown:\n{df_subscriptions['Status'].value_counts()}")
print(f"Change_Type breakdown:\n{df_subscriptions['Change_Type'].value_counts()}")