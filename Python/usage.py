import pandas as pd
from faker import Faker
import random
import datetime
from dateutil.relativedelta import relativedelta

fake = Faker("en_US")
Faker.seed(404)
random.seed(404)

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

usage_rows = []
usage_id = 1

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

    data_limit_mb = plan["Data_Limit_GB"] * 1024
    minutes_included = plan["Minutes_Included"]
    sms_included = plan["SMS_Included"]

    for month in month_starts(start, end):
        usage_ratio = random.uniform(0.3, 1.15)
        data_used = round(data_limit_mb * usage_ratio, 1)
        minutes_used = round(minutes_included * random.uniform(0.2, 1.2))
        sms_used = round(sms_included * random.uniform(0.1, 1.0))
        roaming_mb = round(random.uniform(0, 500), 1) if random.random() < 0.06 else 0

        if random.random() < 0.04:
            data_used = None

        if random.random() < 0.005:
            data_used = round(data_used * 50, 1) if data_used else 999999

        usage_rows.append({
            "Usage_ID": usage_id,
            "Subscription_ID": sub.Subscription_ID,
            "Usage_Month": month,
            "Minutes_Used": minutes_used,
            "SMS_Used": sms_used,
            "Data_Used_MB": data_used,
            "Roaming_MB": roaming_mb
        })
        usage_id += 1

df_usage = pd.DataFrame(usage_rows)

num_duplicates = max(1, int(len(df_usage) * 0.01))
duplicate_rows = df_usage.sample(n=num_duplicates, random_state=404)
df_usage = pd.concat([df_usage, duplicate_rows], ignore_index=True)
df_usage = df_usage.sample(frac=1, random_state=404).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_usage.to_excel(writer, sheet_name="Usage", index=False)

print(f"Usage sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_usage)} (including {num_duplicates} intentional duplicates)")
print(f"Nulls in Data_Used_MB: {df_usage['Data_Used_MB'].isna().sum()}")
print(f"Rows with Roaming_MB > 0: {(df_usage['Roaming_MB'] > 0).sum()}")