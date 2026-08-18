import pandas as pd
from faker import Faker
import random
import datetime
from dateutil.relativedelta import relativedelta

fake = Faker("en_US")
Faker.seed(808)
random.seed(808)

OUTPUT_FILE = "telecom_data.xlsx"

df_towers = pd.read_excel(OUTPUT_FILE, sheet_name="Cell_Towers")
df_towers = df_towers.drop_duplicates(subset="Tower_ID")

START_MONTH = datetime.date(2023, 1, 1)
END_MONTH = datetime.date.today().replace(day=1)

def month_range(start, end):
    current = start
    months = []
    while current <= end:
        months.append(current)
        current = current + relativedelta(months=1)
    return months

months = month_range(START_MONTH, END_MONTH)

tech_profile = {
    "3G": {"uptime": (90, 97), "signal": (-110, -95), "traffic": (500, 2000), "drop_rate": (2, 6)},
    "4G": {"uptime": (95, 99.5), "signal": (-95, -75), "traffic": (2000, 8000), "drop_rate": (0.5, 3)},
    "5G": {"uptime": (97, 99.9), "signal": (-85, -65), "traffic": (8000, 25000), "drop_rate": (0.2, 1.5)},
}

records = []
record_id = 1

for tower in df_towers.itertuples():
    profile = tech_profile[tower.Technology]

    for month in months:
        uptime_pct = round(random.uniform(*profile["uptime"]), 2)
        downtime_hours = round((100 - uptime_pct) / 100 * 24 * 30, 2)
        signal_strength = random.randint(*profile["signal"])
        data_traffic_gb = round(random.uniform(*profile["traffic"]), 1)
        drop_rate_pct = round(random.uniform(*profile["drop_rate"]), 2)
        dropped_calls = round(random.randint(500, 3000) * (drop_rate_pct / 100))

        if random.random() < 0.04:
            signal_strength = None

        records.append({
            "Network_ID": record_id,
            "Tower_ID": tower.Tower_ID,
            "Month": month,
            "Uptime_Percentage": uptime_pct,
            "Downtime_Hours": downtime_hours,
            "Dropped_Calls": dropped_calls,
            "Signal_Strength_dBm": signal_strength,
            "Data_Traffic_GB": data_traffic_gb
        })
        record_id += 1

df_network = pd.DataFrame(records)

num_duplicates = max(1, int(len(df_network) * 0.01))
duplicate_rows = df_network.sample(n=num_duplicates, random_state=808)
df_network = pd.concat([df_network, duplicate_rows], ignore_index=True)
df_network = df_network.sample(frac=1, random_state=808).reset_index(drop=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_network.to_excel(writer, sheet_name="Network_Performance", index=False)

print(f"Network_Performance sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_network)} (including {num_duplicates} intentional duplicates)")
print(f"Nulls in Signal_Strength_dBm: {df_network['Signal_Strength_dBm'].isna().sum()}")
print(f"Average Uptime by Technology:")
print(df_network.merge(df_towers[["Tower_ID", "Technology"]], on="Tower_ID").groupby("Technology")["Uptime_Percentage"].mean())