import pandas as pd
from faker import Faker
import random

fake = Faker("en_US")
Faker.seed(42)
random.seed(42)

OUTPUT_FILE = "telecom_data.xlsx"

egypt_regions = [
    ("Cairo", "Cairo", "Downtown"),
    ("Cairo", "Cairo", "Nasr City"),
    ("Cairo", "Cairo", "Maadi"),
    ("Giza", "Giza", "Dokki"),
    ("Giza", "Giza", "6th of October"),
    ("Alexandria", "Alexandria", "Smouha"),
    ("Alexandria", "Alexandria", "Miami"),
    ("Delta", "Mansoura", "City Center"),
    ("Delta", "Tanta", "City Center"),
    ("Canal", "Suez", "City Center"),
    ("Canal", "Ismailia", "City Center"),
    ("Upper Egypt", "Aswan", "City Center"),
    ("Upper Egypt", "Luxor", "City Center"),
    ("Upper Egypt", "Assiut", "City Center"),
    ("Red Sea", "Hurghada", "City Center"),
]

regions = []
for i, (region, city, area) in enumerate(egypt_regions, start=1):
    regions.append({
        "Region_ID": i,
        "Region": region,
        "City": city,
        "Area": area
    })
df_regions = pd.DataFrame(regions)

branches = []
branch_id = 1
for region in df_regions.itertuples():
    num_branches = random.randint(1, 2)
    for _ in range(num_branches):
        branches.append({
            "Branch_ID": branch_id,
            "Branch_Name": f"{region.City} Branch {branch_id}",
            "Region_ID": region.Region_ID,
            "Employees_Count": random.randint(5, 25)
        })
        branch_id += 1
df_branches = pd.DataFrame(branches)

technologies = ["3G", "4G", "5G"]
tech_weights = [0.15, 0.55, 0.30]

towers = []
tower_id = 1
for region in df_regions.itertuples():
    num_towers = random.randint(3, 8)
    for _ in range(num_towers):
        towers.append({
            "Tower_ID": tower_id,
            "Tower_Name": f"Tower-{region.City}-{tower_id}",
            "Region_ID": region.Region_ID,
            "Technology": random.choices(technologies, weights=tech_weights)[0],
            "Latitude": round(fake.latitude(), 6),
            "Longitude": round(fake.longitude(), 6)
        })
        tower_id += 1
df_towers = pd.DataFrame(towers)

plans_data = [
    ("Basic Prepaid", "Prepaid", 50, 5, 200, 100),
    ("Standard Prepaid", "Prepaid", 100, 15, 500, 300),
    ("Premium Prepaid", "Prepaid", 180, 30, 1000, 500),
    ("Basic Postpaid", "Postpaid", 150, 20, 1000, 500),
    ("Standard Postpaid", "Postpaid", 250, 40, 2000, 1000),
    ("Premium Postpaid", "Postpaid", 400, 80, 5000, 2000),
    ("Unlimited Postpaid", "Postpaid", 600, 150, 99999, 99999),
    ("Student Plan", "Prepaid", 70, 10, 300, 200),
    ("Business Plan", "Postpaid", 500, 100, 3000, 1500),
    ("Youth Plan", "Prepaid", 90, 12, 400, 400),
]

plans = []
for i, (name, ptype, price, data_gb, minutes, sms) in enumerate(plans_data, start=1):
    plans.append({
        "Plan_ID": i,
        "Plan_Name": name,
        "Plan_Type": ptype,
        "Monthly_Price": price,
        "Data_Limit_GB": data_gb,
        "Minutes_Included": minutes,
        "SMS_Included": sms
    })
df_plans = pd.DataFrame(plans)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="w") as writer:
    df_regions.to_excel(writer, sheet_name="Regions", index=False)
    df_branches.to_excel(writer, sheet_name="Branches", index=False)
    df_towers.to_excel(writer, sheet_name="Cell_Towers", index=False)
    df_plans.to_excel(writer, sheet_name="Plans", index=False)

print(f"{OUTPUT_FILE} created successfully")
print(f"Regions: {len(df_regions)} rows")
print(f"Branches: {len(df_branches)} rows")
print(f"Cell_Towers: {len(df_towers)} rows")
print(f"Plans: {len(df_plans)} rows")