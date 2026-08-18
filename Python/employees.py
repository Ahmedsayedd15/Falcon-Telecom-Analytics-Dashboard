import pandas as pd
from faker import Faker
import random

fake = Faker("en_US")
Faker.seed(101)
random.seed(101)

OUTPUT_FILE = "telecom_data.xlsx"

df_branches = pd.read_excel(OUTPUT_FILE, sheet_name="Branches")

roles_clean = [
    "Sales Agent",
    "Customer Support",
    "Network Engineer",
    "Branch Manager",
    "Technician"
]

role_variants = {
    "Sales Agent": ["Sales Agent", "sales agent", "SALES AGENT", "Sales agent"],
    "Customer Support": ["Customer Support", "customer support", "CUSTOMER SUPPORT"],
    "Network Engineer": ["Network Engineer", "network engineer", "Network engineer"],
    "Branch Manager": ["Branch Manager", "branch manager", "BRANCH MANAGER"],
    "Technician": ["Technician", "technician", "TECHNICIAN"]
}

employees = []
employee_id = 1

for branch in df_branches.itertuples():
    num_employees = branch.Employees_Count

    for _ in range(num_employees):
        role_clean = random.choice(roles_clean)
        role_dirty = random.choice(role_variants[role_clean])

        name = fake.name()

        if random.random() < 0.15:
            space_choice = random.choice(["leading", "trailing", "double"])

            if space_choice == "leading":
                name = "  " + name
            elif space_choice == "trailing":
                name = name + "   "
            else:
                name = name.replace(" ", "  ")

        salary_base = {
            "Sales Agent": (6000, 12000),
            "Customer Support": (5500, 10000),
            "Network Engineer": (10000, 20000),
            "Branch Manager": (18000, 30000),
            "Technician": (7000, 13000)
        }[role_clean]

        salary = random.randint(*salary_base)

        target = random.randint(50000, 200000) if role_clean == "Sales Agent" else None
        actual_sales = round(target * random.uniform(0.6, 1.3)) if target else None

        if random.random() < 0.08:
            salary = None

        if role_clean == "Sales Agent" and random.random() < 0.10:
            actual_sales = None

        employees.append({
            "Employee_ID": employee_id,
            "Employee_Name": name,
            "Branch_ID": branch.Branch_ID,
            "Role": role_dirty,
            "Salary": salary,
            "Target": target,
            "Actual_Sales": actual_sales,
            "Hire_Date": fake.date_between(start_date="-6y", end_date="-1M")
        })

        employee_id += 1

df_employees = pd.DataFrame(employees)

num_duplicates = max(1, int(len(df_employees) * 0.02))

duplicate_rows = df_employees.sample(
    n=num_duplicates,
    random_state=101
)

df_employees = pd.concat(
    [df_employees, duplicate_rows],
    ignore_index=True
)

df_employees = df_employees.sample(
    frac=1,
    random_state=101
).reset_index(drop=True)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace"
) as writer:
    df_employees.to_excel(
        writer,
        sheet_name="Employees",
        index=False
    )

print(f"Employees sheet added to {OUTPUT_FILE}")
print(f"Total rows: {len(df_employees)} ({num_duplicates} intentional duplicate rows)")
print(f"Null values in Salary: {df_employees['Salary'].isna().sum()}")
print(f"Null values in Actual_Sales: {df_employees['Actual_Sales'].isna().sum()}")
print(f"Role variants: {df_employees['Role'].unique()[:8].tolist()}")