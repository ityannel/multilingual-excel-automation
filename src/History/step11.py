import pandas as pd
import unicodedata
import dateparser
from datetime import datetime

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

tokyo = pd.read_excel("input/inquiries_tokyo.xlsx")
paris = pd.read_excel("input/inquiries_paris.xlsx")
newyork= pd.read_excel("input/inquiries_newyork.xlsx")

tokyo_columns = {
    "問い合わせID": "id",
    "受付日": "date",
    "氏名": "name",
    "会社名": "company",
    "電話番号": "phone",
    "金額": "amount",
    "問い合わせ内容": "content",
}

paris_columns = {
    "N° demande": "id",
    "Date de réception": "date",
    "Nom complet": "name",
    "Société": "company",
    "Montant": "amount",
    "Message": "content",
}

newyork_columns = {
    "Inquiry ID": "id",
    "Date Received": "date",
    "Full Name": "name",
    "Company": "company",
    "Phone": "phone",
    "Amount": "amount",
    "Message": "content",
}

tokyo = tokyo.rename(columns=tokyo_columns)
tokyo["office"] = "Tokyo"

paris = paris.rename(columns=paris_columns)
paris["office"] = "Paris"

newyork = newyork.rename(columns=newyork_columns)
newyork["office"] = "New York"

required_columns = ["id", "date", "name", "company", "phone", "amount", "content"]

tables = {
    "Tokyo": tokyo,
    "Paris": paris,
    "New York": newyork,
}

issues = []

for office, table in tables.items():
    for col in required_columns:
        if col not in table.columns:
            print(f"{office}: missing column '{col}'")
            issues.append({
                "office": office,
                "id": "(whole file)",
                "problem": f"Missing column '{col}'"
            })

issues_table = pd.DataFrame(issues)

all_data = pd.concat([tokyo, paris, newyork], ignore_index=True)

def clean_text(value):
    if isinstance(value, str):
        return unicodedata.normalize("NFKC", value).strip()
    return value

for col in ["id", "date", "name", "company", "phone", "amount"]:
    all_data[col] = all_data[col].map(clean_text)

date_orders = {
    "Tokyo": "YMD",
    "Paris": "DMY",
    "New York": "MDY",
}

def parse_date(row):
    value = row["date"]
    office = row["office"]

    if isinstance(value, datetime):
        return value
    if not isinstance(value, str):
        return None

    if value.startswith("R"):
        year, month, day = value[1:].split(".")
        return datetime(2018 + int(year), int(month), int(day))

    if value[:4].isdigit():
        order = "YMD"
    else:
        order = date_orders[office]
    
    return dateparser.parse(
        value,
        languages=["ja", "fr", "en"],
        settings={"DATE_ORDER": order},
    )

all_data["date_clean"] = all_data.apply(parse_date, axis=1)

with pd.ExcelWriter("output/master.xlsx", datetime_format="YYYY-MM-DD") as writer:
    all_data.to_excel(writer, sheet_name="Data", index=False)
    issues_table.to_excel(writer, sheet_name="Issues", index=False)

print("Saved: output/master.xlsx")