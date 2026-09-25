import pandas as pd
import unicodedata
import dateparser
from datetime import datetime

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

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

required_columns = ["id", "date", "name", "company", "phone", "amount", "content"]

date_orders = {
    "Tokyo": "YMD",
    "Paris": "DMY",
    "New York": "MDY",
}

default_currencies = {
    "Tokyo": "JPY",
    "Paris": "EUR",
    "New York": "USD",
}


def clean_text(value):
    if isinstance(value, str):
        return unicodedata.normalize("NFKC", value).strip()
    return value


def parse_date(row):
    value = row["date"]
    office = row["office"]
    row_id = row["id"]

    if isinstance(value, datetime):
        return value
    if not isinstance(value, str):
        issues.append({"office": office, "id": row_id,
                       "problem": "Date is missing"})
        return None

    if value.startswith("R"):
        year, month, day = value[1:].split(".")
        return datetime(2018 + int(year), int(month), int(day))

    if value[:4].isdigit():
        order = "YMD"
    else:
        order = date_orders[office]

    result = dateparser.parse(
        value,
        languages=["ja", "fr", "en"],
        settings={"DATE_ORDER": order},
    )

    if result is None:
        issues.append({"office": office, "id": row_id,
                       "problem": f"Could not read date '{value}'"})
        return None

    if str(result.year) not in value:
        issues.append({"office": office, "id": row_id,
                       "problem": f"Year missing or abbreviated in '{value}' (assumed {result.year})"})

    parts = value.split("/")
    if len(parts) == 3 and order in ["MDY", "DMY"]:
        first = int(parts[0])
        if order == "MDY":
            expected = result.month
        else:
            expected = result.day
        if first != expected:
            issues.append({"office": office, "id": row_id,
                           "problem": f"Date '{value}' does not follow the {office} format ({order})"})

    return result


def parse_currency(row):
    office = row["office"]
    amount = row["amount"]

    if pd.isna(amount):
        return None
    if not isinstance(amount, str):
        return default_currencies[office]

    if "¥" in amount or "円" in amount:
        return "JPY"
    elif "€" in amount:
        return "EUR"
    elif "$" in amount or "USD" in amount:
        return "USD"
    else:
        return default_currencies[office]

def parse_amount(row):
    office = row["office"]
    amount = row["amount"]
    row_id = row["id"]

    if pd.isna(amount):
        issues.append({"office": office, "id": row_id, "problem": "Amount is missing"})
        return None
    if not isinstance(amount, str):
        return float(amount)

    number_text = ""
    for c in amount:
        if c in "0123456789,.-":
            number_text = number_text + c

    if "," in number_text and "." in number_text:
        if number_text.rfind(",") > number_text.rfind("."):
            decimal = ","
        else:
            decimal = "."
    elif office == "Paris":
        decimal = ","
    else:
        decimal = "."

    parts = number_text.split(decimal)
    ambiguous = len(parts) == 2 and len(parts[1]) == 3

    if decimal == ",":
        number_text = number_text.replace(".", "")
        number_text = number_text.replace(",", ".")
    else:
        number_text = number_text.replace(",", "")


    try:
        result = float(number_text)
    except ValueError:
        issues.append({"office": office, "id": row_id, "problem": f"could not read amount '{amount}'"})
        return None

    if result < 0:
        issues.append({"office": office, "id": row_id, "problem": f"Amount is negative ({amount})"})
    if ambiguous:
        issues.append({"office": office, "id": row_id, "problem": f"Ambiguous amount '{amount}' (read as {result}). Please check."})

    return result

tokyo = pd.read_excel("input/inquiries_tokyo.xlsx")
paris = pd.read_excel("input/inquiries_paris.xlsx")
newyork = pd.read_excel("input/inquiries_newyork.xlsx")

tokyo = tokyo.rename(columns=tokyo_columns)
tokyo["office"] = "Tokyo"

paris = paris.rename(columns=paris_columns)
paris["office"] = "Paris"

newyork = newyork.rename(columns=newyork_columns)
newyork["office"] = "New York"

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

all_data = pd.concat([tokyo, paris, newyork], ignore_index=True)

for col in ["id", "date", "name", "company", "phone", "amount"]:
    all_data[col] = all_data[col].map(clean_text)

all_data["date_clean"] = all_data.apply(parse_date, axis=1)
all_data["currency"] = all_data.apply(parse_currency, axis=1)
all_data["amount_clean"] = all_data.apply(parse_amount, axis=1)

issues_table = pd.DataFrame(issues)

with pd.ExcelWriter("output/master.xlsx", datetime_format="YYYY-MM-DD") as writer:
    all_data.to_excel(writer, sheet_name="Data", index=False)
    issues_table.to_excel(writer, sheet_name="Issues", index=False)

print("Saved: output/master.xlsx")