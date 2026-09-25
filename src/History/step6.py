import pandas as pd
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

print("phone" in paris.columns)
print("phone" in tokyo.columns)

required_columns = ["id", "date", "name", "company", "phone", "amount", "content"]

tables = {
    "Tokyo": tokyo,
    "Paris": paris,
    "New York": newyork,
}

for office, table in tables.items():
    for col in required_columns:
        if col not in table.columns:
            print(f"{office}: missing column '{col}'")

all_data = pd.concat([tokyo, paris, newyork], ignore_index=True)

all_data.to_excel("output/step3.xlsx", index=False)