default_currencies = {
    "Tokyo": "JPY",
    "Paris": "EUR",
    "New York": "USD",
}

samples = [
    ("¥12,000", "Tokyo"),
    ("8500", "Tokyo"),
    ("990", "Paris"),
    ("2,400", "New York"),
]

for s, office in samples:
    if "¥" in s or "円" in s:
        currency = "JPY"
    elif "€" in s:
        currency = "EUR"
    elif "$" in s or "USD" in s:
        currency = "USD"
    else:
        currency = default_currencies[office]
    print(s, office, "→", currency)

print("-----")

tests = [
    ("1234,50", "Paris"),
    ("2.500,00", "Paris"),
    ("1,250.00", "New York"),
    ("15,000", "Tokyo"),
]

for number_text, office in tests:
    if office == "Paris":
        number_text = number_text.replace(".", "")
        number_text = number_text.replace(",", ".")
    else:
        number_text = number_text.replace(",", "")
    print(office, "→", number_text)


print("-----")

for number_text in ["1234.50", "15000", "", "-450.00"]:
    try:
        amount = float(number_text)
        print(number_text, "→", amount)
    except ValueError:
        print(number_text, "→ could not read")