import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from typesafe_sdk import TypeSafeClient, Choice, TypeSafeError

load_dotenv()
jev = TypeSafeClient()

RUNS = 3

contexts = [
    ("C1", "1", "085", "Please convert our invoice using the EUR/USD exchange rate, meaning how many US dollars one euro buys.", "decimal"),
    ("C2", "1", "879", "Please reimburse the price per liter of diesel paid at the highway station.", "decimal"),
    ("C3", "172", "350", "Please use the EUR/JPY exchange rate, meaning how many yen one euro is worth today.", "decimal"),
    ("C4", "2", "375", "The per-mile rate charged for the truck rental seems too high.", "decimal"),
    ("C5", "2", "450", "This is the monthly rent for our 120 m² office.", "thousands"),
    ("C6", "1", "250", "This is the total for the team lunch for our 25 employees.", "thousands"),
    ("C7", "3", "600", "This is the total for our annual software licence covering 20 seats.", "thousands"),
    ("C8", "4", "860", "Please send the invoice for the hotel: 5 guests, 4 nights each.", "thousands"),
]


def ask_jev_value(office, big_value, small_value, content):
    state = (f"Office: {office}. A number in this message could mean either "
             f"{big_value} or {small_value}. Message: {content}")
    try:
        response = jev.system_one(
            state=state,
            questions={
                "value": Choice(
                    instructions="Which value does the customer mean?",
                    criteria={
                        "thousands": f"{big_value} (a large number)",
                        "decimal": f"{small_value} (a small number with decimals)",
                    },
                ),
            },
        )
    except TypeSafeError as error:
        print(f"Jev error: {error}")
        return None, 0.0, None

    answer = response.answers["value"]
    return answer.choice, answer.confidence, answer.probabilities["thousands"]


records = []
for run in range(1, RUNS + 1):
    print(f"run {run}/{RUNS}")
    for context_id, before, after, content, truth in contexts:
        for office in ["Paris", "New York"]:
            for separator in [",", "."]:
                amount = before + separator + after
                big_value = int(before + after)
                small_value = float(before + "." + after)
                choice, confidence, p_thousands = ask_jev_value(office, big_value, small_value, content)
                records.append({
                    "run": run, "context": context_id, "truth": truth,
                    "office": office, "separator": separator, "amount": amount,
                    "choice": choice, "confidence": confidence, "p_thousands": p_thousands,
                })

results = pd.DataFrame(records)

table = results.pivot_table(index=["context", "truth"], columns=["office", "separator"],
                            values="p_thousands", aggfunc="mean")
print()
print("Probability of 'thousands' (average of runs)")
print(table.round(2))

run_time = datetime.now().strftime("%Y%m%d_%H%M%S")
with pd.ExcelWriter(f"output/minimal_pairs_{run_time}.xlsx") as writer:
    table.round(2).to_excel(writer, sheet_name="Table")
    results.to_excel(writer, sheet_name="All runs", index=False)
print(f"Saved: output/minimal_pairs_{run_time}.xlsx")