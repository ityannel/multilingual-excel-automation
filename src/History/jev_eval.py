import pandas as pd
from dotenv import load_dotenv
from typesafe_sdk import TypeSafeClient, Choice, TypeSafeError
from datetime import datetime

load_dotenv()
jev = TypeSafeClient()


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
        return None, 0.0

    answer = response.answers["value"]
    return answer.choice, answer.confidence


JEV_THRESHOLD = 0.7
RUNS = 5


def read_amount(office, amount):
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

    return float(number_text), ambiguous


datasets = [
    ("mine", pd.read_excel("tests/jev_amount_tests.xlsx", dtype={"amount": str})),
    ("other", pd.read_csv("tests/jev_amount_tests_other.csv", dtype={"amount": str})),
]

records = []
for dataset_name, tests in datasets:
    for run in range(1, RUNS + 1):
        print(f"{dataset_name}: run {run}/{RUNS}")
        for case_id, office, amount, content in zip(tests["case_id"], tests["office"], tests["amount"], tests["content"]):
            rule_value, ambiguous = read_amount(office, amount)
            final_value = rule_value
            choice = None
            confidence = None
            needs_review = False

            if ambiguous:
                big_value = int(round(rule_value * 1000))
                choice, confidence = ask_jev_value(office, big_value, rule_value, content)
                if confidence >= JEV_THRESHOLD:
                    if choice == "thousands":
                        final_value = big_value
                else:
                    needs_review = True

            records.append({
                "dataset": dataset_name, "run": run, "case_id": case_id, "ambiguous": ambiguous,
                "rule_value": rule_value, "jev_choice": choice, "jev_confidence": confidence,
                "final_value": final_value, "needs_review": needs_review,
            })

all_tests = pd.concat([tests for _, tests in datasets], ignore_index=True)

results = pd.merge(pd.DataFrame(records), all_tests, on="case_id")
results["rule_ok"] = (results["rule_value"] - results["truth_value"]).abs() < 0.001
results["final_ok"] = (results["final_value"] - results["truth_value"]).abs() < 0.001
results["silent_error"] = ~results["final_ok"] & ~results["needs_review"]

summary = results.groupby(["dataset", "run"]).agg(
    ambiguous=("ambiguous", "sum"),
    rule_correct=("rule_ok", "sum"),
    final_correct=("final_ok", "sum"),
    needs_review=("needs_review", "sum"),
    silent_errors=("silent_error", "sum"),
)
print()
print(summary)

run_time = datetime.now().strftime("%Y%m%d_%H%M%S")
with pd.ExcelWriter(f"output/jev_eval_v2_{run_time}.xlsx") as writer:
    summary.to_excel(writer, sheet_name="Summary")
    results.to_excel(writer, sheet_name="All runs", index=False)
print(f"Saved: output/jev_eval_v2_{run_time}.xlsx")
