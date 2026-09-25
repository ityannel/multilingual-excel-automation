import pandas as pd
from dotenv import load_dotenv
from typesafe_sdk import TypeSafeClient, Choice, TypeSafeError
from datetime import datetime

load_dotenv()
jev = TypeSafeClient()

JEV_THRESHOLD = 0.7
RUNS = 3


def name_candidates(name):
    candidates = {}
    for i in range(1, len(name)):
        candidates[f"split_{i}"] = f"Family name: {name[:i]} / Given name: {name[i:]}"
    candidates["family_only"] = f"Family name only: {name} (the given name is missing)"
    return candidates


def ask_jev_name(name):
    candidates = name_candidates(name)
    state = f"A Japanese person's name was written without a space between family name and given name: {name}"
    try:
        response = jev.system_one(
            state=state,
            questions={
                "split": Choice(
                    instructions="How should this name be divided into family name and given name?",
                    criteria=candidates,
                ),
            },
        )
    except TypeSafeError as error:
        print(f"Jev error: {error}")
        return None, 0.0

    answer = response.answers["split"]
    return answer.choice, answer.confidence


def truth_key(family_name, given_name):
    if given_name == "":
        return "family_only"
    return f"split_{len(family_name)}"


tests = pd.read_csv("tests/name_tests.tsv", sep="\t", keep_default_na=False)
tests["truth"] = [truth_key(f, g) for f, g in zip(tests["family_name"], tests["given_name"])]

records = []
for run in range(1, RUNS + 1):
    print(f"run {run}/{RUNS}")
    for case_id, name in zip(tests["case_id"], tests["name"]):
        if len(name) == 1:
            choice, confidence = "family_only", 1.0
            asked = False
        else:
            choice, confidence = ask_jev_name(name)
            asked = True

        records.append({
            "run": run, "case_id": case_id, "asked_jev": asked,
            "jev_choice": choice, "jev_confidence": confidence,
            "needs_review": confidence < JEV_THRESHOLD,
        })

results = pd.merge(pd.DataFrame(records), tests, on="case_id")
results["choice_ok"] = results["jev_choice"] == results["truth"]
results["auto_correct"] = results["choice_ok"] & ~results["needs_review"]
results["silent_error"] = ~results["choice_ok"] & ~results["needs_review"]

summary = results.groupby(["type", "run"]).agg(
    cases=("case_id", "count"),
    auto_correct=("auto_correct", "sum"),
    needs_review=("needs_review", "sum"),
    silent_errors=("silent_error", "sum"),
    jev_choice_correct=("choice_ok", "sum"),
)
print()
print(summary)

print()
print("Silent errors and wrong choices:")
wrong = results[~results["choice_ok"]].groupby("case_id").agg(
    name=("name", "first"), type=("type", "first"), truth=("truth", "first"),
    choices=("jev_choice", lambda c: ", ".join(sorted(set(c)))),
    confidence_min=("jev_confidence", "min"), confidence_max=("jev_confidence", "max"),
    times_wrong=("run", "count"),
)
print(wrong)

run_time = datetime.now().strftime("%Y%m%d_%H%M%S")
with pd.ExcelWriter(f"output/name_eval_{run_time}.xlsx") as writer:
    summary.to_excel(writer, sheet_name="Summary")
    wrong.to_excel(writer, sheet_name="Wrong")
    results.to_excel(writer, sheet_name="All runs", index=False)
print(f"Saved: output/name_eval_{run_time}.xlsx")
