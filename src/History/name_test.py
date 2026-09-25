from dotenv import load_dotenv
from typesafe_sdk import TypeSafeClient, Choice, TypeSafeError

load_dotenv()
jev = TypeSafeClient()


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
        return None, 0.0, candidates

    answer = response.answers["split"]
    return answer.choice, answer.confidence, candidates


tests = [
    ("佐藤花子", "split_2"),
    ("田中", "family_only"),
    ("山田太郎", "split_2"),
    ("林さくら", "split_1"),
    ("勅使河原健", "split_4"),
    ("長谷川", "family_only"),
    ("鈴木一郎", "split_2"),
    ("東海林さだお", "split_3"),
    ("中村", "family_only"),
    ("小林愛", "split_2"),
]

correct = 0
for name, truth in tests:
    choice, confidence, candidates = ask_jev_name(name)
    mark = "OK" if choice == truth else "NG"
    if choice == truth:
        correct = correct + 1
    print(f"{mark}  {name} → {candidates.get(choice)}  (confidence {confidence:.2f})")

print(f"{correct}/{len(tests)} correct")