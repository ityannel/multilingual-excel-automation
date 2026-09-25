from pathlib import Path

file_offices = {
    "tokyo": "Tokyo",
    "paris": "Paris",
    "newyork": "New York",
}


def find_office(file_stem):
    for key, office in file_offices.items():
        if key in file_stem.lower():
            return office
    return None


input_files = sorted(Path("input").glob("*.xlsx"))

for path in input_files:
    if path.name.startswith("~$"):
        print("skip:", path.name)
        continue
    print(path.name, "→", find_office(path.stem))