import calendar

def prove_date_order(value):
    parts = value.split("/")
    if len(parts) != 3:
        return None
    if not parts[0].isdigit() or not parts[1].isdigit() or not parts[2].isdigit():
        return None
    if len(parts[2]) != 4:
        return None

    first = int(parts[0])
    second = int(parts[1])
    year = int (parts[2])

    dmy = False
    if 1 <= second <= 12:
        if 1 <= first <= calendar.monthrange(year, second)[1]:
            dmy = True

    mdy = False
    if 1 <= first <= 12:
        if 1 <= second <= calendar.monthrange(year, first)[1]:
            mdy = True

    if dmy and not mdy:
        return "DMY"
    if mdy and not dmy:
        return "MDY"

    return None

def find_file_date_ordders(table):
    orders = {}
    for source_file, group in table.groupby("source_file"):
        found = set()
        for value in group["date"]:
            if isinstance(value, str):
                order = prove_date_order(value)
                if order is not None:
                    found.add(order)
        if len(found) == 1:
            orders[source_file] = found.pop()
        elif len(found) > 1:
            issues.append({"office": office, "id": row_id, "problem": "Can't decide the date order!"})
            orders[source_file] = None
        else:
            orders[source_file] = None
        