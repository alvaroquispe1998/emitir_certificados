import re


def parse_period(value: object) -> int | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    digits = re.findall(r"\d+", text)
    if not digits:
        return None
    return int("".join(digits))


def balance_electives(records: list[dict], years: set[int] | None = None):
    if years is None:
        years = {6, 7, 8}

    if len(records) != 3:
        return None, "El alumno no tiene exactamente 3 electivos aprobados"

    for idx, rec in enumerate(records):
        rec["_idx"] = idx
        rec["TARGET_YEAR"] = rec.get("YEAR_INT")

    counts = {y: 0 for y in years}
    for rec in records:
        year = rec.get("TARGET_YEAR")
        if year in counts:
            counts[year] += 1

    def pick_oldest(year):
        candidates = []
        for rec in records:
            if rec.get("TARGET_YEAR") != year:
                continue
            period = parse_period(rec.get("PERIODO"))
            candidates.append((period is None, period or 0, rec.get("_idx", 0), rec))
        candidates.sort()
        return candidates[0][3] if candidates else None

    max_loops = 10
    while max_loops > 0:
        max_loops -= 1
        deficits = [y for y, c in counts.items() if c < 1]
        surpluses = [y for y, c in counts.items() if c > 1]
        if not deficits and not surpluses:
            break
        if not deficits or not surpluses:
            return None, "No se puede balancear a 1-1-1"

        deficit_year = deficits[0]
        surplus_year = surpluses[0]
        candidate = pick_oldest(surplus_year)
        if candidate is None:
            return None, "No se puede balancear a 1-1-1"

        candidate["TARGET_YEAR"] = deficit_year
        counts[surplus_year] -= 1
        counts[deficit_year] += 1

    if any(counts[y] != 1 for y in years):
        return None, "No se puede balancear a 1-1-1"

    return records, None
