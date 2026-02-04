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


def _parse_correlative(record: dict) -> int | None:
    for key in ("CORRELATIVO", "Z"):
        raw = record.get(key)
        if raw is None:
            continue
        text = str(raw).strip()
        if text.isdigit():
            return int(text)

    code = str(record.get("CODIGO_CURSO") or record.get("CODIGO") or "").strip().upper()
    match = re.search(r"AA(\d+)$", code)
    if not match:
        return None
    try:
        num = int(match.group(1))
    except ValueError:
        return None
    if 1 <= num <= 6:
        return 48 + num
    return None


def balance_electives(records: list[dict], years: set[int] | None = None):
    if years is None:
        years = {6, 7, 8}

    if len(records) != 3:
        return None, "El alumno no tiene exactamente 3 electivos aprobados"

    target_years = sorted(years)
    if len(target_years) != 3:
        return None, "No se puede balancear a 1-1-1"

    for idx, rec in enumerate(records):
        rec["_idx"] = idx
        rec["TARGET_YEAR"] = rec.get("YEAR_INT")

    def sort_key(rec):
        correlative = _parse_correlative(rec)
        period = parse_period(rec.get("PERIODO"))
        return (
            correlative is None,
            correlative or 0,
            period is None,
            period or 0,
            rec.get("_idx", 0),
        )

    ordered = sorted(records, key=sort_key)
    for rec, target_year in zip(ordered, target_years):
        rec["TARGET_YEAR"] = target_year

    return records, None
