from core.electives import balance_electives


def test_balance_electives_moves_oldest():
    records = [
        {"YEAR_INT": 7, "PERIODO": "202101"},
        {"YEAR_INT": 7, "PERIODO": "202001"},
        {"YEAR_INT": 8, "PERIODO": "202102"},
    ]
    balanced, error = balance_electives(records, years={6, 7, 8})
    assert error is None
    years = sorted(rec["TARGET_YEAR"] for rec in balanced)
    assert years == [6, 7, 8]
