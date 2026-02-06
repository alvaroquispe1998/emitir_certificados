from core.electives import balance_electives


def test_balance_electives_orders_by_code():
    records = [
        {"YEAR_INT": 7, "PERIODO": "202301", "CODIGO_CURSO": "21AA4"},
        {"YEAR_INT": 7, "PERIODO": "201901", "CODIGO_CURSO": "21AA1"},
        {"YEAR_INT": 8, "PERIODO": "202102", "CODIGO_CURSO": "21AA6"},
    ]
    balanced, error = balance_electives(records, years={6, 7, 8})
    assert error is None
    mapping = {rec["CODIGO_CURSO"]: rec["TARGET_YEAR"] for rec in balanced}
    assert mapping["21AA1"] == 6
    assert mapping["21AA4"] == 7
    assert mapping["21AA6"] == 8
