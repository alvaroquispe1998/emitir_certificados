from core.grades import grade_to_text


def test_grade_to_text():
    assert grade_to_text(0) == "CERO"
    assert grade_to_text(1) == "UNO"
    assert grade_to_text(20) == "VEINTE"
