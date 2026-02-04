import re
import pandas as pd

REQUIRED_COLUMNS = [
    "DNI",
    "NOMBRE_COMPLETO",
    "CODIGO",
    "NOTA",
    "YEAR",
    "TIPO_CURSO",
    "PERIODO",
    "ESTADO",
]


def normalize_str(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip().upper()


def read_csv(path):
    encodings = ["utf-8", "utf-8-sig", "latin-1", "cp1252"]
    last_error = None
    for enc in encodings:
        try:
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding=enc)
            df.columns = [c.strip() for c in df.columns]
            if "ID_ALUMNO" in df.columns and "CODIGO_ALUMNO" not in df.columns:
                df["CODIGO_ALUMNO"] = df["ID_ALUMNO"]
            return df
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    if last_error:
        raise last_error
    raise UnicodeDecodeError("utf-8", b"", 0, 1, "No se pudo leer el CSV con las codificaciones probadas")


def validate_required_columns(columns) -> list[str]:
    cols = {str(c).strip() for c in columns}
    missing = [c for c in REQUIRED_COLUMNS if c not in cols]
    if "CODIGO_ALUMNO" not in cols and "ID_ALUMNO" not in cols:
        missing.append("CODIGO_ALUMNO/ID_ALUMNO")
    return missing


def extract_course_code(value: object) -> str:
    text = normalize_str(value)
    if not text:
        return ""
    if "-" in text:
        return text.split("-")[-1].strip()
    return text


def parse_grade(value: object) -> int | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", ".")
    try:
        num = float(text)
    except ValueError:
        digits = re.findall(r"\d+", text)
        if not digits:
            return None
        num = float(digits[0])
    if abs(num - round(num)) > 1e-6:
        return None
    grade = int(round(num))
    if 0 <= grade <= 20:
        return grade
    return None


def parse_year(value: object) -> int | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    digits = re.findall(r"\d+", text)
    if not digits:
        return None
    year = int(digits[0])
    return year
