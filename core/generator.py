from __future__ import annotations

import re
import tempfile
from io import BytesIO
from pathlib import Path

import yaml
from openpyxl import load_workbook

from .io import (
    read_csv,
    normalize_str,
    extract_course_code,
    validate_required_columns,
    parse_grade,
    parse_year,
)
from .model_indexer import build_model_index
from .electives import balance_electives, parse_period
from .grades import grade_to_text
from .logger import LogCollector
from .zipout import zip_directory


INVALID_FILENAME_CHARS = r"\\/:*?\"<>|"


def _load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def _is_missing(value: object) -> bool:
    return value is None or str(value).strip() == ""


def _sanitize_filename(text: str) -> str:
    cleaned = re.sub("[{0}]".format(re.escape(INVALID_FILENAME_CHARS)), "", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def _is_valid_dni(dni: str) -> bool:
    return bool(dni) and dni.isdigit()


def _normalize_col_name(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).strip().upper())


def _normalize_spaces(text: str) -> str:
    return " ".join(str(text or "").strip().split())


def _split_words(text: str) -> list[str]:
    return [w for w in _normalize_spaces(text).split(" ") if w]


def _is_roman_numeral(token: str) -> bool:
    token = token.upper()
    if not token:
        return False
    return bool(re.fullmatch(r"M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})", token))


def _format_word(token: str) -> str:
    if not token:
        return token
    if token.isdigit():
        return token
    if _is_roman_numeral(token):
        return token.upper()
    if token.isupper() and len(token) <= 3:
        return token.upper()
    return token[0].upper() + token[1:].lower()


def _format_course_name(name: str) -> str:
    name = _normalize_spaces(name)
    if not name:
        return ""
    parts = re.split(r"([\\s\\-\\/])", name)
    formatted = []
    for part in parts:
        if part in {" ", "-", "/"}:
            formatted.append(part)
            continue
        formatted.append(_format_word(part))
    return "".join(formatted)


def _format_simple_person_name(name: str) -> tuple[str | None, str | None]:
    name = _normalize_spaces(name)
    if not name:
        return None, "Nombre vacío"
    return name.upper(), None


def _load_name_overrides(config: dict, config_path: str) -> tuple[dict[str, int], Path | None]:
    overrides_cfg = config.get("name_overrides") or {}
    path = overrides_cfg.get("path") or "name_overrides.csv"
    overrides_path = Path(path)
    if not overrides_path.is_absolute():
        overrides_path = Path(config_path).parent / overrides_path
    if not overrides_path.exists():
        return {}, overrides_path

    df = read_csv(overrides_path)
    if df.empty:
        return {}, overrides_path

    dni_col = _resolve_column(df, overrides_cfg.get("dni_column"), ["DNI"])
    comma_col = _resolve_column(
        df, overrides_cfg.get("comma_after_column"), ["COMMA_AFTER", "COMA_DESPUES"]
    )
    if not dni_col or not comma_col:
        return {}, overrides_path

    mapping: dict[str, int] = {}
    for rec in df.to_dict("records"):
        dni = str(rec.get(dni_col) or "").strip()
        if not dni:
            continue
        try:
            comma_after = int(str(rec.get(comma_col)).strip())
        except Exception:
            continue
        mapping[dni] = comma_after
    return mapping, overrides_path


def _write_name_overrides(path: Path, rows: list[dict]) -> None:
    if not rows:
        return
    import pandas as pd

    existing = None
    if path.exists():
        try:
            existing = read_csv(path)
        except Exception:
            existing = None

    df_new = pd.DataFrame(rows, columns=["DNI", "NOMBRE_COMPLETO", "COMMA_AFTER"])
    if existing is not None and not existing.empty:
        merged = pd.concat([existing, df_new], ignore_index=True)
        merged = merged.drop_duplicates(subset=["DNI"], keep="first")
        merged.to_csv(path, index=False)
        return
    df_new.to_csv(path, index=False)


def _format_person_name(
    name: str,
    dni: str,
    overrides: dict[str, int],
    ambiguous_out: list[dict],
) -> tuple[str | None, str | None]:
    name = _normalize_spaces(name)
    if not name:
        return None, "Nombre vacío"

    words = _split_words(name)
    upper_name = name.upper()
    has_de_la = "DE LA" in upper_name
    word_count = len(words)

    override_after = overrides.get(dni)
    if override_after:
        if 1 <= override_after < word_count:
            left = " ".join(words[:override_after])
            right = " ".join(words[override_after:])
            return "{0}, {1}".format(left, right), None
        return None, "COMMA_AFTER inválido para DNI: {0}".format(dni)

    if has_de_la or word_count not in (3, 4):
        ambiguous_out.append(
            {"DNI": dni, "NOMBRE_COMPLETO": name, "COMMA_AFTER": ""}
        )
        return None, "Nombre requiere coma manual: {0}".format(name)

    split_after = 1 if word_count == 3 else 2
    left = " ".join(words[:split_after])
    right = " ".join(words[split_after:])
    return "{0}, {1}".format(left, right), None


def _resolve_column(df, preferred: str | None, aliases: list[str]) -> str | None:
    if preferred:
        target = _normalize_col_name(preferred)
        for col in df.columns:
            if _normalize_col_name(col) == target:
                return col
    for alias in aliases:
        target = _normalize_col_name(alias)
        for col in df.columns:
            if _normalize_col_name(col) == target:
                return col
    return None


def _load_course_metadata(config: dict, config_path: str):
    meta_cfg = config.get("course_metadata") or {}
    path = meta_cfg.get("path")
    if not path:
        return {}

    meta_path = Path(path)
    if not meta_path.is_absolute():
        meta_path = Path(config_path).parent / meta_path

    df = read_csv(meta_path)
    if df.empty:
        return {}

    code_col = _resolve_column(df, meta_cfg.get("code_column"), ["CODIGO_CURSO", "CODIGO CURSO", "CODIGO"])
    res_col = _resolve_column(df, meta_cfg.get("resolucion_column"), ["RESOLUCION", "RESOLUCIÓN"])
    x_col = _resolve_column(df, meta_cfg.get("x_column"), ["X"])
    y_col = _resolve_column(df, meta_cfg.get("y_column"), ["Y"])
    z_col = _resolve_column(df, meta_cfg.get("z_column"), ["Z"])

    missing = []
    if not code_col:
        missing.append("CODIGO_CURSO")
    if not res_col:
        missing.append("RESOLUCION")
    if not x_col:
        missing.append("X")
    if not y_col:
        missing.append("Y")
    if not z_col:
        missing.append("Z")
    if missing:
        raise ValueError(
            "Faltan columnas en course_metadata: {0}".format(", ".join(missing))
        )

    mapping: dict[str, dict[str, str]] = {}
    for rec in df.to_dict("records"):
        code = extract_course_code(rec.get(code_col))
        if not code:
            continue
        mapping[code] = {
            "resolucion": str(rec.get(res_col) or "").strip(),
            "x": str(rec.get(x_col) or "").strip(),
            "y": str(rec.get(y_col) or "").strip(),
            "z": str(rec.get(z_col) or "").strip(),
        }

    return mapping


def _get_course_columns(sheet_cfg: dict) -> dict[str, str]:
    cols = sheet_cfg.get("course_columns") or {}

    def _norm_col(value: object) -> str:
        if value is None:
            return ""
        text = str(value).strip().upper()
        return text if text and text != "NONE" else ""

    return {
        "code_col": _norm_col(cols.get("code_col", "B")),
        "curso_col": _norm_col(cols.get("curso_col", "")),
        "nota_col": _norm_col(cols.get("nota_col", "E")),
        "nota_texto_col": _norm_col(cols.get("nota_texto_col", "F")),
        "creditos_col": _norm_col(cols.get("creditos_col", "")),
        "acta_col": _norm_col(cols.get("acta_col", "")),
        "resolucion_col": _norm_col(cols.get("resolucion_col", "")),
        "x_col": _norm_col(cols.get("x_col", "")),
        "y_col": _norm_col(cols.get("y_col", "")),
        "z_col": _norm_col(cols.get("z_col", "")),
    }


def _select_sheet(year_int: int | None, cara_sheet: str, sello_sheet: str) -> str | None:
    if year_int is None:
        return None
    if 1 <= year_int <= 5:
        return cara_sheet
    if 6 <= year_int <= 10:
        return sello_sheet
    return None


def _write_identity(ws, identity, header_cells):
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

    if not header_cells:
        raise ValueError("No se configuraron celdas de cabecera")

    required = {
        "dni": "DNI",
        "nombre": "NOMBRE",
        "facultad": "FACULTAD",
    }

    positions = {}
    missing = []

    for key, identity_key in required.items():
        coord = header_cells.get(key)
        if not coord:
            missing.append(key.upper())
            continue
        try:
            col_letter, row = coordinate_from_string(coord)
            col = column_index_from_string(col_letter)
            positions[identity_key] = (row, col)
        except Exception:
            missing.append(key.upper())

    program_coord = header_cells.get("escuela") or header_cells.get("programa")
    if not program_coord:
        missing.append("ESCUELA")
    else:
        try:
            col_letter, row = coordinate_from_string(program_coord)
            col = column_index_from_string(col_letter)
            positions["PROGRAMA"] = (row, col)
        except Exception:
            missing.append("ESCUELA")

    if missing:
        raise ValueError("No se pudo ubicar campos de cabecera: {0}".format(", ".join(missing)))

    for key, (row, col) in positions.items():
        ws.cell(row=row, column=col).value = identity[key]


def generate_certificates(
    template_path,
    csv_path,
    faculty,
    program,
    config_path="config.yml",
    progress_cb=None,
):
    config = _load_config(config_path)
    optional_codes = {
        extract_course_code(code)
        for code in (config.get("optional_courses") or [])
        if str(code).strip()
    }
    course_metadata = _load_course_metadata(config, config_path)
    name_overrides, name_overrides_path = _load_name_overrides(config, config_path)
    df = read_csv(csv_path)

    missing_cols = validate_required_columns(df.columns)
    if missing_cols:
        raise ValueError("Faltan columnas requeridas: {0}".format(", ".join(missing_cols)))

    df = df.copy()
    df["DNI"] = df["DNI"].fillna("").astype(str).str.strip()
    df["CODIGO_CURSO"] = df["CODIGO"].apply(extract_course_code)
    df["ESTADO_N"] = df["ESTADO"].apply(normalize_str)
    df["TIPO_CURSO_N"] = df["TIPO_CURSO"].apply(normalize_str)
    df["NOMBRE_COMPLETO"] = df["NOMBRE_COMPLETO"].fillna("").astype(str).str.strip()
    df["CODIGO_ALUMNO"] = df["CODIGO_ALUMNO"].fillna("").astype(str).str.strip()

    df = df[df["ESTADO_N"] == "APROBADO"]

    cara_cfg = config.get("cara", {})
    sello_cfg = config.get("sello", {})
    cara_sheet = str(cara_cfg.get("sheet_name", "CARA")).strip() or "CARA"
    sello_sheet = str(sello_cfg.get("sheet_name", "SELLO")).strip() or "SELLO"
    cara_cols = _get_course_columns(cara_cfg)
    sello_cols = _get_course_columns(sello_cfg)

    from openpyxl.utils.cell import column_index_from_string, get_column_letter

    def _col_index_required(col_letter: str, label: str) -> int:
        if not col_letter:
            raise ValueError("Columna faltante en configuración: {0}".format(label))
        return column_index_from_string(col_letter)

    def _col_index_optional(col_letter: str) -> int | None:
        if not col_letter:
            return None
        return column_index_from_string(col_letter)

    columns_by_sheet = {
        cara_sheet: {
            "code_col": _col_index_required(cara_cols["code_col"], "cara.course_columns.code_col"),
            "curso_col": _col_index_required(cara_cols["curso_col"], "cara.course_columns.curso_col"),
            "nota_col": _col_index_required(cara_cols["nota_col"], "cara.course_columns.nota_col"),
            "nota_texto_col": _col_index_required(
                cara_cols["nota_texto_col"], "cara.course_columns.nota_texto_col"
            ),
            "creditos_col": _col_index_optional(cara_cols["creditos_col"]),
            "acta_col": _col_index_optional(cara_cols["acta_col"]),
            "resolucion_col": _col_index_optional(cara_cols["resolucion_col"]),
            "x_col": _col_index_optional(cara_cols["x_col"]),
            "y_col": _col_index_optional(cara_cols["y_col"]),
            "z_col": _col_index_optional(cara_cols["z_col"]),
        },
        sello_sheet: {
            "code_col": _col_index_required(sello_cols["code_col"], "sello.course_columns.code_col"),
            "curso_col": _col_index_required(sello_cols["curso_col"], "sello.course_columns.curso_col"),
            "nota_col": _col_index_required(sello_cols["nota_col"], "sello.course_columns.nota_col"),
            "nota_texto_col": _col_index_required(
                sello_cols["nota_texto_col"], "sello.course_columns.nota_texto_col"
            ),
            "creditos_col": _col_index_optional(sello_cols["creditos_col"]),
            "acta_col": _col_index_optional(sello_cols["acta_col"]),
            "resolucion_col": _col_index_optional(sello_cols["resolucion_col"]),
            "x_col": _col_index_optional(sello_cols["x_col"]),
            "y_col": _col_index_optional(sello_cols["y_col"]),
            "z_col": _col_index_optional(sello_cols["z_col"]),
        },
    }

    def _validate_column_overlaps(sheet_name: str, cols: dict[str, int | None]) -> None:
        used: dict[int, str] = {}
        for key, col_idx in cols.items():
            if col_idx is None:
                continue
            prev = used.get(col_idx)
            if prev and prev != key:
                col_letter = get_column_letter(col_idx)
                raise ValueError(
                    "Columnas en conflicto en {0}: {1} y {2} usan {3}".format(
                        sheet_name, prev, key, col_letter
                    )
                )
            used[col_idx] = key

    _validate_column_overlaps(cara_sheet, columns_by_sheet[cara_sheet])
    _validate_column_overlaps(sello_sheet, columns_by_sheet[sello_sheet])

    slots_by_sheet = build_model_index(template_path, config)

    electives_cfg = config.get("electives", {})
    elective_years = set(electives_cfg.get("years", [6, 7, 8]))

    with open(template_path, "rb") as f:
        template_bytes = f.read()

    out_dir = Path(tempfile.mkdtemp())
    cert_dir = out_dir / "certificados"
    cert_dir.mkdir(parents=True, exist_ok=True)

    logger = LogCollector()

    if df.empty:
        log_path = out_dir / "log.xlsx"
        log_df = logger.to_excel(log_path)
        zip_path = out_dir / "certificados.zip"
        zip_directory(out_dir, zip_path)
        return zip_path, log_df

    total_students = int(df["DNI"].nunique())
    processed_students = 0

    ambiguous_names: list[dict] = []

    for dni, grp in df.groupby("DNI"):
        errors = []
        dni_str = str(dni).strip()

        codigo_alumno = ""
        nombre = ""
        nombre_formateado = ""
        if not grp.empty:
            codigo_alumno = str(grp["CODIGO_ALUMNO"].iloc[0]).strip()
            nombre = str(grp["NOMBRE_COMPLETO"].iloc[0]).strip()
            # Validación de comas deshabilitada temporalmente (ver _format_person_name).
            # nombre_formateado, nombre_error = _format_person_name(
            #     nombre, dni_str, name_overrides, ambiguous_names
            # )
            # if nombre_error:
            #     errors.append(nombre_error)
            nombre_formateado, nombre_error = _format_simple_person_name(nombre)
            if nombre_error:
                errors.append(nombre_error)

        if not _is_valid_dni(dni_str):
            errors.append("DNI vacío o no numérico")

        required_fields = [
            "DNI",
            "NOMBRE_COMPLETO",
            "CODIGO",
            "CURSO",
            "NOTA",
            "YEAR",
            "TIPO_CURSO",
            "PERIODO",
            "ESTADO",
            "CODIGO_ALUMNO",
        ]
        for field in required_fields:
            if grp[field].isna().any() or (grp[field].astype(str).str.strip() == "").any():
                errors.append("Campo faltante: {0}".format(field))

        records = grp.to_dict("records")
        for rec in records:
            rec["CODIGO_CURSO"] = extract_course_code(rec.get("CODIGO"))
            rec["GRADE"] = parse_grade(rec.get("NOTA"))
            rec["YEAR_INT"] = parse_year(rec.get("YEAR"))
            rec["CURSO_NOMBRE"] = str(rec.get("CURSO") or "").strip()
            rec["CREDITOS"] = str(rec.get("CREDITOS") or "").strip()
            rec["ACTA"] = str(rec.get("ACTA") or "").strip()
            meta = course_metadata.get(rec["CODIGO_CURSO"], {})
            rec["RESOLUCION"] = str(meta.get("resolucion") or "").strip()
            rec["X"] = str(meta.get("x") or "").strip()
            rec["Y"] = str(meta.get("y") or "").strip()
            rec["Z"] = str(meta.get("z") or "").strip()

            if _is_missing(rec.get("CODIGO_CURSO")):
                errors.append("Código de curso vacío")
            if rec.get("GRADE") is None:
                errors.append("Nota inválida")
            if rec.get("YEAR_INT") is None:
                errors.append("YEAR inválido")
            if _is_missing(rec.get("CURSO_NOMBRE")):
                errors.append("Curso vacío")
            if _is_missing(rec.get("RESOLUCION")) or _is_missing(rec.get("X")) or _is_missing(rec.get("Y")) or _is_missing(
                rec.get("Z")
            ):
                errors.append(
                    "Sin configuración para curso: {0}".format(rec.get("CODIGO_CURSO"))
                )

        def _deduplicate_records(items: list[dict]) -> list[dict]:
            best: dict[tuple[int | None, str], dict] = {}
            for rec in items:
                year = rec.get("YEAR_INT")
                code = rec.get("CODIGO_CURSO") or ""
                period = parse_period(rec.get("PERIODO"))
                grade = rec.get("GRADE")
                type_rank = 1 if normalize_str(rec.get("TIPO_CURSO")) == "ELECTIVO" else 0
                rank = (
                    type_rank,
                    period if period is not None else -1,
                    grade if grade is not None else -1,
                )
                key = (year, code)
                current = best.get(key)
                if current is None or rank > current["_rank"]:
                    rec["_rank"] = rank
                    best[key] = rec
            deduped = list(best.values())
            for rec in deduped:
                rec.pop("_rank", None)
            return deduped

        if errors:
            reason = "; ".join(sorted(set(errors)))
            logger.add(dni_str, codigo_alumno, nombre, "ERROR", reason, "")
            processed_students += 1
            if progress_cb:
                progress_cb(processed_students, total_students)
            continue

        records = _deduplicate_records(records)

        electives = [
            rec
            for rec in records
            if normalize_str(rec.get("TIPO_CURSO")) == "ELECTIVO"
            and rec.get("YEAR_INT") in elective_years
        ]
        if len(electives) != 3:
            logger.add(
                dni_str,
                codigo_alumno,
                nombre,
                "ERROR",
                "El alumno no tiene exactamente 3 electivos aprobados",
                "",
            )
            processed_students += 1
            if progress_cb:
                progress_cb(processed_students, total_students)
            continue

        balanced, balance_error = balance_electives(electives, years=elective_years)
        if balance_error:
            logger.add(dni_str, codigo_alumno, nombre, "ERROR", balance_error, "")
            processed_students += 1
            if progress_cb:
                progress_cb(processed_students, total_students)
            continue

        for rec in balanced:
            rec["YEAR_INT"] = rec.get("TARGET_YEAR")

        # Los slots configurados se mantienen en config.yml, pero no se usan para escribir.

        missing_reasons = []
        records_by_year: dict[int, list[dict]] = {}
        for rec in records:
            year = rec.get("YEAR_INT")
            target_sheet = _select_sheet(year, cara_sheet, sello_sheet)
            if not target_sheet:
                missing_reasons.append("YEAR fuera de rango (1-10): {0}".format(year))
                continue
            records_by_year.setdefault(year, []).append(rec)

        for year, year_records in records_by_year.items():
            target_sheet = _select_sheet(year, cara_sheet, sello_sheet)
            if not target_sheet:
                continue
            sheet_slots = slots_by_sheet.get(target_sheet, {})
            rows = sheet_slots.get(year)
            if not rows:
                missing_reasons.append(
                    "No hay filas para YEAR={0} en {1}".format(year, target_sheet)
                )
                continue
            if len(year_records) > len(rows):
                missing_reasons.append(
                    "No hay suficientes filas en {0} para YEAR={1}: {2} cursos, {3} filas".format(
                        target_sheet, year, len(year_records), len(rows)
                    )
                )

        if missing_reasons:
            reason = "; ".join(sorted(set(missing_reasons)))
            logger.add(dni_str, codigo_alumno, nombre, "ERROR", reason, "")
            processed_students += 1
            if progress_cb:
                progress_cb(processed_students, total_students)
            continue

        wb = load_workbook(BytesIO(template_bytes))
        identity = {
            "DNI": dni_str,
            "NOMBRE": nombre_formateado or nombre,
            "FACULTAD": faculty,
            "PROGRAMA": program,
        }

        try:
            if cara_sheet not in wb.sheetnames:
                raise ValueError("La hoja '{0}' no existe en la plantilla".format(cara_sheet))
            ws = wb[cara_sheet]
            _write_identity(ws, identity, cara_cfg.get("header_cells", {}))
        except Exception as exc:
            logger.add(dni_str, codigo_alumno, nombre, "ERROR", str(exc), "")
            processed_students += 1
            if progress_cb:
                progress_cb(processed_students, total_students)
            continue

        def _record_sort_key(rec: dict):
            period = parse_period(rec.get("PERIODO"))
            code = rec.get("CODIGO_CURSO") or ""
            return (period is None, period or 0, code)

        for year, year_records in records_by_year.items():
            target_sheet = _select_sheet(year, cara_sheet, sello_sheet)
            if not target_sheet:
                continue
            rows = slots_by_sheet.get(target_sheet, {}).get(year, [])
            if not rows:
                continue

            ws = wb[target_sheet]
            cols = columns_by_sheet[target_sheet]
            sorted_records = sorted(year_records, key=_record_sort_key)
            records_by_code = {rec["CODIGO_CURSO"]: rec for rec in sorted_records}
            used_codes: set[str] = set()

            idx = 0
            for row in rows:
                template_code = extract_course_code(ws.cell(row=row, column=cols["code_col"]).value)
                if template_code and template_code in optional_codes:
                    rec = records_by_code.get(template_code)
                    if rec:
                        used_codes.add(template_code)
                        grade = rec["GRADE"]
                        grade_text = grade_to_text(grade)
                        ws.cell(row=row, column=cols["code_col"]).value = rec["CODIGO_CURSO"]
                        ws.cell(row=row, column=cols["curso_col"]).value = _format_course_name(
                            rec.get("CURSO_NOMBRE", "")
                        )
                        ws.cell(row=row, column=cols["nota_col"]).value = grade
                        ws.cell(row=row, column=cols["nota_texto_col"]).value = grade_text
                        if cols.get("creditos_col"):
                            ws.cell(row=row, column=cols["creditos_col"]).value = rec.get("CREDITOS", "")
                        if cols.get("acta_col"):
                            ws.cell(row=row, column=cols["acta_col"]).value = rec.get("ACTA", "")
                        if cols.get("resolucion_col"):
                            ws.cell(row=row, column=cols["resolucion_col"]).value = rec.get(
                                "RESOLUCION", ""
                            )
                        if cols.get("x_col"):
                            ws.cell(row=row, column=cols["x_col"]).value = rec.get("X", "")
                        if cols.get("y_col"):
                            ws.cell(row=row, column=cols["y_col"]).value = rec.get("Y", "")
                        if cols.get("z_col"):
                            ws.cell(row=row, column=cols["z_col"]).value = rec.get("Z", "")
                    else:
                        ws.cell(row=row, column=cols["nota_col"]).value = None
                        ws.cell(row=row, column=cols["nota_texto_col"]).value = None
                    continue

                while idx < len(sorted_records) and sorted_records[idx]["CODIGO_CURSO"] in used_codes:
                    idx += 1
                if idx >= len(sorted_records):
                    break
                rec = sorted_records[idx]
                idx += 1
                used_codes.add(rec["CODIGO_CURSO"])
                grade = rec["GRADE"]
                grade_text = grade_to_text(grade)

                ws.cell(row=row, column=cols["code_col"]).value = rec["CODIGO_CURSO"]
                ws.cell(row=row, column=cols["curso_col"]).value = _format_course_name(
                    rec.get("CURSO_NOMBRE", "")
                )
                ws.cell(row=row, column=cols["nota_col"]).value = grade
                ws.cell(row=row, column=cols["nota_texto_col"]).value = grade_text
                if cols.get("creditos_col"):
                    ws.cell(row=row, column=cols["creditos_col"]).value = rec.get("CREDITOS", "")
                if cols.get("acta_col"):
                    ws.cell(row=row, column=cols["acta_col"]).value = rec.get("ACTA", "")
                if cols.get("resolucion_col"):
                    ws.cell(row=row, column=cols["resolucion_col"]).value = rec.get("RESOLUCION", "")
                if cols.get("x_col"):
                    ws.cell(row=row, column=cols["x_col"]).value = rec.get("X", "")
                if cols.get("y_col"):
                    ws.cell(row=row, column=cols["y_col"]).value = rec.get("Y", "")
                if cols.get("z_col"):
                    ws.cell(row=row, column=cols["z_col"]).value = rec.get("Z", "")

        safe_name = _sanitize_filename("{0} {1} - {2}".format(dni_str, codigo_alumno, nombre))
        if not safe_name:
            safe_name = dni_str or "alumno"

        out_path = cert_dir / "{0}.xlsx".format(safe_name)
        wb.save(out_path)

        logger.add(dni_str, codigo_alumno, nombre, "OK", "", out_path.name)
        processed_students += 1
        if progress_cb:
            progress_cb(processed_students, total_students)

    log_path = out_dir / "log.xlsx"
    log_df = logger.to_excel(log_path)

    zip_path = out_dir / "certificados.zip"
    zip_directory(out_dir, zip_path)

    if ambiguous_names and name_overrides_path:
        _write_name_overrides(name_overrides_path, ambiguous_names)

    return zip_path, log_df
