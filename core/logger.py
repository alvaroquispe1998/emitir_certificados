import pandas as pd

LOG_COLUMNS = ["DNI", "CODIGO_ALUMNO", "NOMBRE", "STATUS", "REASON", "OUTPUT_FILE"]


class LogCollector:
    def __init__(self):
        self.rows = []

    def add(self, dni, codigo_alumno, nombre, status, reason="", output_file=""):
        self.rows.append(
            {
                "DNI": dni,
                "CODIGO_ALUMNO": codigo_alumno,
                "NOMBRE": nombre,
                "STATUS": status,
                "REASON": reason,
                "OUTPUT_FILE": output_file,
            }
        )

    def to_dataframe(self):
        return pd.DataFrame(self.rows, columns=LOG_COLUMNS)

    def to_excel(self, path):
        df = self.to_dataframe()
        df.to_excel(path, index=False)
        return df
