import tempfile
from pathlib import Path

import streamlit as st

from core.generator import generate_certificates

st.set_page_config(page_title="Generador de certificados", layout="wide")
st.title("Generador de certificados")

st.caption("Carga la plantilla y el CSV para generar certificados por alumno.")

template_file = st.file_uploader("Plantilla .xlsx", type=["xlsx"])
data_file = st.file_uploader("Data .csv", type=["csv"])

faculty = st.text_input("Facultad", value="CIENCIAS DE LA SALUD")
program = st.text_input("Programa/Escuela", value="ENFERMERÍA")

if st.button("Generar certificados", type="primary"):
    if not template_file or not data_file:
        st.error("Debes subir la plantilla .xlsx y el CSV.")
    else:
        with st.spinner("Procesando..."):
            progress_bar = st.progress(0.0)
            progress_text = st.empty()

            def _on_progress(current, total):
                if total <= 0:
                    progress_bar.progress(0.0)
                    progress_text.text("Procesando...")
                    return
                ratio = min(max(current / total, 0.0), 1.0)
                progress_bar.progress(ratio)
                progress_text.text("Procesando {0}/{1}".format(current, total))

            with tempfile.TemporaryDirectory() as temp_dir:
                template_path = Path(temp_dir) / "template.xlsx"
                csv_path = Path(temp_dir) / "data.csv"
                template_path.write_bytes(template_file.getbuffer())
                csv_path.write_bytes(data_file.getbuffer())

                try:
                    zip_path, log_df = generate_certificates(
                        template_path=template_path,
                        csv_path=csv_path,
                        faculty=faculty,
                        program=program,
                        progress_cb=_on_progress,
                    )
                except Exception as exc:
                    st.error("Error: {0}".format(exc))
                    st.stop()

                total = int(log_df.shape[0])
                ok = int((log_df["STATUS"] == "OK").sum())
                err = int((log_df["STATUS"] == "ERROR").sum())

                col1, col2, col3 = st.columns(3)
                col1.metric("Alumnos", total)
                col2.metric("OK", ok)
                col3.metric("ERROR", err)

                st.subheader("Log (preview)")
                st.dataframe(log_df.head(200), use_container_width=True)

                zip_bytes = Path(zip_path).read_bytes()
                st.download_button(
                    label="Descargar ZIP",
                    data=zip_bytes,
                    file_name=Path(zip_path).name,
                    mime="application/zip",
                )
