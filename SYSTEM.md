# Sistema de Certificados — Descripción por Archivo

Este documento explica qué hace cada archivo relevante y cómo fluye el proceso.

**Flujo general**
1. El usuario sube una plantilla Excel y un CSV desde `app.py`.
2. `core/generator.py` lee el CSV, valida datos, balancea electivos, asigna filas por ciclo y escribe en la plantilla.
3. Se genera un Excel por alumno, un `log.xlsx` con OK/ERROR y un ZIP final con todo.

**Archivos principales**
1. `app.py`
   - Interfaz Streamlit.
   - Carga plantilla/CSV y pide Facultad/Programa.
   - Llama a `generate_certificates(...)` y muestra métricas, log y botón de descarga.
   - Muestra progressbar con avance por alumno.
2. `core/generator.py`
   - Orquesta todo el proceso.
   - Lee `config.yml` y el CSV de alumnos.
   - Filtra `ESTADO == APROBADO`.
   - Deduplica por `(YEAR, CODIGO_CURSO)` quedándose con el registro de periodo más reciente y mayor nota.
   - Balancea electivos 6/7/8 a 1‑1‑1 por alumno.
   - Construye slots por ciclo usando `core/model_indexer.py`.
   - Escribe código/curso/nota/nota texto/creditos/resolucion/x/y/z según columnas del `config.yml`.
   - Registra errores (falta de datos, metadata de curso, filas insuficientes, etc.) en `log.xlsx`.
3. `core/model_indexer.py`
   - Escanea la plantilla y detecta filas “CICLO” en la columna B.
   - Construye el rango de filas disponible por ciclo (slots) para CARA y SELLO.
4. `core/io.py`
   - Lectura robusta del CSV (codificaciones múltiples).
   - Validaciones y normalizaciones.
   - Parseo de nota y año.
5. `core/electives.py`
   - Balancea electivos a 1‑1‑1 en años 6/7/8.
   - Usa `PERIODO` para mover el curso más antiguo cuando hay duplicidad.
6. `core/grades.py`
   - Mapea la nota numérica (0..20) a texto en mayúsculas.
7. `core/logger.py`
   - Acumula resultados por alumno y genera `log.xlsx`.
8. `core/zipout.py`
   - Empaqueta todo en un ZIP final.
   - Evita incluir el ZIP dentro de sí mismo.

**Archivos de configuración y datos**
1. `config.yml`
   - Define hojas, columnas y celdas exactas a escribir.
   - Define ubicación de metadata por curso (`course_metadata`).
2. `course_metadata.csv`
   - Tabla por curso con `RESOLUCION`, `X`, `Y`, `Z`.
   - Si un curso del CSV no existe aquí, se registra error.
3. `modelo certificado.xlsx`
   - Plantilla base.
   - Las filas por ciclo se detectan usando las celdas “CICLO” en la columna B.

**Salidas**
1. `certificados/`
   - Un Excel por alumno válido.
2. `log.xlsx`
   - OK/ERROR y razón por alumno.
3. `certificados.zip`
   - Contiene `certificados/` y `log.xlsx`.

**Reglas clave actuales**
1. Solo cursos `APROBADO`.
2. Electivos 6/7/8 quedan 1‑1‑1.
3. Se escribe solo en columnas configuradas en `config.yml`.
4. Si falta metadata de curso en `course_metadata.csv`, el alumno se marca con ERROR.
