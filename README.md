# Generador de certificados

Mini-sistema para generar certificados por alumno a partir de una plantilla Excel (2 hojas: CARA y SELLO) y un CSV de cursos.

## Requisitos
- Python 3.11+

## Instalación
### Windows (PowerShell)
```bash
py -m venv .venv
.\.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

### Alternativa (si ya tienes venv activo)
```bash
py -m pip install -r requirements.txt
```

## Ejecución
```bash
py -m streamlit run app.py
```

## Flujo general
1. Subir la plantilla `.xlsx`.
2. Subir el CSV de datos.
3. Ajustar Facultad y Programa/Escuela si es necesario.
4. (Opcional) Ingresar la fecha en texto libre.
5. Generar certificados y descargar el ZIP (certificados + log).

## Configuración
El archivo `config.yml` permite ajustar celdas de identidad, offsets de columnas y parámetros de indexado.

Campos principales:
- `sheets`: lista de hojas con offsets por defecto y coordenadas para identidad.
- `labels`: etiquetas a buscar en la plantilla para ubicar DNI, NOMBRE, FACULTAD, PROGRAMA.
- `electives`: años válidos y slots opcionales para electivos.

Si las etiquetas no se encuentran, el sistema usa las coordenadas de `identity` por hoja.

## Salidas
- `certificados/` con un `.xlsx` por alumno válido.
- `log.xlsx` con OK/OBSERVADO/ERROR y motivos.
- ZIP con ambos.

## Notas
- DNI se trata siempre como texto.
- Se filtra `ESTADO == "Aprobado"` (case-insensitive).
- Para el código de curso se usa el sufijo luego del último `-`.
- Si el CSV trae columna `RESOLUCION`/`RESOLUCIÓN` con `SI/NO`:
  - `SI`: la columna resolución queda con su valor normal de metadata.
  - `NO`: en la columna resolución se deja una fórmula `=Xfila&Yfila&Zfila` para concatenar `x+y+z`.
  - Las columnas `x`, `y`, `z` no cambian (`PERIODO`, `21`, `CORRELATIVO`).
- Los cuadros de texto de la plantilla (firmas/cabeceras) se conservan en la salida.
- La fecha manual se aplica al cuadro de texto de fecha del modelo.
