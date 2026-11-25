#Para github+streamlit

from pathlib import Path
import tempfile
import re
import io

import pandas as pd
import streamlit as st

from docx import Document
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph


# =====================================================
# CAMPOS DEL MODELO ANECA
# =====================================================

SECCIONES = {
    "A. Breve resumen": "Resumen",
    "Contribución al progreso del conocimiento en el área (10%)": "Conocimiento",
    "Impacto científico (60%)": "Impacto científico",
    "Impacto social (10%)": "Impacto social",
    "Contribución a la ciencia abierta (10%)": "Ciencia abierta"
}


CAMPOS = {
    "Objetivos perseguidos": "Resumen - Objetivos perseguidos",
    "Resultados más relevantes": "Resumen - Resultados más relevantes",
    "Liderazgo": "Resumen - Liderazgo",

    "Grado de originalidad e innovación":
        "Conocimiento - Grado de originalidad e innovación",

    "Prioridad temática":
        "Conocimiento - Prioridad temática",

    "Aportación metodológica":
        "Conocimiento - Aportación metodológica",

    "Uso aportación":
        "Impacto científico - Uso aportación",

    "Citas recibidas por la aportación":
        "Impacto científico - Citas recibidas por la aportación",

    "Calidad del medio de publicación/difusión/exhibición":
        "Impacto científico - Calidad del medio",

    "Menciones y referencias no académicas de la aportación":
        "Impacto social - Menciones y referencias no académicas",

    "Grado de apertura":
        "Ciencia abierta - Grado de apertura",

    "Conjunto de datos":
        "Ciencia abierta - Conjunto de datos",

    "Software":
        "Ciencia abierta - Software",
}


COLUMNAS_BASE = [
    "Aportación",
    "Tipo",
    "Año",
    "Título",
    "Revista",
    "ID Persistente",
    "REF. APA",
    "Fuente iMarina",
    "Fuentes consultadas"
]


# =====================================================
# UTILIDADES
# =====================================================

def limpiar(txt):

    if txt is None:
        return ""

    txt = str(txt)

    txt = txt.replace("\r", "\n")
    txt = re.sub(r"\n+", "\n", txt)
    txt = re.sub(r"[ \t]+", " ", txt)

    return txt.strip()


def iter_block_items(parent):

    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._element

    for child in parent_elm.iterchildren():

        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)

        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


# =====================================================
# NUEVA APORTACIÓN
# =====================================================

def nueva_aportacion():

    registro = {}

    for c in COLUMNAS_BASE:
        registro[c] = ""

    for c in CAMPOS.values():
        registro[c] = ""

    registro["Impacto científico - Otros indicios"] = ""
    registro["Impacto social - Otros indicios"] = ""
    registro["Ciencia abierta - Otros indicios"] = ""

    return registro


# =====================================================
# EXTRACTOR
# =====================================================

def extraer_aportaciones(docx_path):

    doc = Document(docx_path)

    aportaciones = []

    actual = None
    campo_actual = None

    contador_otros_indicios = 0

    for bloque in iter_block_items(doc):

        # TABLAS

        if isinstance(bloque, Table):

            if actual is None:
                continue

            for fila in bloque.rows:

                celdas = [limpiar(c.text) for c in fila.cells]
                celdas = [c for c in celdas if c]

                if not celdas:
                    continue

                if len(celdas) >= 2:

                    clave = celdas[0].replace(":", "")
                    valor = " ".join(celdas[1:])

                    if clave in COLUMNAS_BASE:
                        actual[clave] = valor

        # PÁRRAFOS

        elif isinstance(bloque, Paragraph):

            texto = limpiar(bloque.text)

            if not texto:
                continue

            if re.match(
                r"^ANEXO\s+[IVX0-9]+$",
                texto,
                flags=re.I
            ):
                if actual:
                    aportaciones.append(actual)
                actual = None
                break

            if texto.startswith(
                "Se deben describir de forma clara y concisa"
            ):
                break

            if re.match(
                r"^Aportación\s+\d+",
                texto,
                flags=re.I
            ):

                if actual:
                    aportaciones.append(actual)

                actual = nueva_aportacion()
                actual["Aportación"] = texto

                campo_actual = None
                contador_otros_indicios = 0

                continue

            if actual is None:
                continue

            if texto in [
                "Artículo",
                "Libro",
                "Capítulo de libro",
                "Congreso",
                "Patente"
            ]:

                actual["Tipo"] = texto
                continue

            if texto in CAMPOS:

                campo_actual = CAMPOS[texto]
                continue

            if texto == "Otros indicios":

                contador_otros_indicios += 1

                if contador_otros_indicios == 1:
                    campo_actual = "Impacto científico - Otros indicios"

                elif contador_otros_indicios == 2:
                    campo_actual = "Impacto social - Otros indicios"

                else:
                    campo_actual = "Ciencia abierta - Otros indicios"

                continue

            if campo_actual:

                if texto.upper().startswith("ANEXO"):
                    continue

                actual[campo_actual] = (
                    actual.get(campo_actual, "")
                    + ("\n\n" if actual.get(campo_actual, "") else "")
                    + texto
                )

    if actual:
        aportaciones.append(actual)

    for aportacion in aportaciones:

        for campo, valor in aportacion.items():

            if not isinstance(valor, str):
                continue

            pos = valor.upper().find("ANEXO I")

            if pos >= 0:
                aportacion[campo] = valor[:pos].strip()

    return aportaciones


# =====================================================
# EXCEL
# =====================================================

def generar_excel(uploaded_file):

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=".docx"
    ) as tmp:

        tmp.write(uploaded_file.getbuffer())
        ruta_docx = tmp.name

    aportaciones = extraer_aportaciones(ruta_docx)

    if not aportaciones:
        raise Exception(
            "No se han encontrado aportaciones."
        )

    df = pd.DataFrame(aportaciones)

    salida = io.BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Aportaciones"
        )

    salida.seek(0)

    return df, salida


# =====================================================
# STREAMLIT
# =====================================================

st.set_page_config(
    page_title="Extractor ANECA",
    layout="wide"
)

st.title("📚 Extractor ANECA / Sexenios")

st.markdown(
    """
    Sube un documento Word de aportaciones
    y genera automáticamente un Excel.
    """
)

archivo = st.file_uploader(
    "Documento Word",
    type=["docx"]
)

if archivo is not None:

    if st.button("Procesar"):

        try:

            with st.spinner("Procesando documento..."):

                df, excel_file = generar_excel(
                    archivo
                )

            st.success(
                f"{len(df)} aportaciones exportadas."
            )

            st.dataframe(
                df.head(20),
                use_container_width=True
            )

            st.download_button(
                label="📥 Descargar Excel",
                data=excel_file,
                file_name="aportaciones.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:

            st.error(str(e))