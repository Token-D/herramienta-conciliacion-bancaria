import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations
from dateutil.parser import parse as parse_date
import re
from collections import Counter

# ===============================
# Función para leer extractos según el banco
# ===============================
def leer_extracto_banco(file, banco):
    """
    Lee y normaliza un archivo de extracto bancario según el banco seleccionado.
    Devuelve un DataFrame con al menos las columnas: ['fecha','monto','concepto','numero_movimiento'].
    """
    df = pd.read_excel(file)

    if banco == "Bancolombia":
        # Columna 'VALOR' con separador de miles ","
        df["monto"] = (
            df["VALOR"]
            .astype(str)
            .str.replace(r"[^\d\-\.,]", "", regex=True)  # limpiar símbolos
            .str.replace(",", "", regex=False)  # quitar separadores de miles
            .astype(float)
        )

    elif banco == "Banco de Bogotá":
        # Débitos y Créditos con formato $ 1.234,00
        df["Debitos"] = (
            df["Débitos"]
            .astype(str)
            .str.replace(r"[^\d\-\.,]", "", regex=True)
            .str.replace(".", "", regex=False)  # quitar miles
            .str.replace(",", ".", regex=False)  # convertir decimal
            .astype(float)
        )
        df["Creditos"] = (
            df["Créditos"]
            .astype(str)
            .str.replace(r"[^\d\-\.,]", "", regex=True)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .astype(float)
        )
        df["monto"] = df["Creditos"].fillna(0) - df["Debitos"].fillna(0)

    elif banco == "BBVA":
        # IMPORTE (COP) con formato -1.234,00
        df["monto"] = (
            df["IMPORTE (COP)"]
            .astype(str)
            .str.replace(r"[^\d\-\.,]", "", regex=True)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .astype(float)
        )

    else:
        raise ValueError(f"Banco {banco} no soportado aún")

    # Asegurar columnas mínimas
    if "fecha" not in df.columns:
        df["fecha"] = pd.NaT
    if "concepto" not in df.columns:
        df["concepto"] = ""
    if "numero_movimiento" not in df.columns:
        df["numero_movimiento"] = ""

    return df


# ===============================
# Función para leer libros auxiliares (estructura fija)
# ===============================
def leer_libro_auxiliar(file):
    df = pd.read_excel(file)
    # Se asume que ya viene con columna 'monto' limpia
    return df


# ===============================
# Conciliación (tu lógica original)
# ===============================
def conciliacion_directa(extracto_df, auxiliar_df):
    resultados = []
    extracto_conciliado_idx = set()
    auxiliar_conciliado_idx = set()

    for idx_ext, fila_extracto in extracto_df.iterrows():
        # Buscar coincidencias exactas en el auxiliar
        coincidencias = auxiliar_df[
            (abs(auxiliar_df["monto"] - fila_extracto["monto"]) < 0.01)
            & (~auxiliar_df.index.isin(auxiliar_conciliado_idx))
        ]

        if not coincidencias.empty:
            idx_aux = coincidencias.index[0]
            resultados.append(
                {
                    "extracto_fecha": fila_extracto["fecha"],
                    "extracto_monto": fila_extracto["monto"],
                    "auxiliar_fecha": auxiliar_df.loc[idx_aux, "fecha"],
                    "auxiliar_monto": auxiliar_df.loc[idx_aux, "monto"],
                }
            )
            extracto_conciliado_idx.add(idx_ext)
            auxiliar_conciliado_idx.add(idx_aux)

    return pd.DataFrame(resultados), extracto_conciliado_idx, auxiliar_conciliado_idx


def conciliar_banco_completo(extracto_df, auxiliar_df):
    resultados_directa, extracto_conciliado_idx, auxiliar_conciliado_idx = conciliacion_directa(
        extracto_df, auxiliar_df
    )

    # Marcar conciliados
    extracto_df["conciliado"] = extracto_df.index.isin(extracto_conciliado_idx)
    auxiliar_df["conciliado"] = auxiliar_df.index.isin(auxiliar_conciliado_idx)

    return resultados_directa, extracto_df, auxiliar_df


# ===============================
# Interfaz Streamlit
# ===============================
st.set_page_config(page_title="Conciliación Bancaria", layout="wide")
st.title("📊 Herramienta de Conciliación Bancaria")

# Selección de banco (solo una vez)
banco = st.selectbox(
    "Seleccione el banco del extracto cargado:",
    ["Bancolombia", "Banco de Bogotá", "BBVA"],
)

col1, col2 = st.columns(2)

with col1:
    archivo_extracto = st.file_uploader("Subir extracto bancario", type=["xlsx"])
with col2:
    archivo_auxiliar = st.file_uploader("Subir libro auxiliar", type=["xlsx"])

if archivo_extracto and archivo_auxiliar:
    try:
        extracto_df = leer_extracto_banco(archivo_extracto, banco)
        auxiliar_df = leer_libro_auxiliar(archivo_auxiliar)

        st.success(
            f"✅ Datos cargados correctamente\n\nExtracto bancario: {len(extracto_df)} movimientos\nLibro auxiliar: {len(auxiliar_df)} movimientos"
        )

        resultados_df, extracto_df, auxiliar_df = conciliar_banco_completo(
            extracto_df, auxiliar_df
        )

        st.subheader("Resultados de la conciliación")
        st.dataframe(resultados_df)

        st.subheader("Movimientos del extracto no conciliados")
        st.dataframe(extracto_df[~extracto_df["conciliado"]])

        st.subheader("Movimientos del auxiliar no conciliados")
        st.dataframe(auxiliar_df[~auxiliar_df["conciliado"]])

    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")

else:
    st.info("Por favor, sube ambos archivos para comenzar la conciliación.")
