import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations

# Función para identificar columnas
def identificar_columnas(df, columnas_esperadas):
    """
    Identifica las columnas necesarias en un DataFrame basándose en coincidencias parciales.

    Args:
        df (DataFrame): El DataFrame del archivo de Excel.
        columnas_esperadas (dict): Un diccionario con los nombres esperados y sus posibles variantes.

    Returns:
        dict: Un diccionario con las columnas identificadas.
    """
    columnas_identificadas = {}
    for col_esperada, variantes in columnas_esperadas.items():
        for col in df.columns:
            if any(variante.lower() in col.lower() for variante in variantes):
                columnas_identificadas[col_esperada] = col
                break
        else:
            st.error(f"No se encontró una columna que coincida con: {', '.join(variantes)}")
            st.stop()
    return columnas_identificadas

# Función para normalizar un DataFrame
def normalizar_dataframe(df, columnas_identificadas):
    """
    Normaliza un DataFrame para que use los nombres de columnas esperados.

    Args:
        df (DataFrame): El DataFrame original.
        columnas_identificadas (dict): Un diccionario con las columnas identificadas.

    Returns:
        DataFrame: El DataFrame normalizado.
    """
    return df.rename(columns=columnas_identificadas)

# Función para encontrar combinaciones que sumen un monto específico
def encontrar_combinaciones(movimientos, monto_objetivo, tolerancia=0.01):
    combinaciones_validas = []
    for r in range(1, len(movimientos) + 1):
        for combo in combinations(movimientos, r):
            if abs(sum(combo) - monto_objetivo) <= tolerancia:
                combinaciones_validas.append(combo)
    return combinaciones_validas

# Función para realizar la conciliación por agrupación en el libro auxiliar
def conciliacion_agrupacion_libro_auxiliar(extracto_no_conciliado, auxiliar_no_conciliado):
    resultados = []
    for _, movimiento_extracto in extracto_no_conciliado.iterrows():
        monto_objetivo = movimiento_extracto["monto"]
        movimientos_auxiliar = auxiliar_no_conciliado["monto"].tolist()
        combinaciones = encontrar_combinaciones(movimientos_auxiliar, monto_objetivo)

        if combinaciones:
            for combo in combinaciones:
                indices = auxiliar_no_conciliado[auxiliar_no_conciliado["monto"].isin(combo)].index
                resultados.append({
                    "fecha": movimiento_extracto["fecha"],
                    "monto": monto_objetivo,
                    "origen": "Banco",
                    "estado": "Conciliado",
                    "doc. conciliación": ", ".join(auxiliar_no_conciliado.loc[indices, "doc. num"].astype(str)),
                    "tipo agrupación": "Libro Auxiliar"
                })
                auxiliar_no_conciliado.drop(indices, inplace=True)
    return pd.DataFrame(resultados)

# Función principal de conciliación
def conciliar_banco_excel(extracto_df, auxiliar_df):
    resultados_df = pd.DataFrame()

    # 1. Conciliación Directa (Uno a Uno)
    resultados_directa = pd.merge(
        extracto_df, auxiliar_df, on=["fecha", "monto"], how="outer", suffixes=("_banco", "_auxiliar")
    )
    resultados_directa["origen"] = resultados_directa.apply(
        lambda row: "Banco" if pd.notna(row["concepto"]) else "Libro Auxiliar", axis=1
    )
    resultados_directa["estado"] = resultados_directa.apply(
        lambda row: "Conciliado" if pd.notna(row["concepto"]) and pd.notna(row["nota"]) else "No Conciliado", axis=1
    )
    resultados_directa["doc. conciliación"] = resultados_directa.apply(
        lambda row: row["doc. num_auxiliar"] if pd.notna(row["concepto"]) and pd.notna(row["nota"]) else row["doc. num_banco"] if pd.notna(row["concepto"]) else None, axis=1
    )
    resultados_df = pd.concat([resultados_df, resultados_directa], ignore_index=True)

    # 2. Conciliación por Agrupación en el Libro Auxiliar
    extracto_no_conciliado = resultados_df[resultados_df["estado"] == "No Conciliado"]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df["doc. num"].isin(resultados_df["doc. num_auxiliar"])]
    resultados_agrupacion_libro = conciliacion_agrupacion_libro_auxiliar(extracto_no_conciliado, auxiliar_no_conciliado)
    resultados_df = pd.concat([resultados_df, resultados_agrupacion_libro], ignore_index=True)

    return resultados_df

# Interfaz de Streamlit
st.title("Herramienta de Conciliación Bancaria Automática")

# Cargar archivos Excel
extracto_file = st.file_uploader("Subir Extracto Bancario (Excel)", type=["xlsx"])
auxiliar_file = st.file_uploader("Subir Libro Auxiliar (Excel)", type=["xlsx"])

if extracto_file and auxiliar_file:
    try:
        # Leer la primera hoja de ambos archivos
        extracto_df = pd.read_excel(extracto_file, sheet_name=0)
        auxiliar_df = pd.read_excel(auxiliar_file, sheet_name=0)

        # Definir las columnas esperadas y sus posibles variantes
        columnas_esperadas_extracto = {
            "fecha": ["fecha", "date", "fecha de operación"],
            "monto": ["monto", "importe", "valor", "amount"],
            "concepto": ["concepto", "descripción", "observaciones", "concepto banco"]
        }

        columnas_esperadas_auxiliar = {
            "fecha": ["fecha", "date", "fecha de operación"],
            "monto": ["monto", "importe", "valor", "amount"],
            "nota": ["nota", "nota libro auxiliar", "descripción", "observaciones"]
        }

        # Identificar y normalizar las columnas
        columnas_extracto = identificar_columnas(extracto_df, columnas_esperadas_extracto)
        columnas_auxiliar = identificar_columnas(auxiliar_df, columnas_esperadas_auxiliar)

        extracto_df = normalizar_dataframe(extracto_df, columnas_extracto)
        auxiliar_df = normalizar_dataframe(auxiliar_df, columnas_auxiliar)

        # Realizar conciliación
        resultados_df = conciliar_banco_excel(extracto_df, auxiliar_df)

        # Mostrar resultados
        st.subheader("Resultados de la Conciliación")
        st.write(resultados_df)

        # Generar archivo de resultados
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            resultados_df.to_excel(writer, sheet_name="Resultados", index=False)
        output.seek(0)

        # Botón para descargar resultados
        st.download_button(
            label="Descargar Resultados en Excel",
            data=output,
            file_name="resultados_conciliacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")
else:
    st.info("Por favor, sube ambos archivos para comenzar la conciliación.")