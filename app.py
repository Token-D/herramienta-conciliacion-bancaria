import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations

# Función para encontrar combinaciones que sumen un monto específico
def encontrar_combinaciones(movimientos, monto_objetivo, tolerancia=0.01):
    """
    Encuentra combinaciones de movimientos que sumen el monto objetivo.

    Args:
        movimientos (list): Lista de montos de movimientos.
        monto_objetivo (float): Monto que se desea alcanzar.
        tolerancia (float): Tolerancia para diferencias de redondeo.

    Returns:
        list: Lista de combinaciones que suman el monto objetivo.
    """
    combinaciones_validas = []
    for r in range(1, len(movimientos) + 1):
        for combo in combinations(movimientos, r):
            if abs(sum(combo) - monto_objetivo) <= tolerancia:
                combinaciones_validas.append(combo)
    return combinaciones_validas

# Función para realizar la conciliación por agrupación en el libro auxiliar
def conciliacion_agrupacion_libro_auxiliar(extracto_no_conciliado, auxiliar_no_conciliado):
    """
    Realiza la conciliación por agrupación en el libro auxiliar.

    Args:
        extracto_no_conciliado (DataFrame): Movimientos no conciliados del extracto bancario.
        auxiliar_no_conciliado (DataFrame): Movimientos no conciliados del libro auxiliar.

    Returns:
        DataFrame: Resultados de la conciliación por agrupación.
    """
    resultados = []
    for _, movimiento_extracto in extracto_no_conciliado.iterrows():
        monto_objetivo = movimiento_extracto["Monto"]
        movimientos_auxiliar = auxiliar_no_conciliado["Monto"].tolist()
        combinaciones = encontrar_combinaciones(movimientos_auxiliar, monto_objetivo)

        if combinaciones:
            for combo in combinaciones:
                indices = auxiliar_no_conciliado[auxiliar_no_conciliado["Monto"].isin(combo)].index
                resultados.append({
                    "Fecha": movimiento_extracto["Fecha"],
                    "Monto": monto_objetivo,
                    "Origen": "Banco",
                    "Estado": "Conciliado",
                    "Doc. Conciliación": ", ".join(auxiliar_no_conciliado.loc[indices, "Doc. Num"].astype(str)),
                    "Tipo Agrupación": "Libro Auxiliar"
                })
                auxiliar_no_conciliado.drop(indices, inplace=True)  # Eliminar movimientos ya conciliados
    return pd.DataFrame(resultados)

# Función para realizar la conciliación por agrupación en el extracto bancario
def conciliacion_agrupacion_extracto_bancario(extracto_no_conciliado, auxiliar_no_conciliado):
    """
    Realiza la conciliación por agrupación en el extracto bancario.

    Args:
        extracto_no_conciliado (DataFrame): Movimientos no conciliados del extracto bancario.
        auxiliar_no_conciliado (DataFrame): Movimientos no conciliados del libro auxiliar.

    Returns:
        DataFrame: Resultados de la conciliación por agrupación.
    """
    resultados = []
    for _, movimiento_auxiliar in auxiliar_no_conciliado.iterrows():
        monto_objetivo = movimiento_auxiliar["Monto"]
        movimientos_extracto = extracto_no_conciliado["Monto"].tolist()
        combinaciones = encontrar_combinaciones(movimientos_extracto, monto_objetivo)

        if combinaciones:
            for combo in combinaciones:
                indices = extracto_no_conciliado[extracto_no_conciliado["Monto"].isin(combo)].index
                resultados.append({
                    "Fecha": movimiento_auxiliar["Fecha"],
                    "Monto": monto_objetivo,
                    "Origen": "Libro Auxiliar",
                    "Estado": "Conciliado",
                    "Doc. Conciliación": ", ".join(extracto_no_conciliado.loc[indices, "Doc. Num"].astype(str)),
                    "Tipo Agrupación": "Extracto Bancario"
                })
                extracto_no_conciliado.drop(indices, inplace=True)  # Eliminar movimientos ya conciliados
    return pd.DataFrame(resultados)

# Función principal de conciliación
def conciliar_banco_excel(extracto_df, auxiliar_df):
    """
    Concilia el extracto bancario con el libro auxiliar.

    Args:
        extracto_df (DataFrame): DataFrame del extracto bancario.
        auxiliar_df (DataFrame): DataFrame del libro auxiliar.

    Returns:
        DataFrame: DataFrame con los resultados de la conciliación.
    """
    resultados_df = pd.DataFrame()

    # 1. Conciliación Directa (Uno a Uno)
    resultados_directa = pd.merge(
        extracto_df, auxiliar_df, on=["Fecha", "Monto"], how="outer", suffixes=("_banco", "_auxiliar"))
    resultados_directa["Origen"] = resultados_directa.apply(
        lambda row: "Banco" if pd.notna(row["Concepto Banco"]) else "Libro Auxiliar", axis=1
    )
    resultados_directa["Estado"] = resultados_directa.apply(
        lambda row: "Conciliado" if pd.notna(row["Concepto Banco"]) and pd.notna(row["Nota Libro Auxiliar"]) else "No Conciliado", axis=1
    )
    resultados_directa["Doc. Conciliación"] = resultados_directa.apply(
        lambda row: row["Doc. Num_auxiliar"] if pd.notna(row["Concepto Banco"]) and pd.notna(row["Nota Libro Auxiliar"]) else row["Doc. Num_banco"] if pd.notna(row["Concepto Banco"]) else None, axis=1
    )
    resultados_df = pd.concat([resultados_df, resultados_directa], ignore_index=True)

    # 2. Conciliación por Agrupación en el Libro Auxiliar
    extracto_no_conciliado = resultados_df[resultados_df["Estado"] == "No Conciliado"]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df["Doc. Num"].isin(resultados_df["Doc. Num_auxiliar"])]
    resultados_agrupacion_libro = conciliacion_agrupacion_libro_auxiliar(extracto_no_conciliado, auxiliar_no_conciliado)
    resultados_df = pd.concat([resultados_df, resultados_agrupacion_libro], ignore_index=True)

    # 3. Conciliación por Agrupación en el Extracto Bancario
    resultados_agrupacion_extracto = conciliacion_agrupacion_extracto_bancario(extracto_no_conciliado, auxiliar_no_conciliado)
    resultados_df = pd.concat([resultados_df, resultados_agrupacion_extracto], ignore_index=True)

    return resultados_df

# Interfaz de Streamlit
st.title("Herramienta de Conciliación Bancaria Automática")

# Cargar archivos Excel
extracto_file = st.file_uploader("Subir Extracto Bancario (Excel)", type=["xlsx"])
auxiliar_file = st.file_uploader("Subir Libro Auxiliar (Excel)", type=["xlsx"])

if extracto_file and auxiliar_file:
    # Leer archivos Excel
    extracto_df = pd.read_excel(extracto_file, sheet_name="Extracto")
    auxiliar_df = pd.read_excel(auxiliar_file, sheet_name="Auxiliar")

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
else:
    st.info("Por favor, sube ambos archivos para comenzar la conciliación.")