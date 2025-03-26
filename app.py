import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations

# Función para buscar la fila de encabezados
def buscar_fila_encabezados(df, columnas_esperadas, max_filas=25):
    """
    Busca la fila que contiene los nombres de las columnas esperadas.

    Args:
        df (DataFrame): El DataFrame del archivo de Excel.
        columnas_esperadas (dict): Un diccionario con los nombres esperados y sus posibles variantes.
        max_filas (int): El número máximo de filas para buscar los encabezados.

    Returns:
        int: El índice de la fila que contiene los encabezados.
    """
    # Convertir las variantes de los encabezados a minúsculas
    columnas_esperadas_lower = {col: [variante.lower() for variante in variantes] for col, variantes in columnas_esperadas.items()}

    for idx in range(min(max_filas, len(df))):  # Limitar a max_filas o el número de filas en el DataFrame
        fila = df.iloc[idx]  # Obtener la fila actual
        celdas = [str(valor).lower() for valor in fila if pd.notna(valor)]  # Filtrar celdas no vacías y convertir a minúsculas

        # Variables para verificar coincidencias
        encontrado_fecha = False
        encontrado_monto = False

        # Revisar cada celda en la fila
        for celda in celdas:
            for col, variantes in columnas_esperadas_lower.items():
                if any(variante in celda for variante in variantes):
                    if col == 'fecha':
                        encontrado_fecha = True
                    elif col == 'monto':
                        encontrado_monto = True

        # Si se encuentran ambos encabezados en la misma fila
        if encontrado_fecha and encontrado_monto:
            st.write(f"Encabezados encontrados en la fila {idx + 1}: {fila.tolist()}")  # Mensaje de depuración
            return idx

        st.write(f"Fila {idx + 1} no coincide: {fila.tolist()}")  # Mensaje de depuración

    return None

# Función para leer datos a partir de la fila de encabezados
def leer_datos_desde_encabezados(archivo, columnas_esperadas, nombre_archivo, max_filas=50):
    """
    Lee los datos de un archivo de Excel a partir de la fila que contiene los encabezados.

    Args:
        archivo (UploadedFile): El archivo de Excel cargado en Streamlit.
        columnas_esperadas (dict): Un diccionario con los nombres esperados y sus posibles variantes.
        nombre_archivo (str): El nombre del archivo (para mensajes de error).
        max_filas (int): El número máximo de filas para buscar los encabezados.

    Returns:
        DataFrame: El DataFrame con los datos correctamente cargados.
    """
    # Leer el archivo de Excel sin asumir que los encabezados están en la primera fila
    df = pd.read_excel(archivo, header=None)

    # Mostrar las primeras filas del archivo para depuración
    st.write(f"Vista previa de las primeras {max_filas} filas del archivo {nombre_archivo}:")
    st.write(df.head(max_filas))

    # Buscar la fila de encabezados
    fila_encabezados = buscar_fila_encabezados(df, columnas_esperadas, max_filas)
    if fila_encabezados is None:
        st.error(f"No se encontraron los encabezados necesarios en el archivo {nombre_archivo}.")
        st.error(f"Se buscaron en las primeras {max_filas} filas.")
        st.stop()

    st.success(f"Encabezados encontrados en la fila {fila_encabezados + 1} del archivo {nombre_archivo}.")

    # Leer los datos a partir de la fila de encabezados
    df = pd.read_excel(archivo, header=fila_encabezados)
    

    ## Normalizar las columnas
    df = normalizar_dataframe(df, columnas_esperadas_extracto)

    # Verificar si el DataFrame tiene las columnas esperadas
    for col in columnas_esperadas.keys():
        if col not in df.columns:
            st.error(f"La columna esperada '{col}' no se encontró en los datos leídos del archivo '{nombre_archivo}'.")
            st.stop()
    
    return df

    st.write("Datos leídos correctamente:")
    st.write(df.head())  # Muestra las primeras filas del DataFrame leído

# Función para identificar columnas
def identificar_columnas(df, columnas_esperadas, nombre_archivo):
    """
    Identifica las columnas necesarias en un DataFrame basándose en coincidencias parciales.

    Args:
        df (DataFrame): El DataFrame del archivo de Excel.
        columnas_esperadas (dict): Un diccionario con los nombres esperados y sus posibles variantes.
        nombre_archivo (str): El nombre del archivo (para mensajes de error).

    Returns:
        dict: Un diccionario con las columnas identificadas.
    """
    columnas_identificadas = {}
    for col_esperada, variantes in columnas_esperadas.items():
        for col in df.columns:
            if any(variante.lower() in str(col).lower() for variante in variantes):
                columnas_identificadas[col_esperada] = col
                break
        else:
            st.error(f"No se encontró una columna que coincida con: {', '.join(variantes)} en el archivo {nombre_archivo}.")
            st.error(f"Columnas encontradas en el archivo: {', '.join(df.columns)}")
            st.stop()
    return columnas_identificadas

# Función para normalizar un DataFrame
def normalizar_dataframe(df, columnas_esperadas):
    """
    Normaliza un DataFrame para que use los nombres de columnas esperados.

    Args:
        df (DataFrame): El DataFrame original.
        columnas_esperadas (dict): Un diccionario con los nombres esperados y sus posibles variantes.

    Returns:
        DataFrame: El DataFrame normalizado.
    """
    # Crear un mapeo de nombres de columnas basado en las variantes
    mapeo_columnas = {variante.lower().strip(): col for col, variantes in columnas_esperadas.items() for variante in variantes}

    # Convertir los nombres de las columnas del DataFrame a minúsculas y eliminar espacios
    df.columns = [col.lower().strip() for col in df.columns]

    # Eliminar columnas duplicadas
    df = df.loc[:, ~df.columns.duplicated()]

    # Mostrar el mapeo de columnas para depuración
    st.write("Mapeo de columnas:", mapeo_columnas)

    # Renombrar las columnas según el mapeo
    df.rename(columns=mapeo_columnas, inplace=True)

    # Mostrar el DataFrame después de renombrar las columnas para depuración
    st.write("DataFrame después de renombrar columnas:")
    st.write(df.head())  # Muestra las primeras filas del DataFrame leído
    
    # Opcional: Eliminar columnas no necesarias
    columnas_a_eliminar = [col for col in df.columns if col not in columnas_esperadas.keys()]
    df.drop(columns=columnas_a_eliminar, inplace=True, errors='ignore')
    
    return df

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
        # Definir las columnas esperadas y sus posibles variantes
        columnas_esperadas_extracto = {
            "fecha": ["fecha de operación", "fecha", "date", "fecha_operacion"],
            "monto": ["importe (cop)", "monto", "valor", "amount"],
            "concepto": ["concepto", "descripción", "observaciones", "concepto banco"],
            "numero_movimiento": ["número de movimiento", "numero de movimiento", "movimiento"]
        }

        columnas_esperadas_auxiliar = {
            "fecha": ["fecha", "date", "fecha de operación", "fecha_operacion"],
            "monto": ["monto", "importe", "valor", "amount"],
            "nota": ["nota", "nota libro auxiliar", "descripción", "observaciones"]
        }

        # Leer los datos a partir de la fila de encabezados
        extracto_df = leer_datos_desde_encabezados(extracto_file, columnas_esperadas_extracto, "Extracto Bancario", max_filas=50)
        auxiliar_df = leer_datos_desde_encabezados(auxiliar_file, columnas_esperadas_auxiliar, "Libro Auxiliar", max_filas=50)

        # Identificar y normalizar las columnas
        columnas_extracto = identificar_columnas(extracto_df, columnas_esperadas_extracto, "Extracto Bancario")
        columnas_auxiliar = identificar_columnas(auxiliar_df, columnas_esperadas_auxiliar, "Libro Auxiliar")

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