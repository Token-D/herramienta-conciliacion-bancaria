import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations

# Función para buscar la fila de encabezados
def buscar_fila_encabezados(df, columnas_esperadas, max_filas=25):
    """
    Busca la fila que contiene los nombres de las columnas esperadas.
    """
    # Convertir las variantes de los encabezados a minúsculas
    columnas_esperadas_lower = {col: [variante.lower() for variante in variantes] for col, variantes in columnas_esperadas.items()}

    for idx in range(min(max_filas, len(df))):
        fila = df.iloc[idx]
        celdas = [str(valor).lower() for valor in fila if pd.notna(valor)]

        # Variables para verificar coincidencias
        encontrados = {col: False for col in columnas_esperadas.keys()}

        # Revisar cada celda en la fila
        for celda in celdas:
            for col, variantes in columnas_esperadas_lower.items():
                if any(variante in celda for variante in variantes):
                    encontrados[col] = True

        # Si se encuentran todos los encabezados necesarios en la misma fila
        if all(encontrados.values()):
            return idx

    return None

# Función para leer datos a partir de la fila de encabezados
def leer_datos_desde_encabezados(archivo, columnas_esperadas, nombre_archivo, max_filas=50):
    """
    Lee los datos de un archivo de Excel a partir de la fila que contiene los encabezados.
    """
    # Leer el archivo de Excel sin asumir que los encabezados están en la primera fila
    df = pd.read_excel(archivo, header=None)

    # Buscar la fila de encabezados
    fila_encabezados = buscar_fila_encabezados(df, columnas_esperadas, max_filas)
    if fila_encabezados is None:
        st.error(f"No se encontraron los encabezados necesarios en el archivo {nombre_archivo}.")
        st.error(f"Se buscaron en las primeras {max_filas} filas.")
        st.stop()

    st.success(f"Encabezados encontrados en la fila {fila_encabezados + 1} del archivo {nombre_archivo}.")

    # Leer los datos a partir de la fila de encabezados
    df = pd.read_excel(archivo, header=fila_encabezados)
    
    # Normalizar las columnas
    df = normalizar_dataframe(df, columnas_esperadas)

    # Verificar si el DataFrame tiene las columnas esperadas
    for col in columnas_esperadas.keys():
        if col not in df.columns:
            st.error(f"La columna esperada '{col}' no se encontró en los datos leídos del archivo '{nombre_archivo}'.")
            st.stop()
    
    return df

# Función para normalizar un DataFrame
def normalizar_dataframe(df, columnas_esperadas):
    """
    Normaliza un DataFrame para que use los nombres de columnas esperados.
    """
    # Convertir los nombres de las columnas del DataFrame a minúsculas
    df.columns = [str(col).lower().strip() for col in df.columns]
    
    # Crear un mapeo de nombres de columnas basado en las variantes
    mapeo_columnas = {}
    for col_esperada, variantes in columnas_esperadas.items():
        for variante in variantes:
            variante_lower = variante.lower().strip()
            mapeo_columnas[variante_lower] = col_esperada
    
    # Renombrar las columnas según el mapeo
    nuevo_nombres = []
    columnas_vistas = set()
    
    for col in df.columns:
        if col in mapeo_columnas:
            nuevo_nombre = mapeo_columnas[col]
            # Si ya hemos asignado este nombre antes, añadir un sufijo único
            if nuevo_nombre in columnas_vistas:
                # No renombrar esta columna, la eliminaremos después
                nuevo_nombres.append(col)
            else:
                nuevo_nombres.append(nuevo_nombre)
                columnas_vistas.add(nuevo_nombre)
        else:
            nuevo_nombres.append(col)
    
    # Asignar los nuevos nombres de columnas
    df.columns = nuevo_nombres
    
    # Eliminar columnas duplicadas después de renombrar
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    # Opcional: Eliminar columnas no necesarias
    columnas_a_mantener = list(columnas_esperadas.keys())
    columnas_a_eliminar = [col for col in df.columns if col not in columnas_a_mantener]
    if columnas_a_eliminar:
        df.drop(columns=columnas_a_eliminar, inplace=True, errors='ignore')
    
    return df

# Función para estandarizar el formato de fechas
def estandarizar_fechas(df):
    """
    Convierte la columna 'fecha' a formato datetime64.
    """
    if 'fecha' in df.columns:
        try:
            # Convertir a datetime
            df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
            # Eliminar filas con fechas inválidas
            df = df.dropna(subset=['fecha'])
        except Exception as e:
            st.warning(f"Error al convertir fechas: {e}")
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
    # Asegurarse de que la columna fecha tenga el mismo tipo en ambos DataFrames
    extracto_df_copy = extracto_df.copy()
    auxiliar_df_copy = auxiliar_df.copy()
    
    # Realizar la conciliación uno a uno usando concat en lugar de merge
    # Primero, crear identificadores únicos
    extracto_df_copy['source'] = 'banco'
    auxiliar_df_copy['source'] = 'auxiliar'
    
    # Concatenar los DataFrames
    combined_df = pd.concat([extracto_df_copy, auxiliar_df_copy], ignore_index=True)
    
    # Agrupar por fecha y monto para encontrar coincidencias
    grouped = combined_df.groupby(['fecha', 'monto']).apply(lambda x: {
        'banco': x[x['source'] == 'banco'],
        'auxiliar': x[x['source'] == 'auxiliar'],
        'match': len(x[x['source'] == 'banco']) > 0 and len(x[x['source'] == 'auxiliar']) > 0
    }).reset_index()
    
    # Crear DataFrame de resultados
    resultados = []
    for _, row in grouped.iterrows():
        grupo_banco = row[0]['banco']
        grupo_auxiliar = row[0]['auxiliar']
        match = row[0]['match']
        
        for _, banco_row in grupo_banco.iterrows():
            resultado = {
                'fecha': banco_row['fecha'],
                'monto': banco_row['monto'],
                'concepto': banco_row.get('concepto', None),
                'numero_movimiento': banco_row.get('numero_movimiento', None),
                'origen': 'Banco',
                'estado': 'Conciliado' if match else 'No Conciliado',
                'doc. conciliación': None
            }
            if match and not grupo_auxiliar.empty:
                # Tomar el primer documento del libro auxiliar que coincide
                resultado['doc. conciliación'] = grupo_auxiliar.iloc[0].get('doc. num', None)
            resultados.append(resultado)
        
        # Agregar filas del libro auxiliar que no tienen correspondencia en el banco
        if not match or grupo_banco.empty:
            for _, auxiliar_row in grupo_auxiliar.iterrows():
                resultados.append({
                    'fecha': auxiliar_row['fecha'],
                    'monto': auxiliar_row['monto'],
                    'concepto': None,
                    'numero_movimiento': None,
                    'nota': auxiliar_row.get('nota', None),
                    'origen': 'Libro Auxiliar',
                    'estado': 'No Conciliado',
                    'doc. conciliación': auxiliar_row.get('doc. num', None)
                })
    
    resultados_df = pd.DataFrame(resultados)
    
    # Realizar agrupación de registros no conciliados
    extracto_no_conciliado = resultados_df[(resultados_df['origen'] == 'Banco') & 
                                          (resultados_df['estado'] == 'No Conciliado')]
    
    # Filtrar registros del libro auxiliar que no se han usado en conciliaciones directas
    doc_conciliados = resultados_df[resultados_df['estado'] == 'Conciliado']['doc. conciliación'].dropna().unique()
    auxiliar_no_conciliado = auxiliar_df_copy[~auxiliar_df_copy['doc. num'].isin(doc_conciliados)]
    
    # Realizar conciliación por agrupación
    resultados_agrupacion = conciliacion_agrupacion_libro_auxiliar(extracto_no_conciliado, auxiliar_no_conciliado)
    
    # Combinar resultados
    resultados_finales = pd.concat([resultados_df, resultados_agrupacion], ignore_index=True)
    
    return resultados_finales

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
            "monto": ["debitos", "creditos", "monto", "importe", "valor", "amount"],
            "nota": ["nota", "nota libro auxiliar", "descripción", "observaciones"],
            "doc. num": ["doc num", "doc. num", "documento", "número documento", "numero documento"]
        }

        # Leer los datos a partir de la fila de encabezados
        extracto_df = leer_datos_desde_encabezados(extracto_file, columnas_esperadas_extracto, "Extracto Bancario", max_filas=50)
        auxiliar_df = leer_datos_desde_encabezados(auxiliar_file, columnas_esperadas_auxiliar, "Libro Auxiliar", max_filas=50)

        # Procesar datos del libro auxiliar para combinar débitos y créditos en una sola columna de monto
        if "debitos" in auxiliar_df.columns and "creditos" in auxiliar_df.columns:
            auxiliar_df["monto"] = auxiliar_df["debitos"].fillna(0) - auxiliar_df["creditos"].fillna(0)
            auxiliar_df.drop(columns=["debitos", "creditos"], inplace=True, errors='ignore')
        
        # Estandarizar las fechas en ambos DataFrames
        extracto_df = estandarizar_fechas(extracto_df)
        auxiliar_df = estandarizar_fechas(auxiliar_df)
        
        # Mostrar información sobre los tipos de datos
        st.write("Tipo de dato en columna fecha (Extracto):", extracto_df['fecha'].dtype)
        st.write("Tipo de dato en columna fecha (Auxiliar):", auxiliar_df['fecha'].dtype)

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
        st.exception(e)  # Muestra el traceback completo para depuración
else:
    st.info("Por favor, sube ambos archivos para comenzar la conciliación.")