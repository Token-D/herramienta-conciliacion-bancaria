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
        # Verificar coincidencias aproximadas
        col_encontrada = False
        for variante, nombre_esperado in mapeo_columnas.items():
            if variante in col:
                if nombre_esperado not in columnas_vistas:
                    nuevo_nombres.append(nombre_esperado)
                    columnas_vistas.add(nombre_esperado)
                    col_encontrada = True
                    break
        
        if not col_encontrada:
            nuevo_nombres.append(col)
    
    # Asignar los nuevos nombres de columnas
    df.columns = nuevo_nombres
    
    # Eliminar columnas duplicadas después de renombrar
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
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


# Función para asegurar que los montos sean numéricos
def asegurar_montos_numericos(df):
    """
    Convierte la columna 'monto' a tipo numérico.
    """
    if 'monto' in df.columns:
        try:
            # Convertir a numérico
            df['monto'] = pd.to_numeric(df['monto'], errors='coerce')
            st.success(f"Columna 'monto' convertida a numérico. Tipo de dato: {df['monto'].dtype}")
        except Exception as e:
            st.error(f"Error al convertir montos a numérico: {e}")
    return df

# Función para procesar los montos del libro auxiliar
def procesar_montos_auxiliar(df):
    """
    Procesa las columnas de débitos y créditos para obtener una columna de monto unificada.
    """
    # Verificar si existen las columnas debitos y creditos
    columnas = df.columns.str.lower()
    
    # Mostrar información de diagnóstico
    st.write("Columnas del libro auxiliar:", ", ".join(columnas))
    
    # Buscar columnas de débitos
    cols_debito = [col for col in columnas if "deb" in col or "debe" in col or "cargo" in col]
    # Buscar columnas de créditos
    cols_credito = [col for col in columnas if "cred" in col or "haber" in col or "abono" in col]
    
    st.write(f"Columnas de débito encontradas: {cols_debito}")
    st.write(f"Columnas de crédito encontradas: {cols_credito}")
    
    # Si ya existe una columna de monto, verificar si tiene valores válidos
    if "monto" in columnas:
        if df["monto"].notna().any() and (df["monto"] != 0).any():
            st.success("Columna de monto encontrada con valores válidos.")
            return df
    
    # Si encontramos columnas de débito y crédito
    if cols_debito and cols_credito:
        # Crear nueva columna monto
        df["monto"] = 0.0
        
        # Para cada columna de débito
        for col in cols_debito:
            # Asegurarse de que la columna sea numérica
            try:
                # Intentar convertir a numérico, NaN si falla
                df[col] = pd.to_numeric(df[col], errors='coerce')
                # Reemplazar NaN con 0
                df[col] = df[col].fillna(0)
                # Sumar a la columna monto
                df["monto"] += df[col]
            except Exception as e:
                st.warning(f"Error al procesar la columna de débito '{col}': {e}")
        
        # Para cada columna de crédito
        for col in cols_credito:
            try:
                # Intentar convertir a numérico, NaN si falla
                df[col] = pd.to_numeric(df[col], errors='coerce')
                # Reemplazar NaN con 0
                df[col] = df[col].fillna(0)
                # Restar de la columna monto
                df["monto"] -= df[col]
            except Exception as e:
                st.warning(f"Error al procesar la columna de crédito '{col}': {e}")
            
        # Para verificar, mostrar algunos valores
        st.write("Primeros 5 montos calculados:", df["monto"].head(5).tolist())
        
        # Eliminar columnas originales de débito y crédito si se desea
        # df.drop(columns=cols_debito + cols_credito, inplace=True, errors='ignore')
    else:
        st.warning("No se encontraron columnas de débito y crédito. Puede que los montos no se procesen correctamente.")
    
    return df

# Función para encontrar combinaciones que sumen un monto específico
def encontrar_combinaciones(df, monto_objetivo, tolerancia=0.01, max_combinacion=4):
    """
    Encuentra combinaciones de valores en df['monto'] que sumen aproximadamente monto_objetivo.
    Devuelve lista de índices de las filas que conforman la combinación.
    """
    # Asegurarse de que los montos sean numéricos
    movimientos = []
    indices_validos = []
    
    for idx, valor in zip(df.index, df["monto"]):
        try:
            # Intentar convertir a numérico
            valor_num = float(valor)
            movimientos.append(valor_num)
            indices_validos.append(idx)
        except (ValueError, TypeError):
            # Ignorar valores que no se pueden convertir a flotante
            continue
    
    if not movimientos:
        return []
    
    combinaciones_validas = []
    
    # Convertir monto_objetivo a numérico
    try:
        monto_objetivo = float(monto_objetivo)
    except (ValueError, TypeError):
        return []
    
    # Limitar la búsqueda a combinaciones pequeñas
    for r in range(1, min(max_combinacion, len(movimientos)) + 1):
        for combo_indices in combinations(range(len(movimientos)), r):
            combo_valores = [movimientos[i] for i in combo_indices]
            suma = sum(combo_valores)
            if abs(suma - monto_objetivo) <= tolerancia:
                indices_combinacion = [indices_validos[i] for i in combo_indices]
                combinaciones_validas.append((indices_combinacion, combo_valores))
    
    # Ordenar por tamaño de combinación (preferimos las más pequeñas)
    combinaciones_validas.sort(key=lambda x: len(x[0]))
    
    if combinaciones_validas:
        return combinaciones_validas[0][0]  # Devolver los índices de la mejor combinación
    return []

# Función para la conciliación directa (uno a uno)
def conciliacion_directa(extracto_df, auxiliar_df):
    """
    Realiza la conciliación directa entre el extracto bancario y el libro auxiliar.
    Busca coincidencias exactas en fecha y monto.
    """
    resultados = []
    extracto_conciliado_idx = set()
    auxiliar_conciliado_idx = set()
    
    # Para cada fila en el extracto
    for idx_extracto, fila_extracto in extracto_df.iterrows():
        # Buscar coincidencias en el libro auxiliar
        coincidencias = auxiliar_df[
            (auxiliar_df["fecha"] == fila_extracto["fecha"]) & 
            (abs(auxiliar_df["monto"] - fila_extracto["monto"]) < 0.01)
        ]
        
        if not coincidencias.empty:
            # Tomar la primera coincidencia
            idx_auxiliar = coincidencias.index[0]
            fila_auxiliar = coincidencias.iloc[0]
            
            # Marcar como conciliados
            extracto_conciliado_idx.add(idx_extracto)
            auxiliar_conciliado_idx.add(idx_auxiliar)
            
            # Añadir a resultados
            resultados.append({
                'fecha': fila_extracto["fecha"],
                'monto': fila_extracto["monto"],
                'concepto': fila_extracto.get("concepto", ""),
                'numero_movimiento': fila_extracto.get("numero_movimiento", ""),
                'origen': 'Banco',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Directa',
                'doc_conciliacion': fila_auxiliar.get("doc. num", "")
            })
    
    # Registros no conciliados del extracto bancario
    for idx_extracto, fila_extracto in extracto_df.iterrows():
        if idx_extracto not in extracto_conciliado_idx:
            resultados.append({
                'fecha': fila_extracto["fecha"],
                'monto': fila_extracto["monto"],
                'concepto': fila_extracto.get("concepto", ""),
                'numero_movimiento': fila_extracto.get("numero_movimiento", ""),
                'origen': 'Banco',
                'estado': 'No Conciliado',
                'tipo_conciliacion': '',
                'doc_conciliacion': ''
            })
    
    # Registros no conciliados del libro auxiliar
    for idx_auxiliar, fila_auxiliar in auxiliar_df.iterrows():
        if idx_auxiliar not in auxiliar_conciliado_idx:
            resultados.append({
                'fecha': fila_auxiliar["fecha"],
                'monto': fila_auxiliar["monto"],
                'concepto': fila_auxiliar.get("nota", ""),
                'numero_movimiento': '',
                'origen': 'Libro Auxiliar',
                'estado': 'No Conciliado',
                'tipo_conciliacion': '',
                'doc_conciliacion': fila_auxiliar.get("doc. num", "")
            })
    
    return pd.DataFrame(resultados), extracto_conciliado_idx, auxiliar_conciliado_idx

# Función para la conciliación por agrupación en el libro auxiliar
def conciliacion_agrupacion_auxiliar(extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx):
    """
    Busca grupos de valores en el libro auxiliar que sumen el monto de un movimiento en el extracto.
    """
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    # Filtrar los registros aún no conciliados
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
    
    # Para cada movimiento no conciliado del extracto
    for idx_extracto, fila_extracto in extracto_no_conciliado.iterrows():
        # Buscar combinaciones en el libro auxiliar
        indices_combinacion = encontrar_combinaciones(
            auxiliar_no_conciliado, 
            fila_extracto["monto"],
            tolerancia=0.01
        )
        
        if indices_combinacion:
            # Marcar como conciliados
            nuevos_extracto_conciliado.add(idx_extracto)
            nuevos_auxiliar_conciliado.update(indices_combinacion)
            
            # Obtener números de documento
            docs_conciliacion = auxiliar_no_conciliado.loc[indices_combinacion, "doc. num"].astype(str).tolist()
            docs_conciliacion = [str(doc) for doc in docs_conciliacion]
            
            # Añadir a resultados
            resultados.append({
                'fecha': fila_extracto["fecha"],
                'monto': fila_extracto["monto"],
                'concepto': fila_extracto.get("concepto", ""),
                'numero_movimiento': fila_extracto.get("numero_movimiento", ""),
                'origen': 'Banco',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Agrupación en Libro Auxiliar',
                'doc_conciliacion': ", ".join(docs_conciliacion)
            })
    
    return pd.DataFrame(resultados), nuevos_extracto_conciliado, nuevos_auxiliar_conciliado

# Función para la conciliación por agrupación en el extracto bancario
def conciliacion_agrupacion_extracto(extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx):
    """
    Busca grupos de valores en el extracto que sumen el monto de un movimiento en el libro auxiliar.
    """
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    # Filtrar los registros aún no conciliados
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
    
    # Para cada movimiento no conciliado del libro auxiliar
    for idx_auxiliar, fila_auxiliar in auxiliar_no_conciliado.iterrows():
        # Buscar combinaciones en el extracto
        indices_combinacion = encontrar_combinaciones(
            extracto_no_conciliado, 
            fila_auxiliar["monto"],
            tolerancia=0.01
        )
        
        if indices_combinacion:
            # Marcar como conciliados
            nuevos_auxiliar_conciliado.add(idx_auxiliar)
            nuevos_extracto_conciliado.update(indices_combinacion)
            
            # Obtener números de movimiento
            nums_movimiento = extracto_no_conciliado.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            nums_movimiento = [str(num) for num in nums_movimiento]
            
            # Añadir a resultados
            resultados.append({
                'fecha': fila_auxiliar["fecha"],
                'monto': fila_auxiliar["monto"],
                'concepto': fila_auxiliar.get("nota", ""),
                'numero_movimiento': ", ".join(nums_movimiento),
                'origen': 'Libro Auxiliar',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Agrupación en Extracto Bancario',
                'doc_conciliacion': fila_auxiliar.get("doc. num", "")
            })
    
    return pd.DataFrame(resultados), nuevos_extracto_conciliado, nuevos_auxiliar_conciliado

# Función principal de conciliación
def conciliar_banco_completo(extracto_df, auxiliar_df):
    """
    Implementa la lógica completa de conciliación.
    """
    # 1. Conciliación directa (uno a uno)
    resultados_directa, extracto_conciliado_idx, auxiliar_conciliado_idx = conciliacion_directa(
        extracto_df, auxiliar_df
    )
    st.write(f"Conciliación directa: {len(extracto_conciliado_idx)} movimientos del extracto y {len(auxiliar_conciliado_idx)} movimientos del auxiliar")
    
    # 2. Conciliación por agrupación en el libro auxiliar
    resultados_agrup_aux, nuevos_extracto_conc1, nuevos_auxiliar_conc1 = conciliacion_agrupacion_auxiliar(
        extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx
    )
    st.write(f"Conciliación por agrupación en libro auxiliar: {len(nuevos_extracto_conc1)} movimientos del extracto y {len(nuevos_auxiliar_conc1)} movimientos del auxiliar")
    
    # Actualizar índices de conciliados
    extracto_conciliado_idx.update(nuevos_extracto_conc1)
    auxiliar_conciliado_idx.update(nuevos_auxiliar_conc1)
    
    # 3. Conciliación por agrupación en el extracto bancario
    resultados_agrup_ext, nuevos_extracto_conc2, nuevos_auxiliar_conc2 = conciliacion_agrupacion_extracto(
        extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx
    )
    st.write(f"Conciliación por agrupación en extracto bancario: {len(nuevos_extracto_conc2)} movimientos del extracto y {len(nuevos_auxiliar_conc2)} movimientos del auxiliar")
    
    # Combinar resultados
    resultados_finales = pd.concat([
        resultados_directa,
        resultados_agrup_aux,
        resultados_agrup_ext
    ], ignore_index=True)
    
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
            "fecha": ["fecha de operación", "fecha", "date", "fecha_operacion", "f. operación"],
            "monto": ["importe (cop)", "monto", "valor", "amount", "importe"],
            "concepto": ["concepto", "descripción", "observaciones", "concepto banco", "descripcion"],
            "numero_movimiento": ["número de movimiento", "numero de movimiento", "movimiento", "no. movimiento", "num"]
        }

        columnas_esperadas_auxiliar = {
            "fecha": ["fecha", "date", "fecha de operación", "fecha_operacion", "f. operación"],
            "debitos": ["debitos", "débitos", "debe", "cargo", "cargos", "valor débito"],
            "creditos": ["creditos", "créditos", "haber", "abono", "abonos", "valor crédito"],
            "nota": ["nota", "nota libro auxiliar", "descripción", "observaciones", "descripcion"],
            "doc. num": ["doc num", "doc. num", "documento", "número documento", "numero documento", "nro. documento"]
        }

        # Leer los datos a partir de la fila de encabezados
        extracto_df = leer_datos_desde_encabezados(extracto_file, columnas_esperadas_extracto, "Extracto Bancario", max_filas=50)
        auxiliar_df = leer_datos_desde_encabezados(auxiliar_file, columnas_esperadas_auxiliar, "Libro Auxiliar", max_filas=50)

        # Procesar datos del libro auxiliar
        auxiliar_df = procesar_montos_auxiliar(auxiliar_df)
        
        # Estandarizar las fechas en ambos DataFrames
        extracto_df = estandarizar_fechas(extracto_df)
        auxiliar_df = estandarizar_fechas(auxiliar_df)        
        
        # Mostrar información sobre los tipos de datos
        st.write("Tipo de dato en columna fecha (Extracto):", extracto_df['fecha'].dtype)
        st.write("Tipo de dato en columna fecha (Auxiliar):", auxiliar_df['fecha'].dtype)

        # Mostrar resúmenes de los datos cargados
        st.subheader("Resumen de datos cargados")
        st.write(f"Extracto bancario: {len(extracto_df)} movimientos")
        st.write(f"Libro auxiliar: {len(auxiliar_df)} movimientos")
        
        # Mostrar las primeras filas como ejemplo
        col1, col2 = st.columns(2)
        with col1:
            st.write("Primeras filas del extracto bancario:")
            st.write(extracto_df.head(3))
        with col2:
            st.write("Primeras filas del libro auxiliar:")
            st.write(auxiliar_df.head(3))

        # Realizar conciliación
        resultados_df = conciliar_banco_completo(extracto_df, auxiliar_df)

        # Mostrar resultados
        st.subheader("Resultados de la Conciliación")
        
        # Estadísticas de conciliación
        conciliados = resultados_df[resultados_df['estado'] == 'Conciliado']
        no_conciliados = resultados_df[resultados_df['estado'] == 'No Conciliado']
        
        st.write(f"Total de movimientos: {len(resultados_df)}")
        st.write(f"Movimientos conciliados: {len(conciliados)} ({len(conciliados)/len(resultados_df)*100:.2f}%)")
        st.write(f"Movimientos no conciliados: {len(no_conciliados)} ({len(no_conciliados)/len(resultados_df)*100:.2f}%)")
        
        # Mostrar resultados por tipo de conciliación
        st.write("Distribución por tipo de conciliación:")
        tipo_conciliacion = resultados_df.groupby('tipo_conciliacion').size().reset_index(name='cantidad')
        st.write(tipo_conciliacion)
        
        # Mostrar todos los resultados
        st.write("Detalle de todos los movimientos:")
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
