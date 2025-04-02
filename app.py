import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations

# Función para buscar la fila de encabezados
def buscar_fila_encabezados(df, columnas_esperadas, max_filas=30):
    """
    Busca la fila que contiene al menos 'fecha' y una columna de monto (monto, debitos o creditos).
    Otras columnas son opcionales.
    """
    columnas_esperadas_lower = {col: [variante.lower() for variante in variantes] 
                                for col, variantes in columnas_esperadas.items()}

    for idx in range(min(max_filas, len(df))):
        fila = df.iloc[idx]
        celdas = [str(valor).lower() for valor in fila if pd.notna(valor)]

        # Variables para verificar coincidencias mínimas
        tiene_fecha = False
        tiene_monto = False

        # Revisar cada celda en la fila
        for celda in celdas:
            # Verificar 'fecha'
            if 'fecha' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['fecha']):
                tiene_fecha = True
            # Verificar columnas de monto (monto, debitos o creditos)
            if 'monto' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['monto']):
                tiene_monto = True
            elif 'debitos' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['debitos']):
                tiene_monto = True
            elif 'creditos' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['creditos']):
                tiene_monto = True

        # Si se encuentran los mínimos necesarios (fecha y algún monto)
        if tiene_fecha and tiene_monto:
            return idx

    return None
# Función para leer datos a partir de la fila de encabezados
def leer_datos_desde_encabezados(archivo, columnas_esperadas, nombre_archivo, max_filas=30):
    # Leer el archivo de Excel sin asumir encabezados, leyendo todas las filas por defecto
    df = pd.read_excel(archivo, header=None)
    total_filas_inicial = len(df)
    st.write(f"Total de filas leídas inicialmente en {nombre_archivo}: {total_filas_inicial}")
    
    # Buscar la fila de encabezados
    fila_encabezados = buscar_fila_encabezados(df, columnas_esperadas, max_filas)
    if fila_encabezados is None:
        st.error(f"No se encontraron los encabezados necesarios en el archivo {nombre_archivo}.")
        st.error(f"Se buscaron en las primeras {max_filas} filas. Se requieren al menos 'fecha' y una columna de monto (monto, debitos o creditos).")
        st.stop()

    st.success(f"Encabezados encontrados en la fila {fila_encabezados + 1} del archivo {nombre_archivo}.")

    # Leer los datos a partir de la fila de encabezados, sin limitar filas
    df = pd.read_excel(archivo, header=fila_encabezados)
    total_filas_datos = len(df)
    st.write(f"Filas leídas después de establecer encabezados en {nombre_archivo}: {total_filas_datos}")

    # Buscar la columna 'Doc Num' entre las variantes posibles antes de normalizar
    variantes_doc_num = columnas_esperadas.get('Doc Num', ["Doc Num"])  # Obtener variantes de columnas_esperadas
    doc_num_col = None
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(variante.lower().strip() in col_lower for variante in variantes_doc_num):
            doc_num_col = col
            break
    
    # Filtrar filas donde 'Doc Num' no esté vacío
    if doc_num_col:
        filas_antes = len(df)
        # Eliminar filas donde 'Doc Num' sea NaN, None o cadena vacía
        df = df[df[doc_num_col].notna() & (df[doc_num_col] != '')]
        filas_despues = len(df)
        st.write(f"Filas después de eliminar las que tienen '{doc_num_col}' vacío en {nombre_archivo}: {filas_despues}")
        if filas_antes > filas_despues:
            st.info(f"Se eliminaron {filas_antes - filas_despues} filas con '{doc_num_col}' vacío.")
    else:
        st.warning(f"No se encontró una columna tipo 'Doc Num' en {nombre_archivo} antes de normalizar. No se aplicó filtro.")
    
    # Normalizar las columnas
    df = normalizar_dataframe(df, columnas_esperadas)
    
    # Verificar si el DataFrame tiene al menos las columnas mínimas necesarias
    if 'fecha' not in df.columns:
        st.error(f"La columna obligatoria 'fecha' no se encontró en los datos leídos del archivo '{nombre_archivo}'.")
        st.stop()
    
    # Verificar si existe al menos una columna de monto
    if 'monto' not in df.columns and ('debitos' not in df.columns or 'creditos' not in df.columns):
        st.error(f"No se encontró ninguna columna de monto (monto, debitos o creditos) en el archivo '{nombre_archivo}'.")
        st.stop()
    
    # Mostrar columnas detectadas (para depuración)
    columnas_encontradas = [col for col in columnas_esperadas.keys() if col in df.columns]
    st.info(f"Columnas encontradas en {nombre_archivo} después de normalizar: {columnas_encontradas}")
    
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
def estandarizar_fechas(df, mes_conciliacion):
    """
    Convierte la columna 'fecha' a formato datetime64, considerando el mes de conciliación.
    """
    if 'fecha' in df.columns:
        try:
            # Convertir a string primero para manipular
            df['fecha_original'] = df['fecha'].copy()
            df['fecha_str'] = df['fecha'].astype(str)
            
            # Buscar patrones comunes de fecha en la columna
            fechas_detectadas = []
            for fecha_str in df['fecha_str'].dropna().unique()[:10]:  # Analizar primeras 10 fechas únicas
                if '-' in fecha_str:
                    partes = fecha_str.split('T')[0].split('-')
                    if len(partes) == 3:
                        fechas_detectadas.append(partes)
            
            # Determinar si las fechas parecen estar en formato YYYY-MM-DD o YYYY-DD-MM
            formato_detectado = "desconocido"
            if fechas_detectadas:
                # Si el segundo número (posible mes) coincide mayormente con el mes seleccionado
                coincidencias_mes_en_pos1 = sum(1 for partes in fechas_detectadas if int(partes[1]) == mes_conciliacion)
                coincidencias_mes_en_pos2 = sum(1 for partes in fechas_detectadas if int(partes[2]) == mes_conciliacion)
                
                if coincidencias_mes_en_pos1 > coincidencias_mes_en_pos2:
                    formato_detectado = "YYYY-MM-DD"
                else:
                    formato_detectado = "YYYY-DD-MM"
                
                st.info(f"Formato de fecha detectado: {formato_detectado}")
            
            # Convertir a datetime según el formato detectado
            if formato_detectado == "YYYY-MM-DD":
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce', format='%Y-%m-%d')
            elif formato_detectado == "YYYY-DD-MM":
                # Para fechas en formato YYYY-DD-MM, intercambiamos día y mes
                df['fecha'] = pd.to_datetime(df['fecha_str'].str.replace(r'(\d{4})-(\d{2})-(\d{2})', r'\1-\3-\2', regex=True), 
                                            errors='coerce', format='%Y-%m-%d')
            else:
                # Intentar ambos formatos
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
            
            # Filtrar por mes de conciliación
            df = df[df['fecha'].dt.month == mes_conciliacion]
            
            # Mostrar estadísticas
            st.write(f"Registros para el mes {meses[mes_conciliacion-1]}: {len(df)}")
            
        except Exception as e:
            st.warning(f"Error al convertir fechas: {e}")
    return df

def completar_fechas_sin_anio(extracto_df, auxiliar_df):
    """
    Completa las fechas del extracto que no tengan año utilizando el año predominante
    en el libro auxiliar. Ejemplo: '1/02' -> '01/02/2025' si 2025 es el año predominante.
    """
    if 'fecha' not in extracto_df.columns or 'fecha' not in auxiliar_df.columns:
        st.warning(f"Una de las columnas 'fecha' no está presente. No se ajustaron fechas.")
        return extracto_df
    
    # Obtener el año predominante del libro auxiliar
    año_predominante = auxiliar_df['fecha'].dt.year.mode()
    if año_predominante.empty or pd.isna(año_predominante[0]) or año_predominante[0] <= 1900:
        st.warning("No se pudo determinar un año predominante válido en el libro auxiliar. Usando el año actual.")
        año_predominante = pd.Timestamp.now().year
    else:
        año_predominante = int(año_predominante[0])
        st.info(f"Año predominante detectado en el libro auxiliar: {año_predominante}")
    
    # Guardar una copia de las fechas originales para depuración
    extracto_df['fecha_original'] = extracto_df['fecha'].copy()
    
    # Función para parsear y completar fechas
    def parsear_fecha(fecha, año_default):
        if pd.isna(fecha):
            return pd.NaT
        
        # Convertir a string para analizar
        fecha_str = str(fecha).strip()
        
        # Verificar si la fecha parece no tener año (ejemplo: "1/02", "01/02", "3-2")
        if '/' in fecha_str or '-' in fecha_str:
            partes = fecha_str.replace('/', '-').split('-')
            if len(partes) == 2:  # Formato "día/mes" o "mes/día"
                try:
                    # Intentar parsear como día/mes (formato común en algunos países)
                    dia, mes = map(int, partes)
                    return pd.Timestamp(year=año_default, month=mes, day=dia)
                except (ValueError, TypeError):
                    # Si falla, intentar como mes/día
                    try:
                        mes, dia = map(int, partes)
                        return pd.Timestamp(year=año_default, month=mes, day=dia)
                    except (ValueError, TypeError):
                        return pd.NaT
            elif len(partes) == 3:  # Formato completo "día/mes/año"
                try:
                    dia, mes, año = map(int, partes)
                    return pd.Timestamp(year=año, month=mes, day=dia)
                except (ValueError, TypeError):
                    return pd.NaT
        
        # Si no se detecta formato sin año, intentar parsear directamente
        try:
            fecha_dt = pd.to_datetime(fecha_str, errors='coerce')
            if pd.isna(fecha_dt) or fecha_dt.year <= 1900:
                return pd.NaT
            return fecha_dt
        except:
            return pd.NaT
    
    # Aplicar la conversión a la columna 'fecha'
    filas_antes = len(extracto_df)
    extracto_df['fecha'] = extracto_df['fecha'].apply(lambda x: parsear_fecha(x, año_predominante))
    
    # Mostrar un mensaje si se ajustaron fechas
    fechas_ajustadas = extracto_df['fecha'].dt.year == año_predominante
    if fechas_ajustadas.any():
        st.success(f"Se completaron fechas sin año en el extracto usando el año {año_predominante}.")
    
    return extracto_df
    
def detectar_formato_fecha(df):
    """
    Analiza la columna 'fecha' para detectar el formato más probable.
    """
    if 'fecha' not in df.columns:
        return "desconocido"
    
    # Convertir fechas a texto para análisis
    fechas_str = df['fecha'].astype(str).dropna().tolist()
    
    # Contadores para diferentes formatos
    formatos = {
        "YYYY-MM-DD": 0,
        "YYYY-DD-MM": 0,
        "DD-MM-YYYY": 0,
        "MM-DD-YYYY": 0
    }
    
    for fecha_str in fechas_str:
        # Limpiar la fecha (quitar hora si existe)
        if 'T' in fecha_str:
            fecha_str = fecha_str.split('T')[0]
        
        # Para formatos con guiones
        if '-' in fecha_str:
            partes = fecha_str.split('-')
            if len(partes) == 3:
                # Verificar formato YYYY-MM-DD o YYYY-DD-MM
                if len(partes[0]) == 4:  # Primer parte es año
                    if 1 <= int(partes[1]) <= 12:  # Segunda parte es mes
                        formatos["YYYY-MM-DD"] += 1
                    elif 1 <= int(partes[2]) <= 12:  # Tercera parte es mes
                        formatos["YYYY-DD-MM"] += 1
                # Verificar formato DD-MM-YYYY o MM-DD-YYYY
                elif len(partes[2]) == 4:  # Última parte es año
                    if 1 <= int(partes[1]) <= 12:  # Segunda parte es mes
                        formatos["DD-MM-YYYY"] += 1
                    elif 1 <= int(partes[0]) <= 12:  # Primera parte es mes
                        formatos["MM-DD-YYYY"] += 1
    
    # Determinar el formato más común
    formato_mas_comun = max(formatos.items(), key=lambda x: x[1])
    
    if formato_mas_comun[1] > 0:
        return formato_mas_comun[0]
    else:
        return "desconocido"

def estandarizar_fechas_automatico(df, nombre_archivo):
    """
    Estandariza fechas detectando automáticamente el formato.
    """
    if 'fecha' in df.columns:
        try:
            # Guardar fechas originales
            df['fecha_original'] = df['fecha'].copy()
            
            # Detectar formato
            formato = detectar_formato_fecha(df)
            st.info(f"Formato de fecha detectado en {nombre_archivo}: {formato}")
            
            # Convertir según formato
            if formato == "YYYY-MM-DD":
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce', format='%Y-%m-%d')
            elif formato == "YYYY-DD-MM":
                df['fecha'] = pd.to_datetime(df['fecha'].astype(str).str.replace(r'(\d{4})-(\d{2})-(\d{2})', r'\1-\3-\2', regex=True), 
                                            errors='coerce')
            elif formato == "DD-MM-YYYY":
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce', format='%d-%m-%Y')
            elif formato == "MM-DD-YYYY":
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce', format='%m-%d-%Y')
            else:
                # Intentar inferir el formato
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
            
            # Contar valores nulos
            nulos = df['fecha'].isna().sum()
            if nulos > 0:
                st.warning(f"{nulos} fechas no pudieron ser convertidas en {nombre_archivo}.")
            
        except Exception as e:
            st.warning(f"Error al convertir fechas en {nombre_archivo}: {e}")
    
    return df

# Función para procesar los montos del libro auxiliar
def procesar_montos_auxiliar(df):
    """
    Procesa las columnas de débitos y créditos para obtener una columna de monto unificada.
    """
    # Verificar si existen las columnas debitos y creditos
    columnas = df.columns.str.lower()
    
    
    # Buscar columnas de débitos
    cols_debito = [col for col in columnas if "deb" in col or "debe" in col or "cargo" in col]
    # Buscar columnas de créditos
    cols_credito = [col for col in columnas if "cred" in col or "haber" in col or "abono" in col]
    
    
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
    Busca coincidencias en fecha (solo día, mes, año) y monto.
    """
    resultados = []
    extracto_conciliado_idx = set()
    auxiliar_conciliado_idx = set()
    
    # Crear versiones de las fechas sin hora para comparación
    extracto_df['fecha_solo'] = extracto_df['fecha'].dt.date
    auxiliar_df['fecha_solo'] = auxiliar_df['fecha'].dt.date
    
    # Diagnóstico de fechas
    st.subheader("Diagnóstico de fechas")
    col1, col2 = st.columns(2)
    with col1:
        st.write("Fechas en extracto (primeras 5):")
        st.write(extracto_df[['fecha', 'fecha_solo']].head())
    with col2:
        st.write("Fechas en auxiliar (primeras 5):")
        st.write(auxiliar_df[['fecha', 'fecha_solo']].head())
    
    # Para cada fila en el extracto
    for idx_extracto, fila_extracto in extracto_df.iterrows():
        # Buscar coincidencias en el libro auxiliar usando fecha_solo
        if pd.isna(fila_extracto['fecha_solo']):
            continue
            
        coincidencias = auxiliar_df[
            (auxiliar_df['fecha_solo'] == fila_extracto['fecha_solo']) & 
            (abs(auxiliar_df['monto'] - fila_extracto['monto']) < 0.01)
        ]
        
        if not coincidencias.empty:
            # Tomar la primera coincidencia
            idx_auxiliar = coincidencias.index[0]
            fila_auxiliar = coincidencias.iloc[0]
            
            # Marcar como conciliados
            extracto_conciliado_idx.add(idx_extracto)
            auxiliar_conciliado_idx.add(idx_auxiliar)
            
            # Añadir a resultados - entrada del extracto bancario
            resultados.append({
                'fecha': fila_extracto['fecha'],
                'tercero': '',
                'concepto': fila_extracto.get('concepto', ''),
                'numero_movimiento': fila_extracto.get('numero_movimiento', ''),
                'monto': fila_extracto['monto'],
                'origen': 'Banco',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Directa',
                'doc_conciliacion': fila_auxiliar.get('numero_movimiento', ''),
                'index_original': idx_extracto,  # Guardar índice original
                'tipo_registro': 'extracto'
            })

            # Añadir a resultados - entrada del libro auxiliar
            resultados.append({
                'fecha': fila_auxiliar['fecha'],
                'tercero': fila_auxiliar.get('tercero', ''),
                'concepto': fila_auxiliar.get('nota', ''),
                'numero_movimiento': fila_auxiliar.get('numero_movimiento', ''),
                'monto': fila_auxiliar['monto'],
                'origen': 'Libro Auxiliar',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Directa',
                'doc_conciliacion': fila_extracto.get('numero_movimiento', ''),
                'index_original': idx_auxiliar,  # Guardar índice original
                'tipo_registro': 'auxiliar'
            })
    
    # Registros no conciliados del extracto bancario
    for idx_extracto, fila_extracto in extracto_df.iterrows():
        if idx_extracto not in extracto_conciliado_idx:
            resultados.append({
                'fecha': fila_extracto["fecha"],
                'tercero': '',
                'concepto': fila_extracto.get("concepto", ""),
                'numero_movimiento': fila_extracto.get("numero_movimiento", ""),
                'monto': fila_extracto["monto"],
                'origen': 'Banco',
                'estado': 'No Conciliado',
                'tipo_conciliacion': '',
                'doc_conciliacion': '',
                'index_original': idx_extracto,  # Guardar índice original
                'tipo_registro': 'extracto'
            })
    
    # Registros no conciliados del libro auxiliar
    for idx_auxiliar, fila_auxiliar in auxiliar_df.iterrows():
        if idx_auxiliar not in auxiliar_conciliado_idx:
            resultados.append({
                'fecha': fila_auxiliar["fecha"],
                'tercero': fila_auxiliar.get('tercero', ''),
                'concepto': fila_auxiliar.get("nota", ""),
                'numero_movimiento': fila_auxiliar.get("numero_movimiento", ""),
                'monto': fila_auxiliar["monto"],
                'origen': 'Libro Auxiliar',
                'estado': 'No Conciliado',
                'tipo_conciliacion': '',
                'doc_conciliacion': '',
                'index_original': idx_auxiliar,  # Guardar índice original
                'tipo_registro': 'auxiliar'
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
            docs_conciliacion = auxiliar_no_conciliado.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            docs_conciliacion = [str(doc) for doc in docs_conciliacion]
            
            # Añadir a resultados - Movimiento del extracto
            resultados.append({
                'fecha': fila_extracto["fecha"],
                'tercero': '',
                'concepto': fila_extracto.get("concepto", ""),
                'numero_movimiento': fila_extracto.get("numero_movimiento", ""),
                'monto': fila_extracto["monto"],
                'origen': 'Banco',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Agrupación en Libro Auxiliar',
                'doc_conciliacion': ", ".join(docs_conciliacion),
                'index_original': idx_extracto,
                'tipo_registro': 'extracto'
            })
            
            # Añadir a resultados - Cada movimiento del libro auxiliar en la combinación
            for idx_aux in indices_combinacion:
                fila_aux = auxiliar_no_conciliado.loc[idx_aux]
                resultados.append({
                    'fecha': fila_aux["fecha"],
                    'tercero': fila_aux.get("tercero", ""),
                    'concepto': fila_aux.get("nota", ""),
                    'numero_movimiento': fila_aux.get("numero_movimiento", ""),
                    'monto': fila_aux["monto"],
                    'origen': 'Libro Auxiliar',
                    'estado': 'Conciliado',
                    'tipo_conciliacion': 'Agrupación en Libro Auxiliar',
                    'doc_conciliacion': fila_extracto.get("numero_movimiento", ""),
                    'index_original': idx_aux,
                    'tipo_registro': 'auxiliar'
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
            
            # Añadir a resultados - Movimiento del libro auxiliar
            resultados.append({
                'fecha': fila_auxiliar["fecha"],
                'tercero': fila_auxiliar.get('tercero', ''),
                'concepto': fila_auxiliar.get("nota", ""),
                'numero_movimiento': fila_auxiliar.get("numero_movimiento", ""),
                'monto': fila_auxiliar["monto"],
                'origen': 'Libro Auxiliar',
                'estado': 'Conciliado',
                'tipo_conciliacion': 'Agrupación en Extracto Bancario',
                'doc_conciliacion': ", ".join(nums_movimiento),
                'index_original': idx_auxiliar,
                'tipo_registro': 'auxiliar'
            })
            
            # Añadir a resultados - Cada movimiento del extracto en la combinación
            for idx_ext in indices_combinacion:
                fila_ext = extracto_no_conciliado.loc[idx_ext]
                resultados.append({
                    'fecha': fila_ext["fecha"],
                    'tercero': '',
                    'concepto': fila_ext.get("concepto", ""),
                    'numero_movimiento': fila_ext.get("numero_movimiento", ""),
                    'monto': fila_ext["monto"],
                    'origen': 'Banco',
                    'estado': 'Conciliado',
                    'tipo_conciliacion': 'Agrupación en Extracto Bancario',
                    'doc_conciliacion': fila_auxiliar.get("numero_movimiento", ""),
                    'index_original': idx_ext,
                    'tipo_registro': 'extracto'
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
    
    # Filtrar resultados directos para eliminar los que luego fueron conciliados por agrupación
    if not resultados_directa.empty:
        # Eliminar los registros no conciliados que luego se conciliaron por agrupación
        indices_a_eliminar = []
        for idx, fila in resultados_directa.iterrows():
            if fila['estado'] == 'No Conciliado':
                if (fila['tipo_registro'] == 'extracto' and fila['index_original'] in nuevos_extracto_conc1.union(nuevos_extracto_conc2)) or \
                   (fila['tipo_registro'] == 'auxiliar' and fila['index_original'] in nuevos_auxiliar_conc1.union(nuevos_auxiliar_conc2)):
                    indices_a_eliminar.append(idx)
        
        # Eliminar los registros identificados
        if indices_a_eliminar:
            resultados_directa = resultados_directa.drop(indices_a_eliminar)
    
    # Combinar resultados
    resultados_finales = pd.concat([
        resultados_directa,
        resultados_agrup_aux,
        resultados_agrup_ext
    ], ignore_index=True)
    
    # Eliminar columnas auxiliares antes de devolver los resultados finales
    if 'index_original' in resultados_finales.columns:
        resultados_finales = resultados_finales.drop(['index_original', 'tipo_registro'], axis=1)
    
    return resultados_finales

# Interfaz de Streamlit
st.title("Herramienta de Conciliación Bancaria Automática")

# Configuración del mes a conciliar
st.subheader("Configuración")
meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_seleccionado = st.selectbox("Mes a conciliar:", meses)
num_mes = meses.index(mes_seleccionado) + 1  # 1 para enero, 2 para febrero, etc.
detectar_formato_auto = st.checkbox("Detectar formato de fecha automáticamente", value=True)

# Cargar archivos Excel
extracto_file = st.file_uploader("Subir Extracto Bancario (Excel)", type=["xlsx"])
auxiliar_file = st.file_uploader("Subir Libro Auxiliar (Excel)", type=["xlsx"])

if extracto_file and auxiliar_file:
    try:
        # Definir las columnas esperadas y sus posibles variantes
        columnas_esperadas_extracto = {
            "fecha": ["fecha de operación", "fecha", "date", "fecha_operacion", "f. operación"],
            "monto": ["importe (cop)", "monto", "amount", "importe"],
            "concepto": ["concepto", "descripción", "concepto banco", "descripcion", "transacción", "transaccion"],
            "numero_movimiento": ["número de movimiento", "numero de movimiento", "movimiento", "no. movimiento", "num", "nro. documento"]
        }

        columnas_esperadas_auxiliar = {
            "fecha": ["fecha", "date", "fecha de operación", "fecha_operacion", "f. operación"],
            "debitos": ["debitos", "débitos", "debe", "cargo", "cargos", "valor débito"],
            "creditos": ["creditos", "créditos", "haber", "abono", "abonos", "valor crédito"],
            "nota": ["nota", "nota libro auxiliar", "descripción", "observaciones", "descripcion"],
            "numero_movimiento": ["doc num", "doc. num", "documento", "número documento", "numero documento", "nro. documento"],
            "tercero": ["tercero", "Tercero","proveedor"]
        }

        # Leer los datos a partir de la fila de encabezados
        extracto_df = leer_datos_desde_encabezados(extracto_file, columnas_esperadas_extracto, "Extracto Bancario", max_filas=30)
        auxiliar_df = leer_datos_desde_encabezados(auxiliar_file, columnas_esperadas_auxiliar, "Libro Auxiliar", max_filas=30)

        # Procesar datos del libro auxiliar
        auxiliar_df = procesar_montos_auxiliar(auxiliar_df)
        
        # Completar fechas sin año en el extracto
        extracto_df = completar_fechas_sin_anio(extracto_df, auxiliar_df)
        
        # Estandarizar las fechas en ambos DataFrames
        extracto_df = estandarizar_fechas_automatico(extracto_df, "Extracto Bancario")
        auxiliar_df = estandarizar_fechas_automatico(auxiliar_df, "Libro Auxiliar")

        # Mostrar resúmenes de los datos cargados
        st.subheader("Resumen de datos cargados")
        st.write(f"Extracto bancario: {len(extracto_df)} movimientos")
        st.write(f"Libro auxiliar: {len(auxiliar_df)} movimientos")
        
        # Mostrar las primeras filas como ejemplo
        col1, col2 = st.columns(2)
        with col1:
            st.write("Primeras filas del extracto bancario:")
            st.write(extracto_df.head(30))
        with col2:
            st.write("Primeras filas del libro auxiliar:")
            st.write(auxiliar_df.head(30))

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
