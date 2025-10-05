import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations
from dateutil.parser import parse as parse_date
import re
from collections import Counter
from itertools import combinations

# Función para buscar la fila de encabezados
def buscar_fila_encabezados(df, columnas_esperadas, max_filas=30):
    """
    Busca la fila que contiene al menos 'fecha' y una columna de monto (monto, debitos o creditos).
    Soporta coincidencia exacta si la variante comienza con un asterisco (*).
    Retorna solo el índice de la fila (integer), o None si no se encuentra.
    """
    
    # 1. Normalizar variantes a minúsculas.
    columnas_esperadas_lower = {}
    for col, variantes in columnas_esperadas.items():
        columnas_esperadas_lower[col] = [variante.lower() for variante in variantes]

    for idx in range(min(max_filas, len(df))):
        fila = df.iloc[idx]
        celdas = [str(valor).lower() for valor in fila if pd.notna(valor)]

        tiene_fecha = False
        tiene_monto = False

        # 2. Función helper para verificar la coincidencia (Exacta vs Parcial)
        def check_match(celda, variantes_esperadas):
            for variante in variantes_esperadas:
                if variante.startswith('*'):
                    # Coincidencia EXACTA (comparamos con el texto sin el '*'):
                    if celda == variante[1:]: 
                        return True
                elif variante in celda:
                    # Coincidencia parcial (el nombre esperado está contenido en la celda):
                    return True
            return False

        # 3. Revisar cada celda en la fila
        for celda in celdas:
            # Verificar 'fecha'
            if 'fecha' in columnas_esperadas_lower and check_match(celda, columnas_esperadas_lower['fecha']):
                tiene_fecha = True
            
            # Verificar columnas de monto (monto, debitos o creditos)
            if any(col in columnas_esperadas_lower and check_match(celda, columnas_esperadas_lower[col]) 
                   for col in ['monto', 'debitos', 'creditos']):
                tiene_monto = True

        # Si encontramos una fila que contiene al menos fecha y monto, la retornamos
        if tiene_fecha and tiene_monto:
            return idx # Retornamos solo el índice, resolviendo el ValueError

    # Si no se encuentra encabezado, retornamos None
    return None
    
# Función para leer datos a partir de la fila de encabezados
def leer_datos_desde_encabezados(archivo, columnas_esperadas, nombre_archivo, max_filas=30):
    # Determinar la extensión del archivo
    extension = archivo.name.split('.')[-1].lower()
    
    # Si es .xls, convertir a .xlsx
    if extension == 'xls':
        try:
            # Leer el archivo .xls con xlrd
            df_temp = pd.read_excel(archivo, header=None, engine='xlrd')
            # Guardar como .xlsx en un buffer
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_temp.to_excel(writer, index=False, header=None)
            output.seek(0)
            # Actualizar el archivo a usar
            archivo = output
            st.success(f"Conversión de {nombre_archivo} de .xls a .xlsx completada.")
        except Exception as e:
            st.error(f"Error al convertir {nombre_archivo} de .xls a .xlsx: {e}")
            st.stop()
    elif extension != 'xlsx':
        st.error(f"Formato de archivo no soportado: {extension}. Usa .xls o .xlsx.")
        st.stop()
        
    # Leer el archivo de Excel sin asumir encabezados, leyendo todas las filas por defecto
    df = pd.read_excel(archivo, header=None, engine='openpyxl')
    total_filas_inicial = len(df)
    
    # Buscar la fila de encabezados
    fila_encabezados = buscar_fila_encabezados(df, columnas_esperadas, max_filas)
    if fila_encabezados is None:
        st.error(f"No se encontraron los encabezados necesarios en el archivo {nombre_archivo}.")
        st.error(f"Se buscaron en las primeras {max_filas} filas. Se requieren al menos 'fecha' y una columna de monto (monto, debitos o creditos).")
        st.stop()

    # Leer los datos a partir de la fila de encabezados, sin limitar filas
    df = pd.read_excel(archivo, header=fila_encabezados)
    total_filas_datos = len(df)

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
    
    # Si no se encontró 'numero_movimiento', crearlo vacío para evitar KeyErrors posteriores
    if 'numero_movimiento' not in df.columns:
        # Crea la columna con una cadena vacía o un valor único, dependiendo de lo que esperes
        # Usaremos el índice para garantizar un identificador único en caso de ser necesario.
        df['numero_movimiento'] = 'DOC_' + df.index.astype(str)
        # Opcionalmente, podrías usar una columna de conceptos si es más útil:
        # df['numero_movimiento'] = df.get('concepto', '').fillna('').astype(str).str.slice(0, 50) 

    # Lógica Específica por Banco
    if banco_seleccionado == "Davivienda":
        # Concatenar concepto (asume que "Transacción" fue mapeado a 'transaccion_davivienda' o similar)
        
        # Primero, buscamos la columna original 'Transacción'
        col_transaccion = next((col for col in df.columns if 'transacción' in col.lower()), None)
        
        # Asumimos que 'Descripción motivo' se mapeó a 'concepto'
        if col_transaccion and 'concepto' in df.columns:
            # Concatenar la Transacción a la Descripción motivo (columna 'concepto')
            df['concepto'] = df['concepto'].astype(str) + " (" + df[col_transaccion].astype(str) + ")"
            st.info("Davivienda: Se concatenó la columna Transacción al Concepto.")
        
        # Eliminar la columna 'Valor Cheque' si existe y es inútil (solo en Davivienda)
        col_valor_cheque = next((col for col in df.columns if 'valor cheque' in col.lower()), None)
        if col_valor_cheque:
            df = df.drop(columns=[col_valor_cheque], errors='ignore')

    if 'numero_movimiento' not in df.columns:
        # Crea un identificador único. Si 'Documento' existía, se debe haber renombrado antes.
        df['numero_movimiento'] = 'DOC_' + df.index.astype(str)     
    
    return df

def detectar_formato_fechas(fechas_str, porcentaje_analisis=0.6):
    """
    Analiza un porcentaje de fechas para detectar el formato predominante (DD/MM/AAAA o MM/DD/AAAA).
    Devuelve el formato detectado y si el año está presente.
    """
    # Filtrar fechas válidas (no vacías, no NaN)
    fechas_validas = [f for f in fechas_str if pd.notna(f) and f.strip() and f not in ['nan', 'NaT']]
    if not fechas_validas:
        return "desconocido", False

    # Tomar al menos el 60% de las fechas válidas
    n_analizar = max(1, int(len(fechas_validas) * porcentaje_analisis))
    fechas_muestra = fechas_validas[:n_analizar]

    # Contadores para patrones
    formatos = Counter()
    tiene_año = Counter()

    # Expresión regular para capturar componentes numéricos de la fecha
    patron_fecha = r'^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$'

    for fecha in fechas_muestra:
        match = re.match(patron_fecha, fecha.replace('.', '/'))
        if not match:
            continue

        comp1, comp2, comp3 = match.groups()
        comp1, comp2 = int(comp1), int(comp2)
        año_presente = comp3 is not None
        tiene_año[año_presente] += 1

        # Determinar si el primer componente es mes (1-12) o día (1-31)
        if comp1 <= 12 and comp2 <= 31:
            # Puede ser MM/DD o DD/MM, pero si comp1 <= 12, asumimos MM/DD a menos que comp2 <= 12
            if comp2 <= 12:
                # Ambos pueden ser mes, necesitamos más contexto
                formatos["ambiguo"] += 1
            else:
                formatos["MM/DD/AAAA"] += 1
        elif comp1 <= 31 and comp2 <= 12:
            formatos["DD/MM/AAAA"] += 1
        else:
            formatos["desconocido"] += 1

    # Determinar formato predominante
    formato_predominante = formatos.most_common(1)[0][0] if formatos else "desconocido"
    if formato_predominante == "ambiguo":
        # Resolver ambigüedad asumiendo DD/MM/AAAA (común en muchos países)
        formato_predominante = "DD/MM/AAAA"

    # Determinar si la mayoría tiene año
    año_presente = tiene_año.most_common(1)[0][0] if tiene_año else False

    return formato_predominante, año_presente

def estandarizar_fechas(df, nombre_archivo, mes_conciliacion=None, completar_anio=False, auxiliar_df=None):
    """
    Convierte la columna 'fecha' a datetime64, detectando automáticamente el formato de fecha.
    Opcionalmente completa años faltantes y filtra por mes de conciliación.
    """
    if 'fecha' not in df.columns:
        st.warning(f"No se encontró la columna 'fecha' en {nombre_archivo}.")
        return df

    try:
        # Guardar copia de las fechas originales
        df['fecha_original'] = df['fecha'].copy()
        df['fecha_str'] = df['fecha'].astype(str).str.strip()

        # Determinar el año base para completar fechas sin año
        año_base = None
        if completar_anio and auxiliar_df is not None and 'fecha' in auxiliar_df.columns:
            años_validos = auxiliar_df['fecha'].dropna().apply(lambda x: x.year if pd.notna(x) else None)
            año_base = años_validos.mode()[0] if not años_validos.empty else pd.Timestamp.now().year
        else:
            año_base = pd.Timestamp.now().year

        # Detectar formato predominante solo para extracto
        es_extracto = "Extracto" in nombre_archivo
        formato_fecha = "desconocido"
        año_presente = False
        if es_extracto:
            formato_fecha, año_presente = detectar_formato_fechas(df['fecha_str'])
            st.write(f"Formato de fecha detectado en {nombre_archivo}: {formato_fecha}, Año presente: {año_presente}")

        # Función para parsear fechas
        def parsear_fecha(fecha_str, mes_conciliacion=None, año_base=None, es_extracto=False, formato_fecha="desconocido"):
            if pd.isna(fecha_str) or fecha_str in ['', 'nan', 'NaT']:
                return pd.NaT

            try:
                # Normalizar separadores
                fecha_str = fecha_str.replace('-', '/').replace('.', '/')

                # Para extracto, usar formato detectado
                if es_extracto and formato_fecha != "desconocido":
                    partes = fecha_str.split('/')
                    if len(partes) >= 2:
                        comp1, comp2 = map(int, partes[:2])
                        año = año_base
                        if len(partes) == 3:
                            año = int(partes[2])
                            if len(partes[2]) == 2:
                                año += 2000 if año < 50 else 1900

                        if formato_fecha == "DD/MM/AAAA":
                            dia, mes = comp1, comp2
                        else:  # MM/DD/AAAA
                            dia, mes = comp2, comp1

                        # Forzar mes_conciliacion si está definido
                        if mes_conciliacion and 1 <= mes <= 12:
                            mes = mes_conciliacion

                        if 1 <= dia <= 31 and 1 <= mes <= 12:
                            return pd.Timestamp(year=año, month=mes, day=dia)

                # Para auxiliar o si no se detectó formato, usar dateutil.parser
                parsed = parse_date(fecha_str, dayfirst=True, fuzzy=True)

                # Para extracto, ajustar mes si mes_conciliacion está definido
                if es_extracto and mes_conciliacion and parsed.month != mes_conciliacion:
                    return pd.Timestamp(year=parsed.year, month=mes_conciliacion, day=parsed.day)

                return parsed
            except (ValueError, TypeError):
                # Manejar fechas sin año
                try:
                    partes = fecha_str.split('/')
                    if len(partes) == 2:
                        comp1, comp2 = map(int, partes[:2])
                        if formato_fecha == "DD/MM/AAAA":
                            dia, mes = comp1, comp2
                        else:  # MM/DD/AAAA o desconocido
                            dia, mes = comp2, comp1 if comp2 <= 31 and comp1 <= 12 else comp1, comp2

                        # Forzar mes_conciliacion para extracto
                        if es_extracto and mes_conciliacion:
                            mes = mes_conciliacion

                        if 1 <= dia <= 31 and 1 <= mes <= 12:
                            return pd.Timestamp(year=año_base, month=mes, day=dia)
                    return pd.NaT
                except (ValueError, IndexError):
                    return pd.NaT

        # Aplicar el parseo de fechas
        df['fecha'] = df['fecha_str'].apply(
            lambda x: parsear_fecha(x, mes_conciliacion, año_base, es_extracto, formato_fecha)
        )

        # Reportar fechas inválidas
        fechas_invalidas = df['fecha'].isna().sum()
        if fechas_invalidas > 0:
            st.warning(f"Se encontraron {fechas_invalidas} fechas inválidas en {nombre_archivo}.")
            st.write("Ejemplos de fechas inválidas:")
            st.write(df[df['fecha'].isna()][['fecha_original', 'fecha_str']].head())

        # Depuración: Mostrar fechas parseadas
        st.write(f"Fechas parseadas en {nombre_archivo} (primeras 4):")
        st.write(df[['fecha_original', 'fecha_str', 'fecha']].head(4))

        # Filtrar por mes solo para extracto si se especifica
        if mes_conciliacion and es_extracto:
            filas_antes = len(df)
            df = df[df['fecha'].dt.month == mes_conciliacion]
            if len(df) < filas_antes:
                meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                         'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
                st.info(f"Se filtraron {filas_antes - len(df)} registros fuera del mes {meses[mes_conciliacion-1]} en {nombre_archivo}.")
                if filas_antes - len(df) > 0:
                    st.write(f"Ejemplos de fechas filtradas (no en {meses[mes_conciliacion-1]}):")
                    st.write(df[df['fecha'].dt.month != mes_conciliacion][['fecha_original', 'fecha_str', 'fecha']].head())

        # Limpiar columnas temporales
        df = df.drop(['fecha_str'], axis=1, errors='ignore')

    except Exception as e:
        st.error(f"Error al estandarizar fechas en {nombre_archivo}: {e}")
        return df

    return df
    
# Función para procesar los montos
# Función para procesar los montos (VERSIÓN CON BANCO SELECCIONADO)
def procesar_montos(df, nombre_archivo, es_extracto=False, invertir_signos=False, banco_seleccionado="Generico"):
    """
    Procesa columnas de débitos y créditos para crear una columna 'monto' unificada,
    aplicando lógica específica según el banco seleccionado.
    """
    import pandas as pd
    import streamlit as st # Asegúrate de que st esté accesible

    # --- Función auxiliar de limpieza robusta (latino) ---
    def limpiar_monto_latino(series):
        series_str = series.astype(str).str.strip()
        series_str = series_str.str.replace(r'[^\d\.\,]+', '', regex=True)
        series_str = series_str.str.replace('.', '', regex=False)
        series_str = series_str.str.replace(',', '.', regex=False)
        return pd.to_numeric(series_str, errors='coerce')
    # ---------------------------------------------------

    columnas = df.columns.str.lower()

    # --- Lógica de Manejo de Monto Único ---
    if "monto" in columnas and df["monto"].notna().any() and (df["monto"] != 0).any():
        
        # 1. Limpieza y Conversión
        if es_extracto and banco_seleccionado == "Davivienda":
            # Davivienda: Aplicar limpieza robusta a MONTO ÚNICO
            df["monto"] = limpiar_monto_latino(df["monto"]).fillna(0)
            
            # --- LÓGICA ESPECÍFICA DE SIGNO Y CONCEPTO PARA DAVIVIENDA ---
            if df["monto"].abs().sum() > 0 and 'concepto' in df.columns:
                
                # Concatenar concepto: Lo haremos más adelante en normalizar_dataframe,
                # pero aquí usamos la columna de Transacción si fue mapeada a 'concepto'.
                
                # 1. Definir términos de Débito (salidas -> deberían ser NEGATIVOS)
                terminos_debito = ['débito', 'debito', 'nota débito', 'cargo', 'retiro', 'dcto', 'descuento']
                es_debito_extracto = df['concepto'].astype(str).str.lower().apply(lambda x: any(term in x for term in terminos_debito))

                # 2. Aplicar el signo NEGATIVO a Débitos POSITIVOS (formato de Davivienda)
                if not invertir_signos:
                    df.loc[es_debito_extracto & (df['monto'] > 0), 'monto'] *= -1
                else:
                    df.loc[es_debito_extracto & (df['monto'] < 0), 'monto'] *= -1

                st.success("Davivienda: Lógica de signos aplicada correctamente.")
            # ------------------------------------------------------------
            
        else:
            # BBVA/Bogotá/Auxiliar: Conversión simple
            df["monto"] = pd.to_numeric(df["monto"], errors='coerce').fillna(0)

        # Advertencia final
        if df["monto"].abs().sum() == 0 and df.shape[0] > 0:
             st.warning(f"La columna 'monto' de {nombre_archivo} resultó en ceros. Revise la columna de Monto y el tipo de movimiento.")
        
        return df

    # [BLOQUE 2: MANEJO DE DÉBITOS Y CRÉDITOS SEPARADOS]
    
    # ... (El código de tu lógica original para encontrar y definir signos de débitos/créditos separados) ...
    terminos_debitos = ["deb", "debe", "cargo", "débito", "valor débito"]
    # ... (y el resto del código hasta la definición de signos) ...
    
    # Definir términos para identificar débitos y créditos
    terminos_debitos = ["deb", "debe", "cargo", "débito", "valor débito"]
    terminos_creditos = ["cred", "haber", "abono", "crédito", "valor crédito"]
    cols_debito = [col for col in df.columns if any(term in col.lower() for term in terminos_debitos)]
    cols_credito = [col for col in df.columns if any(term in col.lower() for term in terminos_creditos)]

    # Si no hay columnas de monto, débitos ni créditos, advertir
    if not cols_debito and not cols_credito and "monto" not in columnas:
        st.warning(f"No se encontraron columnas de monto, débitos o créditos en {nombre_archivo}.")
        return df

    # Inicializar columna 'monto'
    df["monto"] = 0.0

    # Definir signos según el tipo de archivo y si se invierten
    if es_extracto:
        signo_debito = 1 if invertir_signos else -1
        signo_credito = -1 if invertir_signos else 1
    else:
        signo_debito = 1
        signo_credito = -1

    # Ciclo de procesamiento (usando la lógica original simple o la robusta si se necesita)
    for col in cols_debito:
        try:
            # 1. INTENTO SIMPLE (Lógica base: funciona para Auxiliar y Extractos limpios)
            simple_conversion = pd.to_numeric(df[col], errors='coerce')
            
            # 2. LÓGICA CONDICIONAL DE DETECCIÓN (Solo se usa si no se seleccionó Davivienda y si la conversión simple falla)
            # Mantenemos esta lógica para la "detección automática" de formatos no estándar.
            valid_count = simple_conversion.notna().sum()
            
            # Si el banco es 'Generico' o no es Davivienda y la conversión falló, aplicamos limpieza robusta.
            if es_extracto and banco_seleccionado != "Davivienda" and valid_count < (len(df) * 0.05):
                st.info(f"Aplicando limpieza robusta (detección automática) a la columna de débito '{col}' en {nombre_archivo}.")
                cleaned_series = limpiar_monto_latino(df[col]).fillna(0)
            else:
                # Caso Auxiliar, BBVA/Bogotá, o Davivienda no seleccionada que no falló la conversión simple
                cleaned_series = simple_conversion.fillna(0)
            
            df[col] = cleaned_series
            df["monto"] += cleaned_series * signo_debito
            
        except Exception as e:
            st.warning(f"Error al procesar columna de débito '{col}' en {nombre_archivo}: {e}")

    # (Repetir lógica similar para cols_credito)

    for col in cols_credito:
        try:
            # 1. INTENTO SIMPLE (Lógica original)
            simple_conversion = pd.to_numeric(df[col], errors='coerce')
            
            # 2. LÓGICA CONDICIONAL DE DETECCIÓN
            valid_count = simple_conversion.notna().sum()
            
            if es_extracto and banco_seleccionado != "Davivienda" and valid_count < (len(df) * 0.05):
                st.info(f"Aplicando limpieza robusta (detección automática) a la columna de crédito '{col}' en {nombre_archivo}.")
                cleaned_series = limpiar_monto_latino(df[col]).fillna(0)
            else:
                cleaned_series = simple_conversion.fillna(0)
            
            df[col] = cleaned_series
            df["monto"] += cleaned_series * signo_credito

        except Exception as e:
            st.warning(f"Error al procesar columna de crédito '{col}' en {nombre_archivo}: {e}")
    
    # [CÓDIGO ORIGINAL - Lógica de verificación final]
    if df["monto"].eq(0).all() and (cols_debito or cols_credito):
        st.warning(f"La columna 'monto' resultó en ceros en {nombre_archivo}. Verifica las columnas de débitos/créditos.")

    return df

# Función para encontrar combinaciones que sumen un monto específico
def encontrar_combinaciones(df, monto_objetivo, tolerancia=0.01, max_combinacion=4):
    """
    Encuentra combinaciones de valores en df['monto'] que sumen aproximadamente monto_objetivo.
    Restringe la búsqueda a valores del MISMO SIGNO que el objetivo.
    Devuelve lista de índices de las filas que conforman la combinación.
    """
    
    # 1. Preparar monto_objetivo y determinar el signo
    try:
        monto_objetivo = float(monto_objetivo)
    except (ValueError, TypeError):
        return []

    # Determinar si el objetivo es positivo o negativo.
    # Usamos la tolerancia para incluir los ceros en el grupo adecuado.
    es_objetivo_positivo = monto_objetivo >= 0 
    
    movimientos = []
    indices_validos = []
    
    # 2. Iterar y filtrar por signo
    for idx, valor in zip(df.index, df["monto"]):
        try:
            valor_num = float(valor)
            
            # --- LÓGICA DE FILTRADO DE SIGNO (NUEVA) ---
            # Si el objetivo es positivo, solo incluimos valores >= -tolerancia (casi cero o positivo)
            if es_objetivo_positivo and valor_num < -tolerancia:
                continue
            
            # Si el objetivo es negativo, solo incluimos valores <= tolerancia (casi cero o negativo)
            if not es_objetivo_positivo and valor_num > tolerancia:
                continue
            # -------------------------------------------

            movimientos.append(valor_num)
            indices_validos.append(idx)
        except (ValueError, TypeError):
            # Ignorar valores que no se pueden convertir a flotante
            continue
    
    if not movimientos:
        return []
        
    combinaciones_validas = []
    
    # 3. Buscar combinaciones (la lógica sigue igual)
    
    # Limitar la búsqueda a combinaciones pequeñas
    for r in range(1, min(max_combinacion, len(movimientos)) + 1):
        for combo_indices in combinations(range(len(movimientos)), r):
            combo_valores = [movimientos[i] for i in combo_indices]
            suma = sum(combo_valores)
            
            # NOTA: La tolerancia en el filtro inicial ya maneja los ceros.
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
    Empareja registros por fecha y monto, asegurando una relación 1:1 sin reutilizar registros.
    El numero_movimiento puede repetirse y no se usa como criterio de unicidad.
    """
    resultados = []
    extracto_conciliado_idx = set()
    auxiliar_conciliado_idx = set()
    
    # Crear copias para no modificar los DataFrames originales
    extracto_df = extracto_df.copy()
    auxiliar_df = auxiliar_df.copy()
    extracto_df['fecha_solo'] = extracto_df['fecha'].dt.date
    auxiliar_df['fecha_solo'] = auxiliar_df['fecha'].dt.date
        
    # Iterar sobre el extracto
    for idx_extracto, fila_extracto in extracto_df.iterrows():
        if idx_extracto in extracto_conciliado_idx or pd.isna(fila_extracto['fecha_solo']):
            continue
        
        # Filtrar registros del auxiliar no conciliados
        auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
        
        # Buscar coincidencias por fecha y monto
        coincidencias = auxiliar_no_conciliado[
            (auxiliar_no_conciliado['fecha_solo'] == fila_extracto['fecha_solo']) & 
            (abs(auxiliar_no_conciliado['monto'] - fila_extracto['monto']) < 0.01)
        ]
        
        if not coincidencias.empty:
            # Tomar el primer registro no conciliado del auxiliar
            idx_auxiliar = coincidencias.index[0]
            fila_auxiliar = coincidencias.iloc[0]
            
            # Marcar como conciliados
            extracto_conciliado_idx.add(idx_extracto)
            auxiliar_conciliado_idx.add(idx_auxiliar)
            
            # Añadir entrada del extracto bancario
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
                'index_original': idx_extracto,
                'tipo_registro': 'extracto'
            })

            # Añadir entrada del libro auxiliar
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
                'index_original': idx_auxiliar,
                'tipo_registro': 'auxiliar'
            })
    
    # Agregar registros no conciliados del extracto
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
                'index_original': idx_extracto,
                'tipo_registro': 'extracto'
            })
    
    # Agregar registros no conciliados del libro auxiliar
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
                'index_original': idx_auxiliar,
                'tipo_registro': 'auxiliar'
            })
    
    resultados_df = pd.DataFrame(resultados)
    return resultados_df, extracto_conciliado_idx, auxiliar_conciliado_idx

def conciliacion_agrupacion_auxiliar(extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx):
    """
    Busca grupos de valores en el libro auxiliar que sumen el monto de un movimiento en el extracto.
    Garantiza que cada registro del auxiliar se concilie solo una vez, manteniendo objetos datetime.
    """
    import pandas as pd
    
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    # Filtrar los registros aún no conciliados
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)].copy()
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)].copy()
    
    # Creamos una lista de índices del extracto para iterar de forma segura
    indices_extracto_a_iterar = extracto_no_conciliado.index.tolist()

    for idx_extracto in indices_extracto_a_iterar:
        if idx_extracto not in extracto_no_conciliado.index:
            continue
            
        fila_extracto = extracto_no_conciliado.loc[idx_extracto]

        # Buscar combinaciones en el libro auxiliar (con filtro de signo incluido)
        indices_combinacion = encontrar_combinaciones(
            auxiliar_no_conciliado, 
            fila_extracto["monto"],
            tolerancia=0.01
        )
        
        if indices_combinacion:
            # 1. Marcar como conciliados (para el set de retorno)
            nuevos_extracto_conciliado.add(idx_extracto)
            nuevos_auxiliar_conciliado.update(indices_combinacion)
            
            # 2. **ACTUALIZACIÓN CRÍTICA (UNICIDAD)**: Eliminar los índices usados del DataFrame de trabajo del auxiliar.
            auxiliar_no_conciliado = auxiliar_no_conciliado.drop(indices_combinacion, errors='ignore')
            
            # 3. **FECHA**: Usamos el objeto datetime original (¡REVERTIDO!)
            fecha_extracto = fila_extracto["fecha"] 

            # 4. Obtener números de documento
            docs_conciliacion = auxiliar_df.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            docs_conciliacion = [str(doc) for doc in docs_conciliacion]
            
            # Añadir a resultados - Movimiento del extracto
            resultados.append({
                'fecha': fecha_extracto, # <--- OBJETO DATETIME
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
                fila_aux = auxiliar_df.loc[idx_aux] # Usamos el auxiliar_df original
                
                # FECHA: Revertimos a usar el objeto datetime original (¡REVERTIDO!)
                fecha_auxiliar = fila_aux["fecha"] 
                
                resultados.append({
                    'fecha': fecha_auxiliar, # <--- OBJETO DATETIME
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
    
def conciliacion_agrupacion_extracto(extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx):
    """
    Busca grupos de valores en el extracto que sumen el monto de un movimiento en el libro auxiliar.
    Garantiza que cada registro del extracto se concilie solo una vez y aplica formato de fecha DD/MM/YYYY.
    """
    import pandas as pd
    
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    # Filtrar los registros aún no conciliados (usamos .copy() para evitar SettingWithCopyWarning)
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)].copy()
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)].copy()
    
    # Creamos una lista de índices del auxiliar para iterar de forma segura
    indices_auxiliar_a_iterar = auxiliar_no_conciliado.index.tolist()
    
    # Para cada movimiento no conciliado del libro auxiliar
    for idx_auxiliar in indices_auxiliar_a_iterar:
        # Si la fila ha sido movida por algún otro proceso (poco probable aquí), saltar
        if idx_auxiliar not in auxiliar_no_conciliado.index:
            continue
            
        fila_auxiliar = auxiliar_no_conciliado.loc[idx_auxiliar]

        # Buscar combinaciones en el extracto (con filtro de signo incluido)
        indices_combinacion = encontrar_combinaciones(
            extracto_no_conciliado, 
            fila_auxiliar["monto"],
            tolerancia=0.01
        )
        
        if indices_combinacion:
            # 1. Marcar como conciliados (para el set de retorno)
            nuevos_auxiliar_conciliado.add(idx_auxiliar)
            nuevos_extracto_conciliado.update(indices_combinacion)
            
            # 2. **ACTUALIZACIÓN CRÍTICA (UNICIDAD)**: Eliminar los índices usados del DataFrame de trabajo del extracto.
            # Esto evita que los registros del extracto se reutilicen en la siguiente iteración del auxiliar.
            extracto_no_conciliado = extracto_no_conciliado.drop(indices_combinacion, errors='ignore')
            
            # 3. **FORMATO DE FECHA**: Aplicar formato de fecha DD/MM/YYYY al registro del auxiliar
            fecha_auxiliar_str = fila_auxiliar["fecha"].strftime('%d/%m/%Y')
            
            # 4. Obtener números de movimiento (usamos el extracto_df original para evitar errores)
            nums_movimiento = extracto_df.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            nums_movimiento = [str(num) for num in nums_movimiento]
            
            # Añadir a resultados - Movimiento del libro auxiliar
            resultados.append({
                'fecha': fecha_auxiliar_str,
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
                fila_ext = extracto_df.loc[idx_ext] # Usamos el extracto_df original
                
                # Formato de fecha para la línea del extracto
                fecha_extracto_str = fila_ext["fecha"].strftime('%d/%m/%Y')
                
                resultados.append({
                    'fecha': fecha_extracto_str,
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
    
    # 2. Conciliación por agrupación en el libro auxiliar
    resultados_agrup_aux, nuevos_extracto_conc1, nuevos_auxiliar_conc1 = conciliacion_agrupacion_auxiliar(
        extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx
    )
    
    # Actualizar índices de conciliados
    extracto_conciliado_idx.update(nuevos_extracto_conc1)
    auxiliar_conciliado_idx.update(nuevos_auxiliar_conc1)
    
    # 3. Conciliación por agrupación en el extracto bancario
    resultados_agrup_ext, nuevos_extracto_conc2, nuevos_auxiliar_conc2 = conciliacion_agrupacion_extracto(
        extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx
    )
    
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

def aplicar_formato_excel(writer, resultados_df):
    """
    Aplica formatos específicos (encabezados, fechas, moneda, no conciliados) 
    al DataFrame de resultados antes de guardarlo en Excel.
    """
    
    # ----------------------------------------------------
    # CAMBIO CRÍTICO: Asegurar que la columna 'fecha' sea datetime y que 
    # interprete el día primero (DD/MM/YYYY) para corregir inconsistencias visuales.
    # ----------------------------------------------------
    try:
        # Intenta convertir la columna 'fecha' al formato datetime de Pandas.
        # Usa errors='coerce' para convertir fechas inválidas a NaT (Not a Time).
        # Se añade dayfirst=True para forzar la interpretación de fechas como DD/MM/YYYY.
        resultados_df['fecha'] = pd.to_datetime(resultados_df['fecha'], errors='coerce', dayfirst=True)
    except KeyError:
        # En caso de que la columna 'fecha' no exista, se ignora (aunque es poco probable)
        pass
    # ----------------------------------------------------

    worksheet = writer.sheets['Resultados']
    workbook = writer.book
    
    formato_encabezado = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'border': 1, 'bg_color': '#D9E1F2'
    })
    # Se usa 'dd/mm/yyyy' para forzar el formato deseado en Excel
    formato_fecha = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    formato_moneda = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'})
    formato_no_conciliado = workbook.add_format({'bg_color': '#FFCCCB'})
    
    anchuras = {'fecha': 12, 'tercero': 30, 'concepto': 30, 'monto': 15}
    for i, col in enumerate(resultados_df.columns):
        ancho = anchuras.get(col.lower(), 14)
        worksheet.set_column(i, i, ancho)
    
    for i, col in enumerate(resultados_df.columns):
        worksheet.write(0, i, col, formato_encabezado)
    
    worksheet.freeze_panes(1, 0)
    
    for i, col in enumerate(resultados_df.columns):
        col_lower = col.lower()
        
        if col_lower == 'fecha':
            for row_num in range(1, len(resultados_df) + 1):
                valor = resultados_df.iloc[row_num-1][col]
                # Verifica si el valor no es NaT (la versión datetime de NaN)
                if pd.isna(valor):
                    worksheet.write(row_num, i, "", formato_fecha)
                else:
                    # Este método ahora funciona porque el valor es garantizado ser un datetime object
                    worksheet.write_datetime(row_num, i, valor, formato_fecha)
        
        elif col_lower == 'monto':
            for row_num in range(1, len(resultados_df) + 1):
                valor = resultados_df.iloc[row_num-1][col]
                if pd.isna(valor):
                    worksheet.write(row_num, i, "", formato_moneda)
                else:
                    worksheet.write_number(row_num, i, valor, formato_moneda)
        
        elif col_lower == 'estado':
            for row_num in range(1, len(resultados_df) + 1):
                estado = resultados_df.iloc[row_num-1][col]
                if estado == 'No Conciliado':
                    for col_idx in range(len(resultados_df.columns)):
                        # Escribir el estado en la celda 'estado'
                        if col_idx == i:
                            worksheet.write(row_num, col_idx, estado, formato_no_conciliado)
                        else:
                            # Aplicar formato de fila 'No Conciliado'
                            col_name = resultados_df.columns[col_idx]
                            valor = resultados_df.iloc[row_num-1][col_name]
                            
                            if col_name.lower() == 'fecha':
                                formato_combinado = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#FFCCCB'})
                                if pd.isna(valor):
                                    worksheet.write(row_num, col_idx, "", formato_combinado)
                                else:
                                    # Usa write_datetime para fechas
                                    worksheet.write_datetime(row_num, col_idx, valor, formato_combinado)
                                    
                            elif col_name.lower() == 'monto':
                                formato_combinado = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'bg_color': '#FFCCCB'})
                                if pd.isna(valor):
                                    worksheet.write(row_num, col_idx, "", formato_combinado)
                                else:
                                    # Usa write_number para montos
                                    worksheet.write_number(row_num, col_idx, valor, formato_combinado)
                            else:
                                # Usa write para otros tipos (texto/general)
                                worksheet.write(row_num, col_idx, valor, formato_no_conciliado)

# Interfaz de Streamlit
st.title("Herramienta de Conciliación Bancaria Automática")

# 1. Selector de Banco (Nuevo)
BANCOS = ["Generico", "BBVA", "Bogotá", "Davivienda", "Bancolombia"]
banco_seleccionado = st.selectbox(
    "Selecciona el Banco:",
    BANCOS,
    key="banco_seleccionado"
)

st.subheader("Configuración")
meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_seleccionado = st.selectbox("Mes a conciliar (opcional):", ["Todos"] + meses)
mes_conciliacion = meses.index(mes_seleccionado) + 1 if mes_seleccionado != "Todos" else None

tipos_aceptados = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # .xlsx
    "application/vnd.ms-excel",  # .xls
    "application/excel",  # Variante .xls
    "application/x-excel",  # Variante .xls
    "application/x-msexcel",  # Variante .xls
    "application/octet-stream"  # Por si el navegador lo detecta genéricamente
]
# Aceptar cualquier tipo y validar manualmente
extracto_file = st.file_uploader("Subir Extracto Bancario (Excel)")  # Sin type=
if extracto_file:
    extension = extracto_file.name.split('.')[-1].lower()
    if extension not in ['xls', 'xlsx']:
        st.error(f"Formato no soportado para Extracto: {extension}. Usa .xls o .xlsx.")
        extracto_file = None

auxiliar_file = st.file_uploader("Subir Libro Auxiliar (Excel)")  # Sin type=
if auxiliar_file:
    extension = auxiliar_file.name.split('.')[-1].lower()
    if extension not in ['xls', 'xlsx']:
        st.error(f"Formato no soportado para Auxiliar: {extension}. Usa .xls o .xlsx.")
        auxiliar_file = None

# Inicializar estado de sesión
if 'invertir_signos' not in st.session_state:
    st.session_state.invertir_signos = False

def realizar_conciliacion(extracto_file, auxiliar_file, mes_conciliacion, invertir_signos, banco_seleccionado):
    # Definir columnas esperadas
    columnas_esperadas_extracto = {
        "fecha": ["fecha de operación", "fecha", "date", "fecha_operacion", "f. operación", "fecha de sistema"],
        "monto": ["importe (cop)", "monto", "amount", "importe", "valor total","*valor","*VALOR"],
        "concepto": ["concepto", "descripción", "concepto banco", "descripcion", "transacción", "transaccion", "descripción motivo"],
        "numero_movimiento": ["número de movimiento", "numero de movimiento", "movimiento", "no. movimiento", "num", "nro. documento", "documento"],
        "debitos": ["debitos", "débitos", "debe", "cargo", "cargos", "valor débito"],
        "creditos": ["creditos", "créditos", "haber", "abono", "abonos", "valor crédito"]
    }

    columnas_esperadas_auxiliar = {
        "fecha": ["fecha", "date", "fecha de operación", "fecha_operacion", "f. operación"],
        "debitos": ["debitos", "débitos", "debe", "cargo", "cargos", "valor débito"],
        "creditos": ["creditos", "créditos", "haber", "abono", "abonos", "valor crédito"],
        "nota": ["nota", "nota libro auxiliar", "descripción", "observaciones", "descripcion"],
        "numero_movimiento": ["doc num", "doc. num", "documento", "número documento", "numero documento", "nro. documento"],
        "tercero": ["tercero", "Tercero", "proveedor"]
    }

    # Leer datos
    extracto_df = leer_datos_desde_encabezados(extracto_file, columnas_esperadas_extracto, "Extracto Bancario")
    auxiliar_df = leer_datos_desde_encabezados(auxiliar_file, columnas_esperadas_auxiliar, "Libro Auxiliar")

    # Procesar montos
    auxiliar_df = procesar_montos(auxiliar_df, "Libro Auxiliar", es_extracto=False, banco_seleccionado="Generico")
    extracto_df = procesar_montos(
        extracto_df, "Extracto Bancario", es_extracto=True, invertir_signos=invertir_signos,
        banco_seleccionado=banco_seleccionado)

    # Estandarizar fechas
    auxiliar_df = estandarizar_fechas(auxiliar_df, "Libro Auxiliar", mes_conciliacion=None)
    extracto_df = estandarizar_fechas(extracto_df, "Extracto Bancario", mes_conciliacion=None, completar_anio=True, auxiliar_df=auxiliar_df)

    st.subheader("🕵️ Análisis de Datos Procesados del Extracto Bancario")
    st.info("Primeros 5 registros del Extracto Bancario después del procesamiento de encabezados, fechas y montos.")
    
    # Seleccionar las columnas clave y las originales de débito/crédito (si existen)
    columnas_clave = ['fecha', 'monto', 'concepto', 'numero_movimiento']
    columnas_opcionales = ['debitos', 'creditos']
    
    columnas_a_mostrar = [col for col in columnas_clave if col in extracto_df.columns]
    columnas_a_mostrar += [col for col in columnas_opcionales if col in extracto_df.columns and col not in columnas_a_mostrar]
    
    # Mostrar el DataFrame, incluyendo el tipo de datos (dtype)
    st.dataframe(
        extracto_df[columnas_a_mostrar].head(5),
        use_container_width=True
    )
    st.write(f"Tipos de datos (Dtypes) del Extracto Bancario: \n{extracto_df[columnas_a_mostrar].dtypes}")

    # Verificar si la columna 'monto' tiene valores diferentes de cero
    monto_cero = (extracto_df['monto'].abs() < 0.01).all() if 'monto' in extracto_df.columns else True
    
    if monto_cero:
        st.warning("⚠️ **Alerta:** La columna 'monto' parece ser cero o muy cercana a cero en todos los registros después de la conversión. Esto indica un posible problema con la interpretación de las columnas de Débitos/Créditos o con la lógica de signos.")

    # Filtrar por mes si se seleccionó
    if mes_conciliacion:
        extracto_df = estandarizar_fechas(extracto_df, "Extracto Bancario", mes_conciliacion=mes_conciliacion)
        auxiliar_df = estandarizar_fechas(auxiliar_df, "Libro Auxiliar", mes_conciliacion=mes_conciliacion)

    # Mostrar resúmenes
    st.subheader("Resumen de datos cargados")
    st.write(f"Extracto bancario: {len(extracto_df)} movimientos")
    st.write(f"Libro auxiliar: {len(auxiliar_df)} movimientos")

    # Realizar conciliación
    resultados_df = conciliar_banco_completo(extracto_df, auxiliar_df)
    
    return resultados_df, extracto_df, auxiliar_df

if extracto_file and auxiliar_file:
    try:
        # Realizar conciliación inicial
        resultados_df, extracto_df, auxiliar_df = realizar_conciliacion(
            extracto_file, auxiliar_file, mes_conciliacion, st.session_state.invertir_signos,
            banco_seleccionado=banco_seleccionado
        )

        # Depurar resultados
        if resultados_df['fecha'].isna().any():
            st.write("Filas con NaT en 'fecha':")
            st.write(resultados_df[resultados_df['fecha'].isna()])

        # Mostrar resultados
        st.subheader("Resultados de la Conciliación")
        conciliados = resultados_df[resultados_df['estado'] == 'Conciliado']
        no_conciliados = resultados_df[resultados_df['estado'] == 'No Conciliado']
        porcentaje_conciliados = len(conciliados) / len(resultados_df) * 100 if len(resultados_df) > 0 else 0
        
        st.write(f"Total de movimientos: {len(resultados_df)}")
        st.write(f"Movimientos conciliados: {len(conciliados)} ({porcentaje_conciliados:.1f}%)")
        st.write(f"Movimientos no conciliados: {len(no_conciliados)} ({len(no_conciliados)/len(resultados_df)*100:.1f}%)")

        # Distribución por tipo de conciliación
        st.write("Distribución por tipo de conciliación:")
        distribucion = resultados_df.groupby(['tipo_conciliacion', 'origen']).size().reset_index(name='subtotal')
        distribucion_pivot = distribucion.pivot_table(
            index='tipo_conciliacion', columns='origen', values='subtotal', fill_value=0
        ).reset_index()
        distribucion_pivot.columns = ['Tipo de Conciliación', 'Extracto Bancario', 'Libro Auxiliar']
        distribucion_pivot['Cantidad Total'] = distribucion_pivot['Extracto Bancario'] + distribucion_pivot['Libro Auxiliar']
        distribucion_pivot = distribucion_pivot[['Tipo de Conciliación', 'Extracto Bancario', 'Libro Auxiliar', 'Cantidad Total']]
        st.write(distribucion_pivot)

        st.write("Detalle de todos los movimientos:")
        st.write(resultados_df)

        # Generar Excel
        def generar_excel(resultados_df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                resultados_df.to_excel(writer, sheet_name="Resultados", index=False)
                aplicar_formato_excel(writer, resultados_df)
            output.seek(0)
            return output

        excel_data = generar_excel(resultados_df)
        st.download_button(
            label="Descargar Resultados en Excel",
            data=excel_data,
            file_name="resultados_conciliacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Mostrar botón si el porcentaje de conciliados es menor al 20%
        if porcentaje_conciliados < 20:
            st.warning("El porcentaje de movimientos conciliados es bajo. ¿Los signos de débitos/créditos están invertidos en el extracto?")
            if st.button("Invertir valores débitos y créditos en Extracto Bancario"):
                st.session_state.invertir_signos = not st.session_state.invertir_signos
                st.rerun()  # Forzar reejecución de la app

    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")
        st.exception(e)
else:
    st.info("Por favor, sube ambos archivos para comenzar la conciliación.")
