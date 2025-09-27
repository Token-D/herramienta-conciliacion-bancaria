import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations
from dateutil.parser import parse as parse_date
import re
from collections import Counter

# -----------------------
# Reglas de limpieza por banco (modular)
# -----------------------
def limpiar_bancolombia(serie):
    """Ej: 2,119,101.00 -> 2119101.00"""
    return (
        serie.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)   # quitar separador de miles
        # punto decimal ya está como '.', no tocar
    )

def limpiar_bogota_bbva(serie):
    """Ej: $14.339.827,00  ó  -2.699.434,00 -> 14339827.00 / -2699434.00"""
    return (
        serie.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(".", "", regex=False)   # quitar separador de miles
        .str.replace(",", ".", regex=False)  # convertir coma decimal a punto
    )

def limpiar_generico(serie):
    """
    Intento genérico: elimina espacios y símbolos comunes y luego intenta convertir.
    Útil como fallback para bancos nuevos hasta definir una regla específica.
    """
    return (
        serie.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace("$", "", regex=False)
        # por defecto quitamos comas para evitar textos con separador de miles
        .str.replace(",", "", regex=False)
    )

# Diccionario de reglas por banco (agrega nuevas funciones aquí)
REGLAS_BANCOS = {
    "Bancolombia": limpiar_bancolombia,
    "Banco de Bogotá": limpiar_bogota_bbva,
    "BBVA": limpiar_bogota_bbva,
    # Agrega nuevos bancos aquí: "Davivienda": limpiar_davivienda
}

# -----------------------
# Autodetección heurística del formato de montos (muestra)
# -----------------------
def detectar_banco_por_muestra(serie):
    """
    Intenta detectar formato (Bancolombia vs Bogotá/BBVA) a partir de una muestra de strings.
    Devuelve nombre del banco representativo ("Bancolombia" o "Banco de Bogotá") o None.
    """
    muestra = serie.dropna().astype(str).head(50).tolist()
    if not muestra:
        return None

    score_bancolombia = 0
    score_bogota = 0

    for s in muestra:
        s = s.strip()
        if not s:
            continue
        if "$" in s:
            score_bogota += 1
        last_dot = s.rfind(".")
        last_comma = s.rfind(",")
        # si ambos no existen, ignorar
        if last_dot == -1 and last_comma == -1:
            continue
        if last_dot > last_comma:
            # ejemplo: '2,119,101.00' -> dot al final -> punto decimal dominando
            score_bancolombia += 1
        elif last_comma > last_dot:
            # ejemplo: '14.339.827,00' -> comma decimal
            score_bogota += 1
        else:
            # empate o estructura extraña: usar conteo de separadores
            if s.count(",") >= 2 and "." in s:
                score_bancolombia += 1
            elif s.count(".") >= 2 and "," in s:
                score_bogota += 1

    if score_bancolombia > score_bogota:
        return "Bancolombia"
    if score_bogota > score_bancolombia:
        return "Banco de Bogotá"
    return None

# -----------------------
# Función para buscar la fila de encabezados
# -----------------------
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

        tiene_fecha = False
        tiene_monto = False

        for celda in celdas:
            if 'fecha' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['fecha']):
                tiene_fecha = True
            if 'monto' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['monto']):
                tiene_monto = True
            elif 'debitos' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['debitos']):
                tiene_monto = True
            elif 'creditos' in columnas_esperadas_lower and any(variante in celda for variante in columnas_esperadas_lower['creditos']):
                tiene_monto = True

        if tiene_fecha and tiene_monto:
            return idx

    return None

# -----------------------
# Función para leer datos a partir de la fila de encabezados
# -----------------------
def leer_datos_desde_encabezados(archivo, columnas_esperadas, nombre_archivo, max_filas=30):
    # Determinar la extensión del archivo
    extension = archivo.name.split('.')[-1].lower()
    
    # Si es .xls, convertir a .xlsx
    if extension == 'xls':
        try:
            df_temp = pd.read_excel(archivo, header=None, engine='xlrd')
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_temp.to_excel(writer, index=False, header=None)
            output.seek(0)
            archivo = output
            st.success(f"Conversión de {nombre_archivo} de .xls a .xlsx completada.")
        except Exception as e:
            st.error(f"Error al convertir {nombre_archivo} de .xls a .xlsx: {e}")
            st.stop()
    elif extension != 'xlsx':
        st.error(f"Formato de archivo no soportado: {extension}. Usa .xls o .xlsx.")
        st.stop()
        
    # Leer el archivo de Excel sin asumir encabezados
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
    variantes_doc_num = columnas_esperadas.get('Doc Num', ["Doc Num"])
    doc_num_col = None
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(variante.lower().strip() in col_lower for variante in variantes_doc_num):
            doc_num_col = col
            break
    
    # Filtrar filas donde 'Doc Num' no esté vacío
    if doc_num_col:
        filas_antes = len(df)
        df = df[df[doc_num_col].notna() & (df[doc_num_col] != '')]
        filas_despues = len(df)
    
    # Normalizar las columnas (renombra según columnas_esperadas)
    df = normalizar_dataframe(df, columnas_esperadas)
    
    # Verificar si el DataFrame tiene al menos las columnas mínimas necesarias
    if 'fecha' not in df.columns:
        st.error(f"La columna obligatoria 'fecha' no se encontró en los datos leídos del archivo '{nombre_archivo}'.")
        st.stop()
    
    if 'monto' not in df.columns and ('debitos' not in df.columns or 'creditos' not in df.columns):
        st.error(f"No se encontró ninguna columna de monto (monto, debitos o creditos) en el archivo '{nombre_archivo}'.")
        st.stop()
    
    return df

# -----------------------
# Función para normalizar un DataFrame
# -----------------------
def normalizar_dataframe(df, columnas_esperadas):
    """
    Normaliza un DataFrame para que use los nombres de columnas esperados.
    """
    df.columns = [str(col).lower().strip() for col in df.columns]
    
    # Crear un mapeo de nombres de columnas basado en las variantes
    mapeo_columnas = {}
    for col_esperada, variantes in columnas_esperadas.items():
        for variante in variantes:
            variante_lower = variante.lower().strip()
            mapeo_columnas[variante_lower] = col_esperada
    
    # Renombrar las columnas según el mapeo (buscando subcadenas)
    nuevo_nombres = []
    columnas_vistas = set()
    
    for col in df.columns:
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
    
    df.columns = nuevo_nombres
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    
    return df

# -----------------------
# Detección y estandarización de fechas
# -----------------------
def detectar_formato_fechas(fechas_str, porcentaje_analisis=0.6):
    fechas_validas = [f for f in fechas_str if pd.notna(f) and f.strip() and f not in ['nan', 'NaT']]
    if not fechas_validas:
        return "desconocido", False

    n_analizar = max(1, int(len(fechas_validas) * porcentaje_analisis))
    fechas_muestra = fechas_validas[:n_analizar]

    formatos = Counter()
    tiene_año = Counter()

    patron_fecha = r'^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$'

    for fecha in fechas_muestra:
        match = re.match(patron_fecha, fecha.replace('.', '/'))
        if not match:
            continue

        comp1, comp2, comp3 = match.groups()
        comp1, comp2 = int(comp1), int(comp2)
        año_presente = comp3 is not None
        tiene_año[año_presente] += 1

        if comp1 <= 12 and comp2 <= 31:
            if comp2 <= 12:
                formatos["ambiguo"] += 1
            else:
                formatos["MM/DD/AAAA"] += 1
        elif comp1 <= 31 and comp2 <= 12:
            formatos["DD/MM/AAAA"] += 1
        else:
            formatos["desconocido"] += 1

    formato_predominante = formatos.most_common(1)[0][0] if formatos else "desconocido"
    if formato_predominante == "ambiguo":
        formato_predominante = "DD/MM/AAAA"

    año_presente = tiene_año.most_common(1)[0][0] if tiene_año else False

    return formato_predominante, año_presente

def estandarizar_fechas(df, nombre_archivo, mes_conciliacion=None, completar_anio=False, auxiliar_df=None):
    if 'fecha' not in df.columns:
        st.warning(f"No se encontró la columna 'fecha' en {nombre_archivo}.")
        return df

    try:
        df['fecha_original'] = df['fecha'].copy()
        df['fecha_str'] = df['fecha'].astype(str).str.strip()

        año_base = None
        if completar_anio and auxiliar_df is not None and 'fecha' in auxiliar_df.columns:
            años_validos = auxiliar_df['fecha'].dropna().apply(lambda x: x.year if pd.notna(x) else None)
            año_base = años_validos.mode()[0] if not años_validos.empty else pd.Timestamp.now().year
        else:
            año_base = pd.Timestamp.now().year

        es_extracto = "Extracto" in nombre_archivo
        formato_fecha = "desconocido"
        año_presente = False
        if es_extracto:
            formato_fecha, año_presente = detectar_formato_fechas(df['fecha_str'])
            st.write(f"Formato de fecha detectado en {nombre_archivo}: {formato_fecha}, Año presente: {año_presente}")

        def parsear_fecha(fecha_str, mes_conciliacion=None, año_base=None, es_extracto=False, formato_fecha="desconocido"):
            if pd.isna(fecha_str) or fecha_str in ['', 'nan', 'NaT']:
                return pd.NaT

            try:
                fecha_str = fecha_str.replace('-', '/').replace('.', '/')

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
                        else:
                            dia, mes = comp2, comp1

                        if mes_conciliacion and 1 <= mes <= 12:
                            mes = mes_conciliacion

                        if 1 <= dia <= 31 and 1 <= mes <= 12:
                            return pd.Timestamp(year=año, month=mes, day=dia)

                parsed = parse_date(fecha_str, dayfirst=True, fuzzy=True)

                if es_extracto and mes_conciliacion and parsed.month != mes_conciliacion:
                    return pd.Timestamp(year=parsed.year, month=mes_conciliacion, day=parsed.day)

                return parsed
            except (ValueError, TypeError):
                try:
                    partes = fecha_str.split('/')
                    if len(partes) == 2:
                        comp1, comp2 = map(int, partes[:2])
                        if formato_fecha == "DD/MM/AAAA":
                            dia, mes = comp1, comp2
                        else:
                            dia, mes = comp2, comp1 if comp2 <= 31 and comp1 <= 12 else comp1, comp2

                        if es_extracto and mes_conciliacion:
                            mes = mes_conciliacion

                        if 1 <= dia <= 31 and 1 <= mes <= 12:
                            return pd.Timestamp(year=año_base, month=mes, day=dia)
                    return pd.NaT
                except (ValueError, IndexError):
                    return pd.NaT

        df['fecha'] = df['fecha_str'].apply(
            lambda x: parsear_fecha(x, mes_conciliacion, año_base, es_extracto, formato_fecha)
        )

        fechas_invalidas = df['fecha'].isna().sum()
        if fechas_invalidas > 0:
            st.warning(f"Se encontraron {fechas_invalidas} fechas inválidas en {nombre_archivo}.")
            st.write("Ejemplos de fechas inválidas:")
            st.write(df[df['fecha'].isna()][['fecha_original', 'fecha_str']].head())

        st.write(f"Fechas parseadas en {nombre_archivo} (primeras 10):")
        st.write(df[['fecha_original', 'fecha_str', 'fecha']].head(10))

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

        df = df.drop(['fecha_str'], axis=1, errors='ignore')

    except Exception as e:
        st.error(f"Error al estandarizar fechas en {nombre_archivo}: {e}")
        return df

    return df

# -----------------------
# Función para procesar montos (modular, aplica limpieza sólo en extractos)
# -----------------------
def procesar_montos(df, nombre_archivo, es_extracto=False, invertir_signos=False, banco=None):
    """
    Versión robusta y modular que:
      - Detecta columnas débito/credito o 'monto'
      - Si es extracto (es_extracto=True) aplica limpieza por banco (si se indicó)
      - Si es auxiliar (es_extracto=False) hace la conversión numérica simple
      - Convierte a numeric (float64) y crea/actualiza columna 'monto'
    """
    # Asegurar nombres como strings
    cols_orig = list(df.columns)
    cols_lower = [str(c).lower() for c in cols_orig]

    # Verificar si ya existe una columna 'monto' válida y numérica
    if "monto" in cols_lower:
        col_monto_name = cols_orig[cols_lower.index("monto")]
        try:
            df[col_monto_name] = pd.to_numeric(df[col_monto_name], errors="coerce")
            if df[col_monto_name].notna().any():
                # Normalizar nombre a 'monto' (minúscula)
                if col_monto_name != "monto":
                    df = df.rename(columns={col_monto_name: "monto"})
                return df
        except Exception:
            pass

    # términos para detectar columnas de debitos/creditos
    terminos_debitos = ["deb", "debe", "cargo", "débito", "valor débito", "debitos", "debitos"]
    terminos_creditos = ["cred", "haber", "abono", "crédito", "valor crédito", "creditos", "creditos"]

    cols_debito = [col for col in df.columns if any(term in str(col).lower() for term in terminos_debitos)]
    cols_credito = [col for col in df.columns if any(term in str(col).lower() for term in terminos_creditos)]

    # Si no hay columnas de monto, débitos ni créditos, advertir
    if not cols_debito and not cols_credito and "monto" not in cols_lower:
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

    # Determinar función de limpieza: sólo aplicable a extractos
    limpiar_func = None
    if es_extracto:
        if banco and banco in REGLAS_BANCOS:
            limpiar_func = REGLAS_BANCOS[banco]
        else:
            # intentar autodetect basándose en la primera columna encontrada
            muestra_col = None
            if cols_debito:
                muestra_col = df[cols_debito[0]].astype(str)
            elif cols_credito:
                muestra_col = df[cols_credito[0]].astype(str)
            elif "monto" in df.columns:
                muestra_col = df["monto"].astype(str)
            if muestra_col is not None:
                detected = detectar_banco_por_muestra(muestra_col)
                if detected and detected in REGLAS_BANCOS:
                    limpiar_func = REGLAS_BANCOS[detected]
    # Si no detectó y es extracto, usar generico como fallback
    if es_extracto and limpiar_func is None:
        limpiar_func = limpiar_generico

    # Procesar débitos
    for col in cols_debito:
        try:
            serie = df[col]
            if es_extracto and limpiar_func:
                serie_limpia = limpiar_func(serie.astype(str))
                serie_num = pd.to_numeric(serie_limpia, errors='coerce').fillna(0)
            else:
                serie_num = pd.to_numeric(serie, errors='coerce').fillna(0)
            df[col] = serie_num
            df["monto"] += serie_num * signo_debito
        except Exception as e:
            st.warning(f"Error al procesar columna de débito '{col}' en {nombre_archivo}: {e}")

    # Procesar créditos
    for col in cols_credito:
        try:
            serie = df[col]
            if es_extracto and limpiar_func:
                serie_limpia = limpiar_func(serie.astype(str))
                serie_num = pd.to_numeric(serie_limpia, errors='coerce').fillna(0)
            else:
                serie_num = pd.to_numeric(serie, errors='coerce').fillna(0)
            df[col] = serie_num
            df["monto"] += serie_num * signo_credito
        except Exception as e:
            st.warning(f"Error al procesar columna de crédito '{col}' en {nombre_archivo}: {e}")

    # Si aún no hay columnas de debitos/creditos pero hay columna 'monto' textual
    if ("monto" in df.columns) and not cols_debito and not cols_credito:
        try:
            serie = df["monto"]
            if es_extracto and limpiar_func:
                serie_limpia = limpiar_func(serie.astype(str))
                df["monto"] = pd.to_numeric(serie_limpia, errors='coerce').fillna(0)
            else:
                df["monto"] = pd.to_numeric(serie, errors='coerce').fillna(0)
        except Exception as e:
            st.warning(f"Error al convertir columna 'monto' en {nombre_archivo}: {e}")

    # Verificar resultado
    if df["monto"].eq(0).all() and (cols_debito or cols_credito):
        st.warning(f"La columna 'monto' resultó en ceros o no se detectaron valores numéricos en {nombre_archivo}.")

    return df

# -----------------------
# Función para encontrar combinaciones que sumen un monto específico
# -----------------------
def encontrar_combinaciones(df, monto_objetivo, tolerancia=0.01, max_combinacion=4):
    movimientos = []
    indices_validos = []
    
    for idx, valor in zip(df.index, df["monto"]):
        try:
            valor_num = float(valor)
            movimientos.append(valor_num)
            indices_validos.append(idx)
        except (ValueError, TypeError):
            continue
    
    if not movimientos:
        return []
    
    combinaciones_validas = []
    
    try:
        monto_objetivo = float(monto_objetivo)
    except (ValueError, TypeError):
        return []
    
    for r in range(1, min(max_combinacion, len(movimientos)) + 1):
        for combo_indices in combinations(range(len(movimientos)), r):
            combo_valores = [movimientos[i] for i in combo_indices]
            suma = sum(combo_valores)
            if abs(suma - monto_objetivo) <= tolerancia:
                indices_combinacion = [indices_validos[i] for i in combo_indices]
                combinaciones_validas.append((indices_combinacion, combo_valores))
    
    combinaciones_validas.sort(key=lambda x: len(x[0]))
    
    if combinaciones_validas:
        return combinaciones_validas[0][0]
    return []

# -----------------------
# Funciones de conciliación (directa y por agrupación)
# -----------------------
def conciliacion_directa(extracto_df, auxiliar_df):
    resultados = []
    extracto_conciliado_idx = set()
    auxiliar_conciliado_idx = set()
    
    extracto_df = extracto_df.copy()
    auxiliar_df = auxiliar_df.copy()
    extracto_df['fecha_solo'] = extracto_df['fecha'].dt.date
    auxiliar_df['fecha_solo'] = auxiliar_df['fecha'].dt.date
        
    for idx_extracto, fila_extracto in extracto_df.iterrows():
        if idx_extracto in extracto_conciliado_idx or pd.isna(fila_extracto['fecha_solo']):
            continue
        
        auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
        
        coincidencias = auxiliar_no_conciliado[
            (auxiliar_no_conciliado['fecha_solo'] == fila_extracto['fecha_solo']) & 
            (abs(auxiliar_no_conciliado['monto'] - fila_extracto['monto']) < 0.01)
        ]
        
        if not coincidencias.empty:
            idx_auxiliar = coincidencias.index[0]
            fila_auxiliar = coincidencias.iloc[0]
            
            extracto_conciliado_idx.add(idx_extracto)
            auxiliar_conciliado_idx.add(idx_auxiliar)
            
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
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
    
    for idx_extracto, fila_extracto in extracto_no_conciliado.iterrows():
        indices_combinacion = encontrar_combinaciones(
            auxiliar_no_conciliado, 
            fila_extracto["monto"],
            tolerancia=0.01
        )
        
        if indices_combinacion:
            nuevos_extracto_conciliado.add(idx_extracto)
            nuevos_auxiliar_conciliado.update(indices_combinacion)
            
            docs_conciliacion = auxiliar_no_conciliado.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            docs_conciliacion = [str(doc) for doc in docs_conciliacion]
            
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

def conciliacion_agrupacion_extracto(extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx):
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
    
    for idx_auxiliar, fila_auxiliar in auxiliar_no_conciliado.iterrows():
        indices_combinacion = encontrar_combinaciones(
            extracto_no_conciliado, 
            fila_auxiliar["monto"],
            tolerancia=0.01
        )
        
        if indices_combinacion:
            nuevos_auxiliar_conciliado.add(idx_auxiliar)
            nuevos_extracto_conciliado.update(indices_combinacion)
            
            nums_movimiento = extracto_no_conciliado.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            nums_movimiento = [str(num) for num in nums_movimiento]
            
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

# ===============================
# ===============================
# Función optimizada para conciliar sin cross merge
# ===============================
def conciliar(extracto, auxiliar, tolerancia=50):
    """
    Conciliación eficiente sin hacer cross merge.
    Busca coincidencias con tolerancia de manera iterativa,
    pero evitando explosión de combinaciones.
    """
    conciliados = []
    usados_aux = set()

    for valor_ext in extracto["VALOR"]:
        # Filtramos solo los auxiliares que no se han usado
        candidatos = auxiliar.loc[~auxiliar.index.isin(usados_aux), "VALOR"]

        # Calculamos diferencia absoluta
        diferencias = (candidatos - valor_ext).abs()

        # Buscamos el mejor match (mínima diferencia dentro de la tolerancia)
        if not diferencias.empty:
            idx_min = diferencias.idxmin()
            if diferencias[idx_min] <= tolerancia:
                conciliados.append((valor_ext, auxiliar.loc[idx_min, "VALOR"]))
                usados_aux.add(idx_min)

    # Conciliados dataframe
    conciliados_df = pd.DataFrame(conciliados, columns=["Extracto", "Auxiliar"])

    # No conciliados
    no_conciliados_ext = extracto.loc[~extracto["VALOR"].isin(conciliados_df["Extracto"])]
    no_conciliados_aux = auxiliar.loc[~auxiliar.index.isin(usados_aux)]

    return conciliados_df, no_conciliados_ext, no_conciliados_aux

# -----------------------
# Formato Excel (para descarga)
# -----------------------
def aplicar_formato_excel(writer, resultados_df):
    worksheet = writer.sheets['Resultados']
    workbook = writer.book
    
    formato_encabezado = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        'border': 1, 'bg_color': '#D9E1F2'
    })
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
                if pd.isna(valor):
                    worksheet.write(row_num, i, "", formato_fecha)
                else:
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
                        if col_idx == i:
                            worksheet.write(row_num, col_idx, estado, formato_no_conciliado)
                        else:
                            valor = resultados_df.iloc[row_num-1][resultados_df.columns[col_idx]]
                            if resultados_df.columns[col_idx].lower() == 'fecha':
                                formato_combinado = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#FFCCCB'})
                                if pd.isna(valor):
                                    worksheet.write(row_num, col_idx, "", formato_combinado)
                                else:
                                    worksheet.write_datetime(row_num, col_idx, valor, formato_combinado)
                            elif resultados_df.columns[col_idx].lower() == 'monto':
                                formato_combinado = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'bg_color': '#FFCCCB'})
                                if pd.isna(valor):
                                    worksheet.write(row_num, col_idx, "", formato_combinado)
                                else:
                                    worksheet.write_number(row_num, col_idx, valor, formato_combinado)
                            else:
                                worksheet.write(row_num, col_idx, valor, formato_no_conciliado)

# -----------------------
# Interfaz de Streamlit (principal)
# -----------------------
st.title("Herramienta de Conciliación Bancaria Automática")

st.subheader("Configuración")
meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_seleccionado = st.selectbox("Mes a conciliar (opcional):", ["Todos"] + meses)
mes_conciliacion = meses.index(mes_seleccionado) + 1 if mes_seleccionado != "Todos" else None

# Selección de banco (solo una vez para el extracto)
banco_extracto = st.selectbox(
    "Banco del extracto (elige o deja 'Auto-detect'):",
    ["Auto-detect", "Bancolombia", "Banco de Bogotá", "BBVA"]
)
# Normalizamos a None si el usuario eligió Auto-detect
banco_extracto = None if banco_extracto == "Auto-detect" else banco_extracto

# Accept file uploads (sin type= para dar flexibilidad, validamos manualmente)
extracto_file = st.file_uploader("Subir Extracto Bancario (Excel)")
if extracto_file:
    extension = extracto_file.name.split('.')[-1].lower()
    if extension not in ['xls', 'xlsx']:
        st.error(f"Formato no soportado para Extracto: {extension}. Usa .xls o .xlsx.")
        extracto_file = None

auxiliar_file = st.file_uploader("Subir Libro Auxiliar (Excel)")
if auxiliar_file:
    extension = auxiliar_file.name.split('.')[-1].lower()
    if extension not in ['xls', 'xlsx']:
        st.error(f"Formato no soportado para Auxiliar: {extension}. Usa .xls o .xlsx.")
        auxiliar_file = None

# Inicializar estado de sesión
if 'invertir_signos' not in st.session_state:
    st.session_state.invertir_signos = False

# -----------------------
# actualizar realizar_conciliacion para aceptar banco_extracto
# -----------------------
def realizar_conciliacion(extracto_file, auxiliar_file, mes_conciliacion, invertir_signos, banco_extracto=None):
    # Definir columnas esperadas
    columnas_esperadas_extracto = {
        "fecha": ["fecha de operación", "fecha", "date", "fecha_operacion", "f. operación", "fecha de sistema"],
        "monto": ["importe (cop)","valor", "monto", "amount", "importe", "valor total"],
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
    # - Auxiliar: no se le aplica limpieza por banco (estructura fija)
    auxiliar_df = procesar_montos(auxiliar_df, "Libro Auxiliar", es_extracto=False, invertir_signos=False, banco=None)
    # - Extracto: aplicar limpieza según banco seleccionado (o autodetect si banco_extracto es None)
    extracto_df = procesar_montos(extracto_df, "Extracto Bancario", es_extracto=True, invertir_signos=invertir_signos, banco=banco_extracto)

    # Estandarizar fechas
    auxiliar_df = estandarizar_fechas(auxiliar_df, "Libro Auxiliar", mes_conciliacion=None)
    extracto_df = estandarizar_fechas(extracto_df, "Extracto Bancario", mes_conciliacion=None, completar_anio=True, auxiliar_df=auxiliar_df)

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

# -----------------------
# Bloque principal: ejecutar conciliación cuando ambos archivos estén cargados
# -----------------------
if extracto_file and auxiliar_file:
    try:
        # Realizar conciliación inicial (pasamos banco_extracto aquí)
        resultados_df, extracto_df, auxiliar_df = realizar_conciliacion(
            extracto_file, auxiliar_file, mes_conciliacion, st.session_state.invertir_signos,
            banco_extracto=banco_extracto
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
        if not distribucion.empty:
            distribucion_pivot = distribucion.pivot_table(
                index='tipo_conciliacion', columns='origen', values='subtotal', fill_value=0
            ).reset_index()
            # Asegurar columnas
            extracto_col = 'Extracto Bancario' if 'Extracto Bancario' in distribucion_pivot.columns else (list(distribucion_pivot.columns[1:2])[0] if distribucion_pivot.shape[1]>1 else None)
            # Construir de forma segura
            try:
                distribucion_pivot.columns = ['Tipo de Conciliación', 'Extracto Bancario', 'Libro Auxiliar']
                distribucion_pivot['Cantidad Total'] = distribucion_pivot['Extracto Bancario'] + distribucion_pivot['Libro Auxiliar']
                distribucion_pivot = distribucion_pivot[['Tipo de Conciliación', 'Extracto Bancario', 'Libro Auxiliar', 'Cantidad Total']]
                st.write(distribucion_pivot)
            except Exception:
                st.write(distribucion)

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
