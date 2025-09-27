import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations
from dateutil.parser import parse as parse_date
import re
from collections import Counter
import numpy as np

# -----------------------------
# Helpers: parseo/normalización
# -----------------------------
def _parse_amount_string(s, banco_key):
    """
    Convierte una cadena de monto a float, según la regla del banco.
    Maneja: signos, paréntesis para negativos, símbolos de moneda, separadores de miles y decimales.
    banco_key: 'Bancolombia', 'Banco de Bogotá', 'BBVA' (BBVA y Banco de Bogotá usan la misma regla).
    """
    if pd.isna(s):
        return np.nan
    s = str(s).strip()

    if s == "" or s.lower() in ["nan", "nat"]:
        return np.nan

    # Detectar negativo por signo o por paréntesis
    negativo = False
    if s.startswith('-'):
        negativo = True
    if '(' in s and ')' in s:
        negativo = True

    # Limpiar caracteres no numéricos salvo separadores . , - ( )
    s_clean = re.sub(r'[^\d\-\.,()]', '', s)
    # eliminar paréntesis si existen
    s_clean = s_clean.replace('(', '').replace(')', '')

    # Si ya es un número simple (por ejemplo '1234.56' o '-1234.56'), intentar parse directo
    try:
        return float(s_clean)
    except:
        pass

    try:
        if banco_key == "Bancolombia":
            # Ejemplo: 2,119,101.00  -> quitar comas de miles, mantener punto decimal
            s_num = s_clean.replace(',', '')
        else:
            # Banco de Bogotá y BBVA (formato con . para miles y , para decimales)
            # Ejemplo: 14.339.827,00 -> quitar puntos de miles, convertir coma decimal a punto
            s_num = s_clean.replace('.', '').replace(',', '.')
        # Si queda algo apto para float:
        val = float(s_num) if s_num not in ["", "-", "."] else np.nan
    except Exception:
        # Fallbacks: intentar heurísticas
        # 1) Si hay coma y no punto, tratar coma como decimal
        if ',' in s_clean and '.' not in s_clean:
            try:
                val = float(s_clean.replace(',', '.'))
            except:
                val = np.nan
        # 2) Si hay punto y no coma, tratar punto como decimal
        elif '.' in s_clean and ',' not in s_clean:
            try:
                val = float(s_clean)
            except:
                val = np.nan
        else:
            # Último recurso: remover todo salvo dígitos y posible signo y parsear (puede perder decimales)
            only_digits = re.sub(r'[^\d\-]', '', s_clean)
            try:
                val = float(only_digits) if only_digits not in ["", "-", "."] else np.nan
            except:
                val = np.nan

    if pd.isna(val):
        return np.nan
    return -abs(val) if negativo and val > 0 else val


def detectar_formato_montos(series):
    """
    Intentar detectar si el formato usa coma decimal (ej: 1.234.567,00)
    o punto decimal (ej: 1,234,567.00) en una muestra de la serie.
    Devuelve 'dot_decimal' o 'comma_decimal' o 'unknown'.
    """
    muestra = series.dropna().astype(str).str.strip()
    muestra = muestra[muestra != ''].head(30)
    if muestra.empty:
        return "unknown"

    conteo = Counter()
    for s in muestra:
        s_clean = re.sub(r'[^\d\.,]', '', s)
        last_dot = s_clean.rfind('.')
        last_comma = s_clean.rfind(',')
        if last_dot == -1 and last_comma == -1:
            continue
        if last_dot > last_comma:
            conteo['dot_decimal'] += 1
        elif last_comma > last_dot:
            conteo['comma_decimal'] += 1

    if not conteo:
        return "unknown"
    return conteo.most_common(1)[0][0]


def limpiar_formato_montos_extracto(df_in, banco_seleccionado="Auto-detect"):
    """
    Normaliza las columnas de monto en el DataFrame del extracto.
    Solo afecta columnas: 'monto', 'debitos', 'creditos' (si existen).
    -> Convierte a float siguiendo la regla del banco seleccionado.
    Si banco_seleccionado == "Auto-detect", intenta inferir la convención.
    """
    df = df_in.copy()
    # Asegurar nombres en minúscula (tu código ya lo hace en normalizar_dataframe pero lo reforzamos)
    df.columns = [str(c).lower().strip() for c in df.columns]

    # Columnas a limpiar (si existen)
    columnas = []
    if 'monto' in df.columns:
        columnas.append('monto')
    if 'debitos' in df.columns:
        columnas.append('debitos')
    if 'creditos' in df.columns:
        columnas.append('creditos')

    if not columnas:
        return df

    # Determinar la "regla bancaria" a usar
    banco_key = banco_seleccionado
    if banco_seleccionado == "Auto-detect":
        # Preferimos usar 'monto' si existe, si no combinar débitos/creditos
        if 'monto' in df.columns:
            fmt = detectar_formato_montos(df['monto'])
        else:
            # concatenar pequeñas muestras de débito/crédito para detectar
            combined = pd.Series(dtype="object")
            if 'debitos' in df.columns:
                combined = combined.append(df['debitos'].dropna().astype(str).head(50), ignore_index=True)
            if 'creditos' in df.columns:
                combined = combined.append(df['creditos'].dropna().astype(str).head(50), ignore_index=True)
            fmt = detectar_formato_montos(combined)

        if fmt == 'dot_decimal':
            banco_key = "Bancolombia"
        elif fmt == 'comma_decimal':
            # suponer BBVA/Banco de Bogotá (mismo manejo)
            banco_key = "Banco de Bogotá"
        else:
            banco_key = "Banco de Bogotá"  # fallback conservador (coma decimal frecuente en Colombia)

    # Aplicar parsing (vectorizado con apply por columna)
    for col in columnas:
        try:
            df[col] = df[col].apply(lambda x: _parse_amount_string(x, banco_key))
        except Exception as e:
            # Como fallback intentar coerción genérica
            try:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[^\d\-,\.]', '', regex=True).str.replace(',', '.', regex=False), errors='coerce')
            except Exception:
                df[col] = pd.to_numeric(df[col], errors='coerce')

    return df

# -----------------------------
# (Tu código original sigue igual, con la limpieza insertada antes de procesar montos)
# -----------------------------

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
        st.write(f"Fechas parseadas en {nombre_archivo} (primeras 10):")
        st.write(df[['fecha_original', 'fecha_str', 'fecha']].head(10))

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
def procesar_montos(df, nombre_archivo, es_extracto=False, invertir_signos=False):
    """
    Procesa columnas de débitos y créditos para crear una columna 'monto' unificada.
    Para extractos: débitos son negativos, créditos positivos.
    Para auxiliar: débitos son positivos, créditos negativos.
    """
    columnas = df.columns.str.lower()

    # Verificar si ya existe una columna 'monto' válida
    if "monto" in columnas and df["monto"].notna().any() and (df["monto"] != 0).any():
        return df

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
        signo_debito = 1  # Auxiliar no invierte signos
        signo_credito = -1

    # Procesar débitos
    for col in cols_debito:
        try:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df["monto"] += df[col] * signo_debito
        except Exception as e:
            st.warning(f"Error al procesar columna de débito '{col}' en {nombre_archivo}: {e}")

    # Procesar créditos
    for col in cols_credito:
        try:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df["monto"] += df[col] * signo_credito
        except Exception as e:
            st.warning(f"Error al procesar columna de crédito '{col}' en {nombre_archivo}: {e}")

    # Verificar resultado
    if df["monto"].eq(0).all() and (cols_debito or cols_credito):
        st.warning(f"La columna 'monto' resultó en ceros en {nombre_archivo}. Verifica las columnas de débitos/créditos.")

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

    return resultados_directa


# ===============================
# Interfaz Streamlit
# ===============================
st.set_page_config(page_title="Conciliación Bancaria", layout="wide")
st.title("📊 Herramienta de Conciliación Bancaria")

st.subheader("Configuración")
meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_seleccionado = st.selectbox("Mes a conciliar (opcional):", ["Todos"] + meses)
mes_conciliacion = meses.index(mes_seleccionado) + 1 if mes_seleccionado != "Todos" else None

# ---------- NUEVO: Selección del banco (solo una vez) ----------
banco_seleccionado = st.selectbox(
    "Seleccione el banco del extracto (esto aplica solo al extracto bancario):",
    ["Auto-detect", "Bancolombia", "Banco de Bogotá", "BBVA"],
    index=0
)
# ----------------------------------------------------------------

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

    # ---------- LIMPIEZA ESPECIAL PARA EXTRACTO segun banco ----------
    try:
        extracto_df = limpiar_formato_montos_extracto(extracto_df, banco_seleccionado)
    except Exception as e:
        st.warning(f"No se pudo normalizar automáticamente los montos del extracto: {e}")
    # ----------------------------------------------------------------

    # Procesar montos
    auxiliar_df = procesar_montos(auxiliar_df, "Libro Auxiliar", es_extracto=False)
    extracto_df = procesar_montos(extracto_df, "Extracto Bancario", es_extracto=True, invertir_signos=invertir_signos)

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

if extracto_file and auxiliar_file:
    try:
        # Realizar conciliación inicial
        resultados_df, extracto_df, auxiliar_df = realizar_conciliacion(
            extracto_file, auxiliar_file, mes_conciliacion, st.session_state.invertir_signos, banco_seleccionado
        )

        # Depurar resultados
        if resultados_df['fecha'].isna().any():
            st.write("Filas con NaT en 'fecha' del extracto:")
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
