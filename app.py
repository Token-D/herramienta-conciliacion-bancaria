import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from itertools import combinations
from dateutil.parser import parse as parse_date
import re
from collections import Counter

# -----------------------------
# Helpers: parseo/normalización
# -----------------------------
def _parse_amount_string(s, banco_key):
    """Convierte una cadena de monto a float, según la regla del banco."""
    if pd.isna(s) or str(s).strip().lower() in ["", "nan", "nat"]:
        return np.nan
    s = str(s).strip()
    negativo = s.startswith('-') or ('(' in s and ')' in s)
    s_clean = re.sub(r'[^\d\-\.,()]', '', s).replace('(', '').replace(')', '')
    
    try:
        if banco_key == "Bancolombia":
            s_num = s_clean.replace(',', '')
        else:
            s_num = s_clean.replace('.', '').replace(',', '.')
        val = float(s_num) if s_num not in ["", "-", "."] else np.nan
    except:
        if ',' in s_clean and '.' not in s_clean:
            val = float(s_clean.replace(',', '.')) if s_clean.replace(',', '.') else np.nan
        elif '.' in s_clean and ',' not in s_clean:
            val = float(s_clean) if s_clean else np.nan
        else:
            only_digits = re.sub(r'[^\d\-]', '', s_clean)
            val = float(only_digits) if only_digits not in ["", "-"] else np.nan
    
    return -abs(val) if negativo and not pd.isna(val) and val > 0 else val

def limpiar_formato_montos_extracto(df_in, banco_seleccionado="Auto-detect"):
    """Normaliza las columnas de monto en el DataFrame del extracto."""
    df = df_in.copy()
    df.columns = [str(c).lower().strip() for c in df.columns]
    columnas = [col for col in ['monto', 'debitos', 'creditos'] if col in df.columns]
    
    if not columnas:
        return df

    # Determinar formato de monto
    banco_key = banco_seleccionado
    if banco_seleccionado == "Auto-detect":
        combined = pd.Series(dtype="object")
        for col in columnas:
            combined = pd.concat([combined, df[col].dropna().astype(str).head(50)], ignore_index=True)
        fmt = detectar_formato_montos(combined)
        banco_key = "Bancolombia" if fmt == 'dot_decimal' else "Banco de Bogotá"

    # Vectorizar la conversión de montos
    for col in columnas:
        try:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(r'[^\d\-\.,]', '', regex=True)
                .str.replace(',', '.' if banco_key != "Bancolombia" else '')
                .str.replace('.', '' if banco_key != "Bancolombia" else '.', regex=False),
                errors='coerce'
            )
            df[col] = df[col].where(~df[col].isna(), df[col].apply(lambda x: _parse_amount_string(x, banco_key)))
        except Exception as e:
            st.warning(f"Error al normalizar columna '{col}': {e}")
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df

def detectar_formato_montos(series):
    """Detecta si el formato usa coma o punto decimal."""
    muestra = series.dropna().astype(str).str.strip().head(30)
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
    
    return conteo.most_common(1)[0][0] if conteo else "unknown"

# -----------------------------
# Funciones de lectura y normalización
# -----------------------------
def buscar_fila_encabezados(df, columnas_esperadas, max_filas=30):
    """Busca la fila que contiene al menos 'fecha' y una columna de monto."""
    columnas_esperadas_lower = {col: [v.lower() for v in variantes] for col, variantes in columnas_esperadas.items()}
    
    for idx in range(min(max_filas, len(df))):
        fila = df.iloc[idx]
        celdas = [str(v).lower() for v in fila if pd.notna(v)]
        tiene_fecha = any(any(v in c for v in columnas_esperadas_lower['fecha']) for c in celdas)
        tiene_monto = any(
            any(v in c for v in columnas_esperadas_lower.get('monto', [])) or
            any(v in c for v in columnas_esperadas_lower.get('debitos', [])) or
            any(v in c for v in columnas_esperadas_lower.get('creditos', []))
            for c in celdas
        )
        if tiene_fecha and tiene_monto:
            return idx
    return None

def leer_datos_desde_encabezados(archivo, columnas_esperadas, nombre_archivo, max_filas=30):
    """Lee el archivo Excel y normaliza los datos."""
    extension = archivo.name.split('.')[-1].lower()
    
    if extension == 'xls':
        try:
            df_temp = pd.read_excel(archivo, header=None, engine='xlrd')
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_temp.to_excel(writer, index=False, header=None)
            archivo = output
            output.seek(0)
            st.success(f"Conversión de {nombre_archivo} de .xls a .xlsx completada.")
        except Exception as e:
            st.error(f"Error al convertir {nombre_archivo}: {e}")
            st.stop()
    elif extension != 'xlsx':
        st.error(f"Formato no soportado: {extension}. Usa .xls o .xlsx.")
        st.stop()
    
    df = pd.read_excel(archivo, header=None, engine='openpyxl')
    fila_encabezados = buscar_fila_encabezados(df, columnas_esperadas, max_filas)
    if fila_encabezados is None:
        st.error(f"No se encontraron encabezados en {nombre_archivo}.")
        st.stop()
    
    df = pd.read_excel(archivo, header=fila_encabezados, engine='openpyxl')
    variantes_doc_num = columnas_esperadas.get('Doc Num', ["Doc Num"])
    doc_num_col = next((col for col in df.columns if any(v.lower().strip() in str(col).lower().strip() for v in variantes_doc_num)), None)
    
    if doc_num_col:
        df = df[df[doc_num_col].notna() & (df[doc_num_col] != '')]
    
    df = normalizar_dataframe(df, columnas_esperadas)
    
    if 'fecha' not in df.columns:
        st.error(f"No se encontró la columna 'fecha' en {nombre_archivo}.")
        st.stop()
    if 'monto' not in df.columns and not ('debitos' in df.columns or 'creditos' in df.columns):
        st.error(f"No se encontró ninguna columna de monto en {nombre_archivo}.")
        st.stop()
    
    return df

def normalizar_dataframe(df, columnas_esperadas):
    """Normaliza los nombres de las columnas."""
    df.columns = [str(col).lower().strip() for col in df.columns]
    mapeo_columnas = {v.lower().strip(): k for k, vs in columnas_esperadas.items() for v in vs}
    nuevo_nombres = []
    columnas_vistas = set()
    
    for col in df.columns:
        col_encontrada = False
        for variante, nombre_esperado in mapeo_columnas.items():
            if variante in col and nombre_esperado not in columnas_vistas:
                nuevo_nombres.append(nombre_esperado)
                columnas_vistas.add(nombre_esperado)
                col_encontrada = True
                break
        if not col_encontrada:
            nuevo_nombres.append(col)
    
    df.columns = nuevo_nombres
    return df.loc[:, ~df.columns.duplicated(keep='first')]

def detectar_formato_fechas(fechas_str, porcentaje_analisis=0.6):
    """Detecta el formato de fecha predominante."""
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
        tiene_año[comp3 is not None] += 1
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
    """Convierte la columna 'fecha' a datetime64."""
    if 'fecha' not in df.columns:
        st.warning(f"No se encontró la columna 'fecha' en {nombre_archivo}.")
        return df
    
    df['fecha_original'] = df['fecha'].copy()
    df['fecha_str'] = df['fecha'].astype(str).str.strip()
    
    año_base = pd.Timestamp.now().year
    if completar_anio and auxiliar_df is not None and 'fecha' in auxiliar_df.columns:
        años_validos = auxiliar_df['fecha'].dropna().apply(lambda x: x.year if pd.notna(x) else None)
        año_base = años_validos.mode()[0] if not años_validos.empty else año_base
    
    es_extracto = "Extracto" in nombre_archivo
    formato_fecha, año_presente = detectar_formato_fechas(df['fecha_str']) if es_extracto else ("desconocido", False)
    
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
                    dia, mes = (comp1, comp2) if formato_fecha == "DD/MM/AAAA" else (comp2, comp1)
                    if mes_conciliacion and 1 <= mes <= 12:
                        mes = mes_conciliacion
                    if 1 <= dia <= 31 and 1 <= mes <= 12:
                        return pd.Timestamp(year=año, month=mes, day=dia)
            parsed = parse_date(fecha_str, dayfirst=True, fuzzy=True)
            if es_extracto and mes_conciliacion and parsed.month != mes_conciliacion:
                return pd.Timestamp(year=parsed.year, month=mes_conciliacion, day=parsed.day)
            return parsed
        except:
            try:
                partes = fecha_str.split('/')
                if len(partes) == 2:
                    comp1, comp2 = map(int, partes[:2])
                    dia, mes = (comp1, comp2) if formato_fecha == "DD/MM/AAAA" else (comp2, comp1 if comp2 <= 31 and comp1 <= 12 else comp1, comp2)
                    if es_extracto and mes_conciliacion:
                        mes = mes_conciliacion
                    if 1 <= dia <= 31 and 1 <= mes <= 12:
                        return pd.Timestamp(year=año_base, month=mes, day=dia)
                return pd.NaT
            except:
                return pd.NaT
    
    df['fecha'] = df['fecha_str'].apply(lambda x: parsear_fecha(x, mes_conciliacion, año_base, es_extracto, formato_fecha))
    
    fechas_invalidas = df['fecha'].isna().sum()
    if fechas_invalidas > 0:
        st.warning(f"Se encontraron {fechas_invalidas} fechas inválidas en {nombre_archivo}.")
        st.write(df[df['fecha'].isna()][['fecha_original', 'fecha_str']].head())
    
    st.write(f"Fechas parseadas en {nombre_archivo} (primeras 10):")
    st.write(df[['fecha_original', 'fecha_str', 'fecha']].head(10))
    
    if mes_conciliacion and es_extracto:
        filas_antes = len(df)
        df = df[df['fecha'].dt.month == mes_conciliacion]
        if len(df) < filas_antes:
            meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            st.info(f"Se filtraron {filas_antes - len(df)} registros fuera del mes {meses[mes_conciliacion-1]} en {nombre_archivo}.")
    
    df = df.drop(['fecha_str'], axis=1, errors='ignore')
    return df

def procesar_montos(df, nombre_archivo, es_extracto=False, invertir_signos=False):
    """Procesa columnas de débitos y créditos para crear una columna 'monto' unificada."""
    columnas = df.columns.str.lower()
    if "monto" in columnas and df["monto"].notna().any() and (df["monto"] != 0).any():
        return df
    
    terminos_debitos = ["deb", "debe", "cargo", "débito", "valor débito"]
    terminos_creditos = ["cred", "haber", "abono", "crédito", "valor crédito"]
    cols_debito = [col for col in df.columns if any(term in col.lower() for term in terminos_debitos)]
    cols_credito = [col for col in df.columns if any(term in col.lower() for term in terminos_creditos)]
    
    if not cols_debito and not cols_credito and "monto" not in columnas:
        st.warning(f"No se encontraron columnas de monto, débitos o créditos en {nombre_archivo}.")
        return df
    
    df["monto"] = 0.0
    signo_debito = 1 if es_extracto and invertir_signos else -1 if es_extracto else 1
    signo_credito = -1 if es_extracto and invertir_signos else 1 if es_extracto else -1
    
    for col in cols_debito:
        try:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df["monto"] += df[col] * signo_debito
        except Exception as e:
            st.warning(f"Error al procesar columna de débito '{col}' en {nombre_archivo}: {e}")
    
    for col in cols_credito:
        try:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df["monto"] += df[col] * signo_credito
        except Exception as e:
            st.warning(f"Error al procesar columna de crédito '{col}' en {nombre_archivo}: {e}")
    
    if df["monto"].eq(0).all() and (cols_debito or cols_credito):
        st.warning(f"La columna 'monto' resultó en ceros en {nombre_archivo}. Verifica las columnas de débitos/créditos.")
    
    return df

def encontrar_combinaciones(df, monto_objetivo, tolerancia=0.01, max_combinacion=4):
    """Encuentra combinaciones de valores en df['monto'] que sumen aproximadamente monto_objetivo."""
    movimientos = []
    indices_validos = []
    
    for idx, valor in zip(df.index, df["monto"]):
        try:
            valor_num = float(valor)
            movimientos.append(valor_num)
            indices_validos.append(idx)
        except:
            continue
    
    if not movimientos:
        return []
    
    try:
        monto_objetivo = float(monto_objetivo)
    except:
        return []
    
    combinaciones_validas = []
    for r in range(1, min(max_combinacion, len(movimientos)) + 1):
        for combo_indices in combinations(range(len(movimientos)), r):
            combo_valores = [movimientos[i] for i in combo_indices]
            if abs(sum(combo_valores) - monto_objetivo) <= tolerancia:
                indices_combinacion = [indices_validos[i] for i in combo_indices]
                combinaciones_validas.append((indices_combinacion, combo_valores))
    
    combinaciones_validas.sort(key=lambda x: len(x[0]))
    return combinaciones_validas[0][0] if combinaciones_validas else []

def conciliacion_directa(extracto_df, auxiliar_df):
    """Realiza conciliación directa (uno a uno)."""
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
    
    return pd.DataFrame(resultados), extracto_conciliado_idx, auxiliar_conciliado_idx

def conciliacion_agrupacion_auxiliar(extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx):
    """Busca grupos de valores en el libro auxiliar que sumen un monto del extracto."""
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
    
    for idx_extracto, fila_extracto in extracto_no_conciliado.iterrows():
        indices_combinacion = encontrar_combinaciones(auxiliar_no_conciliado, fila_extracto["monto"])
        if indices_combinacion:
            nuevos_extracto_conciliado.add(idx_extracto)
            nuevos_auxiliar_conciliado.update(indices_combinacion)
            docs_conciliacion = auxiliar_no_conciliado.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            
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
    """Busca grupos de valores en el extracto que sumen un monto del libro auxiliar."""
    resultados = []
    nuevos_extracto_conciliado = set()
    nuevos_auxiliar_conciliado = set()
    
    extracto_no_conciliado = extracto_df[~extracto_df.index.isin(extracto_conciliado_idx)]
    auxiliar_no_conciliado = auxiliar_df[~auxiliar_df.index.isin(auxiliar_conciliado_idx)]
    
    for idx_auxiliar, fila_auxiliar in auxiliar_no_conciliado.iterrows():
        indices_combinacion = encontrar_combinaciones(extracto_no_conciliado, fila_auxiliar["monto"])
        if indices_combinacion:
            nuevos_auxiliar_conciliado.add(idx_auxiliar)
            nuevos_extracto_conciliado.update(indices_combinacion)
            nums_movimiento = extracto_no_conciliado.loc[indices_combinacion, "numero_movimiento"].astype(str).tolist()
            
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

def conciliar_banco_completo(extracto_df, auxiliar_df):
    """Implementa la lógica completa de conciliación."""
    resultados_directa, extracto_conciliado_idx, auxiliar_conciliado_idx = conciliacion_directa(
        extracto_df, auxiliar_df
    )
    
    resultados_agrup_aux, nuevos_extracto_conc1, nuevos_auxiliar_conc1 = conciliacion_agrupacion_auxiliar(
        extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx
    )
    
    extracto_conciliado_idx.update(nuevos_extracto_conc1)
    auxiliar_conciliado_idx.update(nuevos_auxiliar_conc1)
    
    resultados_agrup_ext, nuevos_extracto_conc2, nuevos_auxiliar_conc2 = conciliacion_agrupacion_extracto(
        extracto_df, auxiliar_df, extracto_conciliado_idx, auxiliar_conciliado_idx
    )
    
    if not resultados_directa.empty:
        indices_a_eliminar = []
        for idx, fila in resultados_directa.iterrows():
            if fila['estado'] == 'No Conciliado':
                if (fila['tipo_registro'] == 'extracto' and fila['index_original'] in nuevos_extracto_conc1.union(nuevos_extracto_conc2)) or \
                   (fila['tipo_registro'] == 'auxiliar' and fila['index_original'] in nuevos_auxiliar_conc1.union(nuevos_auxiliar_conc2)):
                    indices_a_eliminar.append(idx)
        if indices_a_eliminar:
            resultados_directa = resultados_directa.drop(indices_a_eliminar)
    
    resultados_finales = pd.concat([
        resultados_directa,
        resultados_agrup_aux,
        resultados_agrup_ext
    ], ignore_index=True)
    
    if 'index_original' in resultados_finales.columns:
        resultados_finales = resultados_finales.drop(['index_original', 'tipo_registro'], axis=1)
    
    return resultados_finales

def aplicar_formato_excel(writer, resultados_df):
    """Aplica formato al archivo Excel generado."""
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
        worksheet.write(0, i, col, formato_encabezado)
    
    worksheet.freeze_panes(1, 0)
    
    for row_num, row_data in resultados_df.iterrows():
        for col_idx, col_name in enumerate(resultados_df.columns):
            valor = row_data[col_name]
            if col_name.lower() == 'fecha':
                if pd.isna(valor):
                    worksheet.write(row_num + 1, col_idx, "", formato_fecha)
                else:
                    worksheet.write_datetime(row_num + 1, col_idx, valor, formato_fecha)
            elif col_name.lower() == 'monto':
                if pd.isna(valor):
                    worksheet.write(row_num + 1, col_idx, "", formato_moneda)
                else:
                    worksheet.write_number(row_num + 1, col_idx, valor, formato_moneda)
            elif col_name.lower() == 'estado' and valor == 'No Conciliado':
                for c in range(len(resultados_df.columns)):
                    v = resultados_df.iloc[row_num][resultados_df.columns[c]]
                    if resultados_df.columns[c].lower() == 'fecha':
                        formato_combinado = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#FFCCCB'})
                        worksheet.write(row_num + 1, c, v if pd.notna(v) else "", formato_combinado)
                    elif resultados_df.columns[c].lower() == 'monto':
                        formato_combinado = workbook.add_format({'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', 'bg_color': '#FFCCCB'})
                        worksheet.write(row_num + 1, c, v if pd.notna(v) else "", formato_combinado)
                    else:
                        worksheet.write(row_num + 1, c, v if pd.notna(v) else "", formato_no_conciliado)
            else:
                worksheet.write(row_num + 1, col_idx, valor if pd.notna(valor) else "")

# -----------------------------
# Interfaz Streamlit
# -----------------------------
st.set_page_config(page_title="Conciliación Bancaria", layout="wide")
st.title("📊 Herramienta de Conciliación Bancaria")

st.subheader("Configuración")
meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
mes_seleccionado = st.selectbox("Mes a conciliar (opcional):", ["Todos"] + meses)
mes_conciliacion = meses.index(mes_seleccionado) + 1 if mes_seleccionado != "Todos" else None

banco_seleccionado = st.selectbox(
    "Seleccione el banco del extracto:",
    ["Auto-detect", "Bancolombia", "Banco de Bogotá", "BBVA"],
    index=0
)

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

if 'invertir_signos' not in st.session_state:
    st.session_state.invertir_signos = False

def realizar_conciliacion(extracto_file, auxiliar_file, mes_conciliacion, invertir_signos, banco_seleccionado):
    columnas_esperadas_extracto = {
        "fecha": ["fecha de operación", "fecha", "date", "fecha_operacion", "f. operación", "fecha de sistema"],
        "monto": ["importe (cop)", "monto", "amount", "importe", "valor total"],
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

    extracto_df = leer_datos_desde_encabezados(extracto_file, columnas_esperadas_extracto, "Extracto Bancario")
    auxiliar_df = leer_datos_desde_encabezados(auxiliar_file, columnas_esperadas_auxiliar, "Libro Auxiliar")
    
    extracto_df = limpiar_formato_montos_extracto(extracto_df, banco_seleccionado)
    auxiliar_df = limpiar_formato_montos_extracto(auxiliar_df, banco_seleccionado)
    
    auxiliar_df = procesar_montos(auxiliar_df, "Libro Auxiliar", es_extracto=False)
    extracto_df = procesar_montos(extracto_df, "Extracto Bancario", es_extracto=True, invertir_signos=invertir_signos)
    
    auxiliar_df = estandarizar_fechas(auxiliar_df, "Libro Auxiliar", mes_conciliacion=None)
    extracto_df = estandarizar_fechas(extracto_df, "Extracto Bancario", mes_conciliacion=None, completar_anio=True, auxiliar_df=auxiliar_df)
    
    if mes_conciliacion:
        extracto_df = estandarizar_fechas(extracto_df, "Extracto Bancario", mes_conciliacion=mes_conciliacion)
        auxiliar_df = estandarizar_fechas(auxiliar_df, "Libro Auxiliar", mes_conciliacion=mes_conciliacion)
    
    st.subheader("Resumen de datos cargados")
    st.write(f"Extracto bancario: {len(extracto_df)} movimientos")
    st.write(f"Libro auxiliar: {len(auxiliar_df)} movimientos")
    
    resultados_df = conciliar_banco_completo(extracto_df, auxiliar_df)
    return resultados_df, extracto_df, auxiliar_df

if extracto_file and auxiliar_file:
    try:
        resultados_df, extracto_df, auxiliar_df = realizar_conciliacion(
            extracto_file, auxiliar_file, mes_conciliacion, st.session_state.invertir_signos, banco_seleccionado
        )
        
        if resultados_df['fecha'].isna().any():
            st.write("Filas con NaT en 'fecha':")
            st.write(resultados_df[resultados_df['fecha'].isna()])
        
        st.subheader("Resultados de la Conciliación")
        conciliados = resultados_df[resultados_df['estado'] == 'Conciliado']
        no_conciliados = resultados_df[resultados_df['estado'] == 'No Conciliado']
        num_conciliados = len(conciliados) // 2 if len(conciliados) % 2 == 0 else len(conciliados)
        porcentaje_conciliados = (num_conciliados / len(resultados_df)) * 100 if len(resultados_df) > 0 else 0
        
        st.write(f"Total de movimientos: {len(resultados_df)}")
        st.write(f"Movimientos conciliados: {num_conciliados} ({porcentaje_conciliados:.1f}%)")
        st.write(f"Movimientos no conciliados: {len(no_conciliados)} ({len(no_conciliados)/len(resultados_df)*100:.1f}%)")
        
        st.write("Distribución por tipo de conciliación:")
        distribucion = resultados_df.groupby(['tipo_conciliacion', 'origen']).size().reset_index(name='subtotal')
        distribucion_pivot = distribucion.pivot_table(
            index='tipo_conciliacion', columns='origen', values='subtotal', fill_value=0
        ).reset_index()
        distribucion_pivot.columns = ['Tipo de Conciliación', 'Extracto Bancario', 'Libro Auxiliar']
        distribucion_pivot['Cantidad Total'] = distribucion_pivot['Extracto Bancario'] + distribucion_pivot['Libro Auxiliar']
        distribucion_pivot.loc[distribucion_pivot['Tipo de Conciliación'] == 'Directa', 'Cantidad Total'] = distribucion_pivot.loc[distribucion_pivot['Tipo de Conciliación'] == 'Directa', ['Extracto Bancario', 'Libro Auxiliar']].max(axis=1)
        distribucion_pivot = distribucion_pivot[['Tipo de Conciliación', 'Extracto Bancario', 'Libro Auxiliar', 'Cantidad Total']]
        st.write(distribucion_pivot)
        
        st.write("Detalle de todos los movimientos:")
        st.write(resultados_df)
        
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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if porcentaje_conciliados < 20:
            st.warning("El porcentaje de movimientos conciliados es bajo. ¿Los signos de débitos/créditos están invertidos en el extracto?")
            if st.button("Invertir valores débitos y créditos en Extracto Bancario"):
                st.session_state.invertir_signos = not st.session_state.invertir_signos
                st.rerun()
    
    except Exception as e:
        st.error(f"Error al procesar los archivos: {e}")
        st.exception(e)
else:
    st.info("Por favor, sube ambos archivos para comenzar la conciliación.")
