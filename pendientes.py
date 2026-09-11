# -*- coding: utf-8 -*-
"""
pendientes.py
─────────────
Maneja el archivo Excel de movimientos pendientes (partidas abiertas) de conciliación.

El archivo se guarda en una ruta de red compartida y contiene UNA HOJA por empresa.
Cada vez que se genera un reporte FINAL, los movimientos que no fueron conciliados
(solo banco / solo conta) se guardan en la hoja de esa empresa, REEMPLAZANDO los
datos anteriores.

Cada vez que se genera un reporte INICIAL, los movimientos guardados de la conciliación
anterior se cargan automáticamente y se inyectan en los DataFrames de banco y conta,
para que el sistema los tome en cuenta en el nuevo cruce.

Estructura del archivo Excel de pendientes:
  - Una hoja por empresa (nombre normalizado, ej. "STN", "ITS", "CMT", ...)
  - Columnas de banco:  ORIGEN | Fecha | Descripción | Monto | Saldo | Sucursal | # Operación | Hora | Usuario
  - Columnas de conta:  ORIGEN | # Registro | Fecha | Medio Pago | # Operación | Giro | Glosa | Ingreso | Egreso | F. Conciliación | Conciliado
  - Columna ORIGEN: "BANCO" o "CONTA" para distinguir la procedencia de cada fila

Ruta de red:
  \\\\192.168.30.36\\Sig\\Asistentes Contables 2019\\ARCHIVO CONTABLE DIGITAL\\CONCILIACION BANCARIA
"""

import pandas as pd
from pathlib import Path
import unicodedata
import re
import warnings

# ── Ruta de red y nombre del archivo ──────────────────────────────────
RUTA_RED = Path(r"\\192.168.30.36\Sig\Asistentes Contables 2019\ARCHIVO CONTABLE DIGITAL\CONCILIACION BANCARIA")
NOMBRE_ARCHIVO    = "Mov_sin_conciliar.xlsx"
NOMBRE_CONCILIADOS = "Ingreso_a_contanet.xlsx"
RUTA_ARCHIVO      = RUTA_RED / NOMBRE_ARCHIVO

# ── Columnas esperadas para cada origen ──────────────────────────────
COLS_BANCO = [
    'ORIGEN', 'fecha', 'descripcion', 'monto', 'saldo',
    'sucursal', 'nro_operacion', 'hora', 'usuario',
]
COLS_CONTA = [
    'ORIGEN', 'nro_registro', 'fecha_mov', 'medio_pago',
    'nro_operacion', 'giro', 'glosa', 'ingreso', 'egreso',
    'fecha_conciliacion', 'conciliado',
]


# ══════════════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════════════

def _normalizar_empresa(nombre: str) -> str:
    """Normaliza el nombre de empresa para usarlo como nombre de hoja Excel."""
    # Quitar tildes y caracteres especiales
    normalizado = unicodedata.normalize('NFD', nombre)
    sin_tildes  = ''.join(c for c in normalizado if unicodedata.category(c) != 'Mn')
    # Reemplazar caracteres no alfanuméricos (excepto guion) por guion bajo
    limpio = re.sub(r'[^\w\-]', '_', sin_tildes)
    limpio = re.sub(r'_+', '_', limpio).strip('_')
    # Limitar a 31 caracteres (límite de Excel para nombres de hoja)
    return limpio[:31]


def _archivo_accesible() -> bool:
    """Verifica si la ruta de red es accesible."""
    try:
        return RUTA_RED.exists()
    except (OSError, PermissionError):
        return False


# ══════════════════════════════════════════════════════════════════════
# LECTURA — Cargar pendientes al generar reporte INICIAL
# ══════════════════════════════════════════════════════════════════════

def cargar_pendientes(
    empresa: str,
    mes_ref: int | None = None,
    anio_ref: int | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Lee los movimientos pendientes de la conciliación anterior desde el archivo de red.

    Devuelve (df_banco_pend, df_conta_pend), ambos con las columnas internas
    que usa el sistema (igual que los DataFrames sanitizados de main.py).
    Si el archivo no existe o la hoja no existe, devuelve DataFrames vacíos.

    Los DataFrames devueltos tienen una columna extra '_pendiente = True'
    para que el código de main.py pueda identificar su origen.

    Parámetros opcionales:
        mes_ref  : mes del extracto bancario actual (1-12). Si se provee junto con
                   anio_ref, los movimientos cuya fecha corresponde a ese mismo mes
                   y año son EXCLUIDOS (ya pertenecen al período actual y podrían
                   duplicarse). Solo se incluyen los de meses anteriores.
        anio_ref : año del extracto bancario actual (ej. 2026).
    """
    nombre_hoja = _normalizar_empresa(empresa)

    # Silenciar advertencias de openpyxl sobre estilos
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")

        if not _archivo_accesible():
            print(f"  [AVISO] Ruta de red no accesible: {RUTA_RED}")
            print("          Se omitira la carga de movimientos pendientes anteriores.")
            return pd.DataFrame(), pd.DataFrame()

        if not RUTA_ARCHIVO.exists():
            # Primera vez: el archivo no existe aún, es normal
            return pd.DataFrame(), pd.DataFrame()

        try:
            xf = pd.ExcelFile(RUTA_ARCHIVO)
        except Exception as e:
            print(f"  [AVISO] No se pudo abrir el archivo de pendientes: {e}")
            return pd.DataFrame(), pd.DataFrame()

        if nombre_hoja not in xf.sheet_names:
            # Empresa sin pendientes guardados
            return pd.DataFrame(), pd.DataFrame()

        try:
            df = pd.read_excel(RUTA_ARCHIVO, sheet_name=nombre_hoja)
        except Exception as e:
            print(f"  [AVISO] Error al leer la hoja '{nombre_hoja}' del archivo de pendientes: {e}")
            return pd.DataFrame(), pd.DataFrame()

    if df.empty or 'ORIGEN' not in df.columns:
        return pd.DataFrame(), pd.DataFrame()

    df_banco = df[df['ORIGEN'] == 'BANCO'].copy().drop(columns=['ORIGEN'], errors='ignore')
    df_conta = df[df['ORIGEN'] == 'CONTA'].copy().drop(columns=['ORIGEN'], errors='ignore')

    # Convertir tipos de banco
    if not df_banco.empty:
        df_banco['fecha']        = pd.to_datetime(df_banco.get('fecha'),        errors='coerce')
        df_banco['monto']        = pd.to_numeric(df_banco.get('monto'),         errors='coerce').fillna(0)
        df_banco['saldo']        = pd.to_numeric(df_banco.get('saldo'),         errors='coerce').fillna(0)
        df_banco['ingreso']      = df_banco['monto'].clip(lower=0)
        df_banco['egreso']       = df_banco['monto'].clip(upper=0).abs()
        df_banco['nro_operacion'] = df_banco.get('nro_operacion', pd.Series(dtype=str)).astype(str)
        df_banco['descripcion']  = df_banco.get('descripcion', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
        df_banco['_pendiente']   = True

    # Convertir tipos de conta
    if not df_conta.empty:
        df_conta['fecha_mov']          = pd.to_datetime(df_conta.get('fecha_mov'),          errors='coerce')
        df_conta['fecha_conciliacion'] = pd.to_datetime(df_conta.get('fecha_conciliacion'), errors='coerce')
        df_conta['ingreso']  = pd.to_numeric(df_conta.get('ingreso'), errors='coerce').fillna(0)
        df_conta['egreso']   = pd.to_numeric(df_conta.get('egreso'),  errors='coerce').fillna(0)
        df_conta['nro_operacion'] = df_conta.get('nro_operacion', pd.Series(dtype=str)).astype(str)
        for col in ['giro', 'glosa', 'medio_pago']:
            if col in df_conta.columns:
                df_conta[col] = df_conta[col].fillna('').astype(str).str.strip()
        df_conta['conciliado'] = df_conta.get('conciliado', pd.Series(dtype=object)).map(
            lambda v: True if str(v).strip().upper() in ('TRUE', 'SI', '1') else False
        )
        df_conta['_pendiente'] = True

    # ── Filtrar por mes de referencia ─────────────────────────────────
    # Si se proporcionó mes_ref y anio_ref, se excluyen los movimientos cuya
    # fecha corresponde al período actual (ya están en el extracto bancario y
    # podrían duplicarse). Solo se cargan los de meses ANTERIORES.
    if mes_ref is not None and anio_ref is not None:
        if not df_banco.empty:
            fechas_b = pd.to_datetime(df_banco['fecha'], errors='coerce')
            mask_mes_actual_b = (
                (fechas_b.dt.month == mes_ref) & (fechas_b.dt.year == anio_ref)
            )
            omitidos_b = mask_mes_actual_b.sum()
            df_banco = df_banco[~mask_mes_actual_b].copy()
            if omitidos_b > 0:
                print(f"  [INFO] Se omitieron {omitidos_b} pendiente(s) de banco del mes actual "
                      f"({mes_ref:02d}/{anio_ref}) para evitar duplicados.")

        if not df_conta.empty:
            fechas_c = pd.to_datetime(df_conta['fecha_mov'], errors='coerce')
            mask_mes_actual_c = (
                (fechas_c.dt.month == mes_ref) & (fechas_c.dt.year == anio_ref)
            )
            omitidos_c = mask_mes_actual_c.sum()
            df_conta = df_conta[~mask_mes_actual_c].copy()
            if omitidos_c > 0:
                print(f"  [INFO] Se omitieron {omitidos_c} pendiente(s) de contabilidad del mes actual "
                      f"({mes_ref:02d}/{anio_ref}) para evitar duplicados.")

    n_banco = len(df_banco)
    n_conta = len(df_conta)
    if n_banco > 0 or n_conta > 0:
        print(f"  [INFO] Pendientes anteriores cargados: {n_banco} banco, {n_conta} conta "
              f"(empresa: {empresa}, hoja: {nombre_hoja})")

    return df_banco, df_conta


# ══════════════════════════════════════════════════════════════════════
# ESCRITURA — Guardar pendientes al generar reporte FINAL
# ══════════════════════════════════════════════════════════════════════

def guardar_pendientes(
    empresa: str,
    banco_sin_conciliar: pd.DataFrame,
    conta_sin_conciliar: pd.DataFrame,
) -> bool:
    """
    Guarda los movimientos que NO fueron conciliados en la hoja de la empresa
    dentro del archivo Excel de pendientes.

    - Crea el archivo si no existe.
    - Crea la hoja de empresa si no existe.
    - REEMPLAZA el contenido anterior de la hoja (chancará los datos viejos).

    Parámetros:
        empresa               : Nombre de la empresa (ej. "STN", "ITS").
        banco_sin_conciliar   : DataFrame de filas banco que quedaron sin conciliar
                                (banco_sin_marcar del reporte final).
        conta_sin_conciliar   : DataFrame de filas conta que quedaron sin conciliar
                                (conta_sin_marcar del reporte final).

    Devuelve True si se guardó correctamente, False si hubo un error.
    """
    nombre_hoja = _normalizar_empresa(empresa)

    if not _archivo_accesible():
        print(f"  [AVISO] Ruta de red no accesible: {RUTA_RED}")
        print("          No se pudieron guardar los movimientos pendientes.")
        return False

    # ── Preparar datos banco ──────────────────────────────────────────
    filas_banco = _preparar_banco_para_guardar(banco_sin_conciliar)
    # ── Preparar datos conta ──────────────────────────────────────────
    filas_conta = _preparar_conta_para_guardar(conta_sin_conciliar)

    # Combinar en un único DataFrame con columna ORIGEN
    df_nuevo = pd.concat([filas_banco, filas_conta], ignore_index=True)

    # ── Leer el archivo existente (si existe) para preservar otras hojas ─
    hojas_existentes: dict[str, pd.DataFrame] = {}
    if RUTA_ARCHIVO.exists():
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                xf = pd.ExcelFile(RUTA_ARCHIVO)
                for hoja in xf.sheet_names:
                    if hoja != nombre_hoja:
                        hojas_existentes[hoja] = pd.read_excel(RUTA_ARCHIVO, sheet_name=hoja)
        except Exception as e:
            print(f"  [AVISO] No se pudo leer el archivo existente de pendientes: {e}")
            # Continuar: se sobrescribirá el archivo completo

    # ── Escribir archivo con todas las hojas ──────────────────────────
    try:
        with pd.ExcelWriter(RUTA_ARCHIVO, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
            # Primero escribir la hoja actualizada de la empresa
            df_nuevo.to_excel(writer, sheet_name=nombre_hoja, index=False)
            # Luego preservar las demás hojas
            for hoja, df_hoja in hojas_existentes.items():
                df_hoja.to_excel(writer, sheet_name=hoja, index=False)

        n_banco = len(filas_banco)
        n_conta = len(filas_conta)
        print(f"\n  [OK] Pendientes guardados en: {RUTA_ARCHIVO}")
        print(f"       Empresa: {empresa} | Hoja: {nombre_hoja}")
        print(f"       {n_banco} movimiento(s) banco + {n_conta} movimiento(s) conta sin conciliar")
        return True

    except PermissionError:
        print(f"\n  [ERROR] No se pudo guardar el archivo de pendientes.")
        print(f"          El archivo '{RUTA_ARCHIVO.name}' puede estar abierto por otro usuario.")
        print("          Cierrelo e intente nuevamente.")
        return False
    except Exception as e:
        print(f"\n  [ERROR] No se pudo guardar el archivo de pendientes: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════
# PREPARACIÓN DE DATOS PARA GUARDAR
# ══════════════════════════════════════════════════════════════════════

def _preparar_banco_para_guardar(banco_sin_conciliar: pd.DataFrame) -> pd.DataFrame:
    """
    Extrae las columnas relevantes del DataFrame banco_sin_marcar
    y agrega la columna ORIGEN='BANCO'.
    """
    if banco_sin_conciliar.empty:
        return pd.DataFrame(columns=COLS_BANCO)

    df = banco_sin_conciliar.copy()

    # Mapear columnas del formato Anexar1 al formato interno
    mapeo = {
        'Fecha'                 : 'fecha',
        'Banco - Descripción'   : 'descripcion',
        'Monto-Banco'           : 'monto',
        'Banco - Saldo'         : 'saldo',
        'Banco - # Operación'   : 'nro_operacion',
    }

    # Si el DataFrame viene de Anexar1 (tiene columnas con prefijo "Banco -")
    es_formato_anexar = 'Banco - Descripción' in df.columns or 'Monto-Banco' in df.columns

    if es_formato_anexar:
        filas = []
        for _, row in df.iterrows():
            # Ignorar filas sin monto (mov. sin monto o monto 0)
            monto = row.get('Monto-Banco')
            if pd.isna(monto) or float(monto) == 0:
                continue
            filas.append({
                'ORIGEN'        : 'BANCO',
                'fecha'         : row.get('Fecha'),
                'descripcion'   : str(row.get('Banco - Descripción', '') or ''),
                'monto'         : float(monto),
                'saldo'         : float(row.get('Banco - Saldo', 0) or 0),
                'sucursal'      : '',
                'nro_operacion' : str(row.get('Banco - # Operación', '') or ''),
                'hora'          : '',
                'usuario'       : '',
            })
        return pd.DataFrame(filas, columns=COLS_BANCO) if filas else pd.DataFrame(columns=COLS_BANCO)

    else:
        # Formato interno de main.py (banco ya sanitizado)
        filas = []
        for _, row in df.iterrows():
            monto = row.get('monto', 0)
            if pd.isna(monto) or float(monto) == 0:
                continue
            filas.append({
                'ORIGEN'        : 'BANCO',
                'fecha'         : row.get('fecha'),
                'descripcion'   : str(row.get('descripcion', '') or ''),
                'monto'         : float(monto),
                'saldo'         : float(row.get('saldo', 0) or 0),
                'sucursal'      : str(row.get('sucursal', '') or ''),
                'nro_operacion' : str(row.get('nro_operacion', '') or ''),
                'hora'          : str(row.get('hora', '') or ''),
                'usuario'       : str(row.get('usuario', '') or ''),
            })
        return pd.DataFrame(filas, columns=COLS_BANCO) if filas else pd.DataFrame(columns=COLS_BANCO)


def _preparar_conta_para_guardar(conta_sin_conciliar: pd.DataFrame) -> pd.DataFrame:
    """
    Extrae las columnas relevantes del DataFrame conta_sin_marcar
    y agrega la columna ORIGEN='CONTA'.
    """
    if conta_sin_conciliar.empty:
        return pd.DataFrame(columns=COLS_CONTA)

    df = conta_sin_conciliar.copy()

    # Detectar si viene del formato Anexar1 (columnas con prefijo "Conta -")
    es_formato_anexar = 'Conta - # Registro' in df.columns or 'Monto-Conta' in df.columns

    if es_formato_anexar:
        filas = []
        for _, row in df.iterrows():
            monto = row.get('Monto-Conta')
            if pd.isna(monto) or float(monto) == 0:
                continue
            ingreso = float(monto) if float(monto) > 0 else 0.0
            egreso  = abs(float(monto)) if float(monto) < 0 else 0.0
            filas.append({
                'ORIGEN'            : 'CONTA',
                'nro_registro'      : str(row.get('Conta - # Registro', '') or ''),
                'fecha_mov'         : row.get('Fecha'),
                'medio_pago'        : '',
                'nro_operacion'     : str(row.get('Conta - # Operación', '') or ''),
                'giro'              : str(row.get('Conta - Giro', '') or ''),
                'glosa'             : str(row.get('Conta - Glosa', '') or ''),
                'ingreso'           : ingreso,
                'egreso'            : egreso,
                'fecha_conciliacion': None,
                'conciliado'        : False,
            })
        return pd.DataFrame(filas, columns=COLS_CONTA) if filas else pd.DataFrame(columns=COLS_CONTA)

    else:
        # Formato interno de main.py (conta ya sanitizado)
        filas = []
        for _, row in df.iterrows():
            ingreso = float(row.get('ingreso', 0) or 0)
            egreso  = float(row.get('egreso',  0) or 0)
            if ingreso == 0 and egreso == 0:
                continue
            filas.append({
                'ORIGEN'            : 'CONTA',
                'nro_registro'      : str(row.get('nro_registro', '') or ''),
                'fecha_mov'         : row.get('fecha_mov'),
                'medio_pago'        : str(row.get('medio_pago', '') or ''),
                'nro_operacion'     : str(row.get('nro_operacion', '') or ''),
                'giro'              : str(row.get('giro', '') or ''),
                'glosa'             : str(row.get('glosa', '') or ''),
                'ingreso'           : ingreso,
                'egreso'            : egreso,
                'fecha_conciliacion': row.get('fecha_conciliacion'),
                'conciliado'        : bool(row.get('conciliado', False)),
            })
        return pd.DataFrame(filas, columns=COLS_CONTA) if filas else pd.DataFrame(columns=COLS_CONTA)


# ══════════════════════════════════════════════════════════════════════
# INYECCIÓN EN DATAFRAMES DE MAIN.PY
# ══════════════════════════════════════════════════════════════════════

def inyectar_pendientes_en_banco(
    bank: pd.DataFrame,
    df_banco_pend: pd.DataFrame,
    mes_ref: int | None = None,
    anio_ref: int | None = None,
) -> pd.DataFrame:
    """
    Anexa los movimientos pendientes del banco al DataFrame bank ya sanitizado.

    Un pendiente NO se agrega si:
      a) Su fecha corresponde al mes/año de referencia (ya está en el extracto
         actual y podría duplicarse), o
      b) Ya existe en los datos actuales con el mismo número de operación y monto.

    Parámetros:
        mes_ref  : mes del extracto actual (1-12). Si se provee junto con anio_ref,
                   los pendientes con esa misma fecha de mes/año se omiten.
        anio_ref : año del extracto actual (ej. 2026).
    """
    if df_banco_pend.empty:
        return bank

    cols_bank = bank.columns.tolist()

    filas_nuevas = []
    for _, row in df_banco_pend.iterrows():
        fecha_p = row.get('fecha')
        monto_p = float(row.get('monto', 0) or 0)
        op_p    = str(row.get('nro_operacion', '') or '').strip()
        desc_p  = str(row.get('descripcion', '') or '').strip().upper()

        # Filtro por mes/año de referencia: omitir pendientes del período actual
        if mes_ref is not None and anio_ref is not None and pd.notna(fecha_p):
            try:
                fp = pd.to_datetime(fecha_p, errors='coerce')
                if pd.notna(fp) and fp.month == mes_ref and fp.year == anio_ref:
                    continue  # Es del mes actual, podría duplicarse
            except Exception:
                pass

        # Criterio 1: mismo nro_operacion (no vacío/genérico) + mismo monto
        # → el banco ya tiene este movimiento registrado (quizá con otra fecha)
        op_valido = op_p and op_p not in ('nan', '0', 'NAN')
        if op_valido:
            ya_existe = (
                (bank['nro_operacion'].astype(str).str.strip() == op_p) &
                (bank['monto'].fillna(0) == monto_p)
            ).any()
        else:
            # Criterio 2 (fallback): misma fecha + monto + descripción
            ya_existe = (
                (bank['fecha'] == fecha_p) &
                (bank['monto'].fillna(0) == monto_p) &
                (bank['descripcion'].astype(str).str.strip().str.upper() == desc_p)
            ).any()

        if ya_existe:
            continue  # El movimiento ya está en los datos actuales, no agregar

        nueva_fila = {col: None for col in cols_bank}
        nueva_fila['fecha']         = fecha_p
        nueva_fila['descripcion']   = str(row.get('descripcion', '') or '')
        nueva_fila['monto']         = monto_p
        nueva_fila['saldo']         = float(row.get('saldo', 0) or 0)
        nueva_fila['sucursal']      = str(row.get('sucursal', '') or '')
        nueva_fila['nro_operacion'] = op_p
        nueva_fila['hora']          = str(row.get('hora', '') or '')
        nueva_fila['usuario']       = str(row.get('usuario', '') or '')
        nueva_fila['ingreso']       = monto_p if monto_p > 0 else 0
        nueva_fila['egreso']        = abs(monto_p) if monto_p < 0 else 0
        nueva_fila['_pendiente']    = True  # Marcador de origen

        filas_nuevas.append(nueva_fila)

    if not filas_nuevas:
        return bank

    # Asegurar que bank tenga la columna _pendiente antes de concatenar
    if '_pendiente' not in bank.columns:
        bank = bank.copy()
        bank['_pendiente'] = False

    df_pend = pd.DataFrame(filas_nuevas)
    resultado = pd.concat([bank, df_pend], ignore_index=True)
    resultado = resultado.sort_values('fecha').reset_index(drop=True)

    print(f"  [INFO] Se agregaron {len(filas_nuevas)} movimiento(s) pendiente(s) de banco "
          "de la conciliacion anterior.")
    return resultado


def inyectar_pendientes_en_conta(
    conta: pd.DataFrame,
    df_conta_pend: pd.DataFrame,
    mes_ref: int | None = None,
    anio_ref: int | None = None,
) -> pd.DataFrame:
    """
    Anexa los movimientos pendientes de contabilidad al DataFrame conta ya sanitizado.

    Un pendiente NO se agrega si:
      a) Su fecha_mov corresponde al mes/año de referencia (ya está en el extracto
         actual y podría duplicarse), o
      b) El mismo número de operación (y mismo monto) ya existe en los datos actuales.

    Parámetros:
        mes_ref  : mes del extracto actual (1-12). Si se provee junto con anio_ref,
                   los pendientes con esa misma fecha de mes/año se omiten.
        anio_ref : año del extracto actual (ej. 2026).
    """
    if df_conta_pend.empty:
        return conta

    cols_conta = conta.columns.tolist()

    filas_nuevas = []
    for _, row in df_conta_pend.iterrows():
        fecha_p   = row.get('fecha_mov')
        ingreso_p = float(row.get('ingreso', 0) or 0)
        egreso_p  = float(row.get('egreso',  0) or 0)
        op_p      = str(row.get('nro_operacion', '') or '').strip()
        glosa_p   = str(row.get('glosa', '') or '').strip().upper()

        # Filtro por mes/año de referencia: omitir pendientes del período actual
        if mes_ref is not None and anio_ref is not None and pd.notna(fecha_p):
            try:
                fp = pd.to_datetime(fecha_p, errors='coerce')
                if pd.notna(fp) and fp.month == mes_ref and fp.year == anio_ref:
                    continue  # Es del mes actual, podría duplicarse
            except Exception:
                pass

        # Criterio 1: mismo nro_operacion (no vacío/genérico) + mismo monto
        op_valido = op_p and op_p not in ('nan', '0', 'NAN')
        if op_valido:
            ya_existe = (
                (conta['nro_operacion'].astype(str).str.strip() == op_p) &
                (conta['ingreso'].fillna(0) == ingreso_p) &
                (conta['egreso'].fillna(0) == egreso_p)
            ).any()
        else:
            # Criterio 2 (fallback): misma fecha + ingreso + egreso + glosa
            ya_existe = (
                (conta['fecha_mov'] == fecha_p) &
                (conta['ingreso'].fillna(0) == ingreso_p) &
                (conta['egreso'].fillna(0) == egreso_p) &
                (conta['glosa'].astype(str).str.strip().str.upper() == glosa_p)
            ).any()

        if ya_existe:
            continue

        nueva_fila = {col: None for col in cols_conta}
        nueva_fila['nro_registro']       = str(row.get('nro_registro', '') or '')
        nueva_fila['fecha_mov']          = fecha_p
        nueva_fila['medio_pago']         = str(row.get('medio_pago', '') or '')
        nueva_fila['nro_operacion']      = op_p
        nueva_fila['giro']               = str(row.get('giro', '') or '')
        nueva_fila['glosa']              = str(row.get('glosa', '') or '')
        nueva_fila['ingreso']            = ingreso_p
        nueva_fila['egreso']             = egreso_p
        nueva_fila['fecha_conciliacion'] = row.get('fecha_conciliacion')
        nueva_fila['conciliado']         = bool(row.get('conciliado', False))
        nueva_fila['_pendiente']         = True

        filas_nuevas.append(nueva_fila)

    if not filas_nuevas:
        return conta

    if '_pendiente' not in conta.columns:
        conta = conta.copy()
        conta['_pendiente'] = False

    df_pend = pd.DataFrame(filas_nuevas)
    resultado = pd.concat([conta, df_pend], ignore_index=True)
    resultado = resultado.sort_values('fecha_mov').reset_index(drop=True)

    print(f"  [INFO] Se agregaron {len(filas_nuevas)} movimiento(s) pendiente(s) de contabilidad "
          "de la conciliacion anterior.")
    return resultado



# ══════════════════════════════════════════════════════════════════════
# EXPORTACIÓN DE CONCILIADOS PARA CONTANET
# ══════════════════════════════════════════════════════════════════════

RUTA_CONCILIADOS = RUTA_RED / NOMBRE_CONCILIADOS


def guardar_conciliados(empresa: str, df_conciliados: pd.DataFrame) -> bool:
    """
    Guarda en la ruta de red el resumen de movimientos conciliados listos
    para ser ingresados en Contanet.

    Fuente: hoja 'Conciliados' del CBF (reporte final).
    Solo se incluyen filas donde '# Op. Banco' sea un valor distinto de 0
    y no este vacio (tienen numero de operacion bancario identificable).

    Estructura de la tabla guardada:
      Banco | Asiento Contable | Año | Mes | Numero de Operacion

    Donde:
      - Banco             : Descripcion Banco (descripcion del movimiento bancario)
      - Asiento Contable  : # Registro (asiento contable de Contanet)
      - Año               : año extraido de Fecha Banco
      - Mes               : mes (nombre) extraido de Fecha Banco
      - Numero de Operacion: # Op. Banco

    La hoja tendra como nombre el nombre normalizado de la empresa (ej. STN, CMT).
    Si el archivo ya existe, se preservan las demas hojas (otras empresas).

    Devuelve True si se guardo correctamente, False si hubo un error.
    """
    nombre_hoja = _normalizar_empresa(empresa)

    if not _archivo_accesible():
        print(f"  [AVISO] Ruta de red no accesible: {RUTA_RED}")
        print("          No se pudo guardar el archivo de conciliados para Contanet.")
        return False

    if df_conciliados.empty:
        print("  [AVISO] No hay movimientos conciliados para exportar a Contanet.")
        return False

    # ── Normalizar columnas (quitar espacios) ─────────────────────────
    df = df_conciliados.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # ── Filtrar: solo filas con # Op. Banco valido (distinto de 0 y no vacio) ──
    col_op = '# Op. Banco'
    if col_op not in df.columns:
        print(f"  [AVISO] La hoja Conciliados no tiene la columna '{col_op}'.")
        return False

    mask_valido = df[col_op].apply(
        lambda v: (
            pd.notna(v) and
            str(v).strip() not in ('', 'nan', '0', '0.0') and
            str(v).strip() != ''
        )
    )
    df_filtrado = df[mask_valido].copy()

    if df_filtrado.empty:
        print("  [AVISO] Ninguna fila conciliada tiene numero de operacion valido para exportar.")
        return False

    # ── Extraer año y mes de Fecha Banco ─────────────────────────────
    fechas = pd.to_datetime(df_filtrado.get('Fecha Banco'), errors='coerce')
    df_filtrado['_anio'] = fechas.dt.year.where(fechas.notna(), other=None)
    df_filtrado['_mes']  = fechas.dt.month.where(fechas.notna(), other=None)

    # ── Normalizar # Op. Banco: quitar .0 si es numero entero ────────
    df_filtrado['_nro_op'] = df_filtrado[col_op].apply(
        lambda v: str(int(float(v))) if (pd.notna(v) and str(v).strip().endswith('.0')
                  and str(v).strip()[:-2].lstrip('-').isdigit())
                  else str(v).strip()
    )

    # ── Construir tabla final ─────────────────────────────────────────
    tabla = pd.DataFrame({
        'Asiento Contable'    : df_filtrado.get('# Registro', '').fillna('').astype(str).str.strip(),
        'Año'                 : df_filtrado['_anio'],
        'Mes'                 : df_filtrado['_mes'],
        'Número de Operación' : df_filtrado['_nro_op'],
        'Anotación'           : df_filtrado.get('Anotacion', df_filtrado.get('Anotación', '')).fillna('').astype(str).str.strip(),
    }).reset_index(drop=True)

    # ── Leer el archivo existente para preservar otras hojas ─────────
    hojas_existentes: dict[str, pd.DataFrame] = {}
    if RUTA_CONCILIADOS.exists():
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                xf = pd.ExcelFile(RUTA_CONCILIADOS)
                for hoja in xf.sheet_names:
                    if hoja != nombre_hoja:
                        hojas_existentes[hoja] = pd.read_excel(RUTA_CONCILIADOS, sheet_name=hoja)
        except Exception as e:
            print(f"  [AVISO] No se pudo leer el archivo existente de conciliados: {e}")

    # ── Escribir archivo ──────────────────────────────────────────────
    try:
        with pd.ExcelWriter(RUTA_CONCILIADOS, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
            tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)
            for hoja, df_hoja in hojas_existentes.items():
                df_hoja.to_excel(writer, sheet_name=hoja, index=False)

        # Aplicar formato de cabecera al archivo
        _aplicar_formato_conciliados_contanet(RUTA_CONCILIADOS, nombre_hoja)

        n = len(tabla)
        print(f"\n  [OK] Conciliados para Contanet guardados en: {RUTA_CONCILIADOS}")
        print(f"       Empresa: {empresa} | Hoja: {nombre_hoja} | {n} movimiento(s)")
        return True

    except PermissionError:
        print(f"\n  [ERROR] No se pudo guardar el archivo de conciliados para Contanet.")
        print(f"          El archivo '{RUTA_CONCILIADOS.name}' puede estar abierto por otro usuario.")
        print("          Cierrelo e intente nuevamente.")
        return False
    except Exception as e:
        print(f"\n  [ERROR] No se pudo guardar el archivo de conciliados: {e}")
        return False


def _aplicar_formato_conciliados_contanet(ruta: Path, nombre_hoja: str) -> None:
    """Aplica formato visual basico a la hoja de conciliados para Contanet."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.load_workbook(ruta)
        if nombre_hoja not in wb.sheetnames:
            return

        ws = wb[nombre_hoja]
        thin = Side(style='thin', color='D9D9D9')
        borde = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Cabecera: fondo verde oscuro (color Contanet)
        for cell in ws[1]:
            cell.font      = Font(bold=True, color='FFFFFF', size=10)
            cell.fill      = PatternFill('solid', fgColor='375623')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border    = borde
        ws.row_dimensions[1].height = 28

        # Datos: bordes y alineacion
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border    = borde
                cell.alignment = Alignment(vertical='center')
                # Año y Mes centrados
                col_name = str(ws.cell(row=1, column=cell.column).value or '')
                if col_name in ('Año', 'Mes'):
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        # Autoajuste de columnas
        for col in ws.columns:
            max_len = max(
                (len(str(cell.value)) for cell in col if cell.value is not None),
                default=10
            )
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 45)

        ws.freeze_panes = 'A2'
        wb.save(ruta)
    except Exception:
        pass  # El formato es opcional; no fallar si hay error
