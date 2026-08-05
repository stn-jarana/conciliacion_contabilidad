"""
reporte_final.py
────────────────
Lee el Excel trabajado por el especialista (hoja Anexar1 con MAR/SUB
completados manualmente) y genera el reporte final de conciliación.

El especialista vincula filas banco↔conta llenando en Anexar1:
  - MAR          : 'X' en ambas filas (banco y conta) que se corresponden
  - # Operación2 : número de operación del banco en la fila de conta (y viceversa)
  - COD_SUB      : código auxiliar para subcuentas si aplica
  - DIF COMISON  : diferencia por comisión si el monto difiere levemente

El reporte final produce:
  Hoja Conciliados        → pares banco↔conta confirmados por el especialista
  Hoja Conciliacion_Final → formato formal de conciliación bancaria (estilo Formato Propuesto)

Uso:
    from reporte_final import generar_reporte_final
    generar_reporte_final(ruta_excel_trabajado)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════

def generar_reporte_final(
    ruta_inicial: Path,
    ruta_salida: Path | None = None,
    empresa: str = '',
    ruc: str = '',
    moneda: str = '',
) -> Path:
    """
    Lee el Excel inicial trabajado por el especialista y genera el reporte final.
    Devuelve la ruta del archivo creado.

    Parámetros:
        ruta_inicial : Ruta del Excel trabajado por el especialista.
        ruta_salida  : Ruta de salida (opcional, se genera automáticamente).
        empresa      : Nombre de la empresa (tomado del menú de selección).
        ruc          : RUC de la empresa (tomado del menú de selección).
        moneda       : Moneda seleccionada (ej. 'Soles (PEN)', 'Dolares (USD)').
    """
    if ruta_salida is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_salida = Path(f"Conciliacion_Final_{ts}.xlsx")

    # ── Validar que el archivo tenga la hoja Anexar1 ─────────────────
    xl = pd.ExcelFile(ruta_inicial)
    if 'Anexar1' not in xl.sheet_names:
        raise ValueError(
            f"El archivo '{ruta_inicial.name}' no contiene la hoja 'Anexar1'.\n"
            "Asegurese de cargar el reporte inicial ya trabajado por el especialista."
        )

    # ── Leer Anexar1 ──────────────────────────────────────────────────
    anexar1 = pd.read_excel(ruta_inicial, sheet_name='Anexar1')

    # Normalizar nombres de columnas (quitar espacios accidentales)
    anexar1.columns = [str(c).strip() for c in anexar1.columns]

    # Convertir fechas
    anexar1['Fecha'] = pd.to_datetime(anexar1['Fecha'], dayfirst=True, errors='coerce')

    # Normalizar MAR: cualquier valor no nulo se trata como 'X'
    if 'MAR' in anexar1.columns:
        anexar1['MAR'] = anexar1['MAR'].apply(
            lambda v: 'X' if pd.notna(v) and str(v).strip() not in ('', '0', 'nan') else ''
        )
    else:
        anexar1['MAR'] = ''

    # ── 1. PRIORIDAD ABSOLUTA: Vincular filas por códigos comunes (letras, números o combinaciones) ────────
    col_op_conciliar = None
    for posible_col in ['# Operación a Conciliar', 'No Conciliar', 'N° Asiento a conciliar', 'Conciliar']:
        if posible_col in anexar1.columns:
            col_op_conciliar = posible_col
            break

    if col_op_conciliar:
        # 1.a. Agrupar por códigos idénticos ingresados en la columna de vinculación
        # Esto permite que el especialista escriba p.ej. 'A', 'REGLA1', 'X1', '123' tanto en filas de banco como de conta
        mask_codigo = anexar1[col_op_conciliar].apply(
            lambda v: pd.notna(v) and str(v).strip() not in ('', 'nan', '0')
        )
        codigos_unicos = anexar1.loc[mask_codigo, col_op_conciliar].astype(str).str.strip().unique()

        for cod in codigos_unicos:
            filas_cod = anexar1[anexar1[col_op_conciliar].astype(str).str.strip() == cod]
            banco_idxs = filas_cod[filas_cod['Monto-Banco'].notna()].index.tolist()
            conta_idxs = filas_cod[filas_cod['Monto-Conta'].notna()].index.tolist()

            # Si el código une al menos una fila de banco y una de conta
            if banco_idxs and conta_idxs:
                op_link = f"MANUAL-{cod}"
                for b_i in banco_idxs:
                    anexar1.at[b_i, 'MAR'] = 'X'
                    anexar1.at[b_i, '# Operación2'] = op_link
                    anexar1.at[b_i, 'Anotación'] = f"Vinc. Manual ({cod})"

                for c_i in conta_idxs:
                    anexar1.at[c_i, 'MAR'] = 'X'
                    anexar1.at[c_i, '# Operación2'] = op_link
                    anexar1.at[c_i, 'Anotación'] = f"Vinc. Manual ({cod})"

        # 1.b. Buscar coincidencia cuando se ingresa el N° de Asiento / Operación del lado opuesto
        mask_op_manual = anexar1[col_op_conciliar].apply(lambda v: pd.notna(v) and str(v).strip() not in ('', 'nan', '0'))
        for idx_src in anexar1[mask_op_manual].index:
            if anexar1.at[idx_src, 'MAR'] == 'X':
                continue
            val_target = str(anexar1.at[idx_src, col_op_conciliar]).strip()
            es_banco_src = pd.notna(anexar1.at[idx_src, 'Monto-Banco'])

            if es_banco_src:
                mask_target = anexar1['Monto-Conta'].notna() & (
                    (anexar1['Conta - # Registro'].astype(str).str.strip() == val_target) |
                    (anexar1['Conta - # Operación'].astype(str).str.strip() == val_target) |
                    (anexar1['# Operación2'].astype(str).str.strip() == val_target)
                )
            else:
                mask_target = anexar1['Monto-Banco'].notna() & (
                    (anexar1['Banco - # Operación'].astype(str).str.strip() == val_target) |
                    (anexar1['# Operación2'].astype(str).str.strip() == val_target)
                )

            matches_target = anexar1[mask_target]
            if not matches_target.empty:
                all_banco = [idx_src] if es_banco_src else list(matches_target.index)
                all_conta = list(matches_target.index) if es_banco_src else [idx_src]

                op_link = f"MANUAL-{val_target}"
                for b_i in all_banco:
                    anexar1.at[b_i, 'MAR'] = 'X'
                    anexar1.at[b_i, '# Operación2'] = op_link
                    anexar1.at[b_i, 'Anotación'] = f"Manual por # Op ({val_target})"
                for c_i in all_conta:
                    anexar1.at[c_i, 'MAR'] = 'X'
                    anexar1.at[c_i, '# Operación2'] = op_link
                    anexar1.at[c_i, 'Anotación'] = f"Manual por # Op ({val_target})"

    # ── 2. Aprobar sugerencias si el especialista colocó una 'X' ÚNICAMENTE en la columna 'Conciliar' ─────────
    if 'Conciliar' in anexar1.columns and 'Anotación' in anexar1.columns:
        mask_sugerido_conciliar = (
            anexar1['Conciliar'].apply(lambda v: pd.notna(v) and str(v).strip().upper() in ('X', 'SI', '1'))
            & anexar1['# Operación2'].apply(lambda v: pd.notna(v) and str(v).strip() not in ('', 'nan'))
        )
        for op2_group in anexar1.loc[mask_sugerido_conciliar, '# Operación2'].unique():
            # Solo marcar si no fue modificado o ya conciliado manualmente
            mask_grp = (anexar1['# Operación2'] == op2_group) & (anexar1['MAR'] != 'X')
            anexar1.loc[mask_grp, 'MAR'] = 'X'

    # ── Identificar filas 1-a-1 aprobadas (tienen ambos montos en la misma fila)
    mask_combinadas_aprobadas = (
        (anexar1['MAR'] == 'X')
        & anexar1['Monto-Banco'].notna()
        & anexar1['Monto-Conta'].notna()
    )
    df_combinadas = anexar1[mask_combinadas_aprobadas].copy()
    # Excluir esas filas del flujo normal (ya están conciliadas directamente)
    anexar1_normal = anexar1[~mask_combinadas_aprobadas].copy()

    # ── Separar filas banco y filas conta (solo las no-combinadas) ────────
    mask_banco = anexar1_normal['Monto-Banco'].notna()
    mask_conta = anexar1_normal['Monto-Conta'].notna()

    df_banco = anexar1_normal[mask_banco].copy()
    df_conta = anexar1_normal[mask_conta].copy()

    # ── Excluir movimientos con monto 0 (no participan en conciliación ni en partidas) ──
    col_anot = 'Anotación' if 'Anotación' in df_banco.columns else None
    if col_anot:
        df_banco = df_banco[~df_banco[col_anot].astype(str).str.contains('sin monto', case=False, na=False)]
        df_conta = df_conta[~df_conta[col_anot].astype(str).str.contains('sin monto', case=False, na=False)]
    else:
        # Fallback: excluir filas donde el monto es exactamente 0
        df_banco = df_banco[df_banco['Monto-Banco'].fillna(0) != 0]
        df_conta = df_conta[df_conta['Monto-Conta'].fillna(0) != 0]

    # ── Clasificar según MAR ────────────────────────────────────
    banco_marcado    = df_banco[df_banco['MAR'] == 'X'].copy()
    banco_sin_marcar = df_banco[df_banco['MAR'] != 'X'].copy()
    conta_marcado    = df_conta[df_conta['MAR'] == 'X'].copy()
    conta_sin_marcar = df_conta[df_conta['MAR'] != 'X'].copy()

    # ── Construir tabla de conciliados ──────────────────────────────
    conciliados_manual = _cruzar_marcados(banco_marcado, conta_marcado)
    conciliados_1a1    = _conciliados_de_combinadas(df_combinadas)
    conciliados        = pd.concat([conciliados_manual, conciliados_1a1], ignore_index=True)

    # ── Leer metadatos del archivo inicial (empresa, banco, cuenta, moneda, mes) ─────
    meta = _extraer_metadatos(ruta_inicial, anexar1,
                              empresa_param=empresa,
                              ruc_param=ruc,
                              moneda_param=moneda)

    # ── Calcular saldos para el formato final ────────────────────────
    saldos = _calcular_saldos(
        anexar1, banco_marcado, banco_sin_marcar,
        conta_marcado, conta_sin_marcar, conciliados,
        ruta_inicial=ruta_inicial
    )

    # ── Partidas de ajuste ───────────────────────────────────────────
    partidas = _clasificar_partidas(banco_sin_marcar, conta_sin_marcar, anexar1)

    # ── Escribir Excel basado en Formato Propuesto.xlsx ─────────────────
    ruta_plantilla = Path("Formato Propuesto.xlsx")
    if ruta_plantilla.exists():
        wb = openpyxl.load_workbook(ruta_plantilla)
    else:
        wb = openpyxl.Workbook()

    # Si la hoja Conciliados ya existe en la plantilla, eliminarla para escribir la nueva
    if 'Conciliados' in wb.sheetnames:
        del wb['Conciliados']

    # Crear la hoja Conciliados en primer lugar
    ws_conc = wb.create_sheet('Conciliados', 0)

    # Eliminar la pestaña por defecto 'Sheet' creada automáticamente por openpyxl
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    if 'Hoja1' in wb.sheetnames:
        del wb['Hoja1']
    
    # Escribir los datos de conciliados en ws_conc
    df_conc = conciliados if not conciliados.empty else pd.DataFrame(columns=[
        'Fecha Banco', 'Descripcion Banco', 'Monto-Banco', '# Op. Banco',
        'Fecha Conta', '# Registro', 'Monto-Conta', 'DIF COMISON',
        'TIPO', 'Diferencia', 'Anotación'
    ])
    
    headers = list(df_conc.columns)
    ws_conc.append(headers)
    for row in df_conc.itertuples(index=False):
        ws_conc.append(list(row))

    wb.save(ruta_salida)

    # ── Actualizar hoja CONCILIACION FINAL ─────────────────────────────
    _actualizar_hoja_conciliacion_final(
        ruta_salida, meta, saldos, partidas, ruta_inicial=ruta_inicial
    )

    # ── Aplicar formato visual a la hoja Conciliados ─────────────────
    _aplicar_formato_conciliados(ruta_salida)

    # ── Pestañas adicionales: ITF Y COM, INGRESOS, EGRESOS, Anexar1 ──
    # El orden de llamada define el orden de las pestañas (antes de Anexar1).
    _escribir_hoja_itf_com(ruta_salida, ruta_inicial, meta)
    _escribir_hoja_ingresos(ruta_salida, ruta_inicial, meta)
    _escribir_hoja_egresos(ruta_salida, ruta_inicial, meta)
    # Copia exacta de la hoja Anexar1 del archivo inicial (siempre al final)
    _copiar_anexar1(ruta_inicial, ruta_salida)

    # ── Imprimir resumen en consola ───────────────────────────────────
    _imprimir_resumen_consola(saldos, partidas, ruta_salida)

    return ruta_salida


# ══════════════════════════════════════════════════════════════════════
# FUNCIONES INTERNAS – PROCESAMIENTO DE DATOS
# ══════════════════════════════════════════════════════════════════════

def _cruzar_marcados(
    banco_marcado: pd.DataFrame,
    conta_marcado: pd.DataFrame,
) -> pd.DataFrame:
    """
    Construye la tabla de pares conciliados vinculando banco↔conta
    por el campo '# Operación2' que el especialista completó.
    Si no hay vinculación por operación, empareja por monto+fecha.
    """
    if banco_marcado.empty or conta_marcado.empty:
        return pd.DataFrame()

    filas = []
    usados_conta = set()
    usados_banco = set()

    for i, rb in banco_marcado.iterrows():
        op2_b = str(rb.get('# Operación2', '')).strip()
        matched = False

        for j, rc in conta_marcado.iterrows():
            if j in usados_conta:
                continue
            op2_c = str(rc.get('# Operación2', '')).strip()
            op_b  = str(rb.get('Banco - # Operación',  '')).strip()
            op_c  = str(rc.get('Conta - # Operación',  '')).strip()

            if op2_b and op2_c and (op2_b == op2_c or op2_c == op_b or op2_b == op_c):
                filas.append(_fila_conciliada(rb, rc))
                usados_conta.add(j)
                usados_banco.add(i)
                matched = True
                break

        if not matched:
            monto_b = rb.get('Monto-Banco', None)
            if pd.notna(monto_b):
                for j, rc in conta_marcado.iterrows():
                    if j in usados_conta:
                        continue
                    monto_c = rc.get('Monto-Conta', None)
                    if pd.notna(monto_c) and abs(float(monto_b) - float(monto_c)) <= 0.01:
                        filas.append(_fila_conciliada(rb, rc))
                        usados_conta.add(j)
                        usados_banco.add(i)
                        break

    return pd.DataFrame(filas) if filas else pd.DataFrame()


def _conciliados_de_combinadas(df_combinadas: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte filas combinadas (1-a-1 aprobadas por el especialista en la
    columna 'Conciliar') en registros de la tabla de conciliados.
    """
    if df_combinadas.empty:
        return pd.DataFrame()
    filas = []
    for _, row in df_combinadas.iterrows():
        anotacion = str(row.get('Anotación', '')).strip() if pd.notna(row.get('Anotación')) else ''
        filas.append({
            'Fecha Banco'       : row.get('Fecha'),
            'Descripcion Banco' : row.get('Banco - Descripción', ''),
            'Monto-Banco'       : row.get('Monto-Banco'),
            '# Op. Banco'       : row.get('Banco - # Operación', ''),
            'Fecha Conta'       : row.get('Fecha'),
            '# Registro'        : row.get('Conta - # Registro', ''),
            'Monto-Conta'       : row.get('Monto-Conta'),
            'DIF COMISON'       : row.get('DIF COMISON', 0),
            'TIPO'              : row.get('TIPO', ''),
            'Diferencia'        : (row.get('Monto-Banco', 0) or 0) - (row.get('Monto-Conta', 0) or 0),
            'Anotación'         : anotacion,
        })
    return pd.DataFrame(filas)


def _fila_conciliada(rb: pd.Series, rc: pd.Series) -> dict:
    """Construye una fila del resultado de conciliados."""
    anot_b = str(rb.get('Anotación', '')).strip() if pd.notna(rb.get('Anotación')) else ''
    anot_c = str(rc.get('Anotación', '')).strip() if pd.notna(rc.get('Anotación')) else ''
    anotacion = anot_b
    if anot_c and anot_c != anot_b:
        anotacion = f"{anot_b} | {anot_c}" if anot_b else anot_c

    return {
        'Fecha Banco'        : rb.get('Fecha'),
        'Descripcion Banco'  : rb.get('Banco - Descripción', ''),
        'Monto-Banco'        : rb.get('Monto-Banco'),
        '# Op. Banco'        : rb.get('Banco - # Operación', ''),
        'Fecha Conta'        : rc.get('Fecha'),
        '# Registro'         : rc.get('Conta - # Registro', ''),
        'Monto-Conta'        : rc.get('Monto-Conta'),
        'DIF COMISON'        : rc.get('DIF COMISON', 0),
        'TIPO'               : rb.get('TIPO', ''),
        'Diferencia'         : (rb.get('Monto-Banco', 0) or 0) - (rc.get('Monto-Conta', 0) or 0),
        'Anotación'          : anotacion,
    }


def _extraer_metadatos(
    ruta_inicial: Path,
    anexar1: pd.DataFrame,
    empresa_param: str = '',
    ruc_param: str = '',
    moneda_param: str = '',
) -> dict:
    """
    Construye el dict de metadatos para el encabezado del reporte.
    Los valores pasados como parámetros tienen PRIORIDAD sobre los inferidos
    desde el contenido del archivo.
    """
    meta = {
        'empresa'       : empresa_param.strip(),
        'ruc'           : ruc_param.strip(),
        'banco'         : 'BANCO DE CREDITO',
        'cuenta'        : '',
        'moneda'        : moneda_param.strip(),
        'mes'           : '',
        'fecha_reporte' : datetime.now().strftime('%d/%m/%Y'),
    }

    # Obtener el mes de la primera fecha disponible en los datos
    fechas_validas = anexar1['Fecha'].dropna()
    if not fechas_validas.empty:
        primera_fecha = fechas_validas.iloc[0]
        meses_es = {
            1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
            5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
            9: 'SETIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
        }
        meta['mes'] = f"{meses_es.get(primera_fecha.month, '')} {primera_fecha.year}"

    # Si no viene moneda como parámetro, intentar inferirla del Anexar1
    if not meta['moneda']:
        for col in ['MONEDA', 'Moneda', 'moneda']:
            if col in anexar1.columns:
                monedas = anexar1[col].dropna().unique()
                if len(monedas) > 0:
                    meta['moneda'] = str(monedas[0])
                break

    return meta


def _calcular_saldos(
    anexar1, banco_marcado, banco_sin_marcar,
    conta_marcado, conta_sin_marcar, conciliados,
    ruta_inicial: Path | None = None
) -> dict:
    """Calcula los saldos necesarios para el formato de conciliación."""

    def safe_sum(df, col):
        if df.empty or col not in df.columns:
            return 0.0
        return pd.to_numeric(df[col], errors='coerce').fillna(0).sum()

    total_banco = safe_sum(pd.concat([banco_marcado, banco_sin_marcar]), 'Monto-Banco')
    total_conta = safe_sum(pd.concat([conta_marcado, conta_sin_marcar]), 'Monto-Conta')
    monto_solo_banco = safe_sum(banco_sin_marcar, 'Monto-Banco')
    monto_solo_conta = safe_sum(conta_sin_marcar, 'Monto-Conta')

    # Saldo según Libro Bancos: se obtiene del metadato guardado en la hoja CONTANET
    # (fila 2 de la columna 30, etiquetada '__SALDO_CONTABLE_FINAL__')
    saldo_libro_bancos = total_conta  # fallback si no se encuentra el metadato
    if ruta_inicial and ruta_inicial.exists():
        try:
            xl = pd.ExcelFile(ruta_inicial)
            if 'CONTANET' in xl.sheet_names:
                wb_tmp = openpyxl.load_workbook(ruta_inicial, read_only=True, data_only=True)
                ws_tmp = wb_tmp['CONTANET']
                etiqueta = ws_tmp.cell(row=1, column=30).value
                if str(etiqueta).strip() == '__SALDO_CONTABLE_FINAL__':
                    val = ws_tmp.cell(row=2, column=30).value
                    if val is not None:
                        saldo_libro_bancos = float(val)
                wb_tmp.close()
        except Exception:
            pass

    # Saldo según Extracto Bancario: último saldo registrado en el extracto (último movimiento del mes)
    saldo_extracto = total_banco
    if ruta_inicial and ruta_inicial.exists():
        try:
            xl = pd.ExcelFile(ruta_inicial)
            if 'BANCO' in xl.sheet_names:
                df_b_raw = pd.read_excel(ruta_inicial, sheet_name='BANCO')
                if 'Saldo' in df_b_raw.columns:
                    s_vals = df_b_raw['Saldo'].dropna()
                    if not s_vals.empty:
                        saldo_extracto = float(s_vals.iloc[-1])
        except Exception:
            pass

    if saldo_extracto == total_banco:
        # Si no se pudo leer de la hoja BANCO, buscar 'Banco - Saldo' en anexar1
        if 'Banco - Saldo' in anexar1.columns:
            s_vals = anexar1['Banco - Saldo'].dropna()
            if not s_vals.empty:
                saldo_extracto = float(s_vals.iloc[-1])

    return {
        'saldo_libro_bancos'  : round(saldo_libro_bancos, 2),
        'saldo_extracto'      : round(saldo_extracto, 2),
        'total_banco'         : round(total_banco, 2),
        'total_conta'         : round(total_conta, 2),
        'monto_solo_banco'    : round(monto_solo_banco, 2),
        'monto_solo_conta'    : round(monto_solo_conta, 2),
        'n_conciliados'       : len(banco_marcado) + (len(conciliados) if not conciliados.empty else 0),
        'n_solo_banco'        : len(banco_sin_marcar),
        'n_solo_conta'        : len(conta_sin_marcar),
    }


def _clasificar_partidas(
    banco_sin_marcar: pd.DataFrame,
    conta_sin_marcar: pd.DataFrame,
    anexar1: pd.DataFrame,
) -> dict:
    """
    Clasifica las partidas abiertas según el formato propuesto:
      - Abonos en Libros no en Extracto  (conta_sin_marcar con monto positivo)
      - Cargos en Libros no en Extracto  (conta_sin_marcar con monto negativo)
      - Abonos en Extracto no en Libros  (banco_sin_marcar con monto positivo)
      - Cargos en Extracto no en Libros  (banco_sin_marcar con monto negativo)
      - Cheques girados no cobrados       (partidas de conta marcadas como cheques)
    """

    def _desc_conta(row) -> str:
        """Obtiene la mejor descripción disponible para una fila de contabilidad."""
        glosa   = str(row.get('Conta - Glosa', '') or '').strip()
        giro    = str(row.get('Conta - Giro', '')  or '').strip()
        reg     = str(row.get('Conta - # Registro', '') or '').strip()
        return glosa or giro or reg

    def to_lista(df_src, col_monto, col_fecha, col_desc, col_op, positivo=True, es_conta=False, invertir_signo=False, mantener_signo=False):
        """Convierte un DataFrame de partidas a lista de dicts."""
        if df_src.empty:
            return []
        resultado = []
        for _, row in df_src.iterrows():
            monto = float(row.get(col_monto, 0) or 0)
            desc  = _desc_conta(row) if es_conta else str(row.get(col_desc, '') or '')

            if positivo and monto > 0:
                m_val = -monto if invertir_signo else monto
                resultado.append({
                    'fecha'      : row.get(col_fecha),
                    'operacion'  : str(row.get(col_op, '') or ''),
                    'descripcion': desc,
                    'monto'      : m_val,
                    'registro'   : str(row.get('Conta - # Registro', '') or ''),
                })
            elif not positivo and monto < 0:
                if mantener_signo:
                    m_val = monto
                elif invertir_signo:
                    m_val = -monto
                else:
                    m_val = abs(monto)
                resultado.append({
                    'fecha'      : row.get(col_fecha),
                    'operacion'  : str(row.get(col_op, '') or ''),
                    'descripcion': desc,
                    'monto'      : m_val,
                    'registro'   : str(row.get('Conta - # Registro', '') or ''),
                })
        return resultado

    def todas(df_src, col_monto, col_fecha, col_desc, col_op, es_conta=False):
        """Retorna todas las partidas sin filtrar por signo."""
        if df_src.empty:
            return []
        resultado = []
        for _, row in df_src.iterrows():
            monto = float(row.get(col_monto, 0) or 0)
            desc  = _desc_conta(row) if es_conta else str(row.get(col_desc, '') or '')
            resultado.append({
                'fecha'      : row.get(col_fecha),
                'operacion'  : str(row.get(col_op, '') or ''),
                'descripcion': desc,
                'monto'      : abs(monto),
                'registro'   : str(row.get('Conta - # Registro', '') or ''),
                'monto_raw'  : monto,
            })
        return resultado

    # Partidas de contabilidad sin par bancario
    abonos_lib_no_ext = to_lista(
        conta_sin_marcar, 'Monto-Conta', 'Fecha',
        'Conta - Glosa', 'Conta - # Operación', positivo=True, es_conta=True
    )
    cargos_lib_no_ext = to_lista(
        conta_sin_marcar, 'Monto-Conta', 'Fecha',
        'Conta - Glosa', 'Conta - # Operación', positivo=False, es_conta=True, invertir_signo=False, mantener_signo=True
    )

    # Partidas bancarias sin par contable
    abonos_ext_no_lib = to_lista(
        banco_sin_marcar, 'Monto-Banco', 'Fecha',
        'Banco - Descripción', 'Banco - # Operación', positivo=True
    )
    cargos_ext_no_lib = to_lista(
        banco_sin_marcar, 'Monto-Banco', 'Fecha',
        'Banco - Descripción', 'Banco - # Operación', positivo=False, mantener_signo=True
    )

    # Para el formato estándar: todas las partidas de cada categoría
    solo_banco_todas = todas(
        banco_sin_marcar, 'Monto-Banco', 'Fecha',
        'Banco - Descripción', 'Banco - # Operación'
    )
    solo_conta_todas = todas(
        conta_sin_marcar, 'Monto-Conta', 'Fecha',
        'Conta - Glosa', 'Conta - # Operación', es_conta=True
    )

    return {
        'abonos_lib_no_ext' : abonos_lib_no_ext,
        'cargos_lib_no_ext' : cargos_lib_no_ext,
        'abonos_ext_no_lib' : abonos_ext_no_lib,
        'cargos_ext_no_lib' : cargos_ext_no_lib,
        'solo_banco'        : solo_banco_todas,
        'solo_conta'        : solo_conta_todas,
    }


# ══════════════════════════════════════════════════════════════════════
# ESCRITURA DEL FORMATO FORMAL DE CONCILIACIÓN
# ══════════════════════════════════════════════════════════════════════

# Paleta de colores
_COLOR_TITULO    = '1F3864'   # azul marino oscuro
_COLOR_SECCION   = '2E75B6'   # azul medio
_COLOR_SUBSEC    = 'D6E4F0'   # azul muy claro (fondo filas de categoría)
_COLOR_TOTAL     = 'BDD7EE'   # azul claro para totales
_COLOR_CONCIL    = '1F4E79'   # azul oscuro pestaña Conciliados
_COLOR_WHITE     = 'FFFFFF'
_COLOR_NEGRO     = '000000'
_COLOR_GRIS_CLARO = 'F2F2F2'

_THIN   = Side(style='thin',   color='BFBFBF')
_MEDIUM = Side(style='medium', color='2E75B6')
_BORDE_THIN   = Border(left=_THIN,   right=_THIN,   top=_THIN,   bottom=_THIN)
_BORDE_MEDIUM = Border(left=_MEDIUM, right=_MEDIUM, top=_MEDIUM, bottom=_MEDIUM)


def _c(ws, fila, col, valor=None, *, bold=False, italic=False, size=11,
        color=_COLOR_NEGRO, bg=None, align_h='left', align_v='center',
        wrap=False, num_fmt=None, border=None, font_name='Aptos Narrow'):
    """Helper: escribe valor y aplica estilo a una celda."""
    cell = ws.cell(row=fila, column=col)
    if valor is not None:
        cell.value = valor
    cell.font = Font(
        name=font_name, bold=bold, italic=italic, size=size, color=color
    )
    cell.alignment = Alignment(
        horizontal=align_h, vertical=align_v, wrap_text=wrap
    )
    if bg:
        cell.fill = PatternFill('solid', fgColor=bg)
    if num_fmt:
        cell.number_format = num_fmt
    if border:
        cell.border = border
    return cell


def _merge(ws, fila1, col1, fila2, col2):
    ws.merge_cells(
        start_row=fila1, start_column=col1,
        end_row=fila2,   end_column=col2
    )


def _actualizar_hoja_conciliacion_final(
    ruta: Path,
    meta: dict,
    saldos: dict,
    partidas: dict,
    ruta_inicial: Path | None = None,
) -> None:
    """
    Abre el workbook en 'ruta', verifica si tiene la hoja 'CONCILIACION FINAL' del Formato Propuesto.
    Si existe, actualiza sus encabezados, saldos y partidas conservando el formato.
    Si no existe, invoca _escribir_hoja_conciliacion_final como fallback.
    """
    wb = openpyxl.load_workbook(ruta)
    if 'CONCILIACION FINAL' not in wb.sheetnames:
        wb.close()
        _escribir_hoja_conciliacion_final(ruta, meta, saldos, partidas)
        return

    ws = wb['CONCILIACION FINAL']

    # 1. Actualizar Encabezados en la plantilla
    if meta.get('fecha_reporte'):
        ws['E2'] = f"FECHA DE REPORTE DE CONCILIACION: {meta['fecha_reporte']}"
    if meta.get('empresa'):
        ws['B4'] = meta['empresa']
    if meta.get('ruc'):
        ws['B5'] = f"RUC: {meta['ruc']}"
    if meta.get('mes'):
        ws['H4'] = meta['mes']
    if meta.get('banco'):
        ws['C9'] = f"     {meta['banco']}"
    if meta.get('cuenta'):
        ws['F9'] = meta['cuenta']
    if meta.get('moneda'):
        ws['C10'] = meta['moneda']

    # 2. Actualizar Saldos
    ws['H13'] = saldos['saldo_libro_bancos']
    ws['H49'] = saldos['saldo_extracto']

    # 3. Limpiar ítems de ejemplo anteriores y colocar las partidas de la conciliación
    # - Cargos Libros no Extracto (Filas 19-22)
    cargos_lib = partidas.get('cargos_lib_no_ext', [])
    for idx, r in enumerate(range(19, 23)):
        if idx < len(cargos_lib):
            p = cargos_lib[idx]
            fecha_val = p['fecha']
            if isinstance(fecha_val, datetime):
                fecha_str = fecha_val.strftime('%d/%m/%Y')
            else:
                fecha_str = str(fecha_val) if fecha_val else ''
            ws.cell(row=r, column=3, value=fecha_str)
            ws.cell(row=r, column=4, value=p.get('registro') or p.get('operacion'))
            ws.cell(row=r, column=5, value=p.get('descripcion'))
            ws.cell(row=r, column=7, value=p.get('monto'))
        else:
            ws.cell(row=r, column=3, value=None)
            ws.cell(row=r, column=4, value=None)
            ws.cell(row=r, column=5, value=None)
            ws.cell(row=r, column=7, value=None)

    # - Abonos Extracto no Libros (Filas 26-34)
    abonos_ext = partidas.get('abonos_ext_no_lib', [])
    for idx, r in enumerate(range(26, 35)):
        if idx < len(abonos_ext):
            p = abonos_ext[idx]
            fecha_val = p['fecha']
            if isinstance(fecha_val, datetime):
                fecha_str = fecha_val.strftime('%d/%m/%Y')
            else:
                fecha_str = str(fecha_val) if fecha_val else ''
            ws.cell(row=r, column=3, value=fecha_str)
            ws.cell(row=r, column=5, value=p.get('descripcion'))
            ws.cell(row=r, column=7, value=p.get('monto'))
        else:
            ws.cell(row=r, column=3, value=None)
            ws.cell(row=r, column=5, value=None)
            ws.cell(row=r, column=7, value=None)

    # - Cargos Extracto no Libros (Filas 38-39)
    cargos_ext = partidas.get('cargos_ext_no_lib', [])
    for idx, r in enumerate(range(38, 40)):
        if idx < len(cargos_ext):
            p = cargos_ext[idx]
            fecha_val = p['fecha']
            if isinstance(fecha_val, datetime):
                fecha_str = fecha_val.strftime('%d/%m/%Y')
            else:
                fecha_str = str(fecha_val) if fecha_val else ''
            ws.cell(row=r, column=3, value=fecha_str)
            ws.cell(row=r, column=5, value=p.get('descripcion'))
            ws.cell(row=r, column=7, value=p.get('monto'))
        else:
            ws.cell(row=r, column=3, value=None)
            ws.cell(row=r, column=5, value=None)
            ws.cell(row=r, column=7, value=None)

    # Asegurar que las fórmulas de totales sigan intactas
    ws['H16'] = 0
    ws['H18'] = '=SUM(G19:G22)'
    ws['H24'] = '=SUM(G26:G34)'
    ws['H36'] = '=SUM(G38:G39)'
    ws['H41'] = '=SUM(G43:G47)'
    ws['H52'] = '=H13-H16+H18+H24-H36+H41-H49'

    # 4. Poblar las pestañas secundarias (EECC, INGRESOS, EGRESOS, ITF Y COM) con datos reales del banco
    if ruta_inicial and ruta_inicial.exists():
        try:
            xl = pd.ExcelFile(ruta_inicial)
            df_banco_raw = None
            if 'BANCO' in xl.sheet_names:
                df_banco_raw = pd.read_excel(ruta_inicial, sheet_name='BANCO')
            elif 'Anexar1' in xl.sheet_names:
                df_anexar = pd.read_excel(ruta_inicial, sheet_name='Anexar1')
                if 'Monto-Banco' in df_anexar.columns:
                    df_banco_raw = df_anexar[df_anexar['Monto-Banco'].notna()].copy()

            if df_banco_raw is not None and not df_banco_raw.empty:
                # ── EECC (Estado de Cuenta completo del Banco) ───────────
                if 'EECC' in wb.sheetnames:
                    ws_eecc = wb['EECC']
                    if meta.get('empresa') or meta.get('cuenta'):
                        ws_eecc['B1'] = f"{meta.get('cuenta', '')} - {meta.get('empresa', '')}".strip(" -")
                    if meta.get('moneda'):
                        ws_eecc['B2'] = meta.get('moneda', '')

                    if ws_eecc.max_row >= 6:
                        ws_eecc.delete_rows(6, ws_eecc.max_row - 5)

                    for _, row in df_banco_raw.iterrows():
                        fecha_val = row.get('Fecha') or row.get('fecha')
                        if isinstance(fecha_val, datetime):
                            fecha_str = fecha_val.strftime('%d/%m/%Y')
                        else:
                            fecha_str = str(fecha_val) if pd.notna(fecha_val) else ''

                        desc   = str(row.get('Descripción operación') or row.get('Banco - Descripción') or '')
                        monto  = float(row.get('Monto-Banco') if pd.notna(row.get('Monto-Banco')) else row.get('monto', 0) or 0)
                        saldo  = float(row.get('Saldo') if pd.notna(row.get('Saldo')) else row.get('Banco - Saldo', 0) or 0)
                        suc    = str(row.get('Sucursal') or '')
                        nro_op = str(row.get('# Operación') or row.get('Banco - # Operación') or '')
                        hora   = str(row.get('Hora') or '')
                        usr    = str(row.get('Usuario') or '')

                        ws_eecc.append([
                            fecha_str, '', desc, monto, saldo, suc, nro_op, hora, usr
                        ])

                # ── INGRESOS (Movimientos con monto > 0) ──────────────────
                if 'INGRESOS' in wb.sheetnames:
                    ws_ing = wb['INGRESOS']
                    if meta.get('empresa') or meta.get('cuenta'):
                        ws_ing['B1'] = f"{meta.get('cuenta', '')} - {meta.get('empresa', '')}".strip(" -")
                    if meta.get('moneda'):
                        ws_ing['B2'] = meta.get('moneda', '')

                    if ws_ing.max_row >= 6:
                        ws_ing.delete_rows(6, ws_ing.max_row - 5)

                    col_m = 'Monto-Banco' if 'Monto-Banco' in df_banco_raw.columns else 'monto'
                    df_ingres = df_banco_raw[df_banco_raw[col_m].fillna(0) > 0]
                    for _, row in df_ingres.iterrows():
                        fecha_val = row.get('Fecha') or row.get('fecha')
                        if isinstance(fecha_val, datetime):
                            fecha_str = fecha_val.strftime('%d/%m/%Y')
                        else:
                            fecha_str = str(fecha_val) if pd.notna(fecha_val) else ''

                        nro_op = str(row.get('# Operación') or row.get('Banco - # Operación') or '')
                        desc   = str(row.get('Descripción operación') or row.get('Banco - Descripción') or '')
                        monto  = float(row.get('Monto-Banco') if pd.notna(row.get('Monto-Banco')) else row.get('monto', 0) or 0)

                        ws_ing.append([fecha_str, nro_op, desc, monto])

                # ── EGRESOS (Movimientos con monto < 0) ───────────────────
                if 'EGRESOS' in wb.sheetnames:
                    ws_egr = wb['EGRESOS']
                    if meta.get('empresa') or meta.get('cuenta'):
                        ws_egr['B1'] = f"{meta.get('cuenta', '')} - {meta.get('empresa', '')}".strip(" -")
                    if meta.get('moneda'):
                        ws_egr['B2'] = meta.get('moneda', '')

                    if ws_egr.max_row >= 6:
                        ws_egr.delete_rows(6, ws_egr.max_row - 5)

                    col_m = 'Monto-Banco' if 'Monto-Banco' in df_banco_raw.columns else 'monto'
                    df_egres = df_banco_raw[df_banco_raw[col_m].fillna(0) < 0]
                    for _, row in df_egres.iterrows():
                        fecha_val = row.get('Fecha') or row.get('fecha')
                        if isinstance(fecha_val, datetime):
                            fecha_str = fecha_val.strftime('%d/%m/%Y')
                        else:
                            fecha_str = str(fecha_val) if pd.notna(fecha_val) else ''

                        nro_op = str(row.get('# Operación') or row.get('Banco - # Operación') or '')
                        desc   = str(row.get('Descripción operación') or row.get('Banco - Descripción') or '')
                        monto  = float(row.get('Monto-Banco') if pd.notna(row.get('Monto-Banco')) else row.get('monto', 0) or 0)

                        ws_egr.append([fecha_str, nro_op, desc, monto])

                # ── ITF Y COM (Movimientos de ITF y Comisiones) ───────────
                if 'ITF Y COM' in wb.sheetnames:
                    ws_itf = wb['ITF Y COM']
                    if meta.get('empresa') or meta.get('cuenta'):
                        ws_itf['B1'] = f"{meta.get('cuenta', '')} - {meta.get('empresa', '')}".strip(" -")
                    if meta.get('moneda'):
                        ws_itf['B2'] = meta.get('moneda', '')

                    if ws_itf.max_row >= 7:
                        ws_itf.delete_rows(7, ws_itf.max_row - 6)

                    col_desc = 'Descripción operación' if 'Descripción operación' in df_banco_raw.columns else 'Banco - Descripción'
                    if col_desc in df_banco_raw.columns:
                        mask_itf = df_banco_raw[col_desc].astype(str).str.contains('ITF|COM|COMIS|MANTENIM|ENVIO|PORTES', case=False, na=False)
                        df_itf = df_banco_raw[mask_itf]
                        for _, row in df_itf.iterrows():
                            fecha_val = row.get('Fecha') or row.get('fecha')
                            if isinstance(fecha_val, datetime):
                                fecha_str = fecha_val.strftime('%d/%m/%Y')
                            else:
                                fecha_str = str(fecha_val) if pd.notna(fecha_val) else ''

                            nro_op = str(row.get('# Operación') or row.get('Banco - # Operación') or '')
                            desc   = str(row.get(col_desc) or '')
                            monto  = float(row.get('Monto-Banco') if pd.notna(row.get('Monto-Banco')) else row.get('monto', 0) or 0)

                            ws_itf.append([fecha_str, nro_op, desc, monto])
        except Exception as e:
            print(f"  [Aviso] No se pudieron poblar todas las hojas secundarias: {e}")

    wb.save(ruta)


def _escribir_hoja_conciliacion_final(
    ruta: Path,
    meta: dict,
    saldos: dict,
    partidas: dict,
) -> None:
    """
    Abre el workbook ya creado (con la hoja Conciliados) y añade
    la hoja 'Conciliacion_Final' con el formato propuesto.
    """
    wb = openpyxl.load_workbook(ruta)
    if 'Conciliacion_Final' in wb.sheetnames:
        del wb['Conciliacion_Final']
    ws = wb.create_sheet('Conciliacion_Final')

    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    if 'Hoja1' in wb.sheetnames:
        del wb['Hoja1']

    # ── Configurar ancho de columnas (A..J) ───────────────────────────
    anchos = {
        'A': 3.5, 'B': 6.5,  'C': 14.5, 'D': 14.0,
        'E': 30.0, 'F': 16.0, 'G': 16.0, 'H': 14.0,
        'I': 12.0, 'J': 11.0
    }
    for letra, ancho in anchos.items():
        ws.column_dimensions[letra].width = ancho

    fila = 1  # cursor de fila actual

    # ══════════════════════════════════════════════════════════════════
    # BLOQUE DE ENCABEZADO
    # ══════════════════════════════════════════════════════════════════
    ws.row_dimensions[fila].height = 8  # fila espaciadora
    fila += 1

    # Fila fecha de reporte
    ws.row_dimensions[fila].height = 17
    fecha_txt = f"FECHA DE REPORTE DE CONCILIACION: {meta['fecha_reporte']}"
    _c(ws, fila, 5, fecha_txt, bold=True, size=11, align_h='left')
    _merge(ws, fila, 5, fila, 8)
    fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 8
    fila += 1

    # Empresa + Mes
    ws.row_dimensions[fila].height = 16
    empresa = meta.get('empresa') or 'EMPRESA'
    _c(ws, fila, 2, empresa, bold=True, size=12, color=_COLOR_TITULO)
    _merge(ws, fila, 2, fila, 4)
    _c(ws, fila, 6, '          MES:', bold=True, size=11, align_h='right')
    _c(ws, fila, 8, meta.get('mes', ''), bold=True, size=11, align_h='center',
       color=_COLOR_TITULO)
    fila += 1

    # RUC
    ws.row_dimensions[fila].height = 16
    ruc = meta.get('ruc', '')
    if ruc:
        _c(ws, fila, 2, f"RUC: {ruc}", size=10)
        _merge(ws, fila, 2, fila, 3)
    fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 8
    fila += 1

    # Título principal: CONCILIACIÓN BANCARIA
    ws.row_dimensions[fila].height = 22
    _c(ws, fila, 2, 'CONCILIACIÓN BANCARIA', bold=True, size=13,
       color=_COLOR_WHITE, bg=_COLOR_TITULO, align_h='center')
    _merge(ws, fila, 2, fila, 8)
    ws.cell(row=fila, column=2).border = _BORDE_MEDIUM
    fila_titulo = fila
    fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 8
    fila += 1

    # Banco y cuenta
    ws.row_dimensions[fila].height = 15
    _c(ws, fila, 2, 'Banco', bold=True, size=10)
    _c(ws, fila, 3, f"     {meta.get('banco', 'BANCO DE CREDITO')}", size=10)
    _c(ws, fila, 5, 'Cta. Corriente N°', size=10, align_h='right')
    _merge(ws, fila, 5, fila, 6)
    _c(ws, fila, 7, meta.get('cuenta', ''), size=10, bold=True)
    fila += 1

    # Moneda
    ws.row_dimensions[fila].height = 15
    _c(ws, fila, 2, 'Moneda', bold=True, size=10)
    _c(ws, fila, 3, meta.get('moneda', ''), size=10)
    fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 8
    fila += 1

    # ══════════════════════════════════════════════════════════════════
    # SALDO SEGÚN LIBRO BANCOS
    # ══════════════════════════════════════════════════════════════════
    ws.row_dimensions[fila].height = 15
    saldo_libros = saldos['saldo_libro_bancos']

    _c(ws, fila, 2, '       Saldo según Libro Bancos', bold=True, size=11,
       bg=_COLOR_TOTAL, border=_BORDE_THIN)
    _merge(ws, fila, 2, fila, 6)
    _c(ws, fila, 7, None, bg=_COLOR_TOTAL, border=_BORDE_THIN)
    _c(ws, fila, 8, saldo_libros, bold=True, size=11, align_h='right',
       bg=_COLOR_TOTAL, num_fmt='#,##0.00', border=_BORDE_THIN)
    fila_saldo_libros = fila
    fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 4
    fila += 1

    # ══════════════════════════════════════════════════════════════════
    # SECCIÓN: Partidas de ajuste (Libros → Extracto)
    # ══════════════════════════════════════════════════════════════════

    # ── (-) Abonos registrados en Libros y no en Extractos ────────────
    ws.row_dimensions[fila].height = 15
    _c(ws, fila, 2, '(-) Abonos registrados en Libros y no en Extractos',
       bold=True, size=10, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    _merge(ws, fila, 2, fila, 6)
    _c(ws, fila, 7, None, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    total_ab_lib = sum(p['monto'] for p in partidas['abonos_lib_no_ext'])
    _c(ws, fila, 8, total_ab_lib if total_ab_lib else 0, bold=True, size=10,
       align_h='right', bg=_COLOR_SUBSEC, num_fmt='#,##0.00', border=_BORDE_THIN)
    fila_abonos_lib = fila
    fila_abonos_lib_inicio = fila + 1
    fila += 1

    for p in partidas['abonos_lib_no_ext']:
        ws.row_dimensions[fila].height = 15
        fecha_val = p['fecha']
        if isinstance(fecha_val, datetime):
            fecha_str = fecha_val.strftime('%d/%m/%Y')
        else:
            fecha_str = str(fecha_val) if fecha_val else ''
        _c(ws, fila, 3, fecha_str, size=10, bold=True)
        _c(ws, fila, 4, p['registro'] or p['operacion'], size=9)
        _c(ws, fila, 5, p['descripcion'], size=10, wrap=True)
        _c(ws, fila, 7, p['monto'], size=10, align_h='right', num_fmt='#,##0.00')
        fila += 1

    fila_abonos_lib_fin = fila - 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 4
    fila += 1

    # ── (+) Cargos registrados en Libros y no en Extractos ────────────
    ws.row_dimensions[fila].height = 15
    _c(ws, fila, 2, '(+) Cargos registrados en Libros y no en Extractos',
       bold=True, size=10, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    _merge(ws, fila, 2, fila, 6)
    _c(ws, fila, 7, None, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    total_cg_lib = sum(p['monto'] for p in partidas['cargos_lib_no_ext'])
    _c(ws, fila, 8, total_cg_lib if total_cg_lib else 0, bold=True, size=10,
       align_h='right', bg=_COLOR_SUBSEC, num_fmt='#,##0.00', border=_BORDE_THIN)
    fila_cargos_lib_inicio = fila + 1
    fila += 1

    for p in partidas['cargos_lib_no_ext']:
        ws.row_dimensions[fila].height = 15
        fecha_val = p['fecha']
        if isinstance(fecha_val, datetime):
            fecha_str = fecha_val.strftime('%d/%m/%Y')
        else:
            fecha_str = str(fecha_val) if fecha_val else ''
        _c(ws, fila, 3, fecha_str, size=10, bold=True)
        _c(ws, fila, 4, p['registro'] or p['operacion'], size=9)
        _c(ws, fila, 5, p['descripcion'], size=10, wrap=True)
        _c(ws, fila, 7, p['monto'], size=10, align_h='right', num_fmt='#,##0.00')
        fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 4
    fila += 1

    # ══════════════════════════════════════════════════════════════════
    # SECCIÓN: Partidas de ajuste (Extracto → Libros)
    # ══════════════════════════════════════════════════════════════════

    # ── (+) Abonos registrados en Extractos y no en Libros ────────────
    ws.row_dimensions[fila].height = 15
    _c(ws, fila, 2, '(+) Abonos registrados en Extractos y no en Libros',
       bold=True, size=10, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    _merge(ws, fila, 2, fila, 6)
    _c(ws, fila, 7, None, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    total_ab_ext = sum(p['monto'] for p in partidas['abonos_ext_no_lib'])
    _c(ws, fila, 8, total_ab_ext if total_ab_ext else 0, bold=True, size=10,
       align_h='right', bg=_COLOR_SUBSEC, num_fmt='#,##0.00', border=_BORDE_THIN)
    fila_abonos_ext_inicio = fila + 1
    fila += 1

    for p in partidas['abonos_ext_no_lib']:
        ws.row_dimensions[fila].height = 15
        fecha_val = p['fecha']
        if isinstance(fecha_val, datetime):
            fecha_str = fecha_val.strftime('%d/%m/%Y')
        else:
            fecha_str = str(fecha_val) if fecha_val else ''
        _c(ws, fila, 3, fecha_str, size=10, bold=True)
        _c(ws, fila, 5, p['descripcion'], size=10, wrap=True)
        _c(ws, fila, 7, p['monto'], size=10, align_h='right', num_fmt='#,##0.00')
        fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 4
    fila += 1

    # ── (-) Cargos registrados en Extractos y no en Libros ────────────
    ws.row_dimensions[fila].height = 15
    _c(ws, fila, 2, '(-) Cargos registrados en Extractos y no en Libros',
       bold=True, size=10, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    _merge(ws, fila, 2, fila, 6)
    _c(ws, fila, 7, None, bg=_COLOR_SUBSEC, border=_BORDE_THIN)
    total_cg_ext = sum(p['monto'] for p in partidas['cargos_ext_no_lib'])
    _c(ws, fila, 8, total_cg_ext if total_cg_ext else 0, bold=True, size=10,
       align_h='right', bg=_COLOR_SUBSEC, num_fmt='#,##0.00', border=_BORDE_THIN)
    fila_cargos_ext_inicio = fila + 1
    fila += 1

    for p in partidas['cargos_ext_no_lib']:
        ws.row_dimensions[fila].height = 15
        fecha_val = p['fecha']
        if isinstance(fecha_val, datetime):
            fecha_str = fecha_val.strftime('%d/%m/%Y')
        else:
            fecha_str = str(fecha_val) if fecha_val else ''
        _c(ws, fila, 3, fecha_str, size=10, bold=True)
        _c(ws, fila, 5, p['descripcion'], size=10, wrap=True)
        _c(ws, fila, 7, p['monto'], size=10, align_h='right', num_fmt='#,##0.00')
        fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 4
    fila += 1

    # ══════════════════════════════════════════════════════════════════
    # SALDO SEGÚN EXTRACTO BANCARIO
    # ══════════════════════════════════════════════════════════════════
    ws.row_dimensions[fila].height = 15
    saldo_extracto = saldos['saldo_extracto']

    _c(ws, fila, 2, '       Saldo según Extracto Bancario', bold=True, size=11,
       bg=_COLOR_TOTAL, border=_BORDE_THIN)
    _merge(ws, fila, 2, fila+1, 5)
    _c(ws, fila, 6, None, bg=_COLOR_TOTAL, border=_BORDE_THIN)
    _c(ws, fila, 7, None, bg=_COLOR_TOTAL, border=_BORDE_THIN)
    _c(ws, fila, 8, saldo_extracto, bold=True, size=11, align_h='right',
       bg=_COLOR_TOTAL, num_fmt='#,##0.00', border=_BORDE_THIN)
    _merge(ws, fila, 8, fila+1, 8)
    fila_saldo_ext = fila
    fila += 1
    ws.row_dimensions[fila].height = 15
    fila += 1

    # Fila espaciadora
    ws.row_dimensions[fila].height = 10
    fila += 1

    # ══════════════════════════════════════════════════════════════════
    # SALDO CONCILIADO
    # ══════════════════════════════════════════════════════════════════
    ws.row_dimensions[fila].height = 15
    saldo_conciliado = round(
        saldo_libros
        - total_ab_lib
        + total_cg_lib
        + total_ab_ext
        - total_cg_ext,
        2
    )
    _c(ws, fila, 5, 'Saldo Conciliado', bold=True, size=11,
       align_h='right', color=_COLOR_TITULO)
    _merge(ws, fila, 5, fila, 7)
    _c(ws, fila, 8, saldo_conciliado, bold=True, size=11,
       align_h='right', color=_COLOR_TITULO, num_fmt='#,##0.00',
       border=Border(bottom=Side(style='double', color=_COLOR_TITULO)))
    fila += 1

    # ── Freeze ────────────────────────────────────────────────────────
    ws.freeze_panes = 'B4'

    wb.save(ruta)


# ══════════════════════════════════════════════════════════════════════
# FORMATO VISUAL – PESTAÑA CONCILIADOS
# ══════════════════════════════════════════════════════════════════════

def _aplicar_formato_conciliados(ruta: Path) -> None:
    """Aplica formato visual profesional a la hoja Conciliados."""
    wb = openpyxl.load_workbook(ruta)
    thin  = Side(style='thin', color='D9D9D9')
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    if 'Conciliados' in wb.sheetnames:
        ws = wb['Conciliados']
        color_hex = _COLOR_CONCIL

        for cell in ws[1]:
            cell.font      = Font(bold=True, color='FFFFFF', size=10, name='Aptos Narrow')
            cell.fill      = PatternFill('solid', fgColor=color_hex)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border    = borde

        ws.row_dimensions[1].height = 30

        COLS_MONTO = {'Monto-Banco', 'Monto-Conta'}
        monto_col_idxs = set()
        for i, cell in enumerate(ws[1]):
            if str(cell.value or '').strip() in COLS_MONTO:
                monto_col_idxs.add(i)

        for row in ws.iter_rows(min_row=2):
            for idx, cell in enumerate(row):
                cell.alignment = Alignment(vertical='center')
                cell.border    = borde
                if isinstance(cell.value, datetime):
                    cell.number_format = 'DD/MM/YYYY'
                if idx in monto_col_idxs:
                    cell.font = Font(bold=True, name='Aptos Narrow')

        for col in ws.columns:
            max_len    = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    largo = len(str(cell.value)) if cell.value is not None else 0
                    if largo > max_len:
                        max_len = largo
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

        ws.freeze_panes = 'A2'

    wb.save(ruta)


# ══════════════════════════════════════════════════════════════════════
# PESTAÑAS ADICIONALES DEL REPORTE FINAL
# ══════════════════════════════════════════════════════════════════════

# Paleta compartida para las pestañas de movimientos bancarios
_COLOR_BANCO_HDR    = '1F4E79'   # Azul oscuro – encabezado banco
_COLOR_INGRESO_HDR  = '375623'   # Verde oscuro – encabezado ingresos
_COLOR_EGRESO_HDR   = 'C00000'   # Rojo oscuro  – encabezado egresos
_COLOR_ITF_HDR      = '7030A0'   # Morado       – encabezado ITF y com
_COLOR_META_BG      = 'D9E1F2'   # Azul muy claro – fondo filas de metadatos
_COLOR_TOTAL_BG     = 'BDD7EE'   # Azul claro   – fila de total


def _estilo_encabezado_mov(ws, fila_hdr: int, columnas: list[str], color_hex: str) -> None:
    """Aplica estilo profesional a la fila de encabezado de una tabla de movimientos."""
    thin = Side(style='thin', color='BFBFBF')
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, nombre in enumerate(columnas, start=1):
        cell = ws.cell(row=fila_hdr, column=col_idx, value=nombre)
        cell.font = Font(bold=True, color='FFFFFF', size=10, name='Aptos Narrow')
        cell.fill = PatternFill('solid', fgColor=color_hex)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = borde
    ws.row_dimensions[fila_hdr].height = 28


def _fila_metadatos(ws, meta: dict, cuenta_label: str = '') -> int:
    """
    Escribe las 3 filas de metadatos (Cuenta, Moneda, Tipo de Cuenta) al inicio
    de una pestaña de movimientos bancarios, igual que en el Formato Propuesto.
    Devuelve el número de la siguiente fila disponible.
    """
    thin = Side(style='thin', color='BFBFBF')
    borde_meta = Border(left=thin, right=thin, top=thin, bottom=thin)

    cuenta_txt = cuenta_label or (
        f"{meta.get('cuenta', '')} - {meta.get('empresa', '')}".strip(' -')
    )
    moneda_txt  = meta.get('moneda', '')
    tipo_cta    = 'Corriente'

    for fila, (etiq, valor) in enumerate(
        [('Cuenta', cuenta_txt), ('Moneda', moneda_txt), ('Tipo de Cuenta', tipo_cta)],
        start=1
    ):
        ws.cell(row=fila, column=1, value=etiq).font = Font(bold=True, size=10, name='Aptos Narrow')
        ws.cell(row=fila, column=1).fill = PatternFill('solid', fgColor=_COLOR_META_BG)
        ws.cell(row=fila, column=1).border = borde_meta
        ws.cell(row=fila, column=2, value=valor).font = Font(size=10, name='Aptos Narrow')
        ws.cell(row=fila, column=2).border = borde_meta
        ws.row_dimensions[fila].height = 15

    return 5  # fila de inicio del encabezado de tabla (fila 4 vacía, 5 = header)


def _escribir_filas_datos(ws, datos: list[dict], columnas: list[str],
                          fila_inicio: int, color_hex: str) -> int:
    """
    Escribe las filas de datos de movimientos bancarios con estilo alternado.
    Devuelve la siguiente fila disponible (para el total).
    """
    thin = Side(style='thin', color='BFBFBF')
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    COLORES_ALT = ['FFFFFF', 'F2F8FF']  # blanco / azul muy claro alternados

    for i, dato in enumerate(datos):
        fila = fila_inicio + i
        bg = COLORES_ALT[i % 2]
        for col_idx, col_key in enumerate(columnas, start=1):
            valor = dato.get(col_key)
            cell = ws.cell(row=fila, column=col_idx, value=valor)
            cell.font = Font(size=10, name='Aptos Narrow')
            cell.fill = PatternFill('solid', fgColor=bg)
            cell.border = borde
            cell.alignment = Alignment(vertical='center')
            if isinstance(valor, (int, float)) and not isinstance(valor, bool):
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif isinstance(valor, datetime):
                cell.number_format = 'DD/MM/YYYY'
        ws.row_dimensions[fila].height = 15

    return fila_inicio + len(datos)


def _fila_total(ws, fila: int, n_cols: int, col_monto_idx: int,
                fila_inicio: int, color_hex: str) -> None:
    """Escribe la fila de TOTAL al final de una tabla de movimientos."""
    thin = Side(style='thin', color='BFBFBF')
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    col_letra = get_column_letter(col_monto_idx)
    fila_fin   = fila - 1
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=fila, column=col)
        cell.fill   = PatternFill('solid', fgColor=_COLOR_TOTAL_BG)
        cell.border = borde
        if col == 1:
            cell.value     = 'TOTAL'
            cell.font      = Font(bold=True, size=10, name='Aptos Narrow')
            cell.alignment = Alignment(horizontal='left', vertical='center')
        elif col == col_monto_idx:
            ref_inicio = f'{col_letra}{fila_inicio}'
            ref_fin    = f'{col_letra}{fila_fin}'
            cell.value        = f'=SUM({ref_inicio}:{ref_fin})'
            cell.font         = Font(bold=True, size=10, name='Aptos Narrow')
            cell.number_format = '#,##0.00'
            cell.alignment    = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[fila].height = 16


def _ajustar_columnas(ws, columnas: list[str], anchos_min: list[float]) -> None:
    """Ajusta el ancho de columnas con un mínimo definido."""
    for col_idx, (_, ancho_min) in enumerate(zip(columnas, anchos_min), start=1):
        col_letra = get_column_letter(col_idx)
        max_len = ancho_min
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                try:
                    largo = len(str(cell.value)) if cell.value is not None else 0
                    if largo > max_len:
                        max_len = largo
                except Exception:
                    pass
        ws.column_dimensions[col_letra].width = min(max_len + 2, 50)


def _df_banco_desde_inicial(ruta_inicial: Path) -> pd.DataFrame | None:
    """Lee los movimientos bancarios del archivo inicial (hoja BANCO)."""
    try:
        xl = pd.ExcelFile(ruta_inicial)
        if 'BANCO' in xl.sheet_names:
            df = pd.read_excel(ruta_inicial, sheet_name='BANCO')
            # Normalizar nombres de columnas
            df.columns = [str(c).strip() for c in df.columns]
            # Normalizar columna de monto
            for posible in ['Monto-Banco', 'monto', 'Monto']:
                if posible in df.columns:
                    df['_monto'] = pd.to_numeric(df[posible], errors='coerce').fillna(0)
                    break
            else:
                df['_monto'] = 0.0
            # Normalizar columna de descripción
            for posible in ['Descripción operación', 'Banco - Descripción', 'descripcion', 'Descripcion']:
                if posible in df.columns:
                    df['_desc'] = df[posible].astype(str)
                    break
            else:
                df['_desc'] = ''
            # Normalizar columna de operación
            for posible in ['# Operación', 'Banco - # Operación', 'nro_operacion', '# Operacion']:
                if posible in df.columns:
                    df['_op'] = df[posible].astype(str)
                    break
            else:
                df['_op'] = ''
            # Normalizar fecha
            for posible in ['Fecha', 'fecha']:
                if posible in df.columns:
                    df['_fecha'] = pd.to_datetime(df[posible], dayfirst=True, errors='coerce')
                    break
            else:
                df['_fecha'] = pd.NaT
            return df
    except Exception as e:
        print(f"  [Aviso] No se pudo leer BANCO del inicial: {e}")
    return None


def _copiar_anexar1(ruta_inicial: Path, ruta_final: Path) -> None:
    """
    Copia la hoja Anexar1 del archivo inicial al archivo final,
    preservando valores, estilos, anchos de columna y alturas de fila.
    La hoja se agrega con el nombre 'Anexar1' al final del workbook final.
    """
    try:
        wb_src = openpyxl.load_workbook(ruta_inicial)
        wb_dst = openpyxl.load_workbook(ruta_final)

        if 'Anexar1' not in wb_src.sheetnames:
            print("  [Aviso] Hoja Anexar1 no encontrada en el archivo inicial.")
            return

        ws_src = wb_src['Anexar1']

        # Eliminar si ya existe en el destino
        if 'Anexar1' in wb_dst.sheetnames:
            del wb_dst['Anexar1']

        # Crear nueva hoja al final
        ws_dst = wb_dst.create_sheet('Anexar1')

        # ── Copiar celdas (valor + estilo) ────────────────────────────
        from copy import copy
        for row in ws_src.iter_rows():
            for cell in row:
                new_cell = ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:
                    new_cell.font        = copy(cell.font)
                    new_cell.border      = copy(cell.border)
                    new_cell.fill        = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.alignment   = copy(cell.alignment)

        # ── Copiar dimensiones de columnas ────────────────────────────
        for col_letter, cd in ws_src.column_dimensions.items():
            ws_dst.column_dimensions[col_letter].width  = cd.width
            ws_dst.column_dimensions[col_letter].hidden = cd.hidden

        # ── Copiar alturas de filas ────────────────────────────────────
        for row_num, rd in ws_src.row_dimensions.items():
            ws_dst.row_dimensions[row_num].height = rd.height

        # ── Copiar celdas combinadas ───────────────────────────────────
        for merged_range in ws_src.merged_cells.ranges:
            ws_dst.merge_cells(str(merged_range))

        # ── Copiar pane de congelación ─────────────────────────────────
        ws_dst.freeze_panes = ws_src.freeze_panes

        wb_dst.save(ruta_final)
        print("  [OK] Pestaña 'Anexar1' copiada al reporte final.")
    except Exception as e:
        print(f"  [Error] No se pudo copiar Anexar1: {e}")


def _escribir_hoja_itf_com(ruta_final: Path, ruta_inicial: Path, meta: dict) -> None:
    """
    Crea (o reemplaza) la pestaña 'ITF Y COM' en el reporte final con todos los
    movimientos bancarios que corresponden a ITF o comisiones bancarias.
    Criterio: descripción contiene 'ITF', 'COMIS', 'COM', 'MANTENIM', 'ENVIO', 'PORTES'.
    """
    df = _df_banco_desde_inicial(ruta_inicial)
    if df is None or df.empty:
        return

    # Filtrar movimientos ITF / comisiones
    patron_itf = r'ITF|COMIS|MANTEN|ENVIO|PORTES|CARGO\s+PORTES|CARGO\s+MANTEN'
    mask = df['_desc'].str.contains(patron_itf, case=False, na=False, regex=True)
    df_itf = df[mask].copy()

    columnas    = ['Fecha', '# Operación', 'Descripción', 'Monto']
    anchos_min  = [12.0, 16.0, 32.0, 14.0]
    col_monto_i = 4  # índice (1-based) de la columna Monto

    wb = openpyxl.load_workbook(ruta_final)
    if 'ITF Y COM' in wb.sheetnames:
        del wb['ITF Y COM']

    # Insertar antes de 'Anexar1' si existe, sino al final
    pos = len(wb.sheetnames)
    if 'Anexar1' in wb.sheetnames:
        pos = wb.sheetnames.index('Anexar1')
    ws = wb.create_sheet('ITF Y COM', pos)

    # Metadatos
    fila_hdr = _fila_metadatos(ws, meta)

    # Encabezado de tabla
    _estilo_encabezado_mov(ws, fila_hdr, columnas, _COLOR_ITF_HDR)

    # Datos
    datos = []
    for _, row in df_itf.iterrows():
        fecha = row['_fecha']
        datos.append({
            'Fecha'       : fecha.to_pydatetime() if pd.notna(fecha) else None,
            '# Operación' : row['_op'],
            'Descripción' : row['_desc'],
            'Monto'       : float(row['_monto']),
        })

    fila_datos_inicio = fila_hdr + 1
    fila_sig = _escribir_filas_datos(ws, datos, columnas, fila_datos_inicio, _COLOR_ITF_HDR)

    # Total
    if datos:
        _fila_total(ws, fila_sig, len(columnas), col_monto_i, fila_datos_inicio, _COLOR_ITF_HDR)

    # Título de sección (fila 5 del formato propuesto)
    ws.cell(row=fila_hdr - 1, column=3, value='COMISIONES BANCARIAS / ITF').font = Font(
        bold=True, size=11, name='Aptos Narrow', color=_COLOR_ITF_HDR
    )

    _ajustar_columnas(ws, columnas, anchos_min)
    ws.freeze_panes = 'A6'

    wb.save(ruta_final)
    print(f"  [OK] Pestaña 'ITF Y COM' generada ({len(datos)} movimientos).")


def _escribir_hoja_ingresos(ruta_final: Path, ruta_inicial: Path, meta: dict) -> None:
    """
    Crea (o reemplaza) la pestaña 'INGRESOS' en el reporte final con todos los
    movimientos bancarios con monto > 0 (excluye ITF y comisiones).
    """
    df = _df_banco_desde_inicial(ruta_inicial)
    if df is None or df.empty:
        return

    patron_excluir = r'ITF|COMIS|MANTEN|ENVIO|PORTES'
    mask_ingreso = (df['_monto'] > 0) & ~df['_desc'].str.contains(
        patron_excluir, case=False, na=False, regex=True
    )
    df_ing = df[mask_ingreso].copy()

    columnas   = ['Fecha', '# Operación', 'Descripción', 'Monto']
    anchos_min = [12.0, 16.0, 32.0, 14.0]
    col_monto_i = 4

    wb = openpyxl.load_workbook(ruta_final)
    if 'INGRESOS' in wb.sheetnames:
        del wb['INGRESOS']

    pos = len(wb.sheetnames)
    if 'Anexar1' in wb.sheetnames:
        pos = wb.sheetnames.index('Anexar1')
    ws = wb.create_sheet('INGRESOS', pos)

    fila_hdr = _fila_metadatos(ws, meta)
    _estilo_encabezado_mov(ws, fila_hdr, columnas, _COLOR_INGRESO_HDR)

    datos = []
    for _, row in df_ing.iterrows():
        fecha = row['_fecha']
        datos.append({
            'Fecha'       : fecha.to_pydatetime() if pd.notna(fecha) else None,
            '# Operación' : row['_op'],
            'Descripción' : row['_desc'],
            'Monto'       : float(row['_monto']),
        })

    fila_datos_inicio = fila_hdr + 1
    fila_sig = _escribir_filas_datos(ws, datos, columnas, fila_datos_inicio, _COLOR_INGRESO_HDR)

    if datos:
        _fila_total(ws, fila_sig, len(columnas), col_monto_i, fila_datos_inicio, _COLOR_INGRESO_HDR)

    _ajustar_columnas(ws, columnas, anchos_min)
    ws.freeze_panes = 'A6'

    wb.save(ruta_final)
    print(f"  [OK] Pestaña 'INGRESOS' generada ({len(datos)} movimientos).")


def _escribir_hoja_egresos(ruta_final: Path, ruta_inicial: Path, meta: dict) -> None:
    """
    Crea (o reemplaza) la pestaña 'EGRESOS' en el reporte final con todos los
    movimientos bancarios con monto < 0 (excluye ITF y comisiones).
    """
    df = _df_banco_desde_inicial(ruta_inicial)
    if df is None or df.empty:
        return

    patron_excluir = r'ITF|COMIS|MANTEN|ENVIO|PORTES'
    mask_egreso = (df['_monto'] < 0) & ~df['_desc'].str.contains(
        patron_excluir, case=False, na=False, regex=True
    )
    df_egr = df[mask_egreso].copy()

    columnas   = ['Fecha', '# Operación', 'Descripción', 'Monto']
    anchos_min = [12.0, 16.0, 32.0, 14.0]
    col_monto_i = 4

    wb = openpyxl.load_workbook(ruta_final)
    if 'EGRESOS' in wb.sheetnames:
        del wb['EGRESOS']

    pos = len(wb.sheetnames)
    if 'Anexar1' in wb.sheetnames:
        pos = wb.sheetnames.index('Anexar1')
    ws = wb.create_sheet('EGRESOS', pos)

    fila_hdr = _fila_metadatos(ws, meta)
    _estilo_encabezado_mov(ws, fila_hdr, columnas, _COLOR_EGRESO_HDR)

    datos = []
    for _, row in df_egr.iterrows():
        fecha = row['_fecha']
        datos.append({
            'Fecha'       : fecha.to_pydatetime() if pd.notna(fecha) else None,
            '# Operación' : row['_op'],
            'Descripción' : row['_desc'],
            'Monto'       : float(row['_monto']),
        })

    fila_datos_inicio = fila_hdr + 1
    fila_sig = _escribir_filas_datos(ws, datos, columnas, fila_datos_inicio, _COLOR_EGRESO_HDR)

    if datos:
        _fila_total(ws, fila_sig, len(columnas), col_monto_i, fila_datos_inicio, _COLOR_EGRESO_HDR)

    _ajustar_columnas(ws, columnas, anchos_min)
    ws.freeze_panes = 'A6'

    wb.save(ruta_final)
    print(f"  [OK] Pestaña 'EGRESOS' generada ({len(datos)} movimientos).")


# ══════════════════════════════════════════════════════════════════════
# RESUMEN EN CONSOLA
# ══════════════════════════════════════════════════════════════════════

def _imprimir_resumen_consola(saldos: dict, partidas: dict, ruta: Path) -> None:
    print()
    print("=" * 60)
    print("  RESULTADO DE CONCILIACION FINAL")
    print("=" * 60)
    print(f"  {'Saldo Libro Bancos':<40}: {saldos['saldo_libro_bancos']:>15,.2f}")
    print(f"  {'Saldo Extracto Bancario':<40}: {saldos['saldo_extracto']:>15,.2f}")
    print(f"  {'Movimientos solo en banco (sin par conta)':<40}: {saldos['n_solo_banco']:>15,}")
    print(f"  {'Monto solo en banco':<40}: {saldos['monto_solo_banco']:>15,.2f}")
    print(f"  {'Movimientos solo en conta (sin par banco)':<40}: {saldos['n_solo_conta']:>15,}")
    print(f"  {'Monto solo en conta':<40}: {saldos['monto_solo_conta']:>15,.2f}")
    print()
    print(f"  [OK] Reporte final generado: {ruta}")
    print()
