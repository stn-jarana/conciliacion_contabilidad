"""
reporte.py
──────────
Genera el reporte Excel de conciliación bancaria siguiendo la estructura
de la plantilla 'Conciliacion Bancaria Tabla.xlsx':

  Hoja BANCO       → datos completos del banco
  Hoja CONTANET    → datos completos de contabilidad
  Hoja Anexar1     → ambas tablas apiladas (el especialista trabaja aquí)
  Hoja Resumen     → tabla resumen por TIPO/CODIGO con diferencias

Uso:
    from reporte import generar_reporte_inicial
    generar_reporte_inicial(bank, conta, ruta_salida)
"""

import re
<<<<<<< Updated upstream
import time
=======
import unicodedata
>>>>>>> Stashed changes
import pandas as pd
from pathlib import Path
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════

def generar_reporte_inicial(
    bank: pd.DataFrame,
    conta: pd.DataFrame,
    ruta_salida: Path | None = None,
    empresa: str = "Southern Textil",
    moneda: str  = "Dolares (USD)",
    saldo_contable_final: float | None = None,
    banco: str = '',
    cuenta: str = '',
) -> Path:
    """
    Genera el Excel de conciliacion inicial y su PDF de resumen.
    Devuelve la ruta del archivo Excel creado.
    Si no se pasa ruta_salida, crea los archivos en el directorio actual con timestamp.

    Genera UN archivo:
      - CBI_*.xlsx    → columnas técnicas ocultas en Anexar1 (uso del especialista)
    """
    if ruta_salida is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_salida = Path(f"Conciliacion_Inicial_{ts}.xlsx")

    _t0 = time.perf_counter()

    # ── Preparar Tab_Banco ────────────────────────────────────────────
    tab_banco = _preparar_tab_banco(bank)

    # ── Preparar Tab_Contanet ─────────────────────────────────────────
    tab_conta = _preparar_tab_contanet(conta)

    _t1 = time.perf_counter()
    print(f"  [TIEMPO] Preparar tablas: {_t1 - _t0:.2f}s")

    # ── Preparar Anexar1 (union vertical) ─────────────────────────────
    anexar1 = _preparar_anexar1(tab_banco, tab_conta, moneda=moneda, banco=banco)

    _t2 = time.perf_counter()
    print(f"  [TIEMPO] Conciliación (Anexar1): {_t2 - _t1:.2f}s")

    # ── Preparar Resumen ──────────────────────────────────────────────
    resumen = _preparar_resumen(anexar1)

    # ── Escribir Excel INICIAL (uso del especialista) ──────────────────
    with pd.ExcelWriter(ruta_salida, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
        _exportar_banco(bank).to_excel(writer,   sheet_name='BANCO',    index=False)
        _exportar_contanet(conta).to_excel(writer, sheet_name='CONTANET', index=False)
        anexar1.to_excel(writer,                 sheet_name='Anexar1',  index=False)
        resumen.to_excel(writer,                 sheet_name='Resumen',  index=False)
        # Guardar saldo contable final como metadato en celda fija de CONTANET
        if saldo_contable_final is not None:
            wb_ini = writer.book
            ws_ini = wb_ini['CONTANET']
            ws_ini.cell(row=1, column=30).value = '__SALDO_CONTABLE_FINAL__'
            ws_ini.cell(row=2, column=30).value = saldo_contable_final
        # Guardar nombre del banco como metadato
        if banco:
            wb_ini = writer.book
            ws_ini = wb_ini['CONTANET']
            ws_ini.cell(row=1, column=31).value = '__BANCO__'
            ws_ini.cell(row=2, column=31).value = banco
        # Guardar número de cuenta como metadato
        if cuenta:
            wb_ini = writer.book
            ws_ini = wb_ini['CONTANET']
            ws_ini.cell(row=1, column=32).value = '__CUENTA__'
            ws_ini.cell(row=2, column=32).value = cuenta

    _t3 = time.perf_counter()
    print(f"  [TIEMPO] Escribir Excel: {_t3 - _t2:.2f}s")

    # ── Aplicar formato visual Excel ──────────────────────────────────
    _aplicar_formato(ruta_salida)

    _t4 = time.perf_counter()
    print(f"  [TIEMPO] Aplicar formato: {_t4 - _t3:.2f}s")

    # ── Ocultar columnas técnicas en Anexar1 del reporte inicial ──────
    _ocultar_columnas_tecnicas(ruta_salida)

    _t5 = time.perf_counter()
    print(f"  [TIEMPO] TOTAL reporte inicial: {_t5 - _t0:.2f}s")

    print(f"\n  [OK] Reporte inicial generado: {ruta_salida}")

    return ruta_salida



# ══════════════════════════════════════════════════════════════════════
# PREPARACIÓN DE TABLAS
# ══════════════════════════════════════════════════════════════════════

def _exportar_banco(bank: pd.DataFrame) -> pd.DataFrame:
    """Tabla completa del banco con columnas de trabajo."""
    df = bank[['fecha', 'descripcion', 'monto', 'saldo',
               'sucursal', 'nro_operacion', 'hora', 'usuario']].copy()
    df.columns = ['Fecha', 'Descripción operación', 'Monto-Banco', 'Saldo',
                  'Sucursal', '# Operación', 'Hora', 'Usuario']
    df['MAR']           = ''
    df['SUB']           = 0
    df['TIPO']          = df['Monto-Banco'].apply(lambda x: 'INGRESO' if x > 0 else 'EGRESO')
    df['CODIGO']        = df['Monto-Banco']
    # Preserve leading zeros by treating operation numbers as strings
    df['# Operación'] = df['# Operación'].astype(str)
    df['# Operación2']  = df['# Operación']
    return df


def _exportar_contanet(conta: pd.DataFrame) -> pd.DataFrame:
    """Tabla completa de contabilidad con columna Monto-Conta."""
    df = conta[['nro_registro', 'fecha_mov', 'medio_pago', 'nro_operacion',
                'giro', 'glosa', 'ingreso', 'egreso',
                'fecha_conciliacion', 'conciliado']].copy()
    df.columns = ['# Registro', 'Fecha', 'Medio Pago', '# Operación',
                  'Giro', 'Glosa', 'Ingreso', 'Egreso',
                  'F. Conciliación', '¿Conciliado?']
    # Monto-Conta: ingresos positivos, egresos negativos
    df['Monto-Conta']   = df['Ingreso'] - df['Egreso']
    df['MAR']           = ''
    df['SUB']           = 0
    df['TIPO']          = df['Monto-Conta'].apply(lambda x: 'INGRESO' if x > 0 else 'EGRESO')
    df['CODIGO']        = df['Monto-Conta']
    df['DIF COMISON']   = 0
    df['# Operación2']  = df['# Operación']
    df['COD_SUB']       = 0
    return df


def _preparar_tab_banco(bank: pd.DataFrame) -> pd.DataFrame:
    """Tabla completa del banco para el cruce en Anexar1."""
    df = bank[['fecha', 'descripcion', 'monto', 'nro_operacion', 'saldo']].copy()
    df.columns = ['Fecha', 'Banco - Descripción', 'Monto-Banco', 'Banco - # Operación', 'Banco - Saldo']
    df['Banco - Fecha'] = df['Fecha']
    df['MAR']          = ''
    df['TIPO']         = df['Monto-Banco'].apply(lambda x: 'INGRESO' if x > 0 else 'EGRESO')
    df['CODIGO']       = df['Monto-Banco']
    # Ensure operation numbers keep leading zeros
    df['Banco - # Operación'] = df['Banco - # Operación'].astype(str)
    df['# Operación2'] = df['Banco - # Operación']
    # Marcar filas con monto 0 para separarlas al final (no conciliables)
    mask_sin_monto = df['Monto-Banco'].fillna(0) == 0
    df['Anotación'] = ''
    df.loc[mask_sin_monto, 'Anotación'] = 'Mov. sin monto (0)'
    # Propagar bandera de pendiente (movimientos de meses anteriores)
    if '_pendiente' in bank.columns:
        df['_pendiente'] = bank['_pendiente'].fillna(False).values
    else:
        df['_pendiente'] = False
    return df


def _preparar_tab_contanet(conta: pd.DataFrame) -> pd.DataFrame:
    """Tabla completa de contabilidad para el cruce en Anexar1."""
    df = conta[['nro_registro', 'fecha_mov', 'nro_operacion', 'ingreso', 'egreso', 'giro', 'glosa', 'fecha_conciliacion', 'conciliado']].copy()
    df.columns = ['Conta - # Registro', 'Fecha', 'Conta - # Operación', 'Ingreso', 'Egreso', 'Conta - Giro', 'Conta - Glosa', 'Conta - F. Conciliación', 'Conta - ¿Conciliado?']
    df['Conta - Fecha'] = df['Fecha']
    df['Monto-Conta']  = df['Ingreso'] - df['Egreso']
    df['MAR']          = ''
    df['TIPO']         = df['Monto-Conta'].apply(lambda x: 'INGRESO' if x > 0 else 'EGRESO')
    df['CODIGO']       = df['Monto-Conta']
    df['DIF COMISON']  = 0
    df['# Operación2'] = df['Conta - # Operación']
    # Marcar filas sin monto (ingreso=0 y egreso=0) para separarlas al final
    mask_sin_monto = (df['Ingreso'].fillna(0) == 0) & (df['Egreso'].fillna(0) == 0)
    df['Anotación'] = ''
    df.loc[mask_sin_monto, 'Anotación'] = 'Mov. sin monto (0)'
    # Quitar columnas auxiliares de cálculo
    df = df.drop(columns=['Ingreso', 'Egreso'])
    # Propagar bandera de pendiente (movimientos de meses anteriores)
    if '_pendiente' in conta.columns:
        df['_pendiente'] = conta['_pendiente'].fillna(False).values
    else:
        df['_pendiente'] = False
    return df


from itertools import combinations


def _diferencias_comision(moneda: str) -> set[float]:
    """Importes de comisión válidos para la moneda conciliada."""
    if "Soles" in moneda or "PEN" in moneda:
        return {10.50, 69.00, 4.30, 29.00}
    return {1.53, 82.00, 94.00, 69.00, 29.00}


def _conciliacion_automatica(banco_ext: pd.DataFrame, conta_ext: pd.DataFrame, moneda: str = "Dolares (USD)", banco: str = ""):
    """
    Algoritmo de conciliación automática con las reglas:
    1. Código igual → Conciliar automáticamente
    2. Monto + Fecha → Conciliar si la coincidencia es única
    3. Monto repetido → Ignorar
    4. Suma (1 a N) o (N a 1) → Buscar combinaciones y conciliar si es única
    """
    # ── Guardia defensiva: asegurar que todas las columnas requeridas existan ──────
    # Esto protege contra DataFrames de contabilidad o banco vacíos / mal mapeados.
    _cols_banco_req = {
        'Fecha': pd.NaT,
        'Monto-Banco': 0.0,
        'Banco - Descripción': '',
        'Banco - # Operación': '',
        '# Operación2': '',
        'MAR': '',
        'Anotación': '',
        '_pendiente': False,
    }
    _cols_conta_req = {
        'Fecha': pd.NaT,
        'Monto-Conta': 0.0,
        'Conta - # Registro': '',
        'Conta - # Operación': '',
        'Conta - Giro': '',
        'Conta - Glosa': '',
        '# Operación2': '',
        'MAR': '',
        'Anotación': '',
        '_pendiente': False,
        '_suggested': False,
    }
    # Mapear nombres alternativos de banco si vienen directamente de hoja BANCO
    if 'Descripción operación' in banco_ext.columns and 'Banco - Descripción' not in banco_ext.columns:
        banco_ext['Banco - Descripción'] = banco_ext['Descripción operación']
    if 'Descripcion operacion' in banco_ext.columns and 'Banco - Descripción' not in banco_ext.columns:
        banco_ext['Banco - Descripción'] = banco_ext['Descripcion operacion']
    if '# Operación' in banco_ext.columns and 'Banco - # Operación' not in banco_ext.columns:
        banco_ext['Banco - # Operación'] = banco_ext['# Operación']

    # Mapear nombres alternativos de contabilidad si vienen directamente de CONTANET
    if '# Registro' in conta_ext.columns and 'Conta - # Registro' not in conta_ext.columns:
        conta_ext['Conta - # Registro'] = conta_ext['# Registro']
    if 'Giro' in conta_ext.columns and 'Conta - Giro' not in conta_ext.columns:
        conta_ext['Conta - Giro'] = conta_ext['Giro']
    if 'Glosa' in conta_ext.columns and 'Conta - Glosa' not in conta_ext.columns:
        conta_ext['Conta - Glosa'] = conta_ext['Glosa']

    for col, default in _cols_banco_req.items():
        if col not in banco_ext.columns:
            banco_ext[col] = default
    for col, default in _cols_conta_req.items():
        if col not in conta_ext.columns:
            conta_ext[col] = default
    # ─────────────────────────────────────────────────────────────────────────────

    banco_ext['_matched'] = False
    conta_ext['_matched'] = False
    diferencias_set = _diferencias_comision(moneda)

    # Excluir de la conciliación los movimientos de Contanet sin monto (ingreso=0 y egreso=0)
    # Se identifican porque tienen Monto-Conta == 0 y la anotación 'Mov. sin monto (0)'
    mask_sin_monto = conta_ext.get('Anotación', pd.Series([''] * len(conta_ext))).str.contains('sin monto', case=False, na=False)
    conta_ext.loc[mask_sin_monto, '_matched'] = True  # Marcar como ya procesados para ignorarlos

    # Excluir de la conciliación los movimientos de Banco con monto 0
    mask_sin_monto_b = banco_ext.get('Anotación', pd.Series([''] * len(banco_ext))).str.contains('sin monto', case=False, na=False)
    banco_ext.loc[mask_sin_monto_b, '_matched'] = True  # Marcar como ya procesados para ignorarlos

    # Helper para normalizar y comparar números de operación
    def _normalizar_nro_op(val) -> str:
        if val is None or pd.isna(val):
            return ""
        s = str(val).strip()
        if s.lower() == "nan" or s == "0":
            return ""
        if s.endswith(".0"):
            s = s[:-2].strip()
        return s

    def _coincide_nro_op(op_b: str, op_c: str) -> bool:
        import re

        nb = _normalizar_nro_op(op_b)
        nc = _normalizar_nro_op(op_c)
        if not nb or not nc:
            return False
        # Coincidencia por últimos 6 dígitos (aunque los demás dígitos sean distintos)
        db = re.sub(r'\D', '', nb)
        dc = re.sub(r'\D', '', nc)
        if len(db) >= 6 and len(dc) >= 6:
            if db[-6:] == dc[-6:]:
                return True
        # Exactamente iguales con al menos 3 caracteres
        if nb == nc and len(nb) >= 3:
            return True
        sb = nb.lstrip('0') or nb
        sc = nc.lstrip('0') or nc
        if sb == sc and len(sb) >= 3:
            return True
        # Conta es sufijo de Banco (mínimo 4 caracteres)
        variantes_c = [v for v in {nc, sc} if len(v) >= 4]
        for vc in variantes_c:
            if nb.endswith(vc):
                return True
        # Banco es sufijo de Conta (mínimo 4 caracteres)
        variantes_b = [v for v in {nb, sb} if len(v) >= 4]
        for vb in variantes_b:
            if nc.endswith(vb):
                return True
        return False

    # ── Pre-normalizar # Operación UNA VEZ para evitar llamadas repetidas ────────
    # Se calculan en vectores y se almacenan en columnas temporales _op_norm y _op_sufijo6.
    import re as _re_inline

    def _norm_series(ser: pd.Series) -> pd.Series:
        """Normaliza la serie de # Operación a string limpio."""
        s = ser.astype(str).str.strip()
        s = s.where(~s.str.lower().isin(['nan', '0', 'none']), '')
        s = s.where(~s.str.endswith('.0'), s.str[:-2].str.strip())
        return s

    banco_ext['_op_norm'] = _norm_series(banco_ext['Banco - # Operación'])
    conta_ext['_op_norm'] = _norm_series(conta_ext['Conta - # Operación'])

    # Sufijo de 6 dígitos (solo numéricos)
    def _sufijo6(norm_op: str) -> str:
        digits = _re_inline.sub(r'\D', '', norm_op)
        return digits[-6:] if len(digits) >= 6 else ''

    banco_ext['_op_suf6'] = banco_ext['_op_norm'].apply(_sufijo6)
    conta_ext['_op_suf6'] = conta_ext['_op_norm'].apply(_sufijo6)
    banco_ext['_monto_r'] = banco_ext['Monto-Banco'].fillna(0).astype(float).round(2)
    conta_ext['_monto_r'] = conta_ext['Monto-Conta'].fillna(0).astype(float).round(2)

    def _rebuild_suf6_index():
        """Reconstruye el índice sufijo6 → [b_idxs] con los no conciliados actuales."""
        idx: dict[str, list] = {}
        for b_idx in banco_ext.index[~banco_ext['_matched']]:
            suf = banco_ext.at[b_idx, '_op_suf6']
            if suf:
                idx.setdefault(suf, []).append(b_idx)
        return idx

    def _candidatos_por_op(c_idx, suf6_index) -> list:
        """Devuelve b_idxs donde _coincide_nro_op vale True para el conta dado."""
        op_c   = conta_ext.at[c_idx, '_op_norm']
        suf6_c = conta_ext.at[c_idx, '_op_suf6']
        if not op_c:
            return []

        candidatos: list = []
        # Lookup por sufijo6 (hit rápido)
        hits_suf6 = suf6_index.get(suf6_c, []) if suf6_c else []
        for b_idx in hits_suf6:
            if not banco_ext.at[b_idx, '_matched']:
                candidatos.append(b_idx)
        if candidatos:
            return candidatos

        # Fallback: comparación por sufijo textual (cubre nb.endswith, etc.)
        sc_norm = op_c.lstrip('0') or op_c
        variantes_c = [v for v in {op_c, sc_norm} if len(v) >= 4]
        for b_idx in banco_ext.index[~banco_ext['_matched']]:
            op_b = banco_ext.at[b_idx, '_op_norm']
            if not op_b:
                continue
            sb_norm = op_b.lstrip('0') or op_b
            # op_c es sufijo de op_b
            for vc in variantes_c:
                if op_b.endswith(vc):
                    candidatos.append(b_idx)
                    break
            else:
                # op_b es sufijo de op_c
                variantes_b = [v for v in {op_b, sb_norm} if len(v) >= 4]
                for vb in variantes_b:
                    if op_c.endswith(vb):
                        candidatos.append(b_idx)
                        break
        return candidatos

    # 1. Reglas por coincidencia de # Operación (exacto o sufijo)
    # 1.1 Mismo valor (Monto igual) -> Conciliación automática con "Auto: Sufijo de # Operación"
    suf6_idx = _rebuild_suf6_index()
    for c_idx in conta_ext.index[~conta_ext['_matched']]:
        op_c = conta_ext.at[c_idx, '_op_norm']
        if not op_c: continue
        m_c = conta_ext.at[c_idx, '_monto_r']
        if m_c == 0: continue

        candidatos_b = [
            b for b in _candidatos_por_op(c_idx, suf6_idx)
            if abs(banco_ext.at[b, '_monto_r'] - m_c) <= 0.01
        ]

        b_idx_elegido = None
        if len(candidatos_b) == 1:
            b_idx_elegido = candidatos_b[0]
        elif len(candidatos_b) > 1:
            cands_fecha = [b for b in candidatos_b if banco_ext.at[b, 'Fecha'] == conta_ext.at[c_idx, 'Fecha']]
            if len(cands_fecha) == 1:
                b_idx_elegido = cands_fecha[0]
            else:
                b_idx_elegido = candidatos_b[0]

        if b_idx_elegido is not None:
            b_idx = b_idx_elegido
            op_b = banco_ext.at[b_idx, '_op_norm']
            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'MAR'] = 'X'
            conta_ext.at[c_idx, 'MAR'] = 'X'
            op_link = op_b or op_c
            banco_ext.at[b_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, '# Operación2'] = op_link
            banco_ext.at[b_idx, 'Anotación'] = 'Auto: Sufijo de # Operación'
            conta_ext.at[c_idx, 'Anotación'] = 'Auto: Sufijo de # Operación'

    # 1.2 Diferencia en los montos correspondiente a comisiones -> Conciliación automática con "Auto: Sufijo de # Operación"
    suf6_idx = _rebuild_suf6_index()
    for c_idx in conta_ext.index[~conta_ext['_matched']]:
        op_c = conta_ext.at[c_idx, '_op_norm']
        if not op_c: continue
        m_c = conta_ext.at[c_idx, '_monto_r']
        if m_c == 0: continue

        candidatos_b = []
        for b_idx in _candidatos_por_op(c_idx, suf6_idx):
            m_b = banco_ext.at[b_idx, '_monto_r']
            if (m_b > 0) == (m_c > 0):
                dif = round(abs(m_b - m_c), 2)
                if dif in diferencias_set:
                    candidatos_b.append((b_idx, round(m_b - m_c, 2)))

        b_idx_elegido = None
        dif_elegido = 0.0
        if len(candidatos_b) == 1:
            b_idx_elegido, dif_elegido = candidatos_b[0]
        elif len(candidatos_b) > 1:
            cands_fecha = [(b, d) for (b, d) in candidatos_b if banco_ext.at[b, 'Fecha'] == conta_ext.at[c_idx, 'Fecha']]
            if len(cands_fecha) == 1:
                b_idx_elegido, dif_elegido = cands_fecha[0]
            else:
                b_idx_elegido, dif_elegido = candidatos_b[0]

        if b_idx_elegido is not None:
            b_idx = b_idx_elegido
            op_b = banco_ext.at[b_idx, '_op_norm']
            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'MAR'] = 'X'
            conta_ext.at[c_idx, 'MAR'] = 'X'
            op_link = op_b or op_c
            banco_ext.at[b_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, '# Operación2'] = op_link
            banco_ext.at[b_idx, 'Anotación'] = 'Auto: Sufijo de # Operación'
            conta_ext.at[c_idx, 'Anotación'] = 'Auto: Sufijo de # Operación'
            banco_ext.at[b_idx, 'DIF COMISON'] = dif_elegido

    # 1.3 Diferencia distinta a todas las comisiones -> Sugerido para evaluación del especialista
    if '_suggested' not in conta_ext.columns:
        conta_ext['_suggested'] = False

    suf6_idx = _rebuild_suf6_index()
    for c_idx in conta_ext.index[~conta_ext['_matched']]:
        if conta_ext.at[c_idx, '_suggested']: continue
        op_c = conta_ext.at[c_idx, '_op_norm']
        if not op_c: continue
        m_c = conta_ext.at[c_idx, '_monto_r']
        if m_c == 0: continue

        candidatos_b = []
        for b_idx in _candidatos_por_op(c_idx, suf6_idx):
            if str(banco_ext.at[b_idx, 'Anotación']).startswith('Sugerido'): continue
            m_b = banco_ext.at[b_idx, '_monto_r']
            if (m_b > 0) == (m_c > 0):
                dif = round(abs(m_b - m_c), 2)
                if dif > 0.01 and dif not in diferencias_set:
                    candidatos_b.append((b_idx, dif))

        b_idx_elegido = None
        dif_elegido = 0.0
        if len(candidatos_b) == 1:
            b_idx_elegido, dif_elegido = candidatos_b[0]
        elif len(candidatos_b) > 1:
            cands_fecha = [(b, d) for (b, d) in candidatos_b if banco_ext.at[b, 'Fecha'] == row_c['Fecha']]
            if len(cands_fecha) == 1:
                b_idx_elegido, dif_elegido = cands_fecha[0]

        if b_idx_elegido is not None:
            b_idx = b_idx_elegido
            op_b = banco_ext.at[b_idx, '_op_norm']
            op_link = f"SUG-OP-DIF-{op_b or op_c}"
            banco_ext.at[b_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, '# Operación2'] = op_link
            anot = f"Sugerido: Sufijo de # Operación, diferencia de monto {dif_elegido:.2f}"
            banco_ext.at[b_idx, 'Anotación'] = anot
            conta_ext.at[c_idx, 'Anotación'] = anot
            conta_ext.at[c_idx, '_suggested'] = True
            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'DIF COMISON'] = round(float(banco_ext.at[b_idx, 'Monto-Banco']) - float(conta_ext.at[c_idx, 'Monto-Conta']), 2)

    # 1.5 # Operación de Conta contenido en Descripción de Banco + Monto igual
    for c_idx in conta_ext.index[~conta_ext['_matched']]:
        op_c = conta_ext.at[c_idx, '_op_norm']
        # Evitar falsos positivos requiriendo al menos 4 caracteres en la operación
        if not op_c or len(op_c) < 4: continue
        
        m_c = conta_ext.at[c_idx, '_monto_r']
        
        candidates = banco_ext[
            (~banco_ext['_matched']) & 
            (abs(banco_ext['Monto-Banco'].fillna(0) - m_c) <= 0.01) &
            (banco_ext['Banco - Descripción'].astype(str).str.contains(op_c, case=False, na=False, regex=False))
        ]
        
        if len(candidates) == 1:
            b_idx = candidates.index[0]
            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'MAR'] = 'X'
            conta_ext.at[c_idx, 'MAR'] = 'X'
            banco_ext.at[b_idx, '# Operación2'] = op_c
            conta_ext.at[c_idx, '# Operación2'] = op_c
            banco_ext.at[b_idx, 'Anotación'] = 'Auto: # Operación en Glosa Banco'
            conta_ext.at[c_idx, 'Anotación'] = 'Auto: # Operación en Glosa Banco'


    # 1.65 Conciliación por descarte cuando quedan exactamente 1 banco y 1 conta no conciliados
    # con mismo monto y misma fecha (fecha exacta). Cubre el caso donde todos los demás ya
    # fueron emparejados por reglas anteriores y solo queda 1 de cada lado.

    # Guardia defensiva: asegurar que la columna 'Fecha' exista en ambos DataFrames
    if 'Fecha' not in banco_ext.columns:
        banco_ext['Fecha'] = pd.NaT
    if 'Fecha' not in conta_ext.columns:
        conta_ext['Fecha'] = pd.NaT

    fechas_presentes = set(
        banco_ext.loc[~banco_ext['_matched'], 'Fecha'].dropna().unique()
    ).union(
        set(conta_ext.loc[~conta_ext['_matched'], 'Fecha'].dropna().unique())
    )
    for f in fechas_presentes:
        unm_b = banco_ext[(~banco_ext['_matched']) & (banco_ext['Fecha'] == f)]
        unm_c = conta_ext[(~conta_ext['_matched']) & (conta_ext['Fecha'] == f)]

        montos_b = unm_b['Monto-Banco'].apply(lambda v: round(float(v or 0), 2))
        montos_c = unm_c['Monto-Conta'].apply(lambda v: round(float(v or 0), 2))

        # Buscar montos que aparecen exactamente 1 vez en banco y 1 vez en conta
        for m_val in montos_b.unique():
            idxs_b = unm_b[montos_b == m_val].index.tolist()
            idxs_c = unm_c[montos_c == m_val].index.tolist()
            if len(idxs_b) == 1 and len(idxs_c) == 1:
                b_idx = idxs_b[0]
                c_idx = idxs_c[0]
                op_b = str(banco_ext.at[b_idx, 'Banco - # Operación']).strip()
                op_link = op_b if (op_b and op_b != 'nan') else f"AUTO-DESC-{b_idx}"
                banco_ext.at[b_idx, '_matched'] = True
                conta_ext.at[c_idx, '_matched'] = True
                banco_ext.at[b_idx, 'MAR'] = 'X'
                conta_ext.at[c_idx, 'MAR'] = 'X'
                banco_ext.at[b_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, '# Operación2'] = op_link
                banco_ext.at[b_idx, 'Anotación'] = 'Auto: Descarte único por fecha+monto'
                conta_ext.at[c_idx, 'Anotación'] = 'Auto: Descarte único por fecha+monto'

    # 2. Monto + Fecha única

    freq_b = {}
    freq_c = {}
    for b_idx, row_b in banco_ext[~banco_ext['_matched']].iterrows():
        k = (row_b['Fecha'], round(float(row_b.get('Monto-Banco', 0) or 0), 2))
        freq_b[k] = freq_b.get(k, []) + [b_idx]
        
    for c_idx, row_c in conta_ext[~conta_ext['_matched']].iterrows():
        k = (row_c['Fecha'], round(float(row_c.get('Monto-Conta', 0) or 0), 2))
        freq_c[k] = freq_c.get(k, []) + [c_idx]

    for k, b_idxs in freq_b.items():
        if len(b_idxs) == 1 and k in freq_c and len(freq_c[k]) == 1:
            b_idx = b_idxs[0]
            c_idx = freq_c[k][0]
            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'MAR'] = 'X'
            conta_ext.at[c_idx, 'MAR'] = 'X'
            
            op_link = str(banco_ext.at[b_idx, 'Banco - # Operación'])
            if op_link == 'nan' or not op_link.strip(): op_link = f"AUTO-{b_idx}"
            banco_ext.at[b_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, '# Operación2'] = op_link
            
            banco_ext.at[b_idx, 'Anotación'] = 'Auto: Monto+Fecha única'
            conta_ext.at[c_idx, 'Anotación'] = 'Auto: Monto+Fecha única'

    # 3. N Banco = N Conta mismo monto y fecha (sugerencia por emparejamiento posicional)
    # Si hay exactamente N banco y N conta no conciliados con el mismo monto y fecha (N >= 2),
    # no se puede saber cuál es cuál, pero sí que se corresponden 1 a 1.
    # Se emparejan en orden de aparición como sugerencia.
    if '_suggested' not in conta_ext.columns:
        conta_ext['_suggested'] = False

    freq_b2 = {}
    freq_c2 = {}
    for b_idx, row_b in banco_ext[~banco_ext['_matched']].iterrows():
        k = (row_b['Fecha'], round(float(row_b.get('Monto-Banco', 0) or 0), 2))
        freq_b2.setdefault(k, []).append(b_idx)

    for c_idx, row_c in conta_ext[~conta_ext['_matched']].iterrows():
        k = (row_c['Fecha'], round(float(row_c.get('Monto-Conta', 0) or 0), 2))
        freq_c2.setdefault(k, []).append(c_idx)

    for k, b_idxs in freq_b2.items():
        if k not in freq_c2:
            continue
        c_idxs = freq_c2[k]
        n = len(b_idxs)
        # Solo aplica cuando hay el mismo número en ambos lados y es >= 2
        if n < 2 or len(c_idxs) != n:
            continue
        # Emparejar posicionalmente
        for pair_i, (b_idx, c_idx) in enumerate(zip(b_idxs, c_idxs)):
            op_b = str(banco_ext.at[b_idx, 'Banco - # Operación']).strip()
            op_link = op_b if (op_b and op_b != 'nan') else f"SUG-NxN-{b_idx}"
            # Usar op_link único por par para que queden como pares distintos en el Excel
            op_link_par = f"{op_link}-P{pair_i}"
            banco_ext.at[b_idx, '# Operación2'] = op_link_par
            conta_ext.at[c_idx, '# Operación2'] = op_link_par
            anot = 'Sugerido: N Banco = N Conta mismo monto y fecha'
            banco_ext.at[b_idx, 'Anotación'] = anot
            conta_ext.at[c_idx, 'Anotación'] = anot
            conta_ext.at[c_idx, '_suggested'] = True

    diferencias = diferencias_set.union({0.0})
    if '_suggested' not in conta_ext.columns:
        conta_ext['_suggested'] = False
        
    # 4. Sumas N a 1 y 1 a N (COMO SUGERENCIAS)
    fechas = set(banco_ext.loc[~banco_ext['_matched'], 'Fecha']).union(
             set(conta_ext.loc[~conta_ext['_matched'], 'Fecha']))
             
    for f in fechas:
        # 1 Banco = N Conta
        unmatched_b = banco_ext[(~banco_ext['_matched']) & (banco_ext['Fecha'] == f)]
        unmatched_c = conta_ext[(~conta_ext['_matched']) & (conta_ext['Fecha'] == f)]
        pool_b = [(idx, float(r.get('Monto-Banco', 0) or 0)) for idx, r in unmatched_b.iterrows()]
        pool_c = [(idx, float(r.get('Monto-Conta', 0) or 0)) for idx, r in unmatched_c.iterrows() if not r.get('_suggested', False)]
        
        used_c = set()
        for b_idx, b_amt in pool_b:
            avail_c = [x for x in pool_c if x[0] not in used_c]
            # Guard anti-explosión combinatoria: si hay demasiados candidatos, omitir
            if len(avail_c) > 25:
                continue
            valid_combos = []
            for r in range(2, min(5, len(avail_c) + 1)):
                for combo in combinations(avail_c, r):
                    suma_conta = sum(x[1] for x in combo)
                    diff = round(abs(abs(b_amt) - abs(suma_conta)), 2)
                    # La diferencia debe pertenecer al set permitido Y ser menor
                    # que el monto del banco (evita asociar -10.5 con dos -10.5
                    # alegando una "diferencia de comisión" de 10.5)
                    if diff in diferencias and diff < abs(b_amt):
                        valid_combos.append((combo, diff))
                        
            if len(valid_combos) == 1:
                combo_idxs = [x[0] for x in valid_combos[0][0]]
                diff_val = valid_combos[0][1]
                used_c.update(combo_idxs)
                
                op_link = f"SUG-SUM-B{b_idx}"
                banco_ext.at[b_idx, '# Operación2'] = op_link
                
                if diff_val == 0.0:
                    anot = 'Sugerido: 1 Banco = N Conta'
                else:
                    anot = f'Sugerido: 1 Banco = N Conta (Dif {diff_val})'
                    
                banco_ext.at[b_idx, 'Anotación'] = anot
                for c_idx in combo_idxs:
                    conta_ext.at[c_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, 'Anotación'] = anot
                    conta_ext.at[c_idx, '_suggested'] = True
                    
        # N Banco = 1 Conta (refrescar pools)
        unmatched_c2 = conta_ext[(~conta_ext['_matched']) & (conta_ext['Fecha'] == f)]
        unmatched_b2 = banco_ext[(~banco_ext['_matched']) & (banco_ext['Fecha'] == f)]
        pool_c2 = [(idx, float(r.get('Monto-Conta', 0) or 0)) for idx, r in unmatched_c2.iterrows() if not r.get('_suggested', False)]
        pool_b2 = [(idx, float(r.get('Monto-Banco', 0) or 0)) for idx, r in unmatched_b2.iterrows() if not str(banco_ext.at[idx, 'Anotación']).startswith('Sugerido')]
        
        used_b = set()
        for c_idx, c_amt in pool_c2:
            avail_b = [x for x in pool_b2 if x[0] not in used_b]
            # Guard anti-explosión combinatoria: si hay demasiados candidatos, omitir
            if len(avail_b) > 25:
                continue
            valid_combos = []
            for r in range(2, min(5, len(avail_b) + 1)):
                for combo in combinations(avail_b, r):
                    suma_banco = sum(x[1] for x in combo)
                    diff = round(abs(abs(suma_banco) - abs(c_amt)), 2)
                    # La diferencia debe pertenecer al set permitido Y ser menor
                    # que el monto de conta (misma validación que en 1 Banco = N Conta)
                    if diff in diferencias and diff < abs(c_amt):
                        valid_combos.append((combo, diff))
                        
            if len(valid_combos) == 1:
                combo_idxs = [x[0] for x in valid_combos[0][0]]
                diff_val = valid_combos[0][1]
                used_b.update(combo_idxs)
                
                op_link = f"SUG-SUM-C{c_idx}"
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, '_suggested'] = True
                
                if diff_val == 0.0:
                    anot = 'Sugerido: N Banco = 1 Conta'
                else:
                    anot = f'Sugerido: N Banco = 1 Conta (Dif {diff_val})'
                    
                conta_ext.at[c_idx, 'Anotación'] = anot
                for b_idx in combo_idxs:
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    banco_ext.at[b_idx, 'Anotación'] = anot

    # 4.3b Segundo pase de sumas sin restricción de fecha (Opción A)
    # Cubre casos donde N movimientos del banco de distintas fechas suman exactamente
    # 1 registro de contabilidad (o viceversa). Solo se acepta diferencia == 0.0 para
    # evitar falsos positivos al relajar la restricción de fecha.
    # Se ejecuta DESPUÉS del pase por fecha exacta, usando solo los aún no conciliados/sugeridos.
    _pool_b_global = [
        (idx, float(r.get('Monto-Banco', 0) or 0))
        for idx, r in banco_ext[~banco_ext['_matched']].iterrows()
        if not str(banco_ext.at[idx, 'Anotación']).startswith('Sugerido')
    ]
    _pool_c_global = [
        (idx, float(r.get('Monto-Conta', 0) or 0))
        for idx, r in conta_ext[~conta_ext['_matched']].iterrows()
        if not r.get('_suggested', False)
        and not str(conta_ext.at[idx, 'Anotación']).startswith('Sugerido')
    ]

    _used_c_global = set()
    # 1 Banco = N Conta (fechas distintas, monto exacto)
    for b_idx, b_amt in _pool_b_global:
        if b_amt == 0:
            continue
        avail_c = [x for x in _pool_c_global if x[0] not in _used_c_global]
        # Guard anti-explosión combinatoria
        if len(avail_c) > 25:
            continue
        valid_combos = []
        for _r in range(2, min(7, len(avail_c) + 1)):
            for combo in combinations(avail_c, _r):
                suma_conta = sum(x[1] for x in combo)
                diff = round(abs(abs(b_amt) - abs(suma_conta)), 2)
                if diff == 0.0:
                    valid_combos.append((combo, diff))
        if len(valid_combos) == 1:
            combo_idxs = [x[0] for x in valid_combos[0][0]]
            _used_c_global.update(combo_idxs)
            op_link = f"SUG-SUM-B{b_idx}"
            anot = 'Sugerido: 1 Banco = N Conta (fechas distintas)'
            banco_ext.at[b_idx, '# Operación2'] = op_link
            banco_ext.at[b_idx, 'Anotación'] = anot
            for c_idx in combo_idxs:
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, 'Anotación'] = anot
                conta_ext.at[c_idx, '_suggested'] = True

    _used_b_global = set()
    # N Banco = 1 Conta (fechas distintas, monto exacto)
    for c_idx, c_amt in _pool_c_global:
        if c_amt == 0 or c_idx in _used_c_global:
            continue
        avail_b = [x for x in _pool_b_global if x[0] not in _used_b_global]
        # Guard anti-explosión combinatoria
        if len(avail_b) > 25:
            continue
        valid_combos = []
        for _r in range(2, min(7, len(avail_b) + 1)):
            for combo in combinations(avail_b, _r):
                suma_banco = sum(x[1] for x in combo)
                diff = round(abs(abs(suma_banco) - abs(c_amt)), 2)
                if diff == 0.0:
                    valid_combos.append((combo, diff))
        if len(valid_combos) == 1:
            combo_idxs = [x[0] for x in valid_combos[0][0]]
            _used_b_global.update(combo_idxs)
            op_link = f"SUG-SUM-C{c_idx}"
            anot = 'Sugerido: N Banco = 1 Conta (fechas distintas)'
            conta_ext.at[c_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, '_suggested'] = True
            conta_ext.at[c_idx, 'Anotación'] = anot
            for b_idx in combo_idxs:
                banco_ext.at[b_idx, '# Operación2'] = op_link
                banco_ext.at[b_idx, 'Anotación'] = anot

    # 4.35 Regla Scotiabank: TBK-PAGO PLANILLAS en Banco emparejado con PAGO DE HABERES en Conta (si el monto coincide)
    mask_tbk_plan = (
        (~banco_ext['_matched']) &
        (banco_ext['Banco - Descripción'].astype(str).str.contains(r'TBK-PAGO\s+PLANILLAS', case=False, na=False))
    )
    for b_idx in banco_ext[mask_tbk_plan].index:
        m_b = float(banco_ext.at[b_idx, 'Monto-Banco'] or 0)
        if m_b == 0:
            continue
        # Buscar en contabilidad registros con 'PAGO DE HABERES' o relacionado a haberes/planillas en Giro o Glosa
        cands_c = conta_ext[
            (~conta_ext['_matched']) &
            (abs(conta_ext['Monto-Conta'].fillna(0) - m_b) <= 0.01) &
            (
                conta_ext['Conta - Giro'].astype(str).str.contains(r'PAGO\s+DE\s+HABERES|HABERES|PLANILLA', case=False, na=False) |
                conta_ext['Conta - Glosa'].astype(str).str.contains(r'PAGO\s+DE\s+HABERES|HABERES|PLANILLA', case=False, na=False)
            )
        ]
        if not cands_c.empty:
            # Si hay coincidencia por fecha, preferirla; si no, la primera coincidencia de monto
            cands_fecha = cands_c[cands_c['Fecha'] == banco_ext.at[b_idx, 'Fecha']]
            c_idx = cands_fecha.index[0] if not cands_fecha.empty else cands_c.index[0]

            op_b = _normalizar_nro_op(banco_ext.at[b_idx, 'Banco - # Operación'])
            op_c = _normalizar_nro_op(conta_ext.at[c_idx, 'Conta - # Operación'])
            op_link = op_b or op_c or f"AUTO-TBK-PLAN-{b_idx}"

            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'MAR'] = 'X'
            conta_ext.at[c_idx, 'MAR'] = 'X'
            banco_ext.at[b_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, '# Operación2'] = op_link
            banco_ext.at[b_idx, 'Anotación'] = 'Auto: TBK-PAGO PLANILLAS a Haberes'
            conta_ext.at[c_idx, 'Anotación'] = 'Auto: TBK-PAGO PLANILLAS a Haberes'

    # 4.38 Agrupación N Banco = 1 Conta por # Operación (Lotes / Detracciones BN / Monto Exacto)
    # Cubre casos donde un grupo de movimientos en Banco comparten el mismo # Operación (o lote)
    # y su suma total coincide exactamente con un movimiento o asiento de Contabilidad.
    # Diseñado para lotes masivos (ej. cientos de detracciones de Banco de la Nación 'VA 1721')
    # y transferencias en bloque de cualquier banco con monto exacto, sin sufrir explosión combinatoria.
    es_banco_nacion = (
        bool(re.search(r'(?i)\b(BN|NACION|BANCO\s+DE\s+LA\s+NACION)\b', str(banco))) or
        banco_ext['Banco - Descripción'].astype(str).str.contains(r'(?i)\bVA\s+\d+\b', regex=True).any()
    )

    unm_b_mask = (
        (~banco_ext['_matched']) &
        (~banco_ext['Anotación'].astype(str).str.startswith('Sugerido'))
    )

    grupos_b_op = {}
    for b_idx in banco_ext[unm_b_mask].index:
        op_val = _normalizar_nro_op(banco_ext.at[b_idx, 'Banco - # Operación'])
        if not op_val or op_val in ('0', 'nan'):
            desc_b = str(banco_ext.at[b_idx, 'Banco - Descripción'] or '')
            m_va = re.search(r'(?i)\bVA\s+(\d+)\b', desc_b)
            if m_va:
                op_val = m_va.group(1)
        if op_val and op_val not in ('0', 'nan'):
            grupos_b_op.setdefault(op_val, []).append(b_idx)

    for op_b, b_idxs in list(grupos_b_op.items()):
        if len(b_idxs) < 2:
            continue

        b_idxs_validos = [
            i for i in b_idxs
            if not banco_ext.at[i, '_matched'] and not str(banco_ext.at[i, 'Anotación']).startswith('Sugerido')
        ]
        if len(b_idxs_validos) < 2:
            continue

        suma_b = round(sum(float(banco_ext.at[i, 'Monto-Banco'] or 0) for i in b_idxs_validos), 2)
        if suma_b == 0:
            continue

        fechas_b = [banco_ext.at[i, 'Fecha'] for i in b_idxs_validos if pd.notna(banco_ext.at[i, 'Fecha'])]

        cands_c = []
        for c_idx, row_c in conta_ext.iterrows():
            if row_c['_matched'] or row_c.get('_suggested', False) or str(row_c.get('Anotación', '')).startswith('Sugerido'):
                continue
            m_c = float(row_c.get('Monto-Conta', 0) or 0)
            if (m_c > 0) != (suma_b > 0):
                continue
            diff = round(abs(abs(m_c) - abs(suma_b)), 2)
            if diff <= 0.01:
                f_c = row_c.get('Fecha')
                dias_min = None
                if pd.notna(f_c) and fechas_b:
                    dias_min = min(abs((f_c - fb).days) for fb in fechas_b)
                cands_c.append((c_idx, diff, dias_min, row_c))

        elegido_c = None
        if len(cands_c) == 1:
            c_idx, diff, dias_min, row_c = cands_c[0]
            if dias_min is not None and dias_min <= 5:
                elegido_c = c_idx
            elif dias_min is None or es_banco_nacion or abs(suma_b) >= 1000:
                f_c = row_c.get('Fecha')
                if pd.notna(f_c) and fechas_b:
                    if any(f_c.year == fb.year and f_c.month == fb.month for fb in fechas_b):
                        elegido_c = c_idx
                else:
                    elegido_c = c_idx
        elif len(cands_c) > 1:
            cands_glosa = [
                c for c in cands_c
                if (es_banco_nacion and re.search(r'(?i)\bDETRACC\w*\b', str(c[3].get('Conta - Glosa', '')))) or
                   (op_b in str(c[3].get('Conta - Glosa', '')))
            ]
            if len(cands_glosa) == 1:
                elegido_c = cands_glosa[0][0]
            else:
                cands_con_dias = [c for c in cands_c if c[2] is not None]
                if cands_con_dias:
                    cands_con_dias.sort(key=lambda x: x[2])
                    if len(cands_con_dias) == 1 or cands_con_dias[0][2] < cands_con_dias[1][2]:
                        elegido_c = cands_con_dias[0][0]

        if elegido_c is not None:
            c_idx = elegido_c
            glosa_c = str(conta_ext.at[c_idx, 'Conta - Glosa'] or '')
            es_detracc = es_banco_nacion or bool(re.search(r'(?i)\bDETRACC\w*\b', glosa_c))

            op_link = f"SUG-LOTE-{op_b}"
            if es_detracc:
                anot = f"Sugerido: {len(b_idxs_validos)} Banco = 1 Conta (Detracciones Lote {op_b})"
            else:
                anot = f"Sugerido: {len(b_idxs_validos)} Banco = 1 Conta (Lote {op_b})"

            conta_ext.at[c_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, 'Anotación'] = anot
            conta_ext.at[c_idx, '_suggested'] = True

            for b_i in b_idxs_validos:
                banco_ext.at[b_i, '# Operación2'] = op_link
                banco_ext.at[b_i, 'Anotación'] = anot
        else:
            # Buscar coincidencia con un Asiento Contable completo (# Registro)
            asientos_cand = {}
            for c_idx, row_c in conta_ext.iterrows():
                if row_c['_matched'] or row_c.get('_suggested', False) or str(row_c.get('Anotación', '')).startswith('Sugerido'):
                    continue
                reg = _normalizar_nro_op(row_c.get('Conta - # Registro'))
                if reg and reg not in ('0', 'nan'):
                    asientos_cand.setdefault(reg, []).append(c_idx)

            for reg, c_idxs_reg in list(asientos_cand.items()):
                if len(c_idxs_reg) >= 2:
                    suma_reg = round(sum(float(conta_ext.at[i, 'Monto-Conta'] or 0) for i in c_idxs_reg), 2)
                    if (suma_reg > 0) == (suma_b > 0) and round(abs(abs(suma_reg) - abs(suma_b)), 2) <= 0.01:
                        op_link = f"SUG-LOTE-{op_b}"
                        anot = f"Sugerido: {len(b_idxs_validos)} Banco = Asiento Conta {reg} (Lote {op_b})"
                        for i in c_idxs_reg:
                            conta_ext.at[i, '# Operación2'] = op_link
                            conta_ext.at[i, 'Anotación'] = anot
                            conta_ext.at[i, '_suggested'] = True
                        for b_i in b_idxs_validos:
                            banco_ext.at[b_i, '# Operación2'] = op_link
                            banco_ext.at[b_i, 'Anotación'] = anot
                        break

    # 4.4. Sugerencia por ITF (Suma de 'IMPUESTO ITF' en Banco = Movimiento con 'ITF' en Glosa Conta)
    mask_itf_b = (
        (~banco_ext['_matched']) &
        (~banco_ext['Anotación'].astype(str).str.startswith('Sugerido')) &
        (banco_ext['Banco - Descripción'].astype(str).str.contains(r'\bITF\b|IMPUESTO\s+ITF', case=False, na=False))
    )
    if mask_itf_b.any():
        sum_b_itf = round(float(banco_ext.loc[mask_itf_b, 'Monto-Banco'].sum()), 2)
        idxs_b = banco_ext[mask_itf_b].index.tolist()

        candidates_c = conta_ext[
            (~conta_ext['_matched']) &
            (~conta_ext.get('_suggested', pd.Series([False] * len(conta_ext)))) &
            (~conta_ext['Anotación'].astype(str).str.startswith('Sugerido')) &
            (conta_ext['Conta - Glosa'].astype(str).str.contains(r'\bitf\b', case=False, na=False))
        ]

        matched_itf = False
        for c_idx, row_c in candidates_c.iterrows():
            m_c = float(row_c.get('Monto-Conta', 0) or 0)
            if round(abs(abs(sum_b_itf) - abs(m_c)), 2) == 0.0:
                op_link = f"SUG-ITF-{c_idx}"
                anot_itf = "Sugerido: ITF (Suma Banco = Conta)"
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, 'Anotación'] = anot_itf
                conta_ext.at[c_idx, '_suggested'] = True

                for b_idx in idxs_b:
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    banco_ext.at[b_idx, 'Anotación'] = anot_itf
                matched_itf = True
                break

        if not matched_itf and len(candidates_c) > 1:
            sum_c_itf = round(float(candidates_c['Monto-Conta'].sum()), 2)
            if round(abs(abs(sum_b_itf) - abs(sum_c_itf)), 2) == 0.0:
                op_link = "SUG-ITF-MULT"
                anot_itf = "Sugerido: ITF (Suma Banco = Suma Conta)"
                for c_idx in candidates_c.index:
                    conta_ext.at[c_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, 'Anotación'] = anot_itf
                    conta_ext.at[c_idx, '_suggested'] = True
                for b_idx in idxs_b:
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    banco_ext.at[b_idx, 'Anotación'] = anot_itf
                matched_itf = True

    # 4.5. Sugerencias por Factoring (Glosa contiene 'factoring' y montos iguales sin importar la fecha)
    for b_idx, row_b in banco_ext.iterrows():
        if row_b['_matched']: continue
        if str(banco_ext.at[b_idx, 'Anotación']).startswith('Sugerido'): continue
        m_b = float(row_b.get('Monto-Banco', 0) or 0)
        if m_b == 0: continue
        
        for c_idx, row_c in conta_ext.iterrows():
            if row_c['_matched']: continue
            if row_c.get('_suggested', False): continue
            if str(conta_ext.at[c_idx, 'Anotación']).startswith('Sugerido'): continue
            
            glosa_c = str(row_c.get('Conta - Glosa', '')).lower()
            if 'factoring' in glosa_c:
                m_c = float(row_c.get('Monto-Conta', 0) or 0)
                if round(abs(m_b - m_c), 2) == 0.0:
                    op_link = f"SUG-FACT-{b_idx}-{c_idx}"
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, '# Operación2'] = op_link
                    
                    anot_fact = "Sugerido: Factoring"
                    banco_ext.at[b_idx, 'Anotación'] = anot_fact
                    conta_ext.at[c_idx, 'Anotación'] = anot_fact
                    
                    conta_ext.at[c_idx, '_suggested'] = True
                    break

    # 4.6. Sugerencias por Cambio de Moneda (Glosa 'cambio'/'moneda'/'tc' y Descripción 'COMU')
    for b_idx, row_b in banco_ext.iterrows():
        if row_b['_matched']: continue
        if str(banco_ext.at[b_idx, 'Anotación']).startswith('Sugerido'): continue
        
        desc_b = str(row_b.get('Banco - Descripción', '')).upper()
        if 'COMU' in desc_b:
            m_b = float(row_b.get('Monto-Banco', 0) or 0)
            if m_b == 0: continue
            
            for c_idx, row_c in conta_ext.iterrows():
                if row_c['_matched']: continue
                if row_c.get('_suggested', False): continue
                if str(conta_ext.at[c_idx, 'Anotación']).startswith('Sugerido'): continue
                
                glosa_c = str(row_c.get('Conta - Glosa', '')).lower()
                if 'cambio' in glosa_c or 'moneda' in glosa_c or 'tc' in glosa_c:
                    m_c = float(row_c.get('Monto-Conta', 0) or 0)
                    if round(abs(abs(m_b) - abs(m_c)), 2) == 0.0:
                        op_link = f"SUG-CAMBIO-{b_idx}-{c_idx}"
                        banco_ext.at[b_idx, '# Operación2'] = op_link
                        conta_ext.at[c_idx, '# Operación2'] = op_link
                        
                        anot_cambio = "Sugerido: Cambio de Moneda (COMUS)"
                        banco_ext.at[b_idx, 'Anotación'] = anot_cambio
                        conta_ext.at[c_idx, 'Anotación'] = anot_cambio
                        
                        conta_ext.at[c_idx, '_suggested'] = True
                        break

    # 4.7. Sugerencias por Sub-código / Referencia (ej. P07, P08) en Descripción de Banco y Glosa/Giro de Contabilidad + Monto igual
    for c_idx, row_c in conta_ext.iterrows():
        if row_c['_matched']: continue
        if row_c.get('_suggested', False): continue
        if str(conta_ext.at[c_idx, 'Anotación']).startswith('Sugerido'): continue

        glosa_c = f"{row_c.get('Conta - Glosa', '')} {row_c.get('Conta - Giro', '')}"
        m_c = float(row_c.get('Monto-Conta', 0) or 0)
        
        # Buscar patrones alfanuméricos clave de referencia (ej. P07, P08, L-12, DPTO101, etc.)
        tokens_c = set(re.findall(r'\b[A-Z]{1,4}\d{1,4}\b', glosa_c, re.IGNORECASE))
        if not tokens_c: continue

        for tok in tokens_c:
            if len(tok) < 3: continue
            tok_u = tok.upper()
            
            candidates = banco_ext[
                (~banco_ext['_matched']) &
                (~banco_ext['Anotación'].astype(str).str.startswith('Sugerido')) &
                (abs(banco_ext['Monto-Banco'].fillna(0) - m_c) <= 0.01) &
                (banco_ext['Banco - Descripción'].astype(str).str.contains(tok_u, case=False, na=False, regex=False))
            ]
            if len(candidates) == 1:
                b_idx = candidates.index[0]
                op_link = f"SUG-REF-{tok_u}"
                banco_ext.at[b_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, '# Operación2'] = op_link
                
                anot_sub = f"Sugerido: Referencia {tok_u} en Glosa"
                banco_ext.at[b_idx, 'Anotación'] = anot_sub
                conta_ext.at[c_idx, 'Anotación'] = anot_sub
                conta_ext.at[c_idx, '_suggested'] = True
                break

    # 5. Sugerencias por diferencias de montos específicos
    diferencias = diferencias_set
    
    for b_idx, row_b in banco_ext.iterrows():
        if row_b['_matched']: continue
        if str(banco_ext.at[b_idx, 'Anotación']).startswith('Sugerido'): continue
        m_b = float(row_b.get('Monto-Banco', 0) or 0)
        
        for c_idx, row_c in conta_ext.iterrows():
            if row_c['_matched']: continue
            if row_c.get('_suggested', False): continue
            if str(conta_ext.at[c_idx, 'Anotación']).startswith('Sugerido'): continue
            
            m_c = float(row_c.get('Monto-Conta', 0) or 0)
            
            diff = round(abs(abs(m_b) - abs(m_c)), 2)
            if diff in diferencias:
                op_link = f"SUG-DIF-{diff}-{b_idx}"
                banco_ext.at[b_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, '# Operación2'] = op_link
                
                banco_ext.at[b_idx, 'Anotación'] = f"Sugerido: Diferencia {diff}"
                conta_ext.at[c_idx, 'Anotación'] = f"Sugerido: Diferencia {diff}"
                
                conta_ext.at[c_idx, '_suggested'] = True
                break

    # 5.4. Sugerencia por Monto Único en todo el extracto (sin importar la fecha)
    unmatched_b_df = banco_ext[(~banco_ext['_matched']) & (~banco_ext['Anotación'].astype(str).str.startswith('Sugerido'))]
    unmatched_c_df = conta_ext[(~conta_ext['_matched']) & (~conta_ext['_suggested']) & (~conta_ext['Anotación'].astype(str).str.startswith('Sugerido'))]

    # Contar frecuencias globales de cada monto en los no conciliados
    counts_b = unmatched_b_df['Monto-Banco'].apply(lambda v: round(float(v or 0), 2)).value_counts()
    counts_c = unmatched_c_df['Monto-Conta'].apply(lambda v: round(float(v or 0), 2)).value_counts()

    for b_idx, row_b in unmatched_b_df.iterrows():
        if str(banco_ext.at[b_idx, 'Anotación']).startswith('Sugerido'): continue
        m_b = round(float(row_b.get('Monto-Banco', 0) or 0), 2)
        if m_b == 0: continue

        # Debe ser único tanto en banco como en contabilidad
        if counts_b.get(m_b, 0) == 1 and counts_c.get(m_b, 0) == 1:
            c_candidates = unmatched_c_df[
                unmatched_c_df['Monto-Conta'].apply(lambda v: round(float(v or 0), 2)) == m_b
            ]
            if len(c_candidates) == 1:
                c_idx = c_candidates.index[0]
                if not str(conta_ext.at[c_idx, 'Anotación']).startswith('Sugerido') and not conta_ext.at[c_idx, '_suggested']:
                    op_link = f"SUG-M-UNICO-{b_idx}-{c_idx}"
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, '# Operación2'] = op_link

                    anot_u = "Sugerido: Monto único (diferencia de días)"
                    banco_ext.at[b_idx, 'Anotación'] = anot_u
                    conta_ext.at[c_idx, 'Anotación'] = anot_u
                    conta_ext.at[c_idx, '_suggested'] = True

    # 5.5. Sugerencia por residuo único tras descarte (si queda exactamente 1 en Banco y 1 en Conta con mismo monto)
    unmatched_b_rem = [
        idx for idx, r in banco_ext.iterrows() 
        if not r['_matched'] and not str(banco_ext.at[idx, 'Anotación']).startswith('Sugerido')
    ]
    unmatched_c_rem = [
        idx for idx, r in conta_ext.iterrows() 
        if not r['_matched'] and not r.get('_suggested', False) and not str(conta_ext.at[idx, 'Anotación']).startswith('Sugerido')
    ]
    if len(unmatched_b_rem) == 1 and len(unmatched_c_rem) == 1:
        b_idx_rem = unmatched_b_rem[0]
        c_idx_rem = unmatched_c_rem[0]
        m_b_rem = float(banco_ext.at[b_idx_rem, 'Monto-Banco'] or 0)
        m_c_rem = float(conta_ext.at[c_idx_rem, 'Monto-Conta'] or 0)
        
        # Verificar si los montos coinciden
        if round(abs(abs(m_b_rem) - abs(m_c_rem)), 2) == 0.0:
            op_link_rem = f"SUG-RESIDUO-{b_idx_rem}-{c_idx_rem}"
            banco_ext.at[b_idx_rem, '# Operación2'] = op_link_rem
            conta_ext.at[c_idx_rem, '# Operación2'] = op_link_rem
            
            anot_rem = "Sugerido: Pareja resultante por descarte"
            banco_ext.at[b_idx_rem, 'Anotación'] = anot_rem
            conta_ext.at[c_idx_rem, 'Anotación'] = anot_rem
            conta_ext.at[c_idx_rem, '_suggested'] = True

    # ══════════════════════════════════════════════════════════════════════
    # 6. CONCILIACIÓN DE ASIENTOS CONTABLES AGRUPADOS Y VARIOS A VARIOS (N a M)
    # ══════════════════════════════════════════════════════════════════════
    # Regla:
    # 1. Si 2 o más movimientos de conta tienen el mismo asiento contable
    #    ('Conta - # Registro'), se suman sus montos y se cuentan como 1 para la conciliación.
    # 2. Se ejecuta después de todas las conciliaciones 1 a 1 y 1 a muchos.
    # 3. Concilia varios asientos de banco contra asiento(s) de conta,
    #    especialmente ITF y Comisiones/Portes/Gastos Bancarios.

    idxs_unm_b = [
        idx for idx, r in banco_ext.iterrows()
        if not r['_matched'] and not str(banco_ext.at[idx, 'Anotación']).startswith('Sugerido')
    ]
    idxs_unm_c = [
        idx for idx, r in conta_ext.iterrows()
        if not r['_matched'] and not r.get('_suggested', False) and not str(conta_ext.at[idx, 'Anotación']).startswith('Sugerido')
    ]

    # Agrupar conta por asiento contable (Conta - # Registro)
    asientos_map = {}
    for c_idx in idxs_unm_c:
        reg = _normalizar_nro_op(conta_ext.at[c_idx, 'Conta - # Registro'])
        clave = reg if (reg and reg not in ('0', 'nan')) else f"__ROW_{c_idx}__"
        asientos_map.setdefault(clave, []).append(c_idx)

    asientos_conta = []
    for clave, c_idxs in asientos_map.items():
        total_asiento = round(sum(float(conta_ext.at[i, 'Monto-Conta'] or 0) for i in c_idxs), 2)
        nro_reg = clave if not clave.startswith('__ROW_') else ''
        es_gasto = any(
            bool(re.search(r'(?i)\b(ITF|I\.T\.F\.)\b|COMIS|PORTES?|MANTENIMIENTO|GASTOS?\s+BANCARIOS?',
                           f"{conta_ext.at[i, 'Conta - Glosa']} {conta_ext.at[i, 'Conta - Giro']}"))
            for i in c_idxs
        )
        asientos_conta.append({
            'clave': clave,
            'nro_registro': nro_reg,
            'idxs': c_idxs,
            'monto': total_asiento,
            'count': len(c_idxs),
            'es_gasto': es_gasto,
        })

    # 6.1 Varios Banco = Asiento(s) Conta de Gastos Bancarios (ITF, Comisiones, Portes)
    patron_gastos_b = re.compile(
        r'(?i)\b(ITF|I\.T\.F\.)\b|'
        r'IMPUESTO\s+(?:A\s+LOS\s+(?:DEBITOS|CREDITOS)|ITF|TRANSACCIONES)|'
        r'COMIS|PORTES?|MANT\.?\s*CUENTA|MANTENIMIENTO|GASTO\s+BANCARIO|CARGO\s+POR\s+CUENTA',
        re.IGNORECASE
    )

    b_gastos_idxs = [
        b_idx for b_idx in idxs_unm_b
        if patron_gastos_b.search(str(banco_ext.at[b_idx, 'Banco - Descripción'] or ''))
    ]
    c_gastos_asientos = [
        a for a in asientos_conta
        if a['es_gasto'] and a['monto'] != 0
    ]

    if b_gastos_idxs and c_gastos_asientos:
        suma_b_gastos = round(sum(float(banco_ext.at[b, 'Monto-Banco'] or 0) for b in b_gastos_idxs), 2)
        suma_c_gastos = round(sum(a['monto'] for a in c_gastos_asientos), 2)

        if suma_b_gastos != 0 and round(abs(abs(suma_b_gastos) - abs(suma_c_gastos)), 2) <= 0.01:
            regs = [a['nro_registro'] for a in c_gastos_asientos if a['nro_registro']]
            reg_desc = f"Asiento {regs[0]}" if len(regs) == 1 else (f"{len(c_gastos_asientos)} Asientos" if len(c_gastos_asientos) > 1 else "Conta")
            op_link = f"SUG-ITF-COM-{regs[0]}" if (len(regs) == 1) else "SUG-ITF-COM-NxM"
            anot = f"Sugerido: {len(b_gastos_idxs)} Banco = {reg_desc} (ITF y Comisiones)"

            for b_idx in b_gastos_idxs:
                banco_ext.at[b_idx, '# Operación2'] = op_link
                banco_ext.at[b_idx, 'Anotación'] = anot
                if b_idx in idxs_unm_b:
                    idxs_unm_b.remove(b_idx)

            for a in c_gastos_asientos:
                for c_idx in a['idxs']:
                    conta_ext.at[c_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, 'Anotación'] = anot
                    conta_ext.at[c_idx, '_suggested'] = True
                if a in asientos_conta:
                    asientos_conta.remove(a)

    # 6.2 1 Banco = 1 Asiento Conta (cuando el asiento suma 2 o más filas de conta)
    for a in list(asientos_conta):
        if a['count'] < 2 or a['monto'] == 0:
            continue
        m_asiento = a['monto']
        cands_b = [
            b_idx for b_idx in idxs_unm_b
            if abs(float(banco_ext.at[b_idx, 'Monto-Banco'] or 0) - m_asiento) <= 0.01
        ]
        if len(cands_b) == 1:
            b_idx = cands_b[0]
            op_link = f"SUG-ASIENTO-{a['nro_registro'] or b_idx}"
            anot = f"Sugerido: 1 Banco = Asiento Conta {a['nro_registro']} (Suma de {a['count']} filas)"
            banco_ext.at[b_idx, '# Operación2'] = op_link
            banco_ext.at[b_idx, 'Anotación'] = anot
            idxs_unm_b.remove(b_idx)
            for c_idx in a['idxs']:
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, 'Anotación'] = anot
                conta_ext.at[c_idx, '_suggested'] = True
            asientos_conta.remove(a)

    # 6.3 N Banco = 1 Asiento Conta (asiento con 2+ filas o remanente)
    for a in list(asientos_conta):
        m_asiento = a['monto']
        if m_asiento == 0:
            continue
        pool_b_cand = [
            (b_idx, float(banco_ext.at[b_idx, 'Monto-Banco'] or 0))
            for b_idx in idxs_unm_b
            if (float(banco_ext.at[b_idx, 'Monto-Banco'] or 0) > 0) == (m_asiento > 0)
        ]
        if not pool_b_cand or len(pool_b_cand) < 2:
            continue
        # Guard anti-explosión combinatoria: si hay demasiados candidatos de banco,
        # el número de combinaciones sería intratable (C(30,6) ya > 590k). Omitir.
        if len(pool_b_cand) > 25:
            continue

        combos_validos = []
        for r in range(2, min(6, len(pool_b_cand) + 1)):
            for combo in combinations(pool_b_cand, r):
                if round(abs(sum(x[1] for x in combo) - m_asiento), 2) <= 0.01:
                    combos_validos.append(combo)
                    if len(combos_validos) > 1:
                        break
            if combos_validos:
                break

        if len(combos_validos) == 1:
            combo_elegido = combos_validos[0]
            b_idxs_combo = [x[0] for x in combo_elegido]
            op_link = f"SUG-SUM-ASIENTO-{a['nro_registro'] or b_idxs_combo[0]}"
            filas_txt = f" (Suma de {a['count']} filas conta)" if a['count'] > 1 else ""
            anot = f"Sugerido: {len(b_idxs_combo)} Banco = Asiento Conta {a['nro_registro']}{filas_txt}"
            for b_idx in b_idxs_combo:
                banco_ext.at[b_idx, '# Operación2'] = op_link
                banco_ext.at[b_idx, 'Anotación'] = anot
                idxs_unm_b.remove(b_idx)
            for c_idx in a['idxs']:
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, 'Anotación'] = anot
                conta_ext.at[c_idx, '_suggested'] = True
            asientos_conta.remove(a)

    # 6.4 Varios Banco = Varios Asientos Conta (N Banco = M Asientos general)
    if len(asientos_conta) >= 2 and len(idxs_unm_b) >= 2:
        for r_c in range(2, min(5, len(asientos_conta) + 1)):
            matched_nm = False
            for combo_c in combinations(asientos_conta, r_c):
                suma_c_combo = round(sum(a['monto'] for a in combo_c), 2)
                if suma_c_combo == 0:
                    continue
                pool_b_cand = [
                    (b_idx, float(banco_ext.at[b_idx, 'Monto-Banco'] or 0))
                    for b_idx in idxs_unm_b
                    if (float(banco_ext.at[b_idx, 'Monto-Banco'] or 0) > 0) == (suma_c_combo > 0)
                ]
                combos_b_validos = []
                for r_b in range(2, min(8, len(pool_b_cand) + 1)):
                    for combo_b in combinations(pool_b_cand, r_b):
                        if round(abs(sum(x[1] for x in combo_b) - suma_c_combo), 2) <= 0.01:
                            combos_b_validos.append(combo_b)
                            if len(combos_b_validos) > 1:
                                break
                    if combos_b_validos:
                        break
                if len(combos_b_validos) == 1:
                    combo_b_elegido = combos_b_validos[0]
                    b_idxs_combo = [x[0] for x in combo_b_elegido]
                    op_link = "SUG-NxM-ASIENTOS"
                    anot = f"Sugerido: {len(b_idxs_combo)} Banco = {len(combo_c)} Asientos Conta"
                    for b_idx in b_idxs_combo:
                        banco_ext.at[b_idx, '# Operación2'] = op_link
                        banco_ext.at[b_idx, 'Anotación'] = anot
                        idxs_unm_b.remove(b_idx)
                    for a in combo_c:
                        for c_idx in a['idxs']:
                            conta_ext.at[c_idx, '# Operación2'] = op_link
                            conta_ext.at[c_idx, 'Anotación'] = anot
                            conta_ext.at[c_idx, '_suggested'] = True
                        asientos_conta.remove(a)
                    matched_nm = True
                    break
            if matched_nm:
                break

    # ══════════════════════════════════════════════════════════════════════
    # 6.5 ÚLTIMA OPCIÓN: N Banco = 1 Conta / 1 Banco = N Conta sin fecha
    # ══════════════════════════════════════════════════════════════════════
    # Condición de activación: quedan menos de 10 movimientos sin conciliar
    # en al menos uno de los dos lados.
    # Solo acepta diferencia == 0.0 (monto exacto) y solución única para
    # evitar falsos positivos al ignorar la restricción de fecha.
    # Itera eliminando los ya asignados en cada vuelta hasta no haber cambios.
    _cambio_6b = True
    while _cambio_6b:
        _cambio_6b = False

        _unm_b_6b = [
            (idx, round(float(banco_ext.at[idx, 'Monto-Banco'] or 0), 2))
            for idx, r in banco_ext[~banco_ext['_matched']].iterrows()
            if not str(banco_ext.at[idx, 'Anotación']).startswith('Sugerido')
        ]
        _unm_c_6b = [
            (idx, round(float(conta_ext.at[idx, 'Monto-Conta'] or 0), 2))
            for idx, r in conta_ext[~conta_ext['_matched']].iterrows()
            if not r.get('_suggested', False)
            and not str(conta_ext.at[idx, 'Anotación']).startswith('Sugerido')
        ]

        # Activar solo si alguno de los dos lados tiene menos de 10 elementos
        if len(_unm_b_6b) >= 10 and len(_unm_c_6b) >= 10:
            break

        _used_c_6b: set = set()
        _used_b_6b: set = set()

        # N Banco = 1 Conta (un movimiento de conta es cubierto por varios del banco)
        for c_idx, c_amt in _unm_c_6b:
            if c_amt == 0 or c_idx in _used_c_6b:
                continue
            avail_b = [x for x in _unm_b_6b if x[0] not in _used_b_6b]
            if len(avail_b) < 2:
                continue
            # Guard anti-explosión combinatoria
            if len(avail_b) > 30:
                continue
            # Buscar la combinación única que sume exactamente c_amt
            _valid: list = []
            for _r in range(2, min(len(avail_b) + 1, 10)):
                for combo in combinations(avail_b, _r):
                    if round(abs(abs(sum(x[1] for x in combo)) - abs(c_amt)), 2) == 0.0:
                        _valid.append(combo)
                    if len(_valid) > 1:
                        break
                if len(_valid) > 1:
                    break
            if len(_valid) == 1:
                b_idxs_sel = [x[0] for x in _valid[0]]
                _used_b_6b.update(b_idxs_sel)
                _used_c_6b.add(c_idx)
                op_link = f"SUG-LAST-C{c_idx}"
                anot = 'Sugerido: N Banco = 1 Conta (fechas distintas)'
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, 'Anotación'] = anot
                conta_ext.at[c_idx, '_suggested'] = True
                for b_idx in b_idxs_sel:
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    banco_ext.at[b_idx, 'Anotación'] = anot
                _cambio_6b = True

        # 1 Banco = N Conta (un movimiento de banco es cubierto por varios de conta)
        for b_idx, b_amt in _unm_b_6b:
            if b_amt == 0 or b_idx in _used_b_6b:
                continue
            avail_c = [x for x in _unm_c_6b if x[0] not in _used_c_6b]
            if len(avail_c) < 2:
                continue
            # Guard anti-explosión combinatoria
            if len(avail_c) > 30:
                continue
            _valid = []
            for _r in range(2, min(len(avail_c) + 1, 10)):
                for combo in combinations(avail_c, _r):
                    if round(abs(abs(sum(x[1] for x in combo)) - abs(b_amt)), 2) == 0.0:
                        _valid.append(combo)
                    if len(_valid) > 1:
                        break
                if len(_valid) > 1:
                    break
            if len(_valid) == 1:
                c_idxs_sel = [x[0] for x in _valid[0]]
                _used_c_6b.update(c_idxs_sel)
                _used_b_6b.add(b_idx)
                op_link = f"SUG-LAST-B{b_idx}"
                anot = 'Sugerido: 1 Banco = N Conta (fechas distintas)'
                banco_ext.at[b_idx, '# Operación2'] = op_link
                banco_ext.at[b_idx, 'Anotación'] = anot
                for c_idx in c_idxs_sel:
                    conta_ext.at[c_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, 'Anotación'] = anot
                    conta_ext.at[c_idx, '_suggested'] = True
                _cambio_6b = True

    if '_suggested' in conta_ext.columns:
        conta_ext = conta_ext.drop(columns=['_suggested'])

    # Marcar los no conciliados con una anotación por defecto si está vacía
    for idx, row in banco_ext[~banco_ext['_matched']].iterrows():
        if not str(row.get('Anotación', '')).strip():
            banco_ext.at[idx, 'Anotación'] = 'Mov. solo en banco'
            
    for idx, row in conta_ext[~conta_ext['_matched']].iterrows():
        if not str(row.get('Anotación', '')).strip():
            conta_ext.at[idx, 'Anotación'] = 'Mov. solo en conta'

    return banco_ext.drop(columns=['_matched']), conta_ext.drop(columns=['_matched'])


def _sugerencias_meses_anteriores(
    banco_ext: pd.DataFrame,
    conta_ext: pd.DataFrame,
    moneda: str = "Dolares (USD)",
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Segunda pasada de sugerencias para los movimientos que provienen de meses
    anteriores (flag '_pendiente' == True) y que NO fueron conciliados por el
    algoritmo principal.

    Regla: solo se compara por monto (sin importar la fecha), porque el
    movimiento pendiente tiene fecha antigua y no va a coincidir en fecha con
    los del mes actual.

    Casos:
      1. abs(monto_pend) == abs(monto_mes_actual)  →  "Sugerido: Solo monto meses anteriores"
      2. abs(diferencia) en diferencias_comision    →  "Sugerido: Solo monto meses anteriores
                                                         (comisión <DIF>)"
      3. Sin coincidencia                           →  queda como partida abierta
                                                        (Mov. solo en banco / Mov. solo en conta)
    Solo se considera un pendiente si aún NO tiene sugerencia ni está conciliado
    (MAR != 'X' y Anotación no empieza con 'Sugerido').
    """
    diferencias_set = _diferencias_comision(moneda)

    if '_suggested' not in conta_ext.columns:
        conta_ext['_suggested'] = False

    # ── Pendientes de BANCO vs movimientos del mes actual en CONTA ────
    pend_b_mask = (
        banco_ext.get('_pendiente', pd.Series([False] * len(banco_ext))).fillna(False).astype(bool) &
        (banco_ext['MAR'] != 'X') &
        (~banco_ext['Anotación'].astype(str).str.startswith('Sugerido'))
    )
    for b_idx in banco_ext[pend_b_mask].index:
        m_b = float(banco_ext.at[b_idx, 'Monto-Banco'] or 0)
        if m_b == 0:
            continue

        # Candidatos: movimientos del MES ACTUAL (no pendientes), no conciliados
        candidatos_c = conta_ext[
            (~conta_ext.get('_pendiente', pd.Series([False] * len(conta_ext))).fillna(False).astype(bool)) &
            (conta_ext['MAR'] != 'X') &
            (~conta_ext['_suggested']) &
            (~conta_ext['Anotación'].astype(str).str.startswith('Sugerido'))
        ]

        mejor_idx = None
        mejor_dif = None
        for c_idx, row_c in candidatos_c.iterrows():
            m_c = float(row_c.get('Monto-Conta', 0) or 0)
            dif = round(abs(abs(m_b) - abs(m_c)), 2)
            if dif == 0.0:
                mejor_idx = c_idx
                mejor_dif = 0.0
                break  # coincidencia exacta, no buscar más
            if dif in diferencias_set and mejor_idx is None:
                mejor_idx = c_idx
                mejor_dif = dif

        if mejor_idx is not None:
            op_link = f"SUG-PEND-B{b_idx}-C{mejor_idx}"
            if mejor_dif == 0.0:
                anot = 'Sugerido: Solo monto meses anteriores'
            else:
                anot = f'Sugerido: Solo monto meses anteriores (comisión {mejor_dif:.2f})'
            banco_ext.at[b_idx, '# Operación2'] = op_link
            conta_ext.at[mejor_idx, '# Operación2'] = op_link
            banco_ext.at[b_idx, 'Anotación'] = anot
            conta_ext.at[mejor_idx, 'Anotación'] = anot
            conta_ext.at[mejor_idx, '_suggested'] = True

    # ── Pendientes de CONTA vs movimientos del mes actual en BANCO ────
    pend_c_mask = (
        conta_ext.get('_pendiente', pd.Series([False] * len(conta_ext))).fillna(False).astype(bool) &
        (conta_ext['MAR'] != 'X') &
        (~conta_ext['_suggested']) &
        (~conta_ext['Anotación'].astype(str).str.startswith('Sugerido'))
    )
    for c_idx in conta_ext[pend_c_mask].index:
        m_c = float(conta_ext.at[c_idx, 'Monto-Conta'] or 0)
        if m_c == 0:
            continue

        # Candidatos: movimientos del MES ACTUAL (no pendientes), no conciliados
        candidatos_b = banco_ext[
            (~banco_ext.get('_pendiente', pd.Series([False] * len(banco_ext))).fillna(False).astype(bool)) &
            (banco_ext['MAR'] != 'X') &
            (~banco_ext['Anotación'].astype(str).str.startswith('Sugerido'))
        ]

        mejor_idx = None
        mejor_dif = None
        for b_idx, row_b in candidatos_b.iterrows():
            m_b = float(row_b.get('Monto-Banco', 0) or 0)
            dif = round(abs(abs(m_b) - abs(m_c)), 2)
            if dif == 0.0:
                mejor_idx = b_idx
                mejor_dif = 0.0
                break
            if dif in diferencias_set and mejor_idx is None:
                mejor_idx = b_idx
                mejor_dif = dif

        if mejor_idx is not None:
            op_link = f"SUG-PEND-C{c_idx}-B{mejor_idx}"
            if mejor_dif == 0.0:
                anot = 'Sugerido: Solo monto meses anteriores'
            else:
                anot = f'Sugerido: Solo monto meses anteriores (comisión {mejor_dif:.2f})'
            conta_ext.at[c_idx, '# Operación2'] = op_link
            banco_ext.at[mejor_idx, '# Operación2'] = op_link
            conta_ext.at[c_idx, 'Anotación'] = anot
            banco_ext.at[mejor_idx, 'Anotación'] = anot
            conta_ext.at[c_idx, '_suggested'] = True

    return banco_ext, conta_ext


def _preparar_anexar1(tab_banco: pd.DataFrame, tab_conta: pd.DataFrame, moneda: str = "Dolares (USD)", banco: str = "") -> pd.DataFrame:
    """
    Une Tab_Banco y Tab_Contanet verticalmente con todas las columnas combinadas.
    Las columnas que no existen en una tabla quedan en NaN (el especialista las llena).
    """
    # Columnas del resultado final
    # Orden: datos banco | Monto-Banco | Monto-Conta | datos conta  (montos al centro para fácil comparación)
    cols = [
        'Fecha', 'Banco - Fecha', 'Conta - Fecha', 'MAR', 'TIPO', 'CODIGO', '# Operación2', 'DIF COMISON', 'Anotación',
        'Banco - Descripción', 'Banco - # Operación', 'Monto-Banco',
        'Monto-Conta', 'Conta - # Registro', 'Conta - # Operación', 'Conta - Giro', 'Conta - Glosa'
    ]

    banco_ext = tab_banco.copy()
    conta_ext = tab_conta.copy()
    # Elimina filas donde cualquier celda es igual al header o a la frase 'Información anterior' o cabeceras como '# Registro'
    def _clean_df(df):
        # Elimina filas donde cualquier celda es igual al header o a la frase 'Información anterior' o cabeceras como '# Registro'
        words_to_drop = {'INFORMACIÓN ANTERIOR', '# REGISTRO', 'REGISTRO', '# OPERACIÓN', 'GIRO', 'GLOSA', 'FECHA', 'MEDIO PAGO'}
        mask = df.apply(
            lambda row: any(
                str(v).strip() == col or 
                str(v).strip().upper() in words_to_drop or 
                'INFORMACIÓN ANTERIOR' in str(v).strip().upper()
                for col, v in zip(df.columns, row)
            ),
            axis=1
        )
        return df[~mask]
    banco_ext = _clean_df(banco_ext)
    conta_ext = _clean_df(conta_ext)
    
    if 'Anotación' not in banco_ext.columns: banco_ext['Anotación'] = ''
    if 'Anotación' not in conta_ext.columns: conta_ext['Anotación'] = ''
    
    # Asegurar que las columnas acepten strings para evitar errores de tipo al asignar
    for col in ['MAR', '# Operación2', 'Anotación']:
        if col in banco_ext.columns:
            banco_ext[col] = banco_ext[col].astype(object)
        if col in conta_ext.columns:
            conta_ext[col] = conta_ext[col].astype(object)
    
    # ── APLICAR ALGORITMO DE CONCILIACIÓN AUTOMÁTICA ──
    banco_ext, conta_ext = _conciliacion_automatica(banco_ext, conta_ext, moneda=moneda, banco=banco)

    # ── SUGERENCIAS POR MESES ANTERIORES (pendientes de conciliaciones previas) ──
    # Se ejecuta DESPUÉS de la conciliación normal para que los movimientos del
    # mes actual que ya se conciliaron no sean reutilizados por esta segunda pasada.
    banco_ext, conta_ext = _sugerencias_meses_anteriores(banco_ext, conta_ext, moneda=moneda)

    # Limpiar columna auxiliar _suggested si quedó de _sugerencias_meses_anteriores
    if '_suggested' in conta_ext.columns:
        conta_ext = conta_ext.drop(columns=['_suggested'])

    # Asegurar que todas las columnas existan en ambos para evitar warnings
    for c in cols:
        if c not in banco_ext.columns:
            banco_ext[c] = None
        if c not in conta_ext.columns:
            conta_ext[c] = None

    # Separar conciliados y no conciliados
    b_match = banco_ext[banco_ext['MAR'] == 'X'].copy()
    c_match = conta_ext[conta_ext['MAR'] == 'X'].copy()
    b_unmatch = banco_ext[banco_ext['MAR'] != 'X'].copy()
    c_unmatch = conta_ext[conta_ext['MAR'] != 'X'].copy()

    # Fusionar filas conciliadas lado a lado
    matched_rows = []
    grupos = pd.unique(pd.concat([b_match['# Operación2'], c_match['# Operación2']]))
    for g in grupos:
        if str(g) == 'nan' or not str(g).strip(): 
            continue
        bg = b_match[b_match['# Operación2'] == g].reset_index(drop=True)
        cg = c_match[c_match['# Operación2'] == g].reset_index(drop=True)
        
        n = max(len(bg), len(cg))
        for i in range(n):
            row = {}
            for col in cols:
                val_b = bg.at[i, col] if i < len(bg) else None
                val_c = cg.at[i, col] if i < len(cg) else None
                
                def is_valid(v):
                    return pd.notna(v) and str(v) != 'nan' and str(v) != ''
                
                if is_valid(val_b):
                    row[col] = val_b
                elif is_valid(val_c):
                    row[col] = val_c
                else:
                    row[col] = None
                    
            # Si Conta - # Operación está vacío o '0', copiar desde Banco - # Operación
            op_conta = row.get('Conta - # Operación')
            op_banco = row.get('Banco - # Operación')
            
            def is_valid_op(v):
                return pd.notna(v) and str(v) != 'nan' and str(v).strip() != '' and str(v).strip() != '0'
                
            if not is_valid_op(op_conta) and is_valid_op(op_banco):
                row['Conta - # Operación'] = op_banco

            # Si ambos montos están presentes y hay diferencia por comisión, reflejar en DIF COMISON
            mb_val = row.get('Monto-Banco')
            mc_val = row.get('Monto-Conta')
            if pd.notna(mb_val) and pd.notna(mc_val):
                try:
                    dif_calc = round(float(mb_val or 0) - float(mc_val or 0), 2)
                    if round(abs(dif_calc), 2) in _diferencias_comision(moneda):
                        row['DIF COMISON'] = dif_calc
                except (ValueError, TypeError):
                    pass

            matched_rows.append(row)

    df_matched = pd.DataFrame(matched_rows, columns=cols) if matched_rows else pd.DataFrame(columns=cols)

    # Mantener todas las filas no conciliadas (incluyendo sugerencias) separadas
    b_unmatch_filtered = b_unmatch[cols]
    c_unmatch_filtered = c_unmatch[cols]

    # Concatenar todo
    anexar = pd.concat(
        [df_matched, b_unmatch_filtered, c_unmatch_filtered],
        ignore_index=True
    )

    # Columnas para decisión del especialista
    # Filas conciliadas (MAR=='X') → 'Conciliados'; sugeridos → 'X'; resto → ''
    def _valor_conciliar(row):
        if row.get('MAR') == 'X':
            return 'Conciliados'
        if pd.notna(row.get('Anotación')) and 'Sugerido:' in str(row.get('Anotación', '')):
            return 'X'
        return ''
    anexar['Conciliar'] = anexar.apply(_valor_conciliar, axis=1)
    anexar['# Operación a Conciliar'] = ''
    anexar['Anotación-Conta'] = ''

    # ── Calcular DIF COMISON real en pares sugeridos y de Observación ──
    # Para pares 1-a-N: DIF = monto_banco - suma(montos_conta), en la fila banco.
    # Para pares N-a-1: DIF = suma(montos_banco) - monto_conta, en la fila conta.
    # En los demás casos (1-a-1): DIF = monto_banco - monto_conta, en la fila banco.
    # Las filas del lado "múltiple" siempre quedan con DIF COMISON = 0.
    # Se incluyen también los pares de Observación (OBS-COD-DIF-*) para que
    # la diferencia quede registrada en la fila banco y las columnas de
    # clasificación se calculen correctamente después.
    mask_con_dif = (
        anexar['Anotación'].astype(str).str.startswith('Sugerido:') |
        anexar['Anotación'].astype(str).str.startswith('Observación:')
    )
    if mask_con_dif.any():
        grupos_dif = anexar.loc[mask_con_dif, '# Operación2'].dropna().unique()
        for op2 in grupos_dif:
            op2_str = str(op2).strip()
            if not op2_str or op2_str == 'nan':
                continue
            idxs = anexar.index[
                mask_con_dif & (anexar['# Operación2'].astype(str) == op2_str)
            ].tolist()
            idx_banco = [i for i in idxs if pd.notna(anexar.at[i, 'Monto-Banco'])]
            idx_conta = [i for i in idxs if pd.notna(anexar.at[i, 'Monto-Conta'])]

            if not idx_banco or not idx_conta:
                continue

            suma_b = sum(float(anexar.at[i, 'Monto-Banco'] or 0) for i in idx_banco)
            suma_c = sum(float(anexar.at[i, 'Monto-Conta'] or 0) for i in idx_conta)
            dif = round(suma_b - suma_c, 2)

            if len(idx_banco) == 1 and len(idx_conta) >= 1:
                # 1 Banco = N Conta → diferencia en la fila banco
                anexar.at[idx_banco[0], 'DIF COMISON'] = dif
                for i in idx_conta:
                    anexar.at[i, 'DIF COMISON'] = 0
            elif len(idx_banco) >= 1 and len(idx_conta) == 1:
                # N Banco = 1 Conta → diferencia en la fila conta
                for i in idx_banco:
                    anexar.at[i, 'DIF COMISON'] = 0
                anexar.at[idx_conta[0], 'DIF COMISON'] = dif
            else:
                # N-a-N: no aplica, limpiar todo
                for i in idx_banco + idx_conta:
                    anexar.at[i, 'DIF COMISON'] = 0

    # ── 4 columnas de clasificación de diferencias ────────────────────
    # Se calculan DESPUÉS de DIF COMISON para que Comisiones refleje el valor real.
    # ITF        → descripción banco contiene 'ITF'
    # Comisiones → DIF COMISON != 0 Y la diferencia es una comisión conocida
    # Error      → par de Observación (diferencia que NO es comisión) solo en fila banco
    # Otros      → vacío (el especialista marca si aplica)
    _difs_comision = _diferencias_comision(moneda)

    def _marcar_itf(row):
        desc = str(row.get('Banco - Descripción', '') or '').upper().strip()
        if not desc:
            return ''
        if 'ITF' in desc or 'IMPUESTO A LOS CREDITOS' in desc or 'IMPUESTO A LOS DEBITOS' in desc:
            return 'X'
        return ''

    def _es_comision_scotiabank_desc(desc: str) -> bool:
        if not desc:
            return False
        d_upper = desc.upper().strip()
        # Descripciones exactas o contenidas
        conceptos_fijos = [
            'TBK-MANTENIMIENTO',
            'PORTES ESTADO DE CUENTA',
            'MANT TBK CORPO EMP RELAC',
            'COMIS.TRF.CTAS 3ROS BCR',
        ]
        if any(c in d_upper for c in conceptos_fijos):
            return True
        # Comiencen con la palabra Comis y Mant. (soporta COMIS, COMIS., COMISION, MANT, MANT., MANTENIMIENTO)
        if re.match(r'^(?:COMIS|MANT)\b|\bCOMIS\b|\bMANT\b', d_upper):
            return True
        return False

    def _es_comision_bcp_desc(desc: str) -> bool:
        """Detecta comisiones/mantenimiento en el campo Detalle del extracto BCP.

        Patrones reconocidos (insensible a mayúsculas/tildes):
        - COM, COM., COMI, COMIS, COMISION, COMISIÓN, y cualquier variante que
          empiece por COM seguida de letras típicas de 'comisión'.
        - MANT, MANT., MANTENIMIENTO, MANTENIMIENTO. y similares.
        """
        if not desc:
            return False
        d_upper = desc.upper().strip()
        # Eliminar tildes para comparación robusta
        d_norm = ''.join(
            c for c in unicodedata.normalize('NFD', d_upper)
            if unicodedata.category(c) != 'Mn'
        )
        # Patrón: palabra que empiece con COM (comisión, comi, com., com, etc.)
        # o con MANT (mantenimiento, mant., mant, etc.)
        # Se usa \b o fin de palabra/puntuación para evitar falsos positivos
        if re.search(r'\bCOM(?:IS(?:ION)?|I|\.?)?\b', d_norm):
            return True
        if re.search(r'\bMANT(?:ENIMIENTO)?\b', d_norm):
            return True
        return False

    # Resolver clave del banco una sola vez (usada en _marcar_comisiones)
    try:
        from bancos import clave_banco as _clave_banco
        _banco_clave = _clave_banco(banco) if banco else ''
    except Exception:
        _banco_clave = ''

    def _marcar_comisiones(row):
        # Si la fila ya es ITF, no marcar también como comisión
        if _marcar_itf(row) == 'X':
            return ''
        desc = str(row.get('Banco - Descripción', '') or '')
        # Regla explícita por descripción de banco, diferenciada por banco
        if _banco_clave == 'BCP':
            if _es_comision_bcp_desc(desc):
                return 'X'
        else:
            if _es_comision_scotiabank_desc(desc):
                return 'X'

        dif = row.get('DIF COMISON')
        try:
            dif_f = float(dif) if dif is not None else 0.0
        except (ValueError, TypeError):
            return ''
        if dif_f == 0.0:
            return ''
        # Si la anotación es por sufijo de operación y hay diferencia, marcar 'X' en Comisiones
        anot = str(row.get('Anotación', '') or '').lower()
        if 'sufijo de # operación' in anot:
            return 'X'
        # Solo marcar como comisión si el valor absoluto corresponde a una comisión conocida
        return 'X' if round(abs(dif_f), 2) in _difs_comision else ''

    def _marcar_error(row):
        anot = str(row.get('Anotación', '') or '')
        # Si es por sufijo de operación, se clasifica en Comisiones a petición del usuario
        if 'sufijo de # operación' in anot.lower():
            return ''
        # Solo aplica a pares de Observación (diferencia de monto con mismo código)
        if 'diferencia de monto' not in anot.lower():
            return ''
        # Solo en la fila banco (tiene Monto-Banco), no en la fila conta
        if not pd.notna(row.get('Monto-Banco')):
            return ''
        # Si la diferencia ya está capturada como comisión conocida, no es error
        dif = row.get('DIF COMISON')
        try:
            dif_f = float(dif) if dif is not None else 0.0
        except (ValueError, TypeError):
            dif_f = 0.0
        if round(abs(dif_f), 2) in _difs_comision:
            return ''
        return 'X'

    anexar['ITF']        = anexar.apply(_marcar_itf, axis=1)
    anexar['Comisiones'] = anexar.apply(_marcar_comisiones, axis=1)
    anexar['Error']      = anexar.apply(_marcar_error, axis=1)
    anexar['Otros']      = ''

    # Crear una clave de ordenamiento:
    # 0 = Conciliados (MAR == 'X')
    # 1 = Faltan conciliar pero tienen sugerencia ('Sugerido:' en Anotación)
    # 2 = Mov. solo en banco / Mov. solo en conta
    # 3 = Comisiones no sugeridas
    # 4 = ITF no sugerido
    # 5 = Sin monto (ingreso=0 y egreso=0)
    def _orden_grupo(row):
        anotacion = str(row.get('Anotación', ''))
        # Sin monto: siempre al final de todo (aplica a banco y a conta)
        if 'sin monto (0)' in anotacion.lower():
            return 5
        if row.get('MAR') == 'X':
            return 0
        # Sugeridos (faltan conciliar pero tienen sugerencia, incluyendo ITF sugerido)
        if 'Sugerido:' in anotacion:
            return 1
        # Comisiones no sugeridas
        if row.get('Comisiones') == 'X':
            return 3
        # ITF no sugerido
        if row.get('ITF') == 'X':
            return 4
        # Solo en banco o solo en conta
        return 2

    anexar['_orden_conciliado'] = anexar.apply(_orden_grupo, axis=1)

    # Para sugeridos: mantener pares juntos usando # Operación2 como sub-clave
    def get_op2_sort(row, ord_c):
        if ord_c == 1:
            return str(row.get('# Operación2', ''))
        return ''

    anexar['_op2_sort'] = anexar.apply(lambda r: get_op2_sort(r, r['_orden_conciliado']), axis=1)

    # Para grupo 2 (solo en banco y solo en conta), calcular monto para ordenar de mayor a menor
    def get_monto_abs(row):
        mb = row.get('Monto-Banco')
        mc = row.get('Monto-Conta')
        vb = abs(float(mb)) if pd.notna(mb) and str(mb) != 'nan' else 0.0
        vc = abs(float(mc)) if pd.notna(mc) and str(mc) != 'nan' else 0.0
        return max(vb, vc)

    anexar['_monto_abs'] = anexar.apply(get_monto_abs, axis=1)

    # Para grupo 2, queremos de mayor a menor (-_monto_abs)
    anexar['_monto_sort'] = anexar.apply(
        lambda r: -r['_monto_abs'] if r['_orden_conciliado'] == 2 else 0, axis=1
    )

    anexar = anexar.sort_values(
        by=['_orden_conciliado', '_op2_sort', '_monto_sort', 'Fecha']
    ).drop(columns=['_orden_conciliado', '_op2_sort', '_monto_abs', '_monto_sort']).reset_index(drop=True)

    return anexar


def _preparar_resumen(anexar1: pd.DataFrame) -> pd.DataFrame:
    """
    Tabla resumen por TIPO y CODIGO con totales de Monto-Conta, Monto-Banco
    y Diferencia (para que el especialista vea las partidas abiertas).
    """
    # Agrupar por TIPO y CODIGO
    resumen = (
        anexar1
        .groupby(['TIPO', 'CODIGO'], dropna=False)
        .agg(
            **{'Monto-Conta' : ('Monto-Conta',  'sum'),
               'Monto-Banco' : ('Monto-Banco',  'sum')}
        )
        .reset_index()
    )
    print(resumen)
    resumen['Diferencia'] = resumen['Monto-Banco'].fillna(0) - resumen['Monto-Conta'].fillna(0)

    # Ordenar: primero EGRESO, luego INGRESO; dentro de cada tipo por CODIGO
    resumen = resumen.sort_values(['TIPO', 'CODIGO']).reset_index(drop=True)
    return resumen


# ══════════════════════════════════════════════════════════════════════
# FORMATO VISUAL
# ══════════════════════════════════════════════════════════════════════

# Colores de cabecera por hoja
_COLORES = {
    'BANCO'        : '1F4E79',   # azul oscuro
    'CONTANET'     : '375623',   # verde oscuro
    'Anexar1'      : '7030A0',   # morado
    'Resumen'      : 'C55A11',   # naranja
}


def _aplicar_formato(ruta: Path) -> None:
    """Aplica formato básico: cabecera con color, autoajuste de columnas, congelar fila."""
    wb = openpyxl.load_workbook(ruta)

    thin = Side(style='thin', color='D9D9D9')
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Objetos de estilo compartidos (se crean UNA vez, no por celda) ───────
    _font_header     = Font(bold=True, color='FFFFFF', size=10)
    _align_header    = Alignment(horizontal='center', vertical='center', wrap_text=True)
    _align_center    = Alignment(horizontal='center', vertical='center')
    _align_vcenter   = Alignment(vertical='center')
    _num_fmt_date    = 'DD/MM/YYYY'
    _num_fmt_amount  = '#,##0.00'

    # Caché de objetos PatternFill y Font por color/estilo (evita recrear los mismos)
    _fill_cache: dict[str, PatternFill] = {}
    _font_cache: dict[tuple, Font]      = {}

    def _get_fill(color_hex: str) -> PatternFill:
        if color_hex not in _fill_cache:
            _fill_cache[color_hex] = PatternFill('solid', fgColor=color_hex)
        return _fill_cache[color_hex]

    def _get_font(color: str = '000000', bold: bool = False, italic: bool = False, size: int = 11) -> Font:
        key = (color, bold, italic, size)
        if key not in _font_cache:
            _font_cache[key] = Font(color=color, bold=bold, italic=italic, size=size)
        return _font_cache[key]

    _fill_none = _get_fill('FFFFFF')  # blanco (sin relleno efectivo pero evita None)

    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        color_hex = _COLORES.get(nombre_hoja, '404040')

        # Colores específicos para las 4 columnas de clasificación de diferencias
        _COLORES_CLASIF = {
            'ITF'        : 'C00000',  # Rojo oscuro
            'Comisiones' : 'ED7D31',  # Naranja
            'Error'      : 'BF8F00',  # Amarillo ocre
            'Otros'      : '70AD47',  # Verde claro
        }

        # ── Pre-calcular fills de cabecera por columna (solo para Anexar1) ──
        header_fill_map: dict[int, str] = {}  # col_idx → color_hex
        if nombre_hoja == 'Anexar1':
            for i, cell in enumerate(ws[1]):
                col_name = str(cell.value or '')
                if col_name in _COLORES_CLASIF:
                    header_fill_map[i] = _COLORES_CLASIF[col_name]
                elif 'Banco' in col_name:
                    header_fill_map[i] = '1F4E79'
                elif 'Conta' in col_name:
                    header_fill_map[i] = '375623'
                else:
                    header_fill_map[i] = '7030A0'
        else:
            for i in range(ws.max_column):
                header_fill_map[i] = color_hex

        # Formato de cabecera (fila 1)
        for i, cell in enumerate(ws[1]):
            fg = header_fill_map.get(i, color_hex)
            cell.font      = _get_font('FFFFFF', bold=True, size=10)
            cell.fill      = _get_fill(fg)
            cell.alignment = _align_header
            cell.border    = borde

        # Altura de cabecera
        ws.row_dimensions[1].height = 30

        # Encontrar índice de columnas relevantes
        col_mar_idx = None
        col_anot_idx = None
        col_conc_idx = None
        col_no_conc_idx = None
        col_op2_idx = None
        col_banco_desc_idx = None
        # ── Leer nombres de cabecera UNA vez y construir mapa ──────────────
        header_names: list[str] = []
        for cell in ws[1]:
            val = str(cell.value or '').strip()
            header_names.append(val)
            if val == 'MAR':
                col_mar_idx = len(header_names) - 1
            elif val == 'Anotación':
                col_anot_idx = len(header_names) - 1
            elif val == 'Conciliar':
                col_conc_idx = len(header_names) - 1
            elif val in ('# Operación a Conciliar', 'No Conciliar'):
                col_no_conc_idx = len(header_names) - 1
            elif val == '# Operación2':
                col_op2_idx = len(header_names) - 1
            elif val == 'Banco - Descripción':
                col_banco_desc_idx = len(header_names) - 1

        # Pre-calcular para Anexar1: asignar color de fondo alternado por grupo sugerido
        COLORES_SUG = ['FFF2CC', 'FCE4D6']
        sug_group_colors: dict[str, str] = {}
        if nombre_hoja == 'Anexar1' and col_op2_idx is not None and col_anot_idx is not None:
            color_idx = 0
            for row in ws.iter_rows(min_row=2):
                op2_val = row[col_op2_idx].value
                anot_val = str(row[col_anot_idx].value or '')
                if op2_val and 'Sugerido:' in anot_val:
                    key = str(op2_val)
                    if key not in sug_group_colors:
                        sug_group_colors[key] = COLORES_SUG[color_idx % len(COLORES_SUG)]
                        color_idx += 1

        # ── Pre-construir fills para Banco/Conta columnas ──────────────────
        _fill_banco   = _get_fill('D9E1F2')
        _fill_conta   = _get_fill('E2EFDA')
        _fill_sinmonto = _get_fill('F4CCFF')
        _font_sinmonto = _get_font('7F7F7F', italic=True)
        _font_itf      = _get_font('7F7F7F', italic=True)
        _font_rojo     = _get_font('FF0000')
        _font_sugerido = _get_font('7F3F00', bold=True)
        _font_monto    = _get_font('0033CC', bold=True, size=11)

        # Formato de datos fila por fila
        for row in ws.iter_rows(min_row=2):
            es_rojo = False
            es_sugerido = False
            es_itf = False
            es_sin_monto = False
            color_sug_grupo: str | None = None

            if nombre_hoja == 'Anexar1' and col_mar_idx is not None:
                val_mar = row[col_mar_idx].value
                anot_val = str(row[col_anot_idx].value or '') if col_anot_idx is not None else ''

                if 'sin monto (0)' in anot_val.lower():
                    es_sin_monto = True
                elif 'Sugerido:' in anot_val:
                    es_sugerido = True
                    if col_op2_idx is not None:
                        op2_key = str(row[col_op2_idx].value or '')
                        color_sug_grupo = sug_group_colors.get(op2_key)
                elif col_banco_desc_idx is not None and 'ITF' in str(row[col_banco_desc_idx].value or '').upper():
                    es_itf = True
                elif val_mar != 'X':
                    es_rojo = True

            for idx, cell in enumerate(row):
                cell.alignment = _align_vcenter
                cell.border    = borde

                header_name = header_names[idx] if idx < len(header_names) else ''

                # Columnas de clasificación: centrar y colorear la X
                if nombre_hoja == 'Anexar1' and header_name in _COLORES_CLASIF:
                    cell.alignment = _align_center
                    if cell.value == 'X':
                        cell.font = _get_font(_COLORES_CLASIF[header_name], bold=True, size=11)
                    continue

                if es_sin_monto:
                    cell.fill = _fill_sinmonto
                    cell.font = _font_sinmonto
                elif es_itf:
                    cell.font = _font_itf
                    if nombre_hoja == 'Anexar1':
                        if 'Banco' in header_name:
                            cell.fill = _fill_banco
                        elif 'Conta' in header_name:
                            cell.fill = _fill_conta
                elif es_sugerido and color_sug_grupo:
                    cell.fill = _get_fill(color_sug_grupo)
                    cell.font = _font_sugerido
                elif es_rojo:
                    cell.font = _font_rojo
                    if nombre_hoja == 'Anexar1':
                        if 'Banco' in header_name:
                            cell.fill = _fill_banco
                        elif 'Conta' in header_name:
                            cell.fill = _fill_conta
                else:
                    if nombre_hoja == 'Anexar1':
                        if 'Banco' in header_name:
                            cell.fill = _fill_banco
                        elif 'Conta' in header_name:
                            cell.fill = _fill_conta

                if isinstance(cell.value, datetime):
                    cell.number_format = _num_fmt_date

                if header_name in ('Monto-Banco', 'Monto-Conta') and nombre_hoja == 'Anexar1':
                    cell.font = _font_monto
                    cell.number_format = _num_fmt_amount

        # Combinar ÚNICAMENTE celdas de la columna 'Conciliar' para grupos sugeridos
        if (nombre_hoja == 'Anexar1' and col_op2_idx is not None
                and col_conc_idx is not None and col_anot_idx is not None):
            group_rows: dict[str, list[int]] = {}
            for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
                op2_val = row[col_op2_idx].value
                anot_val = row[col_anot_idx].value
                if op2_val and isinstance(anot_val, str) and 'Sugerido:' in anot_val:
                    key = str(op2_val)
                    group_rows.setdefault(key, []).append(row_idx)
            for rows in group_rows.values():
                if len(rows) > 1:
                    ws.merge_cells(
                        start_row=rows[0], start_column=col_conc_idx + 1,
                        end_row=rows[-1],  end_column=col_conc_idx + 1
                    )

        # Autoajuste de columnas — muestrear hasta 200 filas para velocidad
        _MAX_SAMPLE = 200
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            cells = list(col)
            # Tomar cabecera + muestra del resto
            sample = [cells[0]] + cells[1:_MAX_SAMPLE + 1]
            for cell in sample:
                try:
                    largo = len(str(cell.value)) if cell.value is not None else 0
                    if largo > max_len:
                        max_len = largo
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

        # Congelar primera fila
        ws.freeze_panes = 'A2'

    wb.save(ruta)


def _ocultar_columnas_tecnicas(ruta: Path) -> None:
    """
    Oculta en la hoja Anexar1 las columnas técnicas que no necesita ver el especialista.
    Se ocultan: Banco - Fecha, Conta - Fecha, TIPO, CODIGO, # Operación2
    DIF COMISON se mantiene visible para que el especialista pueda ver la diferencia.
    """
    COLS_OCULTAR = {
        'Banco - Fecha', 'Conta - Fecha', 'TIPO', 'CODIGO',
        '# Operación2'
    }

    wb = openpyxl.load_workbook(ruta)
    if 'Anexar1' not in wb.sheetnames:
        wb.save(ruta)
        return

    ws = wb['Anexar1']
    for cell in ws[1]:
        if str(cell.value).strip() in COLS_OCULTAR:
            col_letter = get_column_letter(cell.column)
            ws.column_dimensions[col_letter].hidden = True

    wb.save(ruta)
