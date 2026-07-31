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
) -> Path:
    """
    Genera el Excel de conciliacion inicial y su PDF de resumen.
    Devuelve la ruta del archivo Excel creado.
    Si no se pasa ruta_salida, crea los archivos en el directorio actual con timestamp.
    """
    if ruta_salida is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_salida = Path(f"Conciliacion_Inicial_{ts}.xlsx")

    # ── Preparar Tab_Banco ────────────────────────────────────────────
    tab_banco = _preparar_tab_banco(bank)

    # ── Preparar Tab_Contanet ─────────────────────────────────────────
    tab_conta = _preparar_tab_contanet(conta)

    # ── Preparar Anexar1 (union vertical) ─────────────────────────────
    anexar1 = _preparar_anexar1(tab_banco, tab_conta, moneda=moneda)

    # ── Preparar Resumen ──────────────────────────────────────────────
    resumen = _preparar_resumen(anexar1)

    # ── Escribir Excel ────────────────────────────────────────────────
    with pd.ExcelWriter(ruta_salida, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
        _exportar_banco(bank).to_excel(writer,   sheet_name='BANCO',        index=False)
        _exportar_contanet(conta).to_excel(writer, sheet_name='CONTANET',   index=False)
        anexar1.to_excel(writer,                 sheet_name='Anexar1',      index=False)
        resumen.to_excel(writer,                 sheet_name='Resumen',      index=False)

    # ── Aplicar formato visual Excel ──────────────────────────────────
    _aplicar_formato(ruta_salida)
    print(f"\n  [OK] Reporte inicial generado: {ruta_salida}")

    # ── Generar PDF de resumen (mismo nombre, extensión .pdf) ─────────
    # Generación de PDF deshabilitada.
    # from pdf_resumen import generar_pdf_resumen
    # ruta_pdf = ruta_salida.with_suffix('.pdf')
    # generar_pdf_resumen(
    #     bank=bank,
    #     conta=conta,
    #     anexar1=anexar1,
    #     resumen_df=resumen,
    #     ruta_pdf=ruta_pdf,
    #     empresa=empresa,
    #     moneda=moneda,
    #     ruta_excel=ruta_salida,
    # )

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
    # Quitar columnas auxiliares de cálculo
    df = df.drop(columns=['Ingreso', 'Egreso'])
    return df


from itertools import combinations

def _conciliacion_automatica(banco_ext: pd.DataFrame, conta_ext: pd.DataFrame, moneda: str = "Dolares (USD)"):
    """
    Algoritmo de conciliación automática con las reglas:
    1. Código igual → Conciliar automáticamente
    2. Monto + Fecha → Conciliar si la coincidencia es única
    3. Monto repetido → Ignorar
    4. Suma (1 a N) o (N a 1) → Buscar combinaciones y conciliar si es única
    """
    banco_ext['_matched'] = False
    conta_ext['_matched'] = False
    
    # 1. Código igual
    for b_idx, row_b in banco_ext.iterrows():
        if row_b['_matched']: continue
        op_b = str(row_b.get('Banco - # Operación', '')).strip()
        if not op_b or op_b == 'nan': continue
        m_b = float(row_b.get('Monto-Banco', 0) or 0)
        
        candidates = conta_ext[
            (~conta_ext['_matched']) & 
            (conta_ext['Conta - # Operación'].astype(str).str.strip() == op_b) &
            (abs(conta_ext['Monto-Conta'].fillna(0) - m_b) <= 0.01)
        ]
        if len(candidates) == 1:
            c_idx = candidates.index[0]
            banco_ext.at[b_idx, '_matched'] = True
            conta_ext.at[c_idx, '_matched'] = True
            banco_ext.at[b_idx, 'MAR'] = 'X'
            conta_ext.at[c_idx, 'MAR'] = 'X'
            banco_ext.at[b_idx, '# Operación2'] = op_b
            conta_ext.at[c_idx, '# Operación2'] = op_b
            banco_ext.at[b_idx, 'Anotación'] = 'Auto: Código igual'
            conta_ext.at[c_idx, 'Anotación'] = 'Auto: Código igual'

    # 1.5 # Operación de Conta contenido en Descripción de Banco + Monto igual
    for c_idx, row_c in conta_ext.iterrows():
        if row_c['_matched']: continue
        op_c = str(row_c.get('Conta - # Operación', '')).strip()
        # Evitar falsos positivos requiriendo al menos 4 caracteres en la operación
        if not op_c or op_c == 'nan' or op_c == '0' or len(op_c) < 4: continue
        
        m_c = float(row_c.get('Monto-Conta', 0) or 0)
        
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

    # Definir diferencias según la moneda elegida
    if "Soles" in moneda or "PEN" in moneda:
        diferencias_set = {10.50, 69.00, 4.30, 29.00}
    else:
        diferencias_set = {1.53, 82.00, 94.00, 69.00, 29.00}

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
            valid_combos = []
            for r in range(2, min(5, len(avail_c) + 1)):
                for combo in combinations(avail_c, r):
                    suma_conta = sum(x[1] for x in combo)
                    diff = round(abs(abs(b_amt) - abs(suma_conta)), 2)
                    if diff in diferencias:
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
            valid_combos = []
            for r in range(2, min(5, len(avail_b) + 1)):
                for combo in combinations(avail_b, r):
                    suma_banco = sum(x[1] for x in combo)
                    diff = round(abs(abs(suma_banco) - abs(c_amt)), 2)
                    if diff in diferencias:
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

    # 5. Sugerencias por diferencias de montos específicos
    diferencias = diferencias_set
    
    conta_ext['_suggested'] = False
    
    for b_idx, row_b in banco_ext.iterrows():
        if row_b['_matched']: continue
        m_b = float(row_b.get('Monto-Banco', 0) or 0)
        
        for c_idx, row_c in conta_ext.iterrows():
            if row_c['_matched']: continue
            if row_c.get('_suggested', False): continue
            
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

def _preparar_anexar1(tab_banco: pd.DataFrame, tab_conta: pd.DataFrame, moneda: str = "Dolares (USD)") -> pd.DataFrame:
    """
    Une Tab_Banco y Tab_Contanet verticalmente con todas las columnas combinadas.
    Las columnas que no existen en una tabla quedan en NaN (el especialista las llena).
    """
    # Columnas del resultado final
    cols = [
        'Fecha', 'Banco - Fecha', 'Conta - Fecha', 'MAR', 'TIPO', 'CODIGO', '# Operación2', 'DIF COMISON', 'Anotación',
        'Monto-Banco', 'Banco - Descripción', 'Banco - # Operación',
        'Monto-Conta', 'Conta - # Registro', 'Conta - # Operación', 'Conta - Giro', 'Conta - Glosa'
    ]

    banco_ext = tab_banco.copy()
    conta_ext = tab_conta.copy()
    # Remove placeholder rows that contain header texts like 'Información anterior' or column names
    def _clean_df(df):
        # Drop rows where any cell equals its column header or the phrase 'Información anterior' or headers like '# Registro'
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
    # Mark ITF rows as reconciled and set operation number to 00000000
    itf_mask = banco_ext['Banco - Descripción'].astype(str).str.upper().str.contains('ITF')
    banco_ext.loc[itf_mask, 'MAR'] = 'X'
    banco_ext.loc[itf_mask, 'Banco - # Operación'] = '00000000'
    banco_ext.loc[itf_mask, '# Operación2'] = '00000000'
    
    if 'Anotación' not in banco_ext.columns: banco_ext['Anotación'] = ''
    if 'Anotación' not in conta_ext.columns: conta_ext['Anotación'] = ''
    
    # Asegurar que las columnas acepten strings para evitar errores de tipo al asignar
    for col in ['MAR', '# Operación2', 'Anotación']:
        if col in banco_ext.columns:
            banco_ext[col] = banco_ext[col].astype(object)
        if col in conta_ext.columns:
            conta_ext[col] = conta_ext[col].astype(object)
    
    # ── APLICAR ALGORITMO DE CONCILIACIÓN AUTOMÁTICA ──
    banco_ext, conta_ext = _conciliacion_automatica(banco_ext, conta_ext, moneda=moneda)

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
    anexar['Conciliar'] = ''
    anexar['# Operación a Conciliar'] = ''
    anexar['Anotación-Conta'] = ''

    # Crear una clave de ordenamiento:
    # 0 = Conciliados (MAR == 'X')
    # 1 = Faltan conciliar pero tienen sugerencia ('Sugerido:' en Anotación)
    # 2 = Solo en banco ('Mov. solo en banco')
    # 3 = Solo en conta ('Mov. solo en conta')
    # 4 = ITF (descripción contiene 'ITF')
    def _orden_grupo(row):
        # ITF: siempre al final independientemente del MAR
        if 'ITF' in str(row.get('Banco - Descripción', '')).upper():
            return 4
        # Conciliados en ESTE proceso (único indicador válido: MAR == 'X')
        # NOTA: Conta - ¿Conciliado? indica si estaba conciliado en el sistema
        # contable ANTES de este proceso, NO indica que esté conciliado aquí.
        if row.get('MAR') == 'X':
            return 0
        anotacion = str(row.get('Anotación', ''))
        # Sugeridos (faltan conciliar pero tienen sugerencia)
        if 'Sugerido:' in anotacion:
            return 1
        # Solo en banco
        if 'solo en banco' in anotacion.lower():
            return 2
        # Solo en conta
        if 'solo en conta' in anotacion.lower():
            return 3
        # Cualquier otro caso sin clasificar
        return 2

    anexar['_orden_conciliado'] = anexar.apply(_orden_grupo, axis=1)

    # Para sugeridos: mantener pares juntos usando # Operación2 como sub-clave
    def get_op2_sort(row, ord_c):
        if ord_c == 1:
            return str(row.get('# Operación2', ''))
        return ''

    anexar['_op2_sort'] = anexar.apply(lambda r: get_op2_sort(r, r['_orden_conciliado']), axis=1)

    anexar = anexar.sort_values(
        by=['_orden_conciliado', '_op2_sort', 'Fecha']
    ).drop(columns=['_orden_conciliado', '_op2_sort']).reset_index(drop=True)

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

    for nombre_hoja in wb.sheetnames:
        ws = wb[nombre_hoja]
        color_hex = _COLORES.get(nombre_hoja, '404040')

        # Formato de cabecera (fila 1)
        for cell in ws[1]:
            col_name = str(cell.value)
            
            if nombre_hoja == 'Anexar1':
                if 'Banco' in col_name:
                    fg = '1F4E79'  # Azul oscuro para Banco
                elif 'Conta' in col_name:
                    fg = '375623'  # Verde oscuro para Contanet
                else:
                    fg = '7030A0'  # Morado para columnas de conciliación/comunes
            else:
                fg = color_hex
                
            cell.font      = Font(bold=True, color='FFFFFF', size=10)
            cell.fill      = PatternFill('solid', fgColor=fg)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border    = borde

        # Altura de cabecera
        ws.row_dimensions[1].height = 30

        # Encontrar índice de columnas relevantes
        col_mar_idx = None
        col_anot_idx = None
        col_conc_idx = None
        col_no_conc_idx = None
        col_op2_idx = None
        for i, cell in enumerate(ws[1]):
            val = str(cell.value).strip()
            if val == 'MAR':
                col_mar_idx = i
            elif val == 'Anotación':
                col_anot_idx = i
            elif val == 'Conciliar':
                col_conc_idx = i
            elif val == '# Operación a Conciliar' or val == 'No Conciliar':
                col_no_conc_idx = i
            elif val == '# Operación2':
                col_op2_idx = i

        # Detectar índice de columna Banco - Descripción para ITF
        col_banco_desc_idx = None
        for i, cell in enumerate(ws[1]):
            if str(cell.value).strip() == 'Banco - Descripción':
                col_banco_desc_idx = i
                break

        # Formato de datos
        for row in ws.iter_rows(min_row=2):
            es_rojo = False
            es_sugerido = False
            es_itf = False

            if nombre_hoja == 'Anexar1' and col_mar_idx is not None:
                val_mar = row[col_mar_idx].value
                # Verificar si es ITF
                if col_banco_desc_idx is not None:
                    desc_val = str(row[col_banco_desc_idx].value or '')
                    if 'ITF' in desc_val.upper():
                        es_itf = True
                if not es_itf:
                    if val_mar != 'X':
                        es_rojo = True
                        if col_anot_idx is not None:
                            if 'Sugerido:' in str(row[col_anot_idx].value or ''):
                                es_sugerido = True
                                es_rojo = False

            for idx, cell in enumerate(row):
                cell.alignment = Alignment(vertical='center')
                cell.border    = borde
                
                # Fondo distintivo para columnas en Anexar1: Celeste claro para Banco, Verde claro para Conta
                if nombre_hoja == 'Anexar1':
                    header_name = str(ws.cell(row=1, column=idx + 1).value or '')
                    if 'Banco' in header_name:
                        cell.fill = PatternFill('solid', fgColor='D9E1F2')  # Celeste claro
                    elif 'Conta' in header_name:
                        cell.fill = PatternFill('solid', fgColor='E2EFDA')  # Verde claro

                if es_itf:
                    cell.font = Font(color='7F7F7F', italic=True)  # Gris itálica para ITF
                elif es_rojo:
                    cell.font = Font(color='FF0000')  # Letra roja para no conciliados
                elif es_sugerido:
                    cell.font = Font(color='0070C0', bold=True)  # Azul negrita para sugeridos
                    
                # Fechas en formato legible
                if isinstance(cell.value, datetime):
                    cell.number_format = 'DD/MM/YYYY'

        # Combinar ÚNICAMENTE celdas de la columna 'Conciliar' para grupos sugeridos
        if (nombre_hoja == 'Anexar1' and col_op2_idx is not None
                and col_conc_idx is not None and col_anot_idx is not None):
            group_rows = {}
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

        # Autoajuste de columnas
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
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
