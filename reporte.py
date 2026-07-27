"""
reporte.py
──────────
Genera el reporte Excel de conciliación bancaria siguiendo la estructura
de la plantilla 'Conciliacion Bancaria Tabla.xlsx':

  Hoja BANCO       → datos completos del banco
  Hoja CONTANET    → datos completos de contabilidad
  Hoja Tab_Banco   → tabla simplificada del banco (para cruce)
  Hoja Tab_Contanet→ tabla simplificada de contabilidad (para cruce)
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
    anexar1 = _preparar_anexar1(tab_banco, tab_conta)

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
    from pdf_resumen import generar_pdf_resumen
    ruta_pdf = ruta_salida.with_suffix('.pdf')
    generar_pdf_resumen(
        bank=bank,
        conta=conta,
        anexar1=anexar1,
        resumen_df=resumen,
        ruta_pdf=ruta_pdf,
        empresa=empresa,
        moneda=moneda,
        ruta_excel=ruta_salida,
    )

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
    df['MAR']          = ''
    df['TIPO']         = df['Monto-Banco'].apply(lambda x: 'INGRESO' if x > 0 else 'EGRESO')
    df['CODIGO']       = df['Monto-Banco']
    df['# Operación2'] = df['Banco - # Operación']
    return df


def _preparar_tab_contanet(conta: pd.DataFrame) -> pd.DataFrame:
    """Tabla completa de contabilidad para el cruce en Anexar1."""
    df = conta[['nro_registro', 'fecha_mov', 'nro_operacion', 'ingreso', 'egreso', 'giro', 'glosa', 'fecha_conciliacion', 'conciliado']].copy()
    df.columns = ['Conta - # Registro', 'Fecha', 'Conta - # Operación', 'Ingreso', 'Egreso', 'Conta - Giro', 'Conta - Glosa', 'Conta - F. Conciliación', 'Conta - ¿Conciliado?']
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

def _conciliacion_automatica(banco_ext: pd.DataFrame, conta_ext: pd.DataFrame):
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

    # 4. Sumas N a 1 y 1 a N
    fechas = set(banco_ext.loc[~banco_ext['_matched'], 'Fecha']).union(
             set(conta_ext.loc[~conta_ext['_matched'], 'Fecha']))
             
    for f in fechas:
        # 1 Banco = N Conta
        unmatched_b = banco_ext[(~banco_ext['_matched']) & (banco_ext['Fecha'] == f)]
        unmatched_c = conta_ext[(~conta_ext['_matched']) & (conta_ext['Fecha'] == f)]
        pool_b = [(idx, float(r.get('Monto-Banco', 0) or 0)) for idx, r in unmatched_b.iterrows()]
        pool_c = [(idx, float(r.get('Monto-Conta', 0) or 0)) for idx, r in unmatched_c.iterrows()]
        
        used_c = set()
        for b_idx, b_amt in pool_b:
            avail_c = [x for x in pool_c if x[0] not in used_c]
            valid_combos = []
            for r in range(2, min(5, len(avail_c) + 1)):
                for combo in combinations(avail_c, r):
                    if abs(sum(x[1] for x in combo) - b_amt) <= 0.01:
                        valid_combos.append([x[0] for x in combo])
            if len(valid_combos) == 1:
                combo_idxs = valid_combos[0]
                used_c.update(combo_idxs)
                banco_ext.at[b_idx, '_matched'] = True
                banco_ext.at[b_idx, 'MAR'] = 'X'
                op_link = f"SUM-B{b_idx}"
                banco_ext.at[b_idx, '# Operación2'] = op_link
                banco_ext.at[b_idx, 'Anotación'] = 'Auto: 1 Banco = N Conta'
                for c_idx in combo_idxs:
                    conta_ext.at[c_idx, '_matched'] = True
                    conta_ext.at[c_idx, 'MAR'] = 'X'
                    conta_ext.at[c_idx, '# Operación2'] = op_link
                    conta_ext.at[c_idx, 'Anotación'] = 'Auto: 1 Banco = N Conta'
                    
        # N Banco = 1 Conta (refrescar pools)
        unmatched_c2 = conta_ext[(~conta_ext['_matched']) & (conta_ext['Fecha'] == f)]
        unmatched_b2 = banco_ext[(~banco_ext['_matched']) & (banco_ext['Fecha'] == f)]
        pool_c2 = [(idx, float(r.get('Monto-Conta', 0) or 0)) for idx, r in unmatched_c2.iterrows()]
        pool_b2 = [(idx, float(r.get('Monto-Banco', 0) or 0)) for idx, r in unmatched_b2.iterrows()]
        
        used_b = set()
        for c_idx, c_amt in pool_c2:
            avail_b = [x for x in pool_b2 if x[0] not in used_b]
            valid_combos = []
            for r in range(2, min(5, len(avail_b) + 1)):
                for combo in combinations(avail_b, r):
                    if abs(sum(x[1] for x in combo) - c_amt) <= 0.01:
                        valid_combos.append([x[0] for x in combo])
            if len(valid_combos) == 1:
                combo_idxs = valid_combos[0]
                used_b.update(combo_idxs)
                conta_ext.at[c_idx, '_matched'] = True
                conta_ext.at[c_idx, 'MAR'] = 'X'
                op_link = f"SUM-C{c_idx}"
                conta_ext.at[c_idx, '# Operación2'] = op_link
                conta_ext.at[c_idx, 'Anotación'] = 'Auto: N Banco = 1 Conta'
                for b_idx in combo_idxs:
                    banco_ext.at[b_idx, '_matched'] = True
                    banco_ext.at[b_idx, 'MAR'] = 'X'
                    banco_ext.at[b_idx, '# Operación2'] = op_link
                    banco_ext.at[b_idx, 'Anotación'] = 'Auto: N Banco = 1 Conta'

    # Marcar los no conciliados con una anotación por defecto
    banco_ext.loc[~banco_ext['_matched'], 'Anotación'] = 'Mov. solo en banco'
    conta_ext.loc[~conta_ext['_matched'], 'Anotación'] = 'Mov. solo en conta'

    return banco_ext.drop(columns=['_matched']), conta_ext.drop(columns=['_matched'])

def _preparar_anexar1(tab_banco: pd.DataFrame, tab_conta: pd.DataFrame) -> pd.DataFrame:
    """
    Une Tab_Banco y Tab_Contanet verticalmente con todas las columnas combinadas.
    Las columnas que no existen en una tabla quedan en NaN (el especialista las llena).
    """
    # Columnas del resultado final
    cols = [
        'Fecha', 'MAR', 'TIPO', 'CODIGO', '# Operación2', 'DIF COMISON', 'Anotación',
        'Monto-Banco', 'Banco - Descripción', 'Banco - # Operación', 'Banco - Saldo',
        'Monto-Conta', 'Conta - # Registro', 'Conta - # Operación', 'Conta - Giro', 'Conta - Glosa', 'Conta - F. Conciliación', 'Conta - ¿Conciliado?'
    ]

    banco_ext = tab_banco.copy()
    conta_ext = tab_conta.copy()
    
    if 'Anotación' not in banco_ext.columns: banco_ext['Anotación'] = ''
    if 'Anotación' not in conta_ext.columns: conta_ext['Anotación'] = ''
    
    # Asegurar que las columnas acepten strings para evitar errores de tipo al asignar
    for col in ['MAR', '# Operación2', 'Anotación']:
        if col in banco_ext.columns:
            banco_ext[col] = banco_ext[col].astype(object)
        if col in conta_ext.columns:
            conta_ext[col] = conta_ext[col].astype(object)
    
    # ── APLICAR ALGORITMO DE CONCILIACIÓN AUTOMÁTICA ──
    banco_ext, conta_ext = _conciliacion_automatica(banco_ext, conta_ext)

    # Asegurar que todas las columnas existan en ambos para evitar warnings
    for c in cols:
        if c not in banco_ext.columns:
            banco_ext[c] = None
        if c not in conta_ext.columns:
            conta_ext[c] = None

    anexar = pd.concat(
        [banco_ext[cols], conta_ext[cols]],
        ignore_index=True
    )

    # Crear una clave de ordenamiento para subir los conciliados arriba pero MANTENER LOS PARES JUNTOS
    def _is_matched(row):
        if row.get('MAR') == 'X': return 0
        val = row.get('Conta - ¿Conciliado?')
        if pd.notna(val):
            s = str(val).strip().upper()
            if s in ['SI', 'S', 'X', 'TRUE', '1', 'YES', 'Y', 'VERDADERO']: return 0
        return 1

    anexar['_orden_conciliado'] = anexar.apply(_is_matched, axis=1)
    anexar['_op2_sort'] = anexar['# Operación2'].fillna('').astype(str)
    
    anexar = anexar.sort_values(
        by=['_orden_conciliado', 'Fecha', '_op2_sort']
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

        # Formato de datos
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical='center')
                cell.border    = borde
                # Fechas en formato legible
                if isinstance(cell.value, datetime):
                    cell.number_format = 'DD/MM/YYYY'

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
