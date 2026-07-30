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
  Hoja Conciliados     → pares banco↔conta confirmados por el especialista
  Hoja Solo_Banco      → movimientos bancarios sin par contable
  Hoja Solo_Conta      → movimientos contables sin par bancario
  Hoja Resumen_Final   → tabla por TIPO/CODIGO con diferencias residuales
  Hoja Estadisticas    → totales generales y saldo de diferencias

Uso:
    from reporte_final import generar_reporte_final
    generar_reporte_final(ruta_excel_trabajado)
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

def generar_reporte_final(
    ruta_inicial: Path,
    ruta_salida: Path | None = None,
) -> Path:
    """
    Lee el Excel inicial trabajado por el especialista y genera el reporte final.
    Devuelve la ruta del archivo creado.
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

    # ── Aprobar sugerencias 1-a-1 marcadas por el especialista ───────────
    # Si el especialista puso algo en la columna 'Conciliar' para una fila
    # con anotación 'Sugerido:' y que tiene AMBOS montos (banco y conta en
    if 'Conciliar' in anexar1.columns and 'Anotación' in anexar1.columns:
        mask_sugerido_1a1 = (
            anexar1['Anotación'].astype(str).str.contains('Sugerido:', na=False)
            & anexar1['Conciliar'].apply(lambda v: pd.notna(v) and str(v).strip() not in ('', 'nan'))
            & anexar1['Monto-Banco'].notna()
            & anexar1['Monto-Conta'].notna()
        )
        if mask_sugerido_1a1.any():
            anexar1.loc[mask_sugerido_1a1, 'MAR'] = 'X'
            # Copiar # Operación del banco al campo # Operación2 (usado para el cruce)
            if 'Banco - # Operación' in anexar1.columns and '# Operación2' in anexar1.columns:
                op_banco = anexar1.loc[mask_sugerido_1a1, 'Banco - # Operación'].apply(
                    lambda v: str(v) if pd.notna(v) and str(v) not in ('nan', '') else pd.NA
                )
                anexar1.loc[mask_sugerido_1a1, '# Operación2'] = op_banco


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

    # ── Clasificar según MAR ────────────────────────────────────
    # El especialista pone 'X' en MAR para marcar filas vinculadas
    banco_marcado    = df_banco[df_banco['MAR'] == 'X'].copy()
    banco_sin_marcar = df_banco[df_banco['MAR'] != 'X'].copy()
    conta_marcado    = df_conta[df_conta['MAR'] == 'X'].copy()
    conta_sin_marcar = df_conta[df_conta['MAR'] != 'X'].copy()

    # ── Construir tabla de conciliados ──────────────────────────────
    # a) Pares del especialista en filas separadas
    conciliados_manual = _cruzar_marcados(banco_marcado, conta_marcado)
    # b) Sugerencias 1-a-1 aprobadas directamente (filas combinadas)
    conciliados_1a1 = _conciliados_de_combinadas(df_combinadas)
    conciliados = pd.concat([conciliados_manual, conciliados_1a1], ignore_index=True)

    # ── Solo banco y solo conta ──────────────────────────────────
    solo_banco = banco_sin_marcar[[
        'Fecha', 'Banco - Descripción', 'Monto-Banco', 'Banco - # Operación', 'TIPO', 'CODIGO', 'Anotación'
    ]].copy() if not banco_sin_marcar.empty else pd.DataFrame()

    solo_conta = conta_sin_marcar[[
        'Fecha', 'Conta - # Registro', 'Monto-Conta', 'TIPO', 'CODIGO', 'Anotación'
    ]].copy() if not conta_sin_marcar.empty else pd.DataFrame()

    # ── Resumen final (sobre Anexar1 completo actualizado) ────────────
    resumen_final = _preparar_resumen_final(anexar1)

    # ── Estadísticas ──────────────────────────────────────────────────
    estadisticas = _preparar_estadisticas(
        banco_marcado, banco_sin_marcar,
        conta_marcado, conta_sin_marcar,
        conciliados
    )

    # ── Escribir Excel ────────────────────────────────────────────────
    with pd.ExcelWriter(ruta_salida, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
        if not conciliados.empty:
            conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
        if not solo_banco.empty:
            solo_banco.to_excel(writer, sheet_name='Solo_Banco', index=False)
        if not solo_conta.empty:
            solo_conta.to_excel(writer, sheet_name='Solo_Conta', index=False)
        resumen_final.to_excel(writer, sheet_name='Resumen_Final', index=False)
        estadisticas.to_excel(writer, sheet_name='Estadisticas', index=False)

    # ── Formato visual ────────────────────────────────────────────────
    _aplicar_formato_final(ruta_salida)

    # ── Imprimir resumen en consola ───────────────────────────────────
    _imprimir_resumen_consola(estadisticas, ruta_salida)

    return ruta_salida


# ══════════════════════════════════════════════════════════════════════
# FUNCIONES INTERNAS
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

    # Intentar cruce por # Operación2 (llenado por el especialista)
    usados_conta = set()
    usados_banco = set()

    for i, rb in banco_marcado.iterrows():
        op2_b = str(rb.get('# Operación2', '')).strip()
        matched = False

        # Buscar en conta una fila cuyo # Operación2 coincida con # Operación del banco
        for j, rc in conta_marcado.iterrows():
            if j in usados_conta:
                continue
            op2_c = str(rc.get('# Operación2', '')).strip()
            op_b  = str(rb.get('Banco - # Operación',  '')).strip()

            if op2_b and op2_c and (op2_c == op_b or op2_b == str(rc.get('Conta - # Operación', '')).strip()):
                filas.append(_fila_conciliada(rb, rc))
                usados_conta.add(j)
                usados_banco.add(i)
                matched = True
                break

        if not matched:
            # Fallback: cruce por monto idéntico entre marcados
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
    Cada fila ya contiene tanto los datos del banco como de conta.
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


def _preparar_resumen_final(anexar1: pd.DataFrame) -> pd.DataFrame:
    """Resumen por TIPO/CODIGO con diferencias residuales (igual que el inicial)."""
    resumen = (
        anexar1
        .groupby(['TIPO', 'CODIGO'], dropna=False)
        .agg(
            **{'Monto-Conta' : ('Monto-Conta', 'sum'),
               'Monto-Banco' : ('Monto-Banco', 'sum')}
        )
        .reset_index()
    )
    resumen['Diferencia'] = resumen['Monto-Banco'].fillna(0) - resumen['Monto-Conta'].fillna(0)
    resumen = resumen.sort_values(['TIPO', 'CODIGO']).reset_index(drop=True)
    return resumen


def _preparar_estadisticas(
    banco_marcado, banco_sin_marcar,
    conta_marcado, conta_sin_marcar,
    conciliados,
) -> pd.DataFrame:
    """Tabla de estadísticas generales de la conciliación."""

    def safe_sum(df, col):
        if df.empty or col not in df.columns:
            return 0.0
        return pd.to_numeric(df[col], errors='coerce').fillna(0).sum()

    total_banco_ing   = safe_sum(pd.concat([banco_marcado, banco_sin_marcar]), 'Monto-Banco')
    total_conta_ing   = safe_sum(pd.concat([conta_marcado, conta_sin_marcar]), 'Monto-Conta')
    monto_conciliado  = safe_sum(conciliados, 'Monto-Banco') if not conciliados.empty else 0
    monto_solo_banco  = safe_sum(banco_sin_marcar, 'Monto-Banco')
    monto_solo_conta  = safe_sum(conta_sin_marcar, 'Monto-Conta')

    filas = [
        ('Movimientos banco (total)',           len(banco_marcado) + len(banco_sin_marcar)),
        ('Movimientos conta (total)',           len(conta_marcado) + len(conta_sin_marcar)),
        ('Pares conciliados por especialista',  len(banco_marcado)),
        ('Monto conciliado (banco)',            round(monto_conciliado, 2)),
        ('Movimientos solo en banco',           len(banco_sin_marcar)),
        ('Monto solo en banco',                 round(monto_solo_banco, 2)),
        ('Movimientos solo en contabilidad',    len(conta_sin_marcar)),
        ('Monto solo en contabilidad',          round(monto_solo_conta, 2)),
        ('Total banco',                         round(total_banco_ing, 2)),
        ('Total contabilidad',                  round(total_conta_ing, 2)),
        ('Diferencia neta',                     round(total_banco_ing - total_conta_ing, 2)),
    ]
    return pd.DataFrame(filas, columns=['Concepto', 'Valor'])


# ══════════════════════════════════════════════════════════════════════
# FORMATO VISUAL
# ══════════════════════════════════════════════════════════════════════

_COLORES_FINAL = {
    'Conciliados'   : '1F4E79',   # azul oscuro
    'Solo_Banco'    : 'C55A11',   # naranja (partidas abiertas banco)
    'Solo_Conta'    : '375623',   # verde oscuro (partidas abiertas conta)
    'Resumen_Final' : '7030A0',   # morado
    'Estadisticas'  : '404040',   # gris oscuro
}


def _aplicar_formato_final(ruta: Path) -> None:
    wb = openpyxl.load_workbook(ruta)
    thin  = Side(style='thin', color='D9D9D9')
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    for nombre_hoja in wb.sheetnames:
        ws        = wb[nombre_hoja]
        color_hex = _COLORES_FINAL.get(nombre_hoja, '404040')

        for cell in ws[1]:
            cell.font      = Font(bold=True, color='FFFFFF', size=10)
            cell.fill      = PatternFill('solid', fgColor=color_hex)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border    = borde

        ws.row_dimensions[1].height = 30

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical='center')
                cell.border    = borde
                if isinstance(cell.value, datetime):
                    cell.number_format = 'DD/MM/YYYY'

        for col in ws.columns:
            max_len   = 0
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
# RESUMEN EN CONSOLA
# ══════════════════════════════════════════════════════════════════════

def _imprimir_resumen_consola(estadisticas: pd.DataFrame, ruta: Path) -> None:
    print()
    print("=" * 60)
    print("  RESULTADO DE CONCILIACION FINAL")
    print("=" * 60)
    for _, row in estadisticas.iterrows():
        print(f"  {row['Concepto']:<42}: {row['Valor']:>15,}")
    print()
    print(f"  [OK] Reporte final generado: {ruta}")
    print()
