"""
conciliacion.py
───────────────
Lógica de cruce entre el estado de cuenta bancario y los registros
contables para producir las partidas conciliatorias.

Recibe los DataFrames ya sanitizados (bank, conta) y devuelve:
  - conciliados   : movimientos presentes en ambos lados
  - solo_banco    : movimientos que están en el banco pero no en contabilidad
  - solo_conta    : movimientos que están en contabilidad pero no en el banco
  - resumen       : dict con totales y diferencias
"""

import pandas as pd


# ══════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════

def conciliar(bank: pd.DataFrame, conta: pd.DataFrame) -> dict:
    """
    Cruza bank contra conta por MONTO (ingreso/egreso) y FECHA.
    Devuelve un dict con las cuatro tablas resultado y el resumen.

    Criterio de match:
      - Mismo tipo de movimiento (ingreso o egreso)
      - Mismo importe (tolerancia ±0.01 para redondeos)
      - Misma fecha (o ±1 día para diferencias de valor)
    """

    # ── Preparar banco ────────────────────────────────────────────────
    # Separamos en dos tablas auxiliares: ingresos y egresos del banco
    b_ing = (
        bank[bank['ingreso'] > 0]
        .copy()
        .rename(columns={'ingreso': 'monto_cruce', 'fecha': 'fecha_cruce'})
        [['fecha_cruce', 'monto_cruce', 'descripcion', 'nro_operacion', 'sucursal']]
        .assign(tipo='ingreso')
        .reset_index(drop=True)
    )
    b_egr = (
        bank[bank['egreso'] > 0]
        .copy()
        .rename(columns={'egreso': 'monto_cruce', 'fecha': 'fecha_cruce'})
        [['fecha_cruce', 'monto_cruce', 'descripcion', 'nro_operacion', 'sucursal']]
        .assign(tipo='egreso')
        .reset_index(drop=True)
    )

    # ── Preparar contabilidad ─────────────────────────────────────────
    c_ing = (
        conta[conta['ingreso'] > 0]
        .copy()
        .rename(columns={'ingreso': 'monto_cruce', 'fecha_mov': 'fecha_cruce'})
        [['fecha_cruce', 'monto_cruce', 'nro_registro', 'giro', 'glosa', 'conciliado']]
        .assign(tipo='ingreso')
        .reset_index(drop=True)
    )
    c_egr = (
        conta[conta['egreso'] > 0]
        .copy()
        .rename(columns={'egreso': 'monto_cruce', 'fecha_mov': 'fecha_cruce'})
        [['fecha_cruce', 'monto_cruce', 'nro_registro', 'giro', 'glosa', 'conciliado']]
        .assign(tipo='egreso')
        .reset_index(drop=True)
    )

    # ── Cruce por tipo ────────────────────────────────────────────────
    matches_ing, solo_b_ing, solo_c_ing = _cruzar(b_ing, c_ing)
    matches_egr, solo_b_egr, solo_c_egr = _cruzar(b_egr, c_egr)

    conciliados = pd.concat([matches_ing, matches_egr], ignore_index=True)
    solo_banco  = pd.concat([solo_b_ing,  solo_b_egr],  ignore_index=True)
    solo_conta  = pd.concat([solo_c_ing,  solo_c_egr],  ignore_index=True)

    # ── Resumen ───────────────────────────────────────────────────────
    resumen = {
        'total_banco_ing'   : bank['ingreso'].sum(),
        'total_banco_egr'   : bank['egreso'].sum(),
        'total_conta_ing'   : conta['ingreso'].sum(),
        'total_conta_egr'   : conta['egreso'].sum(),
        'diferencia_ing'    : bank['ingreso'].sum() - conta['ingreso'].sum(),
        'diferencia_egr'    : bank['egreso'].sum()  - conta['egreso'].sum(),
        'n_conciliados'     : len(conciliados),
        'n_solo_banco'      : len(solo_banco),
        'n_solo_conta'      : len(solo_conta),
        'monto_solo_banco'  : solo_banco['monto_cruce'].sum()  if len(solo_banco)  else 0,
        'monto_solo_conta'  : solo_conta['monto_cruce'].sum()  if len(solo_conta)  else 0,
    }

    return {
        'conciliados' : conciliados,
        'solo_banco'  : solo_banco,
        'solo_conta'  : solo_conta,
        'resumen'     : resumen,
    }


# ══════════════════════════════════════════════════════════════════════
# FUNCIÓN INTERNA DE CRUCE
# ══════════════════════════════════════════════════════════════════════

def _cruzar(
    df_banco: pd.DataFrame,
    df_conta: pd.DataFrame,
    tolerancia_monto: float = 0.01,
    tolerancia_dias:  int   = 1,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Cruce greedy entre dos tablas del mismo tipo (ingreso o egreso).
    Cada fila del banco se empareja con la fila más cercana en fecha
    de contabilidad que tenga el mismo monto (±tolerancia).
    Cada fila solo puede usarse una vez (one-to-one).

    Devuelve: (matches, solo_banco, solo_conta)
    """
    banco = df_banco.copy().reset_index(drop=True)
    conta = df_conta.copy().reset_index(drop=True)

    usados_conta = set()
    filas_match  = []
    filas_solo_b = []

    for i, row_b in banco.iterrows():
        mejor_idx   = None
        mejor_delta = None

        for j, row_c in conta.iterrows():
            if j in usados_conta:
                continue

            # Verificar monto
            if abs(row_b['monto_cruce'] - row_c['monto_cruce']) > tolerancia_monto:
                continue

            # Verificar fecha (puede ser NaT)
            if pd.isna(row_b['fecha_cruce']) or pd.isna(row_c['fecha_cruce']):
                delta_dias = 0  # si falta fecha aceptamos igual
            else:
                delta_dias = abs((row_b['fecha_cruce'] - row_c['fecha_cruce']).days)

            if delta_dias > tolerancia_dias:
                continue

            # Candidato válido: preferir el de menor diferencia de fecha
            if mejor_idx is None or delta_dias < mejor_delta:
                mejor_idx   = j
                mejor_delta = delta_dias

        if mejor_idx is not None:
            usados_conta.add(mejor_idx)
            filas_match.append({
                'tipo'              : row_b['tipo'],
                'fecha_banco'       : row_b['fecha_cruce'],
                'fecha_conta'       : conta.at[mejor_idx, 'fecha_cruce'],
                'monto'             : row_b['monto_cruce'],
                'descripcion_banco' : row_b.get('descripcion', ''),
                'nro_op_banco'      : row_b.get('nro_operacion', ''),
                'nro_registro_conta': conta.at[mejor_idx, 'nro_registro'],
                'giro_conta'        : conta.at[mejor_idx, 'giro'],
            })
        else:
            filas_solo_b.append(row_b)

    # Filas de conta sin par
    filas_solo_c = [
        row for j, row in conta.iterrows() if j not in usados_conta
    ]

    matches   = pd.DataFrame(filas_match)
    solo_b    = pd.DataFrame(filas_solo_b).reset_index(drop=True) if filas_solo_b else pd.DataFrame()
    solo_c    = pd.DataFrame(filas_solo_c).reset_index(drop=True) if filas_solo_c else pd.DataFrame()

    return matches, solo_b, solo_c


# ══════════════════════════════════════════════════════════════════════
# IMPRESIÓN DE RESULTADOS
# ══════════════════════════════════════════════════════════════════════

def imprimir_resultados(resultado: dict) -> None:
    """Muestra en consola el resumen de la conciliación."""
    r  = resultado['resumen']
    sb = resultado['solo_banco']
    sc = resultado['solo_conta']

    print()
    print("=" * 60)
    print("  RESULTADO DE CONCILIACIÓN")
    print("=" * 60)

    print(f"\n{'Movimientos cruzados (conciliados)':<38}: {r['n_conciliados']:>6}")
    print(f"{'Movimientos solo en BANCO':<38}: {r['n_solo_banco']:>6}   (S/ {r['monto_solo_banco']:,.2f})")
    print(f"{'Movimientos solo en CONTABILIDAD':<38}: {r['n_solo_conta']:>6}   (S/ {r['monto_solo_conta']:,.2f})")

    print()
    print(f"{'':─<60}")
    print(f"  {'':30} {'BANCO':>12}  {'CONTA':>12}")
    print(f"{'':─<60}")
    print(f"  {'Total ingresos':<30} {r['total_banco_ing']:>12,.2f}  {r['total_conta_ing']:>12,.2f}  Δ {r['diferencia_ing']:+,.2f}")
    print(f"  {'Total egresos':<30} {r['total_banco_egr']:>12,.2f}  {r['total_conta_egr']:>12,.2f}  Δ {r['diferencia_egr']:+,.2f}")
    print(f"{'':─<60}")

    if r['n_solo_banco'] > 0:
        print("\n── Movimientos en BANCO sin registro contable ──")
        print(sb[['tipo', 'fecha_cruce', 'monto_cruce', 'descripcion']].to_string(index=False))

    if r['n_solo_conta'] > 0:
        print("\n── Movimientos en CONTABILIDAD sin registro bancario ──")
        cols_sc = [c for c in ['tipo', 'fecha_cruce', 'monto_cruce', 'nro_registro', 'giro'] if c in sc.columns]
        print(sc[cols_sc].to_string(index=False))
