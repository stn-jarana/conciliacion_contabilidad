import pandas as pd
from pathlib import Path

# ══════════════════════════════════════════════════════════
# CONFIGURACIÓN POR EMPRESA Y MONEDA
# ══════════════════════════════════════════════════════════

EMPRESAS = [
    {
        "nombre"  : "Southern Textil Network (STN)",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "STN DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "STN SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Integrated Textile Solutions (ITS)",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "ITS DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "ITS SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "CMT del Sur",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "CMT DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "CMT SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Dynamitex",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "DYNAMITEX DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "DYNAMITEX SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "DINSURA",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "DINSURA DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "DINSURA SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Perú Commerce",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "P.COMMERCE",   "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Inversiones Forestales del Sur (INFOSUR)",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "INFOSUR DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "INFOSUR SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Thimble Sourcing / TST",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "TST",         "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Reforestadora Iñaupari",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "REF.IÑAPARI",  "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "TECA Peruvian Group",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "TECA",        "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "DIONISO",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "DIONISO",     "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
]


# ══════════════════════════════════════════════════════════
# FUNCIONES DE MENÚ
# ══════════════════════════════════════════════════════════

def seleccionar_empresa() -> dict:
    """Muestra el menú de empresas y monedas; devuelve la configuración elegida."""
    while True:
        print("=" * 60)
        print("  CONCILIACION BANCARIA")
        print("=" * 60)
        print("\nSeleccione la empresa:\n")
        for i, emp in enumerate(EMPRESAS, 1):
            print(f"  {i}. {emp['nombre']}")
        print("  0. Salir")
        print()
        opcion = input("Opcion: ").strip()

        if opcion == "0":
            print("\n  Hasta luego.\n")
            raise SystemExit(0)

        if not opcion.isdigit() or not (1 <= int(opcion) <= len(EMPRESAS)):
            print("  X Opcion invalida, intente de nuevo.\n")
            continue

        empresa = EMPRESAS[int(opcion) - 1]

        while True:
            print(f"\nEmpresa: {empresa['nombre']}")
            print("Seleccione la moneda:\n")
            for j, mon in enumerate(empresa['monedas'], 1):
                estado = "" if mon['habilitado'] else "  [no disponible]"
                print(f"  {j}. {mon['nombre']}{estado}")
            print("  0. Volver")
            print()

            opcion_mon = input("Opcion: ").strip()

            if opcion_mon == "0":
                break

            if not opcion_mon.isdigit() or not (1 <= int(opcion_mon) <= len(empresa['monedas'])):
                print("  X Opcion invalida, intente de nuevo.\n")
                continue

            moneda = empresa['monedas'][int(opcion_mon) - 1]
            if not moneda['habilitado']:
                print(f"\n  X La opcion '{moneda['nombre']}' aun no esta disponible.\n")
                continue

            print(f"\n  >> {empresa['nombre']} - {moneda['nombre']}\n")
            # Incluir nombre de empresa en el config para el reporte
            return {**moneda, 'empresa': empresa['nombre']}


def seleccionar_accion() -> str:
    """Pregunta qué desea hacer: reporte inicial o reporte final."""
    while True:
        print("-" * 60)
        print("  Que desea hacer?\n")
        print("  1. Generar reporte inicial")
        print("     (carga archivos banco + contabilidad y genera el Excel de trabajo)")
        print()
        print("  2. Generar reporte final")
        print("     (carga el Excel ya trabajado por el especialista)")
        print()
        print("  0. Salir")
        print()
        opcion = input("Opcion: ").strip()

        if opcion == "0":
            print("\n  Hasta luego.\n")
            raise SystemExit(0)
        if opcion in ("1", "2"):
            return opcion
        print("  X Opcion invalida, intente de nuevo.\n")


def pedir_archivo(mensaje: str) -> Path:
    """Solicita una ruta de archivo hasta que exista y sea .xlsx/.xls."""
    while True:
        ruta = input(mensaje).strip().strip('"').strip("'")
        path = Path(ruta)
        if not path.exists():
            print(f"  X No se encontro el archivo: {ruta}")
        elif path.suffix.lower() not in ('.xlsx', '.xls'):
            print("  X El archivo debe ser Excel (.xlsx o .xls)")
        else:
            return path


# ══════════════════════════════════════════════════════════
# FLUJO — REPORTE INICIAL
# ══════════════════════════════════════════════════════════

def flujo_reporte_inicial(config: dict) -> None:
    """Carga, sanitiza y genera el reporte inicial de conciliacion."""
    from reporte import generar_reporte_inicial

    print("-" * 60)
    file_bank  = pedir_archivo("Ruta del estado de cuenta del BANCO (.xlsx): ")
    file_conta = pedir_archivo("Ruta del reporte de CONTABILIDAD (.xlsx)  : ")
    print()

    # ── Carga raw ─────────────────────────────────────────────────────
    bank_raw  = pd.read_excel(file_bank,  sheet_name=config['sheet_bank'], skiprows=config['skip_bank'])
    conta_raw = pd.read_excel(file_conta, skiprows=config['skip_conta'])

    # ── Sanitización banco ────────────────────────────────────────────
    bank = bank_raw.copy()
    bank.columns = [
        'fecha', 'fecha_valuta', 'descripcion', 'monto', 'saldo',
        'sucursal', 'nro_operacion', 'hora', 'usuario', 'utc', 'referencia'
    ]
    bank['fecha']        = pd.to_datetime(bank['fecha'],        format='%d/%m/%Y', errors='coerce')
    bank['fecha_valuta'] = pd.to_datetime(bank['fecha_valuta'], format='%d/%m/%Y', errors='coerce')
    bank['ingreso']      = bank['monto'].clip(lower=0)
    bank['egreso']       = bank['monto'].clip(upper=0).abs()
    bank = bank.drop(columns=['fecha_valuta', 'referencia'])
    bank['descripcion']  = bank['descripcion'].str.strip()
    bank = bank.sort_values('fecha').reset_index(drop=True)

    # ── Sanitización contabilidad ─────────────────────────────────────
    conta = conta_raw.copy()
    conta.columns = [
        'col_vacia', 'nro_registro', 'fecha_mov', 'medio_pago',
        'nro_operacion', 'giro', 'glosa', 'ingreso', 'egreso',
        'fecha_conciliacion', 'conciliado'
    ]
    conta = conta.drop(columns=['col_vacia'])
    conta = conta[
        conta['nro_registro'].notna() &
        ~conta['nro_registro'].astype(str).str.startswith('Saldo', na=False) &
        ~conta['nro_registro'].astype(str).str.contains('Registro|N°|No|Nro', case=False, na=False) &
        ~conta['fecha_mov'].astype(str).str.contains('Fecha', case=False, na=False)
    ]
    conta['fecha_mov']          = pd.to_datetime(conta['fecha_mov'],          dayfirst=True, errors='coerce')
    conta['fecha_conciliacion'] = pd.to_datetime(conta['fecha_conciliacion'], dayfirst=True, errors='coerce')
    for col in ['ingreso', 'egreso']:
        conta[col] = pd.to_numeric(
            conta[col].astype(str).str.replace(',', '', regex=False).str.strip(),
            errors='coerce'
        ).fillna(0)
    conta['conciliado'] = conta['conciliado'].str.upper().str.strip().map({'SI': True, 'NO': False})
    for col in ['giro', 'glosa', 'medio_pago']:
        conta[col] = conta[col].str.strip()
    conta = conta.sort_values('fecha_mov').reset_index(drop=True)

    # ── Generar reporte ───────────────────────────────────────────────
    generar_reporte_inicial(
        bank,
        conta,
        empresa=config.get('empresa', 'Southern Textil'),
        moneda=config['nombre'],
    )


# ══════════════════════════════════════════════════════════
# FLUJO — REPORTE FINAL
# ══════════════════════════════════════════════════════════

def flujo_reporte_final() -> None:
    """Carga el Excel trabajado por el especialista y genera el reporte final."""
    from reporte_final import generar_reporte_final

    print("-" * 60)
    print("Cargue el Excel de conciliacion ya trabajado por el especialista.")
    print("(Es el archivo 'Conciliacion_Inicial_...' con la hoja Anexar1 completada)\n")

    file_inicial = pedir_archivo("Ruta del Excel trabajado (.xlsx): ")
    print()

    try:
        generar_reporte_final(file_inicial)
    except ValueError as e:
        print(f"\n  X Error: {e}\n")


# ══════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════

config = seleccionar_empresa()
accion = seleccionar_accion()

if accion == "1":
    flujo_reporte_inicial(config)
elif accion == "2":
    flujo_reporte_final()
