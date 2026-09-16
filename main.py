import pandas as pd
from pathlib import Path
import re as _re_main
import unicodedata

# ══════════════════════════════════════════════════════════
# UTILIDADES DE NOMBRE DE ARCHIVO
# ══════════════════════════════════════════════════════════

def _limpiar_nombre(texto: str) -> str:
    """Elimina tildes, caracteres especiales y espacios para usar en nombres de archivo."""
    # Normalizar unicode (quitar tildes)
    normalizado = unicodedata.normalize('NFD', texto)
    sin_tildes  = ''.join(c for c in normalizado if unicodedata.category(c) != 'Mn')
    # Reemplazar caracteres no alfanuméricos (excepto guion) por guion bajo
    limpio = _re_main.sub(r'[^\w\-]', '_', sin_tildes)
    # Colapsar guiones bajos múltiples
    limpio = _re_main.sub(r'_+', '_', limpio).strip('_')
    return limpio


def _extraer_nombre_banco(file_conta: Path, skip_conta: int) -> str:
    """
    Lee las filas de cabecera del archivo de contabilidad (antes del skip)
    y extrae el nombre corto del banco.
    Contanet exporta en las primeras filas algo como:
      'Banco: BANCO DE CREDITO M.N. 194-1162203-0-23'
    Devuelve solo el nombre (ej. 'BCP', 'BBVA', 'SCOTIABANK', etc.)
    o 'BANCO' si no lo puede determinar.
    """
    try:
        import pandas as pd
        header_raw = pd.read_excel(file_conta, header=None, nrows=skip_conta)
        for _, row in header_raw.iterrows():
            for val in row:
                texto = str(val) if val is not None else ''
                if 'banco' in texto.lower() and ':' in texto:
                    contenido = texto.split(':', 1)[1].strip().upper()
                    # Mapear nombres conocidos a abreviatura
                    _MAP_BANCOS = {
                        'CREDITO'   : 'BCP',
                        'BCP'       : 'BCP',
                        'BBVA'      : 'BBVA',
                        'CONTINENTAL': 'BBVA',
                        'SCOTIABANK': 'SCOTIABANK',
                        'SCOTIAN'   : 'SCOTIABANK',
                        'INTERBANK' : 'INTERBANK',
                        'IBK'       : 'INTERBANK',
                        'PICHINCHA' : 'PICHINCHA',
                        'BANBIF'    : 'BANBIF',
                        'MIBANCO'   : 'MIBANCO',
                        'GNB'       : 'GNB',
                        'FALABELLA' : 'FALABELLA',
                        'RIPLEY'    : 'RIPLEY',
                        'ALFIN'     : 'ALFIN',
                        'CITIBANK'  : 'CITIBANK',
                        'NACION'    : 'BN',
                    }
                    for clave, abrev in _MAP_BANCOS.items():
                        if clave in contenido:
                            return abrev
                    # Si no hay coincidencia, tomar las primeras 3 palabras
                    palabras = [p for p in contenido.split() if len(p) > 2]
                    return '_'.join(palabras[:2]) if palabras else 'BANCO'
    except Exception:
        pass
    return 'BANCO'


def _construir_nombre_archivo(
    prefijo: str,           # 'CBI' o 'CBF'
    empresa: str,
    banco: str,
    moneda: str,            # 'Soles (PEN)' o 'Dolares (USD)'
    mes_anio: str,          # 'ENERO_2026' ya formateado
) -> Path:
    """
    Construye el nombre de archivo con el formato:
      CBI_[empresa]_[banco]_[moneda]_[mes]_[año]_[YYYYMMDD].xlsx
    """
    from datetime import datetime
    # Normalizar moneda → 'Soles' o 'Dolares'
    mon_limpia = 'Dolares' if 'USD' in moneda.upper() or 'DOL' in moneda.upper() else 'Soles'
    emp_limpia = _limpiar_nombre(empresa)
    ban_limpio = _limpiar_nombre(banco)
    fecha_hoy  = datetime.now().strftime('%Y%m%d')
    nombre = f"{prefijo}_{emp_limpia}_{ban_limpio}_{mon_limpia}_{mes_anio}_{fecha_hoy}.xlsx"
    return Path(nombre)


def _mes_anio_desde_bank(bank_df) -> str:
    """Extrae el mes y año del DataFrame del banco ya sanitizado (columna 'fecha')."""
    try:
        import pandas as pd
        fechas = pd.to_datetime(bank_df['fecha'], errors='coerce').dropna()
        if not fechas.empty:
            primera = fechas.iloc[0]
            meses = {
                1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
                5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
                9: 'SETIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
            }
            return f"{meses.get(primera.month, 'MES')}_{primera.year}"
    except Exception:
        pass
    return 'MES_ANIO'


def _mes_anio_desde_inicial(ruta_inicial: Path) -> str:
    """Extrae mes y año desde el nombre del archivo inicial o desde su hoja Anexar1."""
    try:
        import pandas as pd
        xl = pd.ExcelFile(ruta_inicial)
        if 'Anexar1' in xl.sheet_names:
            df = pd.read_excel(ruta_inicial, sheet_name='Anexar1', usecols=['Fecha'], nrows=20)
            fechas = pd.to_datetime(df['Fecha'], errors='coerce').dropna()
            if not fechas.empty:
                primera = fechas.iloc[0]
                meses = {
                    1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
                    5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
                    9: 'SETIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
                }
                return f"{meses.get(primera.month, 'MES')}_{primera.year}"
    except Exception:
        pass
    return 'MES_ANIO'


# ══════════════════════════════════════════════════════════
# CONFIGURACIÓN POR EMPRESA Y MONEDA
# ══════════════════════════════════════════════════════════

EMPRESAS = [
    {
        "nombre"  : "STN",
        "ruc"     : "20376729126",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": ("STN DOL", "STN USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True,
             "cuentas": {"BCP": "194-1162203-0-23", "SCOTIABANK": "001-0137451"}},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("STN SOL", "STN SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True,
             "cuentas": {}},
        ],
    },
    {
        "nombre"  : "ITS",
        "ruc"     : "20601910603",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": ("ITS DOL", "ITS USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("ITS SOL", "ITS SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "CMT",
        "ruc"     : "20537658471",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": ("CMT DOL", "CMT USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("CMT SOL", "CMT SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True,
             "cuentas": {"BN": "00-000-513598"}},
        ],
    },
    {
        "nombre"  : "Dynamitex",
        "ruc"     : "20600995761",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": ("DYNAMITEX DOL", "DYNAMITEX USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("DYNAMITEX SOL", "DYNAMITEX SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "DINSURA",
        "ruc"     : "20603964571",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": ("DINSURA DOL", "DINSURA USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("DINSURA SOL", "DINSURA SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Perú Commerce",
        "ruc"     : "20601234567",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": ("P.COMMERCE", "P.COMMERCE SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "INFOSUR",
        "ruc"     : "20600567890",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": ("INFOSUR DOL", "INFOSUR USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("INFOSUR SOL", "INFOSUR SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Thimble Sourcing / TST",
        "ruc"     : "20601987654",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": ("TST", "TST SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Iñaupari",
        "ruc"     : "20601345678",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": ("REF. IÑAPARI", "REF. IÑAPARI SOLES"),  "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "TECA",
        "ruc"     : "20601456789",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": ("TECA", "TECA SOLES"),        "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "DIONISO",
        "ruc"     : "20601567890",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": ("INV. DIONISO", "INV. DIONISO SOLES"),     "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
]


# ══════════════════════════════════════════════════════════
# FUNCIONES DE MENÚ Y CARGA DE ARCHIVOS
# ══════════════════════════════════════════════════════════

BANCOS_SOPORTADOS = [
    {"id": "BCP",        "nombre": "BCP (Excel .xlsx)",                               "formato": ".xlsx"},
    {"id": "SCOTIABANK", "nombre": "Scotiabank (PDF .pdf)",                           "formato": ".pdf"},
    {"id": "BN",         "nombre": "Banco de la Nación (PDF .pdf / Excel .xlsx)",     "formato": ".pdf"},
]


def seleccionar_banco() -> dict | None:
    """Muestra el menú de bancos soportados; devuelve el banco elegido o None si se cancela."""
    while True:
        print("\nSeleccione el BANCO:\n")
        for k, b in enumerate(BANCOS_SOPORTADOS, 1):
            print(f"  {k}. {b['nombre']}")
        print("  0. Volver")
        print()
        opcion = input("Opcion: ").strip()
        if opcion == "0":
            return None
        if opcion.isdigit() and 1 <= int(opcion) <= len(BANCOS_SOPORTADOS):
            return BANCOS_SOPORTADOS[int(opcion) - 1]
        print("  X Opcion invalida, intente de nuevo.\n")


def seleccionar_empresa() -> dict:
    """Muestra el menú de empresas, monedas y bancos; devuelve la configuración elegida."""
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

            banco_sel = seleccionar_banco()
            if banco_sel is None:
                continue

            print(f"\n  >> {empresa['nombre']} - {moneda['nombre']} - {banco_sel['nombre']}\n")
            # Incluir nombre de empresa, RUC y banco en el config para el reporte
            return {
                **moneda,
                'empresa': empresa['nombre'],
                'ruc'    : empresa.get('ruc', ''),
                'banco'  : banco_sel['id'],
            }


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


def pedir_archivo_banco(banco: str) -> Path:
    """Solicita una ruta de archivo para el banco (.xlsx para BCP, .pdf o .xlsx para Scotia y BN)."""
    banco_upper = (banco or '').upper()
    if 'SCOTIA' in banco_upper:
        extensiones = ('.pdf', '.xlsx', '.xls')
        mensaje = "Ruta del estado de cuenta de SCOTIABANK (.pdf / .xlsx): "
        error_msg = "  X El archivo debe ser PDF (.pdf) o Excel (.xlsx)"
    elif 'BN' in banco_upper or 'NACION' in banco_upper:
        extensiones = ('.pdf', '.xlsx', '.xls')
        mensaje = "Ruta del estado de cuenta de BANCO DE LA NACIÓN (.pdf / .xlsx): "
        error_msg = "  X El archivo debe ser PDF (.pdf) o Excel (.xlsx)"
    else:
        extensiones = ('.xlsx', '.xls')
        mensaje = "Ruta del estado de cuenta de BCP (.xlsx): "
        error_msg = "  X El archivo debe ser Excel (.xlsx o .xls)"

    while True:
        ruta = input(mensaje).strip().strip('"').strip("'")
        path = Path(ruta)
        if not path.exists():
            print(f"  X No se encontro el archivo: {ruta}")
        elif path.suffix.lower() not in extensiones:
            print(error_msg)
        else:
            return path


def _cargar_banco_scotiabank(ruta_banco: Path) -> tuple[pd.DataFrame, float | None]:
    """
    Carga el estado de cuenta de Scotiabank desde PDF (o Excel como fallback).
    Devuelve (bank_df, saldo_final). bank_df tiene las columnas internas del motor.
    """
    if ruta_banco.suffix.lower() == '.pdf':
        from SCOTIABANK.extract_data_pdf import leer_pdf_scotiabank
        resultado = leer_pdf_scotiabank(ruta_banco)
        return resultado['movimientos'], resultado['saldo_final']
    else:
        from bancos import leer_estado_bancario
        df, _ = leer_estado_bancario(ruta_banco, 'SCOTIABANK')
        return df, None


def _cargar_banco_bn(ruta_banco: Path) -> tuple[pd.DataFrame, float | None]:
    """
    Carga el estado de cuenta de Banco de la Nación desde PDF (o Excel como fallback).
    Devuelve (bank_df, saldo_final). bank_df tiene las columnas internas del motor.
    """
    if ruta_banco.suffix.lower() == '.pdf':
        from BN.extract_data_bn import leer_pdf_bn
        resultado = leer_pdf_bn(ruta_banco)
        return resultado['movimientos'], resultado['saldo_final']
    else:
        from bancos import leer_estado_bancario
        df, _ = leer_estado_bancario(ruta_banco, 'BN')
        return df, None



def _cargar_banco_bcp(file_bank: Path, config: dict) -> pd.DataFrame:
    """Carga y sanitiza el estado de cuenta BCP desde un archivo Excel."""
    xls = pd.ExcelFile(file_bank)

    sheet_bank = config.get("sheet_bank", "")
    if isinstance(sheet_bank, (tuple, list)):
        hoja_encontrada = next(
            (h for h in sheet_bank if h in xls.sheet_names),
            None
        )
    else:
        hoja_encontrada = sheet_bank

    if hoja_encontrada is None:
        if xls.sheet_names:
            hoja_encontrada = xls.sheet_names[0]
        else:
            raise ValueError(f"No se encontró ninguna hoja válida en {file_bank}")

    bank_raw = pd.read_excel(
        file_bank,
        sheet_name=hoja_encontrada,
        skiprows=config.get("skip_bank", 4)
    )

    bank = bank_raw.copy()
    bank.columns = [
        'fecha', 'fecha_valuta', 'descripcion', 'monto', 'saldo',
        'sucursal', 'nro_operacion', 'hora', 'usuario', 'utc', 'referencia'
    ]
    bank['fecha']        = pd.to_datetime(bank['fecha'],        format='%d/%m/%Y', errors='coerce')
    bank['fecha_valuta'] = pd.to_datetime(bank['fecha_valuta'], format='%d/%m/%Y', errors='coerce')
    bank['ingreso']      = bank['monto'].clip(lower=0)
    bank['egreso']       = bank['monto'].clip(upper=0).abs()
    bank = bank.drop(columns=['fecha_valuta', 'referencia'], errors='ignore')
    bank['descripcion']  = bank['descripcion'].astype(str).str.strip()
    bank = bank.sort_values('fecha').reset_index(drop=True)
    return bank


# ══════════════════════════════════════════════════════════
# FLUJO — REPORTE INICIAL
# ══════════════════════════════════════════════════════════

def flujo_reporte_inicial(config: dict) -> None:
    """Carga, sanitiza y genera el reporte inicial de conciliacion."""
    from reporte import generar_reporte_inicial

    print("-" * 60)
    banco_config = config.get('banco', 'BCP')
    file_bank  = pedir_archivo_banco(banco_config)
    file_conta = pedir_archivo("Ruta del reporte de CONTABILIDAD (.xlsx)  : ")
    print()

    # ── Carga y sanitización banco ────────────────────────────────────
    if banco_config == 'SCOTIABANK':
        bank, saldo_banco_final = _cargar_banco_scotiabank(file_bank)
        banco_nombre = 'Scotia'
    elif banco_config == 'BN':
        bank, saldo_banco_final = _cargar_banco_bn(file_bank)
        banco_nombre = 'BN'
    else:
        bank = _cargar_banco_bcp(file_bank, config)
        saldo_banco_final = None
        banco_nombre = _extraer_nombre_banco(file_conta, config.get('skip_conta', 11))
        if banco_nombre in ('', 'BANCO'):
            banco_nombre = 'BCP'

    # ── Carga contabilidad ────────────────────────────────────────────
    conta_raw = pd.read_excel(file_conta, skiprows=config['skip_conta'])

    # ── Sanitización contabilidad ─────────────────────────────────────
    # Mapeo dinámico por nombre de columna (normalizado sin tildes/mayúsculas).
    # Soporta 11 columnas (formato estándar), 14 (con MAR/SUB/MONTO de Scotiabank)
    # o cualquier variante futura. Columnas faltantes se llenan con None.
    import unicodedata as _ud
    import re as _re

    def _norm_col(txt: str) -> str:
        s = str(txt).strip().lower()
        return ''.join(c for c in _ud.normalize('NFD', s) if _ud.category(c) != 'Mn')

    # Reglas de mapeo: (nombre_interno, predicado sobre nombre normalizado)
    _REGLAS_CONTA = [
        ('nro_registro',       lambda n: 'registro' in n and 'fecha' not in n),
        ('fecha_mov',          lambda n: 'movimiento' in n or ('f.' in n and 'mov' in n)),
        ('medio_pago',         lambda n: 'medio' in n or 'pago' in n),
        ('nro_operacion',      lambda n: 'operac' in n),
        ('giro',               lambda n: n == 'giro'),
        ('glosa',              lambda n: 'glosa' in n or 'concepto' in n),
        ('ingreso',            lambda n: n == 'ingreso'),
        ('egreso',             lambda n: n == 'egreso'),
        ('fecha_conciliacion', lambda n: 'concilia' in n and ('f.' in n or 'fecha' in n)),
        ('conciliado',         lambda n: 'conciliad' in n),
    ]

    col_map: dict[str, str] = {}   # nombre_interno → nombre_real_en_df
    for col_real in conta_raw.columns:
        norm = _norm_col(str(col_real))
        for nombre_interno, predicado in _REGLAS_CONTA:
            if nombre_interno not in col_map and predicado(norm):
                col_map[nombre_interno] = col_real
                break

    # Construir DataFrame con columnas internas; las que falten quedan en None
    # (no bloquean la conciliación)
    import pandas as _pd_local
    conta = _pd_local.DataFrame(index=conta_raw.index)
    for nombre_interno, _ in _REGLAS_CONTA:
        if nombre_interno in col_map:
            conta[nombre_interno] = conta_raw[col_map[nombre_interno]]
        else:
            conta[nombre_interno] = None

    # ── Extraer Saldo Contable Final ANTES de filtrar filas de saldo ──
    saldo_contable_final = None

    # Patrón 1: "Saldo Contable (Final): XXXXX" en la columna nro_registro
    mask_saldo_texto = conta['nro_registro'].astype(str).str.contains(
        r'Saldo\s+Contable\s*\(Final\)', case=False, na=False
    )
    filas_saldo_texto = conta[mask_saldo_texto]
    if not filas_saldo_texto.empty:
        try:
            texto = str(filas_saldo_texto['nro_registro'].iloc[0])
            m = _re.search(r':\s*([\d,\.]+)', texto)
            if m:
                saldo_contable_final = float(m.group(1).replace(',', ''))
        except Exception:
            pass

    # Patrón 1b: buscar en cualquier columna del raw (cuando nro_registro no mapea)
    if saldo_contable_final is None:
        for _col_raw in conta_raw.columns:
            _mask = conta_raw[_col_raw].astype(str).str.contains(
                r'Saldo\s+Contable\s*\(Final\)', case=False, na=False
            )
            _filas = conta_raw[_mask]
            if not _filas.empty:
                try:
                    texto = str(_filas[_col_raw].iloc[0])
                    m = _re.search(r':\s*([\d,\.]+)', texto)
                    if m:
                        saldo_contable_final = float(m.group(1).replace(',', ''))
                        break
                except Exception:
                    pass

    # Patrón 2 (fallback): "Saldo" en nro_registro y "Final" en fecha_mov (formato antiguo)
    if saldo_contable_final is None:
        mask_saldo_final = (
            conta['nro_registro'].astype(str).str.contains('Saldo', case=False, na=False) &
            conta['fecha_mov'].astype(str).str.contains('Final', case=False, na=False)
        )
        filas_saldo = conta[mask_saldo_final]
        if not filas_saldo.empty:
            try:
                val_ing = pd.to_numeric(
                    filas_saldo['ingreso'].astype(str).str.replace(',', '', regex=False).str.strip(),
                    errors='coerce'
                ).fillna(0).iloc[0]
                val_eg = pd.to_numeric(
                    filas_saldo['egreso'].astype(str).str.replace(',', '', regex=False).str.strip(),
                    errors='coerce'
                ).fillna(0).iloc[0]
                saldo_contable_final = float(val_ing - val_eg) if (val_ing != 0 or val_eg != 0) else None
            except Exception:
                saldo_contable_final = None

    conta = conta[
        conta['nro_registro'].notna() &
        ~conta['nro_registro'].astype(str).str.startswith('Saldo', na=False) &
        ~conta['nro_registro'].astype(str).str.contains('Registro|N°|No|Nro', case=False, na=False) &
        ~conta['nro_registro'].astype(str).str.contains(r'Informaci[oó]n\s+anterior', case=False, na=False) &
        ~conta['fecha_mov'].astype(str).str.contains('Fecha', case=False, na=False)
    ]
    conta['fecha_mov']          = pd.to_datetime(conta['fecha_mov'],          dayfirst=True, errors='coerce')
    conta['fecha_conciliacion'] = pd.to_datetime(conta['fecha_conciliacion'], dayfirst=True, errors='coerce')
    for col in ['ingreso', 'egreso']:
        conta[col] = pd.to_numeric(
            conta[col].astype(str).str.replace(',', '', regex=False).str.strip(),
            errors='coerce'
        ).fillna(0)
    conta['conciliado'] = conta['conciliado'].astype(str).str.upper().str.strip().map(
        {'SI': True, 'NO': False}
    ).fillna(False)
    for col in ['giro', 'glosa', 'medio_pago', 'nro_operacion', 'nro_registro']:
        conta[col] = conta[col].fillna('').astype(str).str.strip()

    # ── Filtrar movimientos de otros bancos si aparecen en Giro ───────
    # Si la contabilidad incluye explícitamente otros bancos en la columna Giro,
    # se omiten para evitar cruces indebidos, conservando siempre a proveedores/terceros.
    try:
        from bancos import BANCOS, _contiene_alias, clave_banco
        clave_act = clave_banco(banco_nombre)
        otros_bancos = [cb for cb in BANCOS if cb != clave_act]
        mascara_otro_banco = conta['giro'].map(
            lambda val: any(_contiene_alias(val, ob) for ob in otros_bancos)
        )
        if mascara_otro_banco.any():
            descartados = int(mascara_otro_banco.sum())
            conta = conta[~mascara_otro_banco].reset_index(drop=True)
            print(f"  [INFO] Se omitieron {descartados} registro(s) de otros bancos en la columna 'Giro'.")
    except Exception:
        pass

    # ── Filtrar registros de meses anteriores ("Información anterior") ─────
    # El mes de referencia lo determinamos a partir del banco ya sanitizado.
    # Solo se conservan registros cuya fecha_mov pertenezca al mismo mes y año.
    try:
        fechas_banco = pd.to_datetime(bank['fecha'], errors='coerce').dropna()
        if not fechas_banco.empty:
            mes_ref  = fechas_banco.iloc[0].month
            anio_ref = fechas_banco.iloc[0].year
            filas_antes = len(conta)
            conta = conta[
                conta['fecha_mov'].isna() |  # mantener filas sin fecha (se descartan luego)
                (
                    (conta['fecha_mov'].dt.month == mes_ref) &
                    (conta['fecha_mov'].dt.year  == anio_ref)
                )
            ]
            filas_descartadas = filas_antes - len(conta)
            if filas_descartadas > 0:
                print(f"  [INFO] Se omitieron {filas_descartadas} registro(s) de contabilidad "
                      f"fuera del período {mes_ref:02d}/{anio_ref} (sección 'Información anterior').")
    except Exception:
        pass  # Si falla el filtro por mes, continuar sin él

    conta = conta.sort_values('fecha_mov').reset_index(drop=True)

    # ── Cargar movimientos pendientes de la conciliación anterior ─────
    # Calcular mes y año de referencia ANTES de inyectar pendientes para que
    # el nombre del archivo use el mes correcto (no el de los pendientes inyectados).
    mes_ref  = None
    anio_ref = None
    try:
        fechas_banco_ref = pd.to_datetime(bank['fecha'], errors='coerce').dropna()
        if not fechas_banco_ref.empty:
            mes_ref  = int(fechas_banco_ref.iloc[0].month)
            anio_ref = int(fechas_banco_ref.iloc[0].year)
    except Exception:
        pass

    try:
        from pendientes import cargar_pendientes, inyectar_pendientes_en_banco, inyectar_pendientes_en_conta

        # mes_ref / anio_ref ya calculados arriba

        df_banco_pend, df_conta_pend = cargar_pendientes(
            config.get('empresa', ''),
            banco=banco_nombre,
            moneda=config.get('nombre', ''),
            mes_ref=mes_ref,
            anio_ref=anio_ref,
        )
        bank  = inyectar_pendientes_en_banco(bank,  df_banco_pend, mes_ref=mes_ref, anio_ref=anio_ref)
        conta = inyectar_pendientes_en_conta(conta, df_conta_pend, mes_ref=mes_ref, anio_ref=anio_ref)
    except Exception as e_pend:
        print(f"  [AVISO] No se pudieron cargar los pendientes anteriores: {e_pend}")

    # ── Generar reporte ───────────────────────────────────────────────
    # NOTA: el mes/año se determina a partir de los datos originales del banco
    # (antes de inyectar pendientes), usando mes_ref/anio_ref ya calculados arriba.
    # Si se usara _mes_anio_desde_bank(bank) aquí, los pendientes del mes anterior
    # (ordenados por fecha) harían que iloc[0] apunte a un mes incorrecto.
    if mes_ref is not None and anio_ref is not None:
        _meses = {
            1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
            5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
            9: 'SETIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
        }
        mes_anio = f"{_meses.get(mes_ref, 'MES')}_{anio_ref}"
    else:
        mes_anio = _mes_anio_desde_bank(bank)
    ruta_salida  = _construir_nombre_archivo(
        'CBI',
        empresa=config.get('empresa', 'Empresa'),
        banco=banco_nombre,
        moneda=config['nombre'],
        mes_anio=mes_anio,
    )
    # Obtener número de cuenta según banco seleccionado
    cuenta_num = config.get('cuentas', {}).get(banco_config, '')
    generar_reporte_inicial(
        bank,
        conta,
        ruta_salida=ruta_salida,
        empresa=config.get('empresa', 'Southern Textil'),
        moneda=config['nombre'],
        saldo_contable_final=saldo_contable_final,
        banco=banco_nombre,
        cuenta=cuenta_num,
    )


# ══════════════════════════════════════════════════════════
# FLUJO — REPORTE FINAL
# ══════════════════════════════════════════════════════════

def flujo_reporte_final(config: dict) -> None:
    """Carga el Excel trabajado por el especialista y genera el reporte final."""
    from reporte_final import generar_reporte_final

    print("-" * 60)
    print("Cargue el Excel de conciliacion ya trabajado por el especialista.")
    print("(Es el archivo 'CBI_...' con la hoja Anexar1 completada)\n")

    file_inicial = pedir_archivo("Ruta del Excel trabajado (.xlsx): ")
    print()

    # Extraer banco y mes/año desde config o desde el archivo inicial para construir el nombre CBF
    banco_nombre = config.get('banco') or 'BANCO'
    try:
        from bancos import nombre_corto_banco
        if banco_nombre in ('SCOTIABANK', 'BCP', 'BN'):
            banco_nombre = nombre_corto_banco(banco_nombre)
    except Exception:
        pass
    # Intentar leer banco desde la hoja CONTANET (metadato guardado en el reporte inicial)
    try:
        import openpyxl as _opxl
        _wb_tmp = _opxl.load_workbook(file_inicial, read_only=True, data_only=True)
        if 'CONTANET' in _wb_tmp.sheetnames:
            _ws_tmp = _wb_tmp['CONTANET']
            # La columna de banco está en la cabecera del archivo conta original,
            # pero aquí usamos el metadato de banco guardado si existe
            for col_idx in range(1, 40):
                etiq = _ws_tmp.cell(row=1, column=col_idx).value
                if str(etiq).strip() == '__BANCO__':
                    val = _ws_tmp.cell(row=2, column=col_idx).value
                    if val:
                        banco_nombre = str(val).strip()
                    break
        _wb_tmp.close()
    except Exception:
        pass

    if banco_nombre in ('', 'BANCO'):
        nombre_arch = file_inicial.name.upper()
        if 'BCP' in nombre_arch or 'CREDITO' in nombre_arch:
            banco_nombre = 'BCP'
        elif 'SCOTIA' in nombre_arch:
            banco_nombre = 'Scotia'
        elif 'BN' in nombre_arch or 'NACION' in nombre_arch:
            banco_nombre = 'BN'

    mes_anio = _mes_anio_desde_inicial(file_inicial)
    ruta_salida = _construir_nombre_archivo(
        'CBF',
        empresa=config.get('empresa', 'Empresa'),
        banco=banco_nombre,
        moneda=config.get('nombre', ''),
        mes_anio=mes_anio,
    )

    # Obtener número de cuenta desde el config (fallback si el CBI no tiene __CUENTA__)
    banco_id = config.get('banco', '').upper()  # ej. 'SCOTIABANK', 'BCP', 'BN'
    cuenta_config = config.get('cuentas', {}).get(banco_id, '')

    try:
        generar_reporte_final(
            file_inicial,
            ruta_salida=ruta_salida,
            empresa=config.get('empresa', ''),
            ruc=config.get('ruc', ''),
            moneda=config.get('nombre', ''),
            banco=banco_nombre,
            cuenta=cuenta_config,
        )
    except ValueError as e:
        print(f"\n  X Error: {e}\n")


# ══════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════


if __name__ == '__main__':
    config = seleccionar_empresa()
    accion = seleccionar_accion()

    if accion == "1":
        flujo_reporte_inicial(config)
    elif accion == "2":
        flujo_reporte_final(config)
