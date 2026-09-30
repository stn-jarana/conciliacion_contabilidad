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
            {"nombre": "Dolares (USD)", "sheet_bank": ("Thimble DOL", "Thimble USD"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": ("Thimble", "Thimble SOLES"), "skip_bank": 4, "skip_conta": 11, "habilitado": True},
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


def extraer_datos_conciliacion_inicial(ruta_archivo: Path | str) -> dict:
    """
    Extrae empresa, banco, moneda, mes y año a partir del nombre del archivo
    de conciliación inicial (ej: CBI_STN_Scotia_Dolares_JUNIO_2026_20260918.xlsx).
    También lee metadatos de la hoja CONTANET si están disponibles como respaldo.
    """
    path = Path(ruta_archivo)
    nombre = path.stem.strip()

    # 1. Mes y Año
    MESES = {
        'ENERO': 'ENERO', 'FEBRERO': 'FEBRERO', 'MARZO': 'MARZO', 'ABRIL': 'ABRIL',
        'MAYO': 'MAYO', 'JUNIO': 'JUNIO', 'JULIO': 'JULIO', 'AGOSTO': 'AGOSTO',
        'SETIEMBRE': 'SETIEMBRE', 'SEPTIEMBRE': 'SETIEMBRE', 'OCTUBRE': 'OCTUBRE',
        'NOVIEMBRE': 'NOVIEMBRE', 'DICIEMBRE': 'DICIEMBRE'
    }
    patron_meses = '|'.join(MESES.keys())
    m_mes_anio = _re_main.search(rf'(?:^|_)(?:MES_)?({patron_meses})_(\d{{4}})(?:_|$)', nombre, flags=_re_main.IGNORECASE)
    if m_mes_anio:
        mes_str = MESES[m_mes_anio.group(1).upper()]
        anio_str = m_mes_anio.group(2)
        mes_anio = f"{mes_str}_{anio_str}"
    else:
        mes_anio = _mes_anio_desde_inicial(path)

    # 2. Moneda
    if _re_main.search(r'(?:^|_)(?:DOLARES|USD|DOL)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        moneda = "Dolares (USD)"
        mon_tipo = "Dolares"
    elif _re_main.search(r'(?:^|_)(?:SOLES|PEN|SOL)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        moneda = "Soles (PEN)"
        mon_tipo = "Soles"
    else:
        if any(x in nombre.upper() for x in ('DOL', 'USD')):
            moneda = "Dolares (USD)"
            mon_tipo = "Dolares"
        else:
            moneda = "Soles (PEN)"
            mon_tipo = "Soles"

    # 3. Banco
    banco_nombre = ''
    banco_id = ''
    if _re_main.search(r'(?:^|_)(?:SCOTIABANK|SCOTIA)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'Scotia'
        banco_id = 'SCOTIABANK'
    elif _re_main.search(r'(?:^|_)(?:BN|NACION|BANCO_DE_LA_NACION)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'BN'
        banco_id = 'BN'
    elif _re_main.search(r'(?:^|_)(?:BCP|CREDITO)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'BCP'
        banco_id = 'BCP'
    elif _re_main.search(r'(?:^|_)(?:BBVA|CONTINENTAL)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'BBVA'
        banco_id = 'BBVA'
    elif _re_main.search(r'(?:^|_)(?:INTERBANK|IBK)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'INTERBANK'
        banco_id = 'INTERBANK'
    elif _re_main.search(r'(?:^|_)(?:BANBIF)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'BANBIF'
        banco_id = 'BANBIF'
    elif _re_main.search(r'(?:^|_)(?:PICHINCHA)(?:_|$)', nombre, flags=_re_main.IGNORECASE):
        banco_nombre = 'PICHINCHA'
        banco_id = 'PICHINCHA'

    # Metadatos del archivo Excel (si existe y tiene CONTANET)
    banco_leido = ''
    cuenta_leida = ''
    if path.exists() and path.suffix.lower() in ('.xlsx', '.xls'):
        try:
            import openpyxl as _opxl
            wb_meta = _opxl.load_workbook(path, read_only=True, data_only=True)
            if 'CONTANET' in wb_meta.sheetnames:
                ws_meta = wb_meta['CONTANET']
                headers_meta = {
                    str(ws_meta.cell(row=1, column=c).value).strip(): c
                    for c in range(1, 40)
                }
                if '__BANCO__' in headers_meta:
                    val = ws_meta.cell(row=2, column=headers_meta['__BANCO__']).value
                    if val:
                        banco_leido = str(val).strip()
                if '__CUENTA__' in headers_meta:
                    val = ws_meta.cell(row=2, column=headers_meta['__CUENTA__']).value
                    if val:
                        cuenta_leida = str(val).strip()
            wb_meta.close()
        except Exception:
            pass

    if not banco_nombre:
        if banco_leido:
            banco_nombre = banco_leido
        else:
            banco_nombre = 'BANCO'
    if not banco_id:
        try:
            from bancos import clave_banco
            banco_id = clave_banco(banco_nombre) if banco_nombre else 'BCP'
        except Exception:
            banco_id = 'BCP'

    # 4. Empresa
    empresa_token = ''
    patron_std = _re_main.match(
        rf'^(?:CBI|CBF)_(.+?)_(?:Scotia|Scotiabank|BCP|BN|Nacion|BBVA|Interbank|Banbif|Pichincha|Banco)_(?:Dolares|Soles|USD|PEN|DOL|SOL)_(?:{patron_meses})_(\d{{4}})',
        nombre,
        flags=_re_main.IGNORECASE
    )
    if patron_std:
        empresa_token = patron_std.group(1).strip()
    else:
        partes = nombre.split('_')
        if len(partes) >= 2 and partes[0].upper() in ('CBI', 'CBF'):
            empresa_token = partes[1]

    empresa_detectada = ''
    empresa_cfg = None

    for emp in EMPRESAS:
        nombre_limpio = _limpiar_nombre(emp['nombre']).upper()
        if empresa_token:
            token_upper = _limpiar_nombre(empresa_token).upper()
            if token_upper == nombre_limpio or token_upper in nombre_limpio or nombre_limpio in token_upper:
                empresa_cfg = emp
                empresa_detectada = emp['nombre']
                break
        if nombre_limpio in nombre.upper():
            empresa_cfg = emp
            empresa_detectada = emp['nombre']
            break

    if not empresa_detectada:
        empresa_detectada = empresa_token if empresa_token else 'Empresa'

    # 5. RUC y Cuenta
    ruc = ''
    cuentas = {}
    if empresa_cfg:
        ruc = empresa_cfg.get('ruc', '')
        mon_cfg = next(
            (m for m in empresa_cfg.get('monedas', []) if mon_tipo.upper() in m['nombre'].upper()),
            empresa_cfg.get('monedas', [{}])[0] if empresa_cfg.get('monedas') else {}
        )
        cuentas = mon_cfg.get('cuentas', {})

    cuenta = cuenta_leida or cuentas.get(banco_id, cuentas.get(banco_nombre, ''))

    return {
        'empresa': empresa_detectada,
        'ruc': ruc,
        'banco_nombre': banco_nombre,
        'banco_id': banco_id,
        'moneda': moneda,
        'mes_anio': mes_anio,
        'cuenta': cuenta,
        'cuentas': cuentas,
    }


def seleccionar_empresa() -> dict | None:
    """Muestra el menú de empresas, monedas y bancos; devuelve la configuración elegida o None si cancela."""
    while True:
        print("-" * 60)
        print("  CONCILIACION INICIAL - SELECCION DE EMPRESA")
        print("-" * 60)
        print("\nSeleccione la empresa:\n")
        for i, emp in enumerate(EMPRESAS, 1):
            print(f"  {i}. {emp['nombre']}")
        print("  0. Volver")
        print()
        opcion = input("Opcion: ").strip()

        if opcion == "0":
            return None

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
    """Pregunta qué desea hacer: conciliación inicial o conciliación final."""
    while True:
        print("=" * 60)
        print("  CONCILIACION BANCARIA")
        print("=" * 60)
        print("  Que desea hacer?\n")
        print("  1. Generar reporte inicial")
        print("     (selecciona empresa, moneda y banco para generar el Excel de trabajo)")
        print()
        print("  2. Generar reporte final")
        print("     (carga el Excel inicial trabajado; detecta empresa, banco, moneda y mes)")
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


def _extraer_cuenta_bcp(file_bank: Path, config: dict) -> str:
    """Lee las filas de cabecera del Excel BCP (antes del skip_bank) y extrae
    el número de cuenta corriente.

    El extracto BCP incluye en sus primeras filas una celda con el formato
    'N° Cuenta: 194-XXXXXXX-X-XX' o simplemente el número desnudo (p.ej.
    '194-1162203-0-23'). La función busca ambos patrones.
    Devuelve la cuenta como string o '' si no la encuentra.
    """
    import re as _re_bcp
    try:
        xls = pd.ExcelFile(file_bank)
        sheet_bank = config.get("sheet_bank", "")
        if isinstance(sheet_bank, (tuple, list)):
            hoja = next((h for h in sheet_bank if h in xls.sheet_names), None)
        else:
            hoja = sheet_bank if sheet_bank in xls.sheet_names else None
        if hoja is None and xls.sheet_names:
            hoja = xls.sheet_names[0]
        if hoja is None:
            return ''

        n_skip = config.get("skip_bank", 4)
        # Leer solo las filas de cabecera sin skip
        cabecera_df = pd.read_excel(file_bank, sheet_name=hoja, header=None, nrows=n_skip)

        # Patrón de número de cuenta BCP: dígitos separados por guiones, ej. 194-1162203-0-23
        _pat_cuenta = _re_bcp.compile(r'\b(\d{3}-\d{5,}-\d-\d{2})\b')

        for _, fila in cabecera_df.iterrows():
            for celda in fila:
                texto = str(celda).strip() if celda is not None else ''
                if not texto or texto.lower() in ('nan', 'none'):
                    continue
                m = _pat_cuenta.search(texto)
                if m:
                    return m.group(1)
                # También buscar después de etiquetas como "N° Cuenta:", "Cuenta:", "Cta:"
                m2 = _re_bcp.search(
                    r'(?:N[°º]?\s*Cuenta|Cuenta|Cta)[^:]*:\s*([\d\-]+)',
                    texto, _re_bcp.IGNORECASE
                )
                if m2:
                    return m2.group(1).strip()
    except Exception:
        pass
    return ''


def _cargar_banco_scotiabank(ruta_banco: Path) -> tuple[pd.DataFrame, float | None, str]:
    """
    Carga el estado de cuenta de Scotiabank desde PDF (o Excel como fallback).
    Devuelve (bank_df, saldo_final, cuenta). bank_df tiene las columnas internas del motor.
    """
    if ruta_banco.suffix.lower() == '.pdf':
        from SCOTIABANK.extract_data_pdf import leer_pdf_scotiabank
        resultado = leer_pdf_scotiabank(ruta_banco)
        return resultado['movimientos'], resultado['saldo_final'], resultado.get('cuenta', '')
    else:
        from bancos import leer_estado_bancario
        df, _ = leer_estado_bancario(ruta_banco, 'SCOTIABANK')
        return df, None, ''


def _cargar_banco_bn(ruta_banco: Path) -> tuple[pd.DataFrame, float | None, str]:
    """
    Carga el estado de cuenta de Banco de la Nación desde PDF (o Excel como fallback).
    Devuelve (bank_df, saldo_final, cuenta). bank_df tiene las columnas internas del motor.
    """
    if ruta_banco.suffix.lower() == '.pdf':
        from BN.extract_data_bn import leer_pdf_bn
        resultado = leer_pdf_bn(ruta_banco)
        return resultado['movimientos'], resultado['saldo_final'], resultado.get('cuenta', '')
    else:
        from bancos import leer_estado_bancario
        df, _ = leer_estado_bancario(ruta_banco, 'BN')
        return df, None, ''



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
        bank, saldo_banco_final, cuenta_banco = _cargar_banco_scotiabank(file_bank)
        banco_nombre = 'Scotia'
    elif banco_config == 'BN':
        bank, saldo_banco_final, cuenta_banco = _cargar_banco_bn(file_bank)
        banco_nombre = 'BN'
    else:
        bank = _cargar_banco_bcp(file_bank, config)
        saldo_banco_final = None
        cuenta_banco = _extraer_cuenta_bcp(file_bank, config)
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
    # Número de cuenta: se obtiene directamente del extracto bancario.
    # Como fallback se consulta la configuración estática (EMPRESAS) para compatibilidad.
    cuenta_num = cuenta_banco or config.get('cuentas', {}).get(banco_config, '')
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
# GENERACIÓN DE ASIENTO ITF Y COMISIONES
# ══════════════════════════════════════════════════════════

# Mapa de plantillas por banco (clave normalizada) y moneda ('Soles'/'Dolares').
# Ajustar las rutas si las plantillas están en un subdirectorio distinto.
_PLANTILLAS_ASIENTO: dict[tuple[str, str], str] = {
    ("BCP",        "Dolares"): "Asiento_ITF_DOL_1.xlsm",
    ("BCP",        "Soles"):   "Asiento_ITF_SOL_1.xlsm",
    ("SCOTIABANK", "Dolares"): "Asiento_ITF_DOL_1.xlsm",
    ("SCOTIABANK", "Soles"):   "Asiento_ITF_SOL_1.xlsm",
    ("Scotia",     "Dolares"): "Asiento_ITF_DOL_1.xlsm",
    ("Scotia",     "Soles"):   "Asiento_ITF_SOL_1.xlsm",
    ("BN",         "Soles"):   "Asiento_ITF_SOL_1.xlsm",
}

# Meses en español tal como aparecen en el nombre del CBF
_MESES_NUM: dict[str, int] = {
    'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4,
    'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8,
    'SETIEMBRE': 9, 'SEPTIEMBRE': 9, 'OCTUBRE': 10,
    'NOVIEMBRE': 11, 'DICIEMBRE': 12,
}


def _generar_asiento_itf_tras_cbf(
    ruta_cbf: Path,
    banco: str,
    moneda: str,
) -> None:
    """Consulta el TC Venta del último día del mes y genera el asiento ITF/comisiones.

    Se invoca automáticamente después de crear el archivo CBF (conciliación final).
    El nombre del asiento de salida es ``Asiento_ITF_Comis_<stem_del_cbf>.xlsm``,
    lo que permite identificar de qué conciliación (banco, moneda, mes) proviene.

    Parameters
    ----------
    ruta_cbf:
        Ruta del archivo CBF recién creado.
    banco:
        Nombre o ID del banco (p.ej. "BCP", "Scotia", "BN").
    moneda:
        Moneda tal como viene de la configuración (p.ej. "Soles (PEN)", "Dolares (USD)").
    """
    import re as _re_itf

    print()
    print("─" * 60)
    print("  GENERANDO ASIENTO ITF Y COMISIONES")
    print("─" * 60)

    # ── 1. Determinar moneda corta ────────────────────────────────
    moneda_upper = moneda.upper()
    if any(x in moneda_upper for x in ("DOL", "USD")):
        mon_tipo = "Dolares"
    else:
        mon_tipo = "Soles"

    # ── 2. Determinar mes y año desde el nombre del CBF ───────────
    stem = ruta_cbf.stem.upper()  # e.g. "CBF_STN_BCP_DOLARES_JULIO_2026_20260925"
    patron_meses = "|".join(_MESES_NUM.keys())
    m = _re_itf.search(rf'({patron_meses})_(\d{{4}})', stem)
    if m:
        nombre_mes = m.group(1)
        anio = int(m.group(2))
        mes = _MESES_NUM.get(nombre_mes, 0)
    else:
        # Fallback: leer la primera fecha de la hoja ITF y COM del propio CBF
        try:
            from Asiento_ITF_Comis import leer_movimientos_itf_comisiones
            movimientos = leer_movimientos_itf_comisiones(ruta_cbf)
            primera_fecha = movimientos[0].fecha
            mes = primera_fecha.month
            anio = primera_fecha.year
        except Exception as e_fallback:
            print(f"  [AVISO] No se pudo determinar el mes/año del CBF: {e_fallback}")
            print("  Se omite la generación del asiento ITF y comisiones.")
            return

    # ── 3. Consultar TC Venta del último día del mes ──────────────
    from calendar import monthrange
    from datetime import date
    ultimo_dia = date(anio, mes, monthrange(anio, mes)[1])
    print(f"  >> Consultando TC Venta de Contanet para el {ultimo_dia.strftime('%d/%m/%Y')}...")

    try:
        from operaciones_sql import obtener_tipo_cambio
        tc_venta = obtener_tipo_cambio(ultimo_dia)
        print(f"  >> TC Venta obtenido: {tc_venta}")
    except Exception as e_tc:
        print(f"  [ERROR] No se pudo obtener el TC Venta de Contanet: {e_tc}")
        print("  Se omite la generación del asiento ITF y comisiones.")
        return

    # ── 4. Seleccionar plantilla según banco y moneda ─────────────
    # Normalizar el banco para buscar en el diccionario de plantillas
    banco_key = banco.upper().strip()
    plantilla_nombre = (
        _PLANTILLAS_ASIENTO.get((banco_key, mon_tipo))
        or _PLANTILLAS_ASIENTO.get((banco, mon_tipo))
        or _PLANTILLAS_ASIENTO.get(("BCP", mon_tipo))  # fallback genérico
    )
    ruta_plantilla = Path(plantilla_nombre)
    if not ruta_plantilla.is_file():
        print(f"  [ERROR] Plantilla no encontrada: {ruta_plantilla}")
        print("  Se omite la generación del asiento ITF y comisiones.")
        return

    # ── 5. Construir ruta de salida con el mismo stem del CBF ─────
    #   Asiento_ITF_Comis_CBF_STN_BCP_Dolares_JULIO_2026_20260925.xlsm
    nombre_salida = f"Asiento_ITF_Comis_{ruta_cbf.stem}{ruta_plantilla.suffix}"
    ruta_salida_asiento = ruta_cbf.parent / nombre_salida

    # ── 6. Determinar banco normalizado para Asiento_ITF_Comis ────
    # Asiento_ITF_Comis acepta "BCP", "Scotia" o "BCP Miami"
    banco_asiento = banco

    # ── 7. Generar el asiento ─────────────────────────────────────
    try:
        from Asiento_ITF_Comis import generar_asientos_itf_comisiones
        salida = generar_asientos_itf_comisiones(
            ruta_cbf,
            ruta_plantilla,
            banco=banco_asiento,
            moneda=mon_tipo,
            tipo_cambio_venta=tc_venta,
            ruta_salida=ruta_salida_asiento,
        )
        print(f"  >> Asiento ITF y comisiones generado: {salida}")
    except Exception as e_asiento:
        print(f"  [ERROR] No se pudo generar el asiento: {e_asiento}")


# ══════════════════════════════════════════════════════════
# FLUJO — REPORTE FINAL
# ══════════════════════════════════════════════════════════

def flujo_reporte_final(config: dict | None = None) -> None:
    """Carga el Excel trabajado por el especialista y genera el reporte final."""
    from reporte_final import generar_reporte_final

    print("-" * 60)
    print("  CONCILIACION FINAL")
    print("-" * 60)
    print("Cargue el Excel de conciliacion ya trabajado por el especialista.")
    print("(Es el archivo 'CBI_...' con la hoja Anexar1 completada)\n")

    file_inicial = pedir_archivo("Ruta del Excel inicial trabajado (.xlsx): ")
    print()

    datos = extraer_datos_conciliacion_inicial(file_inicial)

    if config:
        empresa = config.get('empresa') or datos['empresa']
        ruc = config.get('ruc') or datos['ruc']
        moneda = config.get('nombre') or datos['moneda']
        banco_nombre = datos['banco_nombre'] if datos['banco_nombre'] not in ('', 'BANCO') else (config.get('banco') or 'BANCO')
        banco_id = datos.get('banco_id') or config.get('banco', '').upper() or 'BCP'
        cuentas = config.get('cuentas') or datos.get('cuentas', {})
    else:
        empresa = datos['empresa']
        ruc = datos['ruc']
        moneda = datos['moneda']
        banco_nombre = datos['banco_nombre']
        banco_id = datos['banco_id']
        cuentas = datos.get('cuentas', {})

    try:
        from bancos import nombre_corto_banco
        if banco_nombre in ('SCOTIABANK', 'BCP', 'BN'):
            banco_nombre = nombre_corto_banco(banco_nombre)
    except Exception:
        pass

    mes_anio = datos['mes_anio']
    cuenta_final = datos.get('cuenta') or cuentas.get(banco_id, '')

    print("  >> Datos detectados del archivo inicial:")
    print(f"     - Empresa : {empresa}")
    print(f"     - Banco   : {banco_nombre}")
    print(f"     - Moneda  : {moneda}")
    print(f"     - Periodo : {mes_anio}")
    if cuenta_final:
        print(f"     - Cuenta  : {cuenta_final}")
    print()

    ruta_salida = _construir_nombre_archivo(
        'CBF',
        empresa=empresa,
        banco=banco_nombre,
        moneda=moneda,
        mes_anio=mes_anio,
    )

    try:
        generar_reporte_final(
            file_inicial,
            ruta_salida=ruta_salida,
            empresa=empresa,
            ruc=ruc,
            moneda=moneda,
            banco=banco_nombre,
            cuenta=cuenta_final,
        )
    except ValueError as e:
        print(f"\n  X Error: {e}\n")
        return

    # ── Generar asiento ITF y comisiones tras la conciliación final ────
    _generar_asiento_itf_tras_cbf(
        ruta_cbf=ruta_salida,
        banco=banco_nombre,
        moneda=moneda,
    )


# ══════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ══════════════════════════════════════════════════════════


if __name__ == '__main__':
    while True:
        accion = seleccionar_accion()

        if accion == "1":
            config = seleccionar_empresa()
            if config is not None:
                flujo_reporte_inicial(config)
                break
        elif accion == "2":
            flujo_reporte_final()
            break
