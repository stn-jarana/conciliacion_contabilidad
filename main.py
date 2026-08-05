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
        "nombre"  : "Southern Textil Network (STN)",
        "ruc"     : "20376729126",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "STN DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "STN SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Integrated Textile Solutions (ITS)",
        "ruc"     : "20601910603",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "ITS DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "ITS SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "CMT del Sur",
        "ruc"     : "20537658471",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "CMT DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "CMT SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Dynamitex",
        "ruc"     : "20600995761",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "DYNAMITEX DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "DYNAMITEX SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "DINSURA",
        "ruc"     : "20603964571",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "DINSURA DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "DINSURA SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Perú Commerce",
        "ruc"     : "20601234567",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "P.COMMERCE",   "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Inversiones Forestales del Sur (INFOSUR)",
        "ruc"     : "20600567890",
        "monedas" : [
            {"nombre": "Dolares (USD)", "sheet_bank": "INFOSUR DOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
            {"nombre": "Soles (PEN)",   "sheet_bank": "INFOSUR SOL", "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Thimble Sourcing / TST",
        "ruc"     : "20601987654",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "TST",         "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "Reforestadora Iñaupari",
        "ruc"     : "20601345678",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "REF. IÑAPARI",  "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "TECA Peruvian Group",
        "ruc"     : "20601456789",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "TECA",        "skip_bank": 4, "skip_conta": 11, "habilitado": True},
        ],
    },
    {
        "nombre"  : "DIONISO",
        "ruc"     : "20601567890",
        "monedas" : [
            {"nombre": "Soles (PEN)",   "sheet_bank": "INV. DIONISO",     "skip_bank": 4, "skip_conta": 11, "habilitado": True},
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
            # Incluir nombre de empresa y RUC en el config para el reporte
            return {
                **moneda,
                'empresa': empresa['nombre'],
                'ruc'    : empresa.get('ruc', ''),
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

    # Extraer el Saldo Contable Final ANTES de filtrar las filas de saldo
    # Contanet exporta el saldo en el texto: "Saldo Contable (Final): 175,002.61"
    # que cae en la columna 'nro_registro' (col B del Excel) con fecha_mov = NaN.
    import re as _re
    saldo_contable_final = None

    # Patrón 1: "Saldo Contable (Final): XXXXX" en la columna nro_registro
    mask_saldo_texto = conta['nro_registro'].astype(str).str.contains(
        r'Saldo\s+Contable\s*\(Final\)', case=False, na=False
    )
    filas_saldo_texto = conta[mask_saldo_texto]
    if not filas_saldo_texto.empty:
        try:
            texto = str(filas_saldo_texto['nro_registro'].iloc[0])
            # Extraer el número que sigue al ':'
            m = _re.search(r':\s*([\d,\.]+)', texto)
            if m:
                saldo_contable_final = float(m.group(1).replace(',', ''))
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
    banco_nombre = _extraer_nombre_banco(file_conta, config['skip_conta'])
    mes_anio     = _mes_anio_desde_bank(bank)
    ruta_salida  = _construir_nombre_archivo(
        'CBI',
        empresa=config.get('empresa', 'Empresa'),
        banco=banco_nombre,
        moneda=config['nombre'],
        mes_anio=mes_anio,
    )
    generar_reporte_inicial(
        bank,
        conta,
        ruta_salida=ruta_salida,
        empresa=config.get('empresa', 'Southern Textil'),
        moneda=config['nombre'],
        saldo_contable_final=saldo_contable_final,
        banco=banco_nombre,
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

    # Extraer banco y mes/año desde el archivo inicial para construir el nombre CBF
    banco_nombre = 'BANCO'
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

    mes_anio = _mes_anio_desde_inicial(file_inicial)
    ruta_salida = _construir_nombre_archivo(
        'CBF',
        empresa=config.get('empresa', 'Empresa'),
        banco=banco_nombre,
        moneda=config.get('nombre', ''),
        mes_anio=mes_anio,
    )

    try:
        generar_reporte_final(
            file_inicial,
            ruta_salida=ruta_salida,
            empresa=config.get('empresa', ''),
            ruc=config.get('ruc', ''),
            moneda=config.get('nombre', ''),
        )
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
    flujo_reporte_final(config)
