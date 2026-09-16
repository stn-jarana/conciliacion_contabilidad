"""
BN/extract_data_bn.py
─────────────────────
Parsea estados de cuenta del Banco de la Nación en formato PDF y los convierte
al DataFrame interno del motor de conciliacion bancaria.

Uso como modulo (importado desde main.py o reporte.py):
    from BN.extract_data_bn import leer_pdf_bn
    resultado = leer_pdf_bn(Path("ESTADO DE CUENTA CMT B.NACION 08-2026.pdf"))
    df_banco   = resultado["movimientos"]   # pd.DataFrame
    saldo_fin  = resultado["saldo_final"]   # float

Uso directo (linea de comandos):
    uv run python -m BN.extract_data_bn
    → pregunta la(s) ruta(s) de los PDF de forma interactiva.
"""

from __future__ import annotations

import re
import datetime
from pathlib import Path

import pandas as pd
import pymupdf


# =====================================================================
# CATALOGO DE EMPRESAS (RUC / DOI → nombre corto)
# =====================================================================

COMPANIES: dict[str, str] = {
    "20376729126": "STN",
    "20601910603": "ITS",
    "20504334041": "ITS",
    "20537658471": "CMT",
    "20506883301": "CMT",
    "20600995761": "Dynamitex",
    "20514016624": "Dynamitex",
    "20603964571": "DINSURA",
    "20494530865": "DINSURA",
    "20601234567": "Perú Commerce",
    "20600567890": "INFOSUR",
    "20601987654": "Thimble Sourcing / TST",
    "20601345678": "Iñaupari",
    "20601456789": "TECA",
    "20601567890": "DIONISO",
}

# Nombres y alias de empresas frecuentes para resolución cuando el PDF
# no incluye RUC numérico explícito (común en formatos del Banco de la Nación)
COMPANY_NAMES: dict[str, tuple[str, str]] = {
    "CMT DEL SUR": ("CMT", "20537658471"),
    "CORPORACION TEXTIL DEL SUR": ("CMT", "20537658471"),
    "CMT": ("CMT", "20537658471"),
    "SOUTHERN TEXTIL": ("STN", "20376729126"),
    "STN": ("STN", "20376729126"),
    "INVERSIONES TEXTILES DEL SUR": ("ITS", "20601910603"),
    "ITS": ("ITS", "20601910603"),
    "DYNAMITEX": ("Dynamitex", "20600995761"),
    "DINSURA": ("DINSURA", "20603964571"),
    "PERU COMMERCE": ("Perú Commerce", "20601234567"),
    "INFOSUR": ("INFOSUR", "20600567890"),
    "THIMBLE SOURCING": ("Thimble Sourcing / TST", "20601987654"),
    "TST": ("Thimble Sourcing / TST", "20601987654"),
    "INAPARI": ("Iñaupari", "20601345678"),
    "IÑAUPARI": ("Iñaupari", "20601345678"),
    "TECA": ("TECA", "20601456789"),
    "DIONISO": ("DIONISO", "20601567890"),
}


# =====================================================================
# FUNCIONES INTERNAS DE PARSEO
# =====================================================================

def _parse_monto(texto: object) -> float:
    """Convierte texto numérico con separadores de miles o comas a float."""
    if texto is None:
        return 0.0
    t = str(texto).replace(",", "").replace("*", "").replace(" ", "").strip()
    if not t or t in ("-", "--", ".", "nan", "none"):
        return 0.0
    if t.startswith("(") and t.endswith(")"):
        t = "-" + t[1:-1]
    try:
        return float(t)
    except ValueError:
        return 0.0


def _get_company(texto_completo: str, ruta_pdf: Path | None = None) -> tuple[str, str]:
    """Extrae el DOI (RUC) y el nombre corto de empresa del texto del PDF o su nombre."""
    # 1. Buscar RUC explícito de 11 dígitos
    ruc_match = re.search(r'\b(20\d{9})\b', texto_completo)
    if ruc_match:
        doi = ruc_match.group(1).strip()
        empresa = COMPANIES.get(doi, doi)
        return doi, empresa

    # 2. Buscar por NOMBRE DE CUENTA en la cabecera
    cta_match = re.search(r'NOMBRE DE CUENTA\s*:\s*([^\n\r]+)', texto_completo, re.IGNORECASE)
    if cta_match:
        nom_cuenta = cta_match.group(1).upper()
        for clave, (emp, ruc_val) in COMPANY_NAMES.items():
            if clave in nom_cuenta:
                return ruc_val, emp

    # 3. Buscar en todo el texto del documento
    txt_upper = texto_completo.upper()
    for clave, (emp, ruc_val) in COMPANY_NAMES.items():
        if clave in txt_upper:
            return ruc_val, emp

    # 4. Fallback al nombre del archivo
    if ruta_pdf:
        nombre_arch = Path(ruta_pdf).stem.upper()
        for clave, (emp, ruc_val) in COMPANY_NAMES.items():
            if clave in nombre_arch:
                return ruc_val, emp

    return "", ""


def _cuenta_corriente(texto_completo: str) -> str:
    """Extrae el número de cuenta corriente del texto del estado de cuenta."""
    match_cta = re.search(r'NUMERO\s+DE\s+CUENTA\s*:\s*([\d\-]+)', texto_completo, re.IGNORECASE)
    if match_cta:
        return match_cta.group(1).strip()

    match_alt = re.search(r'CTA\.?\s*(?:CTE\.?)?\s*No\.?\s*([\d\-]+)', texto_completo, re.IGNORECASE)
    if match_alt:
        return match_alt.group(1).strip()

    return ""


def _moneda_interna(texto_completo: str) -> str:
    """
    Deduce la moneda del extracto.
    Las cuentas corrientes de detracciones (D.LEG. 940) en Banco de la Nación
    son en moneda nacional (Soles). Si se detecta M.E., dólares o USD, usa Dólares.
    """
    t = texto_completo.upper()
    if any(k in t for k in ("M.E.", "MONEDA EXTRANJERA", "DOLARES", "DOLAR", "USD")):
        return "Dolares (USD)"
    return "Soles (PEN)"


def _inferir_anio(ruta_pdf: Path, texto_completo: str = "") -> int:
    """
    Infiere el año del estado de cuenta a partir del nombre del archivo
    o de la cabecera 'ESTADO DE CTA CTE AL : dd/mm/yyyy'.
    """
    # 1. Desde la cabecera de periodo
    m_periodo = re.search(r'ESTADO DE CTA CTE AL\s*:[^\n\r\d]*\d{1,2}/\d{1,2}/(20\d{2})', texto_completo, re.IGNORECASE)
    if m_periodo:
        return int(m_periodo.group(1))

    # 2. Desde el nombre del archivo
    nombre = Path(ruta_pdf).stem
    m_nom = re.search(r'(20\d{2})', nombre)
    if m_nom:
        return int(m_nom.group(1))

    # 3. Cualquier fecha completa en el texto
    m_fecha = re.search(r'\b\d{2}/\d{2}/(20\d{2})\b', texto_completo)
    if m_fecha:
        return int(m_fecha.group(1))

    return datetime.datetime.now().year


def _extraer_datos_directos_pymupdf(ruta_pdf: Path) -> dict | None:
    """
    Extrae los datos y movimientos usando PyMuPDF find_tables().
    Detecta automáticamente:
      - Encabezados y metadatos (empresa, cuenta, periodo)
      - Fila de SALDO ANTERIOR (saldo inicial)
      - Filas de transacciones con DIA, CODIFICACION, CARGOS, ABONOS, SALDOS
      - Resumen final (TOTAL CARGOS, TOTAL ABONOS, SALDO ACTUAL)
    """
    try:
        doc = pymupdf.open(ruta_pdf)
    except Exception:
        return None

    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"

    doi, empresa = _get_company(full_text, ruta_pdf)
    cuenta = _cuenta_corriente(full_text)
    moneda = _moneda_interna(full_text)
    anio_defecto = _inferir_anio(ruta_pdf, full_text)

    movimientos: list[dict] = []
    saldo_anterior: float | None = None
    saldo_actual: float = 0.0
    saldo_final_resumen: float | None = None

    for p_idx, page in enumerate(doc):
        tabs = page.find_tables()
        if not tabs.tables:
            continue

        df_tab = tabs.tables[0].to_pandas()
        if df_tab.empty:
            continue

        # Verificar si es la página final de totales
        cols_upper = [str(c).upper() for c in df_tab.columns]
        if any("SALDO ACTUAL" in c for c in cols_upper):
            # Fila de totales finales
            for _, r_tot in df_tab.iterrows():
                for c_name in df_tab.columns:
                    if "SALDO ACTUAL" in str(c_name).upper():
                        val_sa = _parse_monto(r_tot[c_name])
                        if val_sa != 0.0 or "0" in str(r_tot[c_name]):
                            saldo_final_resumen = val_sa
            continue

        # Identificar índices de columnas
        idx_cod = 0
        idx_cargos = 1
        idx_abonos = 2
        idx_saldos = 3
        idx_dia = 4

        for c_i, c_val in enumerate(df_tab.columns):
            c_norm = str(c_val).upper()
            if "CODIFICACION" in c_norm or "CHEQUE" in c_norm:
                idx_cod = c_i
            elif "CARGO" in c_norm:
                idx_cargos = c_i
            elif "ABONO" in c_norm:
                idx_abonos = c_i
            elif "SALDO" in c_norm:
                idx_saldos = c_i
            elif "DIA" in c_norm or "FECHA" in c_norm:
                idx_dia = c_i

        for _, row in df_tab.iterrows():
            if len(row) <= max(idx_cod, idx_cargos, idx_abonos, idx_dia):
                continue

            cod = str(row.iloc[idx_cod] or "").strip()
            if not cod:
                continue

            # Detectar fila de saldo anterior
            if "SALDO ANTERIOR" in cod.upper():
                s_ant = _parse_monto(row.iloc[idx_saldos])
                if s_ant == 0.0:
                    # En algunos casos el saldo anterior está en la columna abonos/cargos
                    s_ant = _parse_monto(row.iloc[idx_abonos])
                saldo_anterior = s_ant
                saldo_actual = s_ant
                continue

            # Movimiento normal
            cargos_val = abs(_parse_monto(row.iloc[idx_cargos]))
            abonos_val = abs(_parse_monto(row.iloc[idx_abonos]))
            saldos_val = _parse_monto(row.iloc[idx_saldos]) if idx_saldos < len(row) else 0.0
            dia_str = str(row.iloc[idx_dia] or "").strip()

            # Si no hay monto ni fecha, descartar
            if cargos_val == 0.0 and abonos_val == 0.0 and not dia_str:
                continue

            # Parsear fecha
            fecha_dt = None
            if dia_str:
                m_f = re.search(r'(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?', dia_str)
                if m_f:
                    dia_n, mes_n, anio_n = m_f.groups()
                    a_usar = int(anio_n) if anio_n else anio_defecto
                    if a_usar < 100:
                        a_usar += 2000
                    try:
                        fecha_dt = datetime.datetime(a_usar, int(mes_n), int(dia_n))
                    except ValueError:
                        fecha_dt = None

            # Monto neto con signo: Abonos (+) y Cargos (-)
            monto_neto = round(abonos_val - cargos_val, 2)

            # Actualizar saldo acumulado
            saldo_actual = round(saldo_actual + monto_neto, 2)
            if saldos_val > 0.0:
                saldo_fila = saldos_val
                saldo_actual = saldos_val
            else:
                saldo_fila = saldo_actual

            # Extraer número de operación si aplica
            m_op = re.search(r'\d+', cod)
            nro_op = m_op.group(0) if m_op else cod

            movimientos.append({
                "fecha": fecha_dt,
                "descripcion": cod,
                "nro_operacion": nro_op,
                "monto": monto_neto,
                "saldo": saldo_fila,
            })

    # Si encontramos el saldo final en la tabla resumen, usarlo; de lo contrario, el acumulado
    saldo_final = saldo_final_resumen if saldo_final_resumen is not None else (saldo_actual if movimientos else None)

    return {
        "empresa": empresa,
        "doi": doi,
        "moneda": moneda,
        "cuenta": cuenta,
        "saldo_anterior": saldo_anterior,
        "saldo_final": saldo_final,
        "movimientos": movimientos,
    }


def _extraer_datos_texto_fallback(ruta_pdf: Path) -> dict:
    """
    Fallback basado en expresiones regulares sobre el texto completo
    cuando find_tables() no detecte la estructura de tabla.
    """
    import pymupdf4llm
    md = pymupdf4llm.to_markdown(str(ruta_pdf))
    anio_defecto = _inferir_anio(ruta_pdf, md)

    doi, empresa = _get_company(md, ruta_pdf)
    cuenta = _cuenta_corriente(md)
    moneda = _moneda_interna(md)

    movimientos: list[dict] = []
    saldo_final: float | None = None

    # Buscar saldo actual final
    m_sfin = re.search(r'SALDO ACTUAL\s*[:\*]*\s*([\d,\.]+)', md, re.IGNORECASE)
    if m_sfin:
        saldo_final = _parse_monto(m_sfin.group(1))

    # Parsear líneas de tabla en Markdown
    for linea in md.splitlines():
        linea = linea.strip()
        if not linea.startswith("|"):
            continue

        cols = [c.strip() for c in linea.strip("|").split("|")]
        if len(cols) < 4:
            continue

        # Descartar cabeceras o filas de separación
        if any(h in linea.upper() for h in ("CARGOS", "ABONOS", "SALDO ACTUAL", "TOTAL")):
            continue

        # Buscar fecha en la última columna o en alguna de las columnas
        fecha_encontrada = None
        dia_col = cols[-1]
        m_dia = re.search(r'(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?', dia_col)
        if m_dia:
            dia_n, mes_n, anio_n = m_dia.groups()
            a_usar = int(anio_n) if anio_n else anio_defecto
            if a_usar < 100:
                a_usar += 2000
            try:
                fecha_encontrada = datetime.datetime(a_usar, int(mes_n), int(dia_n))
            except ValueError:
                fecha_encontrada = None

        if not fecha_encontrada:
            continue

        cod = cols[0]
        cargos = abs(_parse_monto(cols[1])) if len(cols) > 1 else 0.0
        abonos = abs(_parse_monto(cols[2])) if len(cols) > 2 else 0.0
        saldos = _parse_monto(cols[3]) if len(cols) > 3 else 0.0

        monto_neto = round(abonos - cargos, 2)
        m_op = re.search(r'\d+', cod)
        nro_op = m_op.group(0) if m_op else cod

        movimientos.append({
            "fecha": fecha_encontrada,
            "descripcion": cod,
            "nro_operacion": nro_op,
            "monto": monto_neto,
            "saldo": saldos,
        })

    return {
        "empresa": empresa,
        "doi": doi,
        "moneda": moneda,
        "cuenta": cuenta,
        "saldo_anterior": None,
        "saldo_final": saldo_final,
        "movimientos": movimientos,
    }


# =====================================================================
# FUNCIÓN PRINCIPAL
# =====================================================================

def leer_pdf_bn(ruta_pdf: Path | str) -> dict:
    """
    Parsea un estado de cuenta del Banco de la Nación en PDF y devuelve un dict
    con los metadatos y el DataFrame de movimientos listo para el motor de conciliación.

    Retorna:
      'empresa'     : str    – nombre corto de empresa (ej. 'CMT')
      'doi'         : str    – RUC de la empresa
      'moneda'      : str    – 'Soles (PEN)' o 'Dolares (USD)'
      'cuenta'      : str    – número de cuenta bancaria
      'saldo_final' : float  – saldo final del periodo
      'movimientos' : pd.DataFrame con columnas internas del motor:
                        fecha, descripcion, monto, ingreso, egreso,
                        saldo, nro_operacion, sucursal, hora, usuario

    Raises:
        FileNotFoundError  si el PDF no existe.
        ValueError         si no se pueden extraer movimientos válidos.
    """
    ruta_pdf = Path(ruta_pdf)
    if not ruta_pdf.exists():
        raise FileNotFoundError(f"No se encontro el archivo: {ruta_pdf}")

    print(f"  Leyendo PDF Banco de la Nación: {ruta_pdf.name} ...", end=" ", flush=True)

    # 1. Intentar extracción directa con PyMuPDF (precisa y rápida)
    res = _extraer_datos_directos_pymupdf(ruta_pdf)
    if res and res.get("movimientos"):
        filas = res["movimientos"]
        empresa = res["empresa"]
        doi = res["doi"]
        moneda = res["moneda"]
        cuenta = res["cuenta"]
        saldo_final = res["saldo_final"]
        print(f"OK ({len(filas)} movimientos encontrados)")
    else:
        # 2. Fallback a extracción basada en texto / markdown
        res_fb = _extraer_datos_texto_fallback(ruta_pdf)
        filas = res_fb.get("movimientos", [])
        empresa = res_fb["empresa"]
        doi = res_fb["doi"]
        moneda = res_fb["moneda"]
        cuenta = res_fb["cuenta"]
        saldo_final = res_fb["saldo_final"]
        print(f"OK (fallback, {len(filas)} movimientos encontrados)")

    if not filas:
        raise ValueError(
            f"No se encontraron movimientos en el PDF '{ruta_pdf.name}'. "
            "Verifique que el archivo sea un estado de cuenta de Banco de la Nación valido."
        )

    df = pd.DataFrame(filas)
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["monto"] = pd.to_numeric(df["monto"], errors="coerce").fillna(0.0)
    df["ingreso"] = df["monto"].clip(lower=0)
    df["egreso"] = df["monto"].clip(upper=0).abs()
    if "saldo" not in df.columns:
        df["saldo"] = 0.0
    else:
        df["saldo"] = pd.to_numeric(df["saldo"], errors="coerce").fillna(0.0)
    df["sucursal"] = ""
    df["hora"] = ""
    df["usuario"] = ""

    # Ordenar por fecha cronológica manteniendo estabilidad
    df = df.sort_values("fecha", kind="mergesort").reset_index(drop=True)

    # Columnas en el mismo orden que el motor de conciliación espera
    columnas_internas = [
        "fecha", "descripcion", "monto", "ingreso", "egreso",
        "saldo", "nro_operacion", "sucursal", "hora", "usuario",
    ]
    df = df[columnas_internas]

    return {
        "empresa": empresa,
        "doi": doi,
        "moneda": moneda,
        "cuenta": cuenta,
        "saldo_final": saldo_final,
        "movimientos": df,
    }


# =====================================================================
# PUNTO DE ENTRADA INTERACTIVO (ejecución directa)
# =====================================================================

def _pedir_archivos_pdf() -> list[Path]:
    """Solicita rutas de archivos PDF por consola hasta ingresar una línea vacía."""
    print("=" * 60)
    print("  EXTRACCION DE ESTADO DE CUENTA BANCO DE LA NACION (PDF)")
    print("=" * 60)
    print("\nIngrese la ruta de cada archivo PDF y presione ENTER.")
    print("Cuando haya terminado de agregar archivos, presione ENTER sin escribir nada.\n")

    rutas: list[Path] = []
    while True:
        entrada = input(f"  Archivo PDF #{len(rutas) + 1}: ").strip().strip('"').strip("'")
        if not entrada:
            if not rutas:
                print("  X Debe ingresar al menos un archivo.\n")
                continue
            break
        ruta = Path(entrada)
        if not ruta.exists():
            print(f"  X No se encontro el archivo: {entrada}\n")
        elif ruta.suffix.lower() != ".pdf":
            print("  X El archivo debe ser un PDF (.pdf)\n")
        else:
            rutas.append(ruta)
            print(f"  OK Agregado: {ruta.name}\n")

    return rutas


if __name__ == "__main__":
    import sys
    # Si se pasó un archivo como argumento, procesarlo directamente
    if len(sys.argv) > 1:
        rutas_entrada = [Path(p) for p in sys.argv[1:] if Path(p).exists()]
    else:
        # Si existe el archivo de ejemplo en el directorio actual, mostrar opción o pedir
        ejemplo = Path("ESTADO DE CUENTA CMT B.NACION 08-2026.pdf")
        if ejemplo.exists():
            print(f"  [Auto-detectado] {ejemplo.name}")
            rutas_entrada = [ejemplo]
        else:
            rutas_entrada = _pedir_archivos_pdf()

    print(f"\n  Procesando {len(rutas_entrada)} archivo(s)...\n")
    for ruta in rutas_entrada:
        print("-" * 60)
        try:
            resultado = leer_pdf_bn(ruta)
            df_mov = resultado["movimientos"]

            print(f"  Empresa    : {resultado['empresa']}  (DOI: {resultado['doi']})")
            print(f"  Moneda     : {resultado['moneda']}")
            print(f"  Cuenta     : {resultado['cuenta']}")
            print(f"  Saldo final: {resultado['saldo_final']}")
            print(f"  Movimientos: {len(df_mov)} registros")
            print()
            print("  Primeros 5 movimientos:")
            print(df_mov.head(5).to_string(index=False))
            print()
            print("  Ultimos 5 movimientos:")
            print(df_mov.tail(5).to_string(index=False))
        except (FileNotFoundError, ValueError) as e:
            print(f"  [ERROR] {e}")
        print()

    print("=" * 60)
    print("  Extraccion completada.")
    print("=" * 60)
