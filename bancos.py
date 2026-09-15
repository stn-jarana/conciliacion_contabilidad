"""Reglas compartidas para identificar, filtrar y normalizar bancos.

El motor de conciliación trabaja con una estructura interna única.  Este módulo
adapta los extractos de BCP, Scotiabank y Banco de la Nación a esa estructura,
y evita que los registros de Contanet de bancos distintos se crucen entre sí.
"""

from __future__ import annotations

from pathlib import Path
import re
import unicodedata

import pandas as pd


BANCOS: dict[str, dict[str, tuple[str, ...] | str]] = {
    "BCP": {
        "nombre": "BCP",
        "alias_contanet": ("BANCO DE CREDITO", "BANCO DE CREDITO DEL PERU", "BCP"),
    },
    "SCOTIABANK": {
        "nombre": "Scotiabank",
        "alias_contanet": ("SCOTIABANK", "SCOTIA"),
    },
    "BN": {
        "nombre": "Banco de la Nación",
        "alias_contanet": ("BANCO DE LA NACION", "BANCO NACION", "BN"),
    },
}


def normalizar_texto(valor: object) -> str:
    """Convierte texto a una forma comparable, sin tildes ni espacios dobles."""
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = texto.upper().replace("N°", " NUMERO ").replace("Nº", " NUMERO ").replace("#", " NUMERO ")
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", texto)).strip()


def clave_banco(banco: str) -> str:
    """Devuelve BCP, SCOTIABANK o BN a partir de una clave, alias o nombre."""
    texto = normalizar_texto(banco)
    if texto in BANCOS:
        return texto
    for clave, datos in BANCOS.items():
        if texto == normalizar_texto(datos["nombre"]):
            return clave
        if any(texto == normalizar_texto(alias) for alias in datos["alias_contanet"]):
            return clave
    raise ValueError(f"Banco no soportado: {banco}. Opciones: BCP, Scotiabank o BN.")


def nombre_banco(banco: str) -> str:
    """Nombre de presentación del banco seleccionado."""
    return str(BANCOS[clave_banco(banco)]["nombre"])


def nombre_corto_banco(banco: str) -> str:
    """Devuelve la abreviatura estándar del banco para nombres de archivo: BCP, BN o Scotia."""
    clave = clave_banco(banco)
    if clave == "BCP":
        return "BCP"
    elif clave == "BN":
        return "BN"
    elif clave == "SCOTIABANK":
        return "Scotia"
    return clave


def nombre_archivo_banco(banco: str, sufijo: str) -> str:
    """Construye el nombre de archivo a partir del banco seleccionado (BCP, BN o Scotia)."""
    return f"{nombre_corto_banco(banco)}_{sufijo}.xlsx"


def _contiene_alias(texto: object, banco: str) -> bool:
    texto_normalizado = normalizar_texto(texto)
    for alias in BANCOS[clave_banco(banco)]["alias_contanet"]:
        alias_normalizado = normalizar_texto(alias)
        if len(alias_normalizado) <= 3:
            if re.search(rf"(?<![A-Z0-9]){re.escape(alias_normalizado)}(?![A-Z0-9])", texto_normalizado):
                return True
        elif alias_normalizado in texto_normalizado:
            return True
    return False


def filtrar_contabilidad_por_banco(conta: pd.DataFrame, banco: str) -> pd.DataFrame:
    """Conserva solo los registros Contanet cuyo *Giro* corresponde al banco.

    El export de Contanet identifica la cuenta bancaria en la columna ``giro``.
    Filtrar ahí —en vez de por importe o glosa— evita que movimientos de otro
    banco queden disponibles para una conciliación equivocada.
    """
    banco = clave_banco(banco)
    if "giro" not in conta.columns:
        raise ValueError("El reporte de contabilidad no contiene la columna 'Giro' necesaria para separar el banco.")

    mascara = conta["giro"].map(lambda valor: _contiene_alias(valor, banco))
    resultado = conta[mascara].copy()
    if resultado.empty:
        raise ValueError(
            f"No se encontraron movimientos de {nombre_banco(banco)} en la columna 'Giro' del reporte Contanet."
        )
    return resultado.reset_index(drop=True)


def _a_numero(serie: pd.Series) -> pd.Series:
    """Convierte importes con separadores locales o símbolo monetario a número."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")

    def convertir(valor: object) -> float | None:
        if pd.isna(valor):
            return None
        if isinstance(valor, (int, float)):
            return float(valor)
        texto = str(valor).strip()
        if not texto or texto.lower() in {"nan", "none", "-"}:
            return None
        negativo = texto.startswith("(") and texto.endswith(")")
        texto = texto.strip("()")
        texto = re.sub(r"[^0-9,.-]", "", texto)
        if not texto or texto in {"-", ".", ","}:
            return None

        coma, punto = texto.rfind(","), texto.rfind(".")
        if coma >= 0 and punto >= 0:
            if coma > punto:  # 1.234,56
                texto = texto.replace(".", "").replace(",", ".")
            else:  # 1,234.56
                texto = texto.replace(",", "")
        elif coma >= 0:
            decimales = len(texto) - coma - 1
            texto = texto.replace(",", ".") if decimales in (1, 2) else texto.replace(",", "")
        elif punto >= 0:
            decimales = len(texto) - punto - 1
            if decimales not in (1, 2):
                texto = texto.replace(".", "")
        try:
            numero = float(texto)
            return -abs(numero) if negativo else numero
        except ValueError:
            return None

    return serie.map(convertir)


def _a_texto(serie: pd.Series | None) -> pd.Series:
    if serie is None:
        return pd.Series(dtype="object")

    def convertir(valor: object) -> str:
        if pd.isna(valor):
            return ""
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()

    return serie.map(convertir)


def _buscar_columna(columnas: list[str], *patrones: str) -> str | None:
    """Encuentra una columna por nombre exacto y luego por contenido."""
    normalizadas = {columna: normalizar_texto(columna) for columna in columnas}
    for patron in patrones:
        patron_normalizado = normalizar_texto(patron)
        for columna, nombre in normalizadas.items():
            if nombre == patron_normalizado:
                return columna
    for patron in patrones:
        patron_normalizado = normalizar_texto(patron)
        for columna, nombre in normalizadas.items():
            if patron_normalizado in nombre:
                return columna
    return None


def _fila_cabecera(datos: pd.DataFrame) -> int | None:
    """Ubica la cabecera en los primeros renglones de un extracto Excel."""
    claves = ("FECHA", "DESCRIPCION", "DETALLE", "CONCEPTO", "MONTO", "IMPORTE", "DEBITO", "CREDITO", "CARGO", "ABONO", "SALDO", "OPERACION")
    mejor_fila: int | None = None
    mejor_puntaje = 0
    for indice in range(min(len(datos), 40)):
        celdas = [normalizar_texto(valor) for valor in datos.iloc[indice].tolist()]
        puntaje = sum(any(clave in celda for celda in celdas) for clave in claves)
        tiene_fecha = any("FECHA" in celda for celda in celdas)
        tiene_importe = any(clave in celda for celda in celdas for clave in ("MONTO", "IMPORTE", "DEBITO", "CREDITO", "CARGO", "ABONO"))
        if tiene_fecha and tiene_importe and puntaje > mejor_puntaje:
            mejor_fila, mejor_puntaje = indice, puntaje
    return mejor_fila if mejor_puntaje >= 2 else None


def _hoja_bancaria(xls: pd.ExcelFile, hojas_preferidas: object) -> str:
    """Usa la hoja de empresa/moneda si existe; si no, la primera del extracto."""
    if isinstance(hojas_preferidas, (tuple, list)):
        encontrada = next((hoja for hoja in hojas_preferidas if hoja in xls.sheet_names), None)
        if encontrada:
            return encontrada
    elif isinstance(hojas_preferidas, str) and hojas_preferidas in xls.sheet_names:
        return hojas_preferidas
    if not xls.sheet_names:
        raise ValueError("El estado de cuenta no contiene hojas para leer.")
    return xls.sheet_names[0]


def leer_estado_bancario(
    archivo: Path,
    banco: str,
    hojas_preferidas: object = (),
) -> tuple[pd.DataFrame, str]:
    """Lee un extracto de los tres bancos y lo convierte al formato interno.

    Soporta una columna de importe firmada (el formato BCP actual) o columnas
    separadas de cargos/débitos y abonos/créditos (comunes en Scotia y BN).
    """
    clave_banco(banco)  # validar antes de abrir el archivo
    with pd.ExcelFile(archivo) as xls:
        hoja = _hoja_bancaria(xls, hojas_preferidas)
        bruto = pd.read_excel(xls, sheet_name=hoja, header=None)
    fila_cabecera = _fila_cabecera(bruto)
    if fila_cabecera is None:
        raise ValueError(
            f"No se pudo identificar la cabecera del extracto de {nombre_banco(banco)} en la hoja '{hoja}'. "
            "Debe incluir Fecha y Monto/Importe o Débito/Crédito."
        )

    cabecera = [str(valor).strip() if pd.notna(valor) else f"columna_{i}" for i, valor in enumerate(bruto.iloc[fila_cabecera])]
    datos = bruto.iloc[fila_cabecera + 1:].copy()
    datos.columns = cabecera
    datos = datos.dropna(axis=1, how="all")

    columnas = list(datos.columns)
    col_fecha = _buscar_columna(columnas, "fecha", "fecha operacion", "fecha transaccion", "fecha movimiento")
    col_descripcion = _buscar_columna(columnas, "descripcion", "detalle", "concepto", "glosa", "movimiento")
    col_monto = _buscar_columna(columnas, "monto", "importe", "monto operacion", "importe operacion")
    col_credito = _buscar_columna(columnas, "credito", "abono", "haber", "ingreso", "deposito")
    col_debito = _buscar_columna(columnas, "debito", "cargo", "debe", "egreso", "retiro")
    col_saldo = _buscar_columna(columnas, "saldo", "balance")
    col_operacion = _buscar_columna(columnas, "nro operacion", "numero operacion", "operacion", "nro documento", "numero documento", "documento", "referencia", "secuencia")
    col_fecha_valuta = _buscar_columna(columnas, "fecha valuta", "fecha valor")
    col_sucursal = _buscar_columna(columnas, "sucursal", "agencia", "oficina")
    col_hora = _buscar_columna(columnas, "hora")
    col_usuario = _buscar_columna(columnas, "usuario", "cliente")

    if col_fecha is None or col_descripcion is None or (col_monto is None and col_credito is None and col_debito is None):
        raise ValueError(
            f"La hoja '{hoja}' no tiene las columnas mínimas para conciliar {nombre_banco(banco)}. "
            f"Columnas detectadas: {', '.join(map(str, columnas))}"
        )

    if col_credito is not None or col_debito is not None:
        credito = _a_numero(datos[col_credito]).fillna(0) if col_credito else pd.Series(0.0, index=datos.index)
        debito = _a_numero(datos[col_debito]).fillna(0) if col_debito else pd.Series(0.0, index=datos.index)
        monto = credito - debito
    else:
        monto = _a_numero(datos[col_monto])

    resultado = pd.DataFrame({
        "fecha": pd.to_datetime(datos[col_fecha], dayfirst=True, errors="coerce"),
        "fecha_valuta": pd.to_datetime(datos[col_fecha_valuta], dayfirst=True, errors="coerce") if col_fecha_valuta else pd.NaT,
        "descripcion": _a_texto(datos[col_descripcion]),
        "monto": monto,
        "saldo": _a_numero(datos[col_saldo]) if col_saldo else 0.0,
        "sucursal": _a_texto(datos[col_sucursal]) if col_sucursal else "",
        "nro_operacion": _a_texto(datos[col_operacion]) if col_operacion else "",
        "hora": _a_texto(datos[col_hora]) if col_hora else "",
        "usuario": _a_texto(datos[col_usuario]) if col_usuario else "",
    })
    resultado = resultado[resultado["fecha"].notna() | resultado["monto"].notna()].copy()
    resultado["monto"] = resultado["monto"].fillna(0.0)
    resultado["saldo"] = resultado["saldo"].fillna(0.0)
    resultado["ingreso"] = resultado["monto"].clip(lower=0)
    resultado["egreso"] = resultado["monto"].clip(upper=0).abs()
    return resultado.sort_values("fecha").reset_index(drop=True), hoja
