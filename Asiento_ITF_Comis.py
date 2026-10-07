"""Genera asientos de ITF y comisiones desde una conciliación final.

El archivo fuente debe contener una hoja ``ITF Y COM`` (también se aceptan
variantes como ``ITF y Comisiones``) con las secciones de comisiones e ITF.
El resultado es una copia de la plantilla Contanet, conservando sus macros,
formatos y valores por defecto. Los asientos se escriben en la hoja visible
``CONTABILIDAD``; la hoja técnica ``ImportCONTABILIDAD`` no se modifica.

El tipo de cambio no se adivina: debe ser el *TC Venta* consultado en
Contanet para el último día de cada mes informado. Puede entregarse como un
valor único para un único período, un diccionario por fecha de cierre, o una
función que realice la consulta autenticada de Contanet.

Ejemplo de ejecución (un solo período)::

    python Asiento_ITF_Comis.py CBF_STN_BCP_Dolares_JULIO_2026.xlsx \
        "Plantilla comisiones Bcp Dolares.xlsm" \
        --banco BCP --moneda ME --tc-venta 3.415 --salida Asiento_ITF.xlsm

Para más de un período, se debe indicar un TC Venta por cada cierre mensual::

    --tc-venta 2026-06-30=3.415 --tc-venta 2026-07-31=3.400
"""

from __future__ import annotations

import argparse
from calendar import monthrange
from collections.abc import Callable, Iterable, Mapping
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import json
from pathlib import Path
import re
import unicodedata
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


# ── Carga de cuentas desde cuentas_banco.json ────────────────────────────────
# El JSON tiene la estructura: empresas[empresa][banco_norm][moneda_codigo]
# con campos: cuenta_comision, cuenta_itf, cuenta_cce, cuenta_banco.
# Se mantiene el dict CUENTAS como fallback para compatibilidad con código
# existente que no pase empresa (solo banco+moneda, igual que antes).

_RUTA_JSON_CUENTAS = Path(__file__).parent / "cuentas_banco.json"


def cargar_cuentas_banco() -> dict:
    """Carga el archivo cuentas_banco.json con las cuentas contables por empresa/banco/moneda.

    Returns
    -------
    dict
        Contenido completo del JSON. La clave principal es ``'empresas'``.
        Devuelve un dict vacío si el archivo no existe o está mal formado.
    """
    if not _RUTA_JSON_CUENTAS.is_file():
        return {}
    try:
        with open(_RUTA_JSON_CUENTAS, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


# Diccionario de fallback para empresas no registradas en el JSON
# (mantiene compatibilidad con el flujo anterior de una única empresa).
CUENTAS: dict[tuple[str, str], dict[str, str]] = {
    ("BCP", "MN"): {
        "empresa": "97.7.9.2003",
        "empresa-ITF": "94.4.1.2001",
        "banco": "10.4.1.1002",
    },
    ("BCP", "ME"): {
        "empresa": "97.7.9.2003",
        "empresa-ITF": "94.4.1.2001",
        "banco": "10.4.1.2001",
    },
    ("SCOTIA", "MN"): {
        "empresa": "",
        "empresa-ITF": "",
        "banco": "",
    },
    ("SCOTIA", "ME"): {
        "empresa": "97.7.9.2003",
        "empresa-ITF": "94.4.1.2001",
        "banco": "10.4.1.2002",
    },
    ("BCP MIAMI", "ME"): {
        "empresa": "97.7.9.1001",
        "empresa-ITF": "10.4.1.2004",
        "banco": "10.4.1.2004",
    },
}

# La plantilla tiene además una hoja técnica oculta llamada
# ``ImportCONTABILIDAD`` que Contanet lee directamente durante la importación.
NOMBRE_HOJA_DESTINO = "CONTABILIDAD"
NOMBRE_HOJA_IMPORT = "ImportCONTABILIDAD"
NOMBRES_HOJA_FUENTE = ("itf y com", "itf y comisiones", "itf comisiones")
CENTAVOS = Decimal("0.01")


class ErrorAsientoITFComisiones(ValueError):
    """Error de datos o configuración al generar el asiento."""


class TipoCambioNoDisponibleError(ErrorAsientoITFComisiones):
    """No se recibió el TC Venta de Contanet para un cierre mensual."""


@dataclass(frozen=True)
class MovimientoITFComision:
    """Movimiento válido leído de la hoja ITF y Comisiones."""

    fecha: date
    descripcion: str
    monto: Decimal
    tipo: str  # ``COMISION`` o ``ITF``
    fila_origen: int


TipoCambioVenta = (
    Decimal
    | float
    | int
    | str
    | Mapping[date | datetime | str, Decimal | float | int | str]
    | Callable[[date], Decimal | float | int | str]
)


def _normalizar(texto: Any) -> str:
    """Normaliza texto para comparar encabezados, bancos y nombres de hoja."""
    valor = "" if texto is None else str(texto)
    valor = "".join(
        caracter
        for caracter in unicodedata.normalize("NFD", valor)
        if unicodedata.category(caracter) != "Mn"
    )
    return " ".join(valor.upper().replace("º", "O").replace("°", "O").split())


def _normalizar_moneda(moneda: str) -> str:
    valor = _normalizar(moneda)
    if valor in {"MN", "PEN", "SOL", "SOLES", "SOLES (PEN)"} or "SOL" in valor:
        return "MN"
    if valor in {"ME", "USD", "DOLAR", "DOLARES", "DOLARES (USD)"} or "USD" in valor or "DOL" in valor:
        return "ME"
    raise ErrorAsientoITFComisiones(
        f"Moneda no reconocida: {moneda!r}. Use MN/Soles o ME/Dólares."
    )


def _normalizar_banco(banco: str) -> str:
    valor = _normalizar(banco)
    if "MIAMI" in valor and "BCP" in valor:
        return "BCP MIAMI"
    if valor in {"SCOTIA", "SCOTIABANK"} or "SCOTIA" in valor:
        return "SCOTIA"
    if valor in {"BCP", "BANCO DE CREDITO"} or "CREDITO" in valor:
        return "BCP"
    return valor


def _a_decimal(valor: Any, *, campo: str) -> Decimal:
    """Convierte un importe o TC a Decimal, aceptando formatos ES y EN."""
    if isinstance(valor, Decimal):
        return valor
    if isinstance(valor, bool) or valor is None:
        raise ErrorAsientoITFComisiones(f"{campo} no es un número válido: {valor!r}.")
    if isinstance(valor, (int, float)):
        try:
            return Decimal(str(valor))
        except InvalidOperation as exc:
            raise ErrorAsientoITFComisiones(
                f"{campo} no es un número válido: {valor!r}."
            ) from exc

    texto = str(valor).strip()
    if not texto:
        raise ErrorAsientoITFComisiones(f"{campo} está vacío.")
    negativo_por_parentesis = texto.startswith("(") and texto.endswith(")")
    texto = texto.strip("()")
    texto = re.sub(r"[^0-9,.-]", "", texto)

    # Si hay ambos separadores, el último es el decimal. Un único punto se
    # considera decimal: permite ingresar correctamente TC Venta como 3.415.
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif texto.count(",") == 1:
        entero, decimal = texto.split(",")
        texto = entero + decimal if len(decimal) == 3 else entero + "." + decimal
    elif texto.count(".") > 1:
        partes = texto.split(".")
        texto = "".join(partes[:-1]) + "." + partes[-1]

    try:
        numero = Decimal(texto)
    except InvalidOperation as exc:
        raise ErrorAsientoITFComisiones(
            f"{campo} no es un número válido: {valor!r}."
        ) from exc
    return -numero if negativo_por_parentesis and numero > 0 else numero


def _a_fecha(valor: Any) -> date | None:
    """Convierte fechas Excel, datetime y textos frecuentes a ``date``."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
    return None


def ultimo_dia_mes(fecha_movimiento: date) -> date:
    """Devuelve el último día natural del mes de una operación."""
    return date(
        fecha_movimiento.year,
        fecha_movimiento.month,
        monthrange(fecha_movimiento.year, fecha_movimiento.month)[1],
    )


def _encabezado_equivale(valor: Any, *opciones: str) -> bool:
    normalizado = _normalizar(valor)
    return normalizado in {_normalizar(opcion) for opcion in opciones}


def _buscar_hoja_fuente(libro) -> Worksheet:
    for hoja in libro.worksheets:
        if _normalizar(hoja.title).lower() in NOMBRES_HOJA_FUENTE:
            return hoja
    for hoja in libro.worksheets:
        nombre = _normalizar(hoja.title)
        if "ITF" in nombre and ("COM" in nombre or "COMISION" in nombre):
            return hoja
    raise ErrorAsientoITFComisiones(
        "No se encontró la hoja 'ITF Y COM' (o 'ITF y Comisiones') "
        "en la conciliación final."
    )


def _tipo_seccion(valor: Any) -> str | None:
    texto = _normalizar(valor)
    if not texto or "TOTAL" in texto:
        return None
    if "ITF" in texto:
        return "ITF"
    if "COMIS" in texto:
        return "COMISION"
    return None


def _indice_encabezados_fuente(fila: Iterable[Cell]) -> dict[str, int] | None:
    """Identifica los encabezados Fecha, Descripción y Monto de una sección."""
    resultado: dict[str, int] = {}
    for celda in fila:
        valor = _normalizar(celda.value)
        if valor in {"FECHA", "FECHA MOVIMIENTO"}:
            resultado["fecha"] = celda.column
        elif valor in {"DESCRIPCION", "GLOSA", "DETALLE"}:
            resultado["descripcion"] = celda.column
        elif valor in {"MONTO", "IMPORTE"}:
            resultado["monto"] = celda.column
    return resultado if {"fecha", "descripcion", "monto"} <= resultado.keys() else None


def leer_movimientos_itf_comisiones(ruta_conciliacion_final: str | Path) -> list[MovimientoITFComision]:
    """Lee los movimientos de las secciones ITF y comisiones del reporte final."""
    ruta = Path(ruta_conciliacion_final).expanduser().resolve()
    if not ruta.is_file():
        raise FileNotFoundError(f"No existe la conciliación final: {ruta}")

    libro = load_workbook(ruta, read_only=True, data_only=True)
    try:
        hoja = _buscar_hoja_fuente(libro)
        seccion: str | None = None
        columnas: dict[str, int] | None = None
        movimientos: list[MovimientoITFComision] = []

        for numero_fila, fila in enumerate(hoja.iter_rows(), start=1):
            primer_valor = fila[0].value if fila else None
            nueva_seccion = _tipo_seccion(primer_valor)
            if nueva_seccion is not None:
                seccion = nueva_seccion
                columnas = None
                continue

            encabezados = _indice_encabezados_fuente(fila)
            if encabezados is not None:
                columnas = encabezados
                continue

            if seccion is None or columnas is None:
                continue

            fecha = _a_fecha(fila[columnas["fecha"] - 1].value)
            if fecha is None:
                continue  # Totales, espacios u otros textos del reporte.
            descripcion = str(fila[columnas["descripcion"] - 1].value or "").strip()
            if not descripcion:
                raise ErrorAsientoITFComisiones(
                    f"La fila {numero_fila} de '{hoja.title}' no tiene descripción."
                )
            monto = _a_decimal(
                fila[columnas["monto"] - 1].value,
                campo=f"Monto de la fila {numero_fila} de '{hoja.title}'",
            )
            if monto == 0:
                continue
            movimientos.append(
                MovimientoITFComision(
                    fecha=fecha,
                    descripcion=descripcion,
                    monto=monto,
                    tipo=seccion,
                    fila_origen=numero_fila,
                )
            )
    finally:
        libro.close()

    if not movimientos:
        raise ErrorAsientoITFComisiones(
            "La hoja ITF y Comisiones no contiene movimientos con fecha y monto."
        )
    return movimientos


def _buscar_fila_encabezados_destino(hoja: Worksheet) -> tuple[int, dict[str, int]]:
    """Ubica los encabezados requeridos en la plantilla Contanet."""
    requeridos: dict[str, tuple[str, ...]] = {
        "correlativo": ("Correlativo",),
        "relacionado": ("Relacionado",),
        "ejercicio": ("Ejercicio",),
        "periodo": ("Periodo", "Período"),
        "nro_cuenta": ("Numero Cuenta", "Número Cuenta", "Nro Cuenta", "N° Cuenta"),
        "glosa": ("Glosa",),
        "fecha_emision": ("Fecha Emision Doc", "Fecha Emisión Doc"),
        "fecha_vencimiento": ("Fecha Vencimiento Doc", "Fecha Vencimiento Doc."),
        "fecha_movimiento": ("Fecha Movimiento",),
        "fecha_cbr": ("Fecha Cbr", "Fecha Cbr."),
        "fecha_registro": ("Fecha Registro",),
        "monto_debe": ("Monto Debe",),
        "monto_haber": ("Monto Haber",),
        "monto_debe_me": ("Monto Debe ME",),
        "monto_haber_me": ("Monto Haber ME",),
        "cambio_moneda": ("Cambio Moneda",),
    }
    centros_costo: dict[str, tuple[str, ...]] = {
        "cod_centro": (
            "Cod. Centro", "Cod Centro", "Cod. Centro C.", "Cod Centro C",
            "Codigo Centro Costo", "Código Centro Costo", "Codigo Centro", "Código Centro",
            "Cod. Centro Costo", "Cod Centro Costo", "Centro Costo", "Centro de Costo",
        ),
        "cod_sub_centro": (
            "Cod. Sub. Centro", "Cod Sub Centro", "Cod. Sub Centro",
            "Cod. Sub. Centro C.", "Cod Sub Centro C", "Cod. Sub Centro C",
            "Codigo Sub Centro Costo", "Código Sub Centro Costo",
            "Codigo Sub Centro", "Código Sub Centro",
            "Cod. Sub. Centro Costo", "Cod Sub Centro Costo",
            "Sub Centro Costo", "Sub Centro",
        ),
    }

    for numero_fila in range(1, min(hoja.max_row, 30) + 1):
        hallados: dict[str, int] = {}
        for numero_columna in range(1, hoja.max_column + 1):
            valor = hoja.cell(numero_fila, numero_columna).value
            for clave, alias in requeridos.items():
                if clave not in hallados and _encabezado_equivale(valor, *alias):
                    hallados[clave] = numero_columna
                    break
        if set(requeridos) <= set(hallados):
            # Buscar columnas de Centro y Sub Centro de Costo (ej. Cod. Centro C. / Codigo Centro Costo)
            for clave, alias in centros_costo.items():
                if clave not in hallados:
                    for fila_chk in (numero_fila, numero_fila - 1, numero_fila - 2, numero_fila + 1):
                        if 1 <= fila_chk <= hoja.max_row:
                            for numero_columna in range(1, hoja.max_column + 1):
                                valor = hoja.cell(fila_chk, numero_columna).value
                                if _encabezado_equivale(valor, *alias):
                                    hallados[clave] = numero_columna
                                    break
                            if clave in hallados:
                                break
            # Fallback a posiciones estándar en plantilla Contanet si no se ubicaron por encabezado
            if "cod_centro" not in hallados and hoja.max_column >= 20:
                hallados["cod_centro"] = 20
            if "cod_sub_centro" not in hallados and hoja.max_column >= 21:
                hallados["cod_sub_centro"] = 21

            return numero_fila, hallados

    faltantes = ", ".join(requeridos)
    raise ErrorAsientoITFComisiones(
        "La plantilla no contiene todos los encabezados requeridos: " + faltantes
    )


def _copiar_fila_modelo(hoja: Worksheet, fila_modelo: int, fila_destino: int) -> None:
    """Copia estilo, comentarios, vínculos y valores por defecto de la plantilla."""
    hoja.row_dimensions[fila_destino].height = hoja.row_dimensions[fila_modelo].height
    hoja.row_dimensions[fila_destino].hidden = hoja.row_dimensions[fila_modelo].hidden
    for columna in range(1, hoja.max_column + 1):
        origen = hoja.cell(fila_modelo, columna)
        destino = hoja.cell(fila_destino, columna)
        destino.value = origen.value
        if origen.has_style:
            destino._style = copy(origen._style)
        if origen.alignment:
            destino.alignment = copy(origen.alignment)
        if origen.protection:
            destino.protection = copy(origen.protection)
        if origen.comment:
            destino.comment = copy(origen.comment)
        if origen.hyperlink:
            destino._hyperlink = copy(origen.hyperlink)


def _limpiar_importacion(hoja: Worksheet, primera_fila_a_eliminar: int) -> None:
    """Elimina las filas de ejemplo posteriores a la fila modelo."""
    if hoja.max_row >= primera_fila_a_eliminar:
        hoja.delete_rows(
            primera_fila_a_eliminar,
            hoja.max_row - primera_fila_a_eliminar + 1,
        )


_ALIAS_EMPRESA: dict[str, tuple[str, ...]] = {
    "THIMBLE": (
        "THIMBLE", "TST", "THIMBLE SOURCING",
        "THIMBLE SOURCING / TST", "THIMBLE SOURCING/TST",
        "THIMBLE SOURCING_TST", "THIMBLE_SOURCING_TST",
    ),
    "STN": ("STN", "SOUTHERN TEXTIL", "SOUTHERN", "SOUTHERN_TEXTIL"),
    "CMT": ("CMT", "CMT DEL SUR", "CMT_DEL_SUR"),
    "ITS": ("ITS", "INTEGRATED TEXTILE", "INTEGRATED_TEXTILE"),
}


def _resolver_clave_empresa(
    empresa: str | None, empresas_disponibles: Iterable[str]
) -> str | None:
    """Resuelve la clave de empresa en el JSON a partir de cualquier variante o alias."""
    if not empresa:
        return None
    emp_upper = _normalizar(empresa).upper()
    empresas_list = list(empresas_disponibles)

    # 1. Coincidencia exacta insensible a mayúsculas
    for emp_clave in empresas_list:
        if emp_clave.upper() == emp_upper:
            return emp_clave

    # 2. Coincidencia por alias conocidos
    for emp_clave, alias_list in _ALIAS_EMPRESA.items():
        for alias in alias_list:
            norm_alias = _normalizar(alias).upper()
            if norm_alias == emp_upper or norm_alias in emp_upper or emp_upper in norm_alias:
                for candidata in empresas_list:
                    if candidata.upper() == emp_clave:
                        return candidata

    # 3. Substring: si el nombre clave está dentro del nombre dado (ej. "THIMBLE" en "THIMBLE SOURCING / TST")
    for emp_clave in empresas_list:
        if emp_clave.upper() in emp_upper:
            return emp_clave

    # 4. Token matching: si algún token del nombre coincide con la clave
    tokens = [t for t in re.split(r"[_\W]+", emp_upper) if len(t) >= 2]
    for emp_clave in empresas_list:
        if any(t == emp_clave.upper() for t in tokens):
            return emp_clave

    return None


def _cuentas_para(
    banco: str,
    moneda: str,
    cuentas_personalizadas: Mapping[tuple[str, str], Mapping[str, str]] | None = None,
    empresa: str | None = None,
) -> Mapping[str, str]:
    """Devuelve las cuentas contables para un banco, moneda y empresa.

    Orden de búsqueda:
    1. ``cuentas_personalizadas`` si se proporcionan (API de compatibilidad).
    2. ``cuentas_banco.json`` buscando por empresa + banco + moneda.
    3. Diccionario ``CUENTAS`` hardcodeado (fallback legacy).

    Parameters
    ----------
    banco:
        Nombre del banco (ej. ``"BCP"``, ``"SCOTIABANK"``).
    moneda:
        Código de moneda (``"MN"`` o ``"ME"``).
    cuentas_personalizadas:
        Mapa opcional con la misma estructura que ``CUENTAS``.
    empresa:
        Abreviatura o nombre de la empresa (ej. ``"STN"``, ``"CMT"``,
        ``"Thimble Sourcing / TST"``). Si se indica, se busca en
        ``cuentas_banco.json`` antes del fallback.
    """
    banco_normalizado = _normalizar_banco(banco)
    moneda_normalizada = _normalizar_moneda(moneda)

    # ── 1. Buscar configuración en cuentas_banco.json por empresa ─────────
    clave_empresa = None
    cfg_empresa = None
    cc_empresa = ""
    scc_empresa = ""
    if empresa:
        datos_json = cargar_cuentas_banco()
        empresas_json: dict = datos_json.get("empresas", {})
        clave_empresa = _resolver_clave_empresa(empresa, empresas_json.keys())
        if clave_empresa:
            cfg_empresa = empresas_json.get(clave_empresa)
            if cfg_empresa is not None:
                cc_empresa = str(cfg_empresa.get("CC") or cfg_empresa.get("cc") or "").strip()
                scc_empresa = str(cfg_empresa.get("SCC") or cfg_empresa.get("scc") or "").strip()
        else:
            raise ErrorAsientoITFComisiones(
                f"Empresa {empresa!r} no encontrada en cuentas_banco.json. "
                f"Empresas registradas: {list(empresas_json.keys())}"
            )

    # ── 2. cuentas_personalizadas (compatibilidad) ────────────────────────
    if cuentas_personalizadas is not None:
        cuentas = cuentas_personalizadas.get((banco_normalizado, moneda_normalizada))
        if cuentas is not None:
            return _validar_y_extraer_cuentas(
                cuentas, banco_normalizado, moneda_normalizada, cc=cc_empresa, scc=scc_empresa
            )

    # ── 3. cuentas_banco.json (fuente principal) ──────────────────────────
    if cfg_empresa is not None:
        # El JSON usa bancos normalizados: BCP, SCOTIABANK, BN.
        # _normalizar_banco devuelve "SCOTIA" pero el JSON puede tener "SCOTIABANK".
        # Se prueban todos los alias para garantizar la coincidencia.
        _ALIAS_BANCO: dict[str, tuple[str, ...]] = {
            "SCOTIA": ("SCOTIABANK", "SCOTIBANK", "SCOTIA"),
            "SCOTIABANK": ("SCOTIABANK", "SCOTIBANK", "SCOTIA"),
        }
        claves_banco = _ALIAS_BANCO.get(banco_normalizado, (banco_normalizado,))
        cfg_banco = None
        for clave_banco_json in claves_banco:
            cfg_banco = cfg_empresa.get(clave_banco_json)
            if cfg_banco is not None:
                break
        if cfg_banco is not None:
            cfg = cfg_banco.get(moneda_normalizada)  # "MN" o "ME"
            if cfg is not None:
                cuenta_comision = cfg.get("cuenta_comision", "")
                cuenta_itf = cfg.get("cuenta_itf", "")
                cuenta_cce = cfg.get("cuenta_cce", "")  # puede no existir (Scotiabank)
                cuenta_banco = cfg.get("cuenta_banco", "")
                faltantes = [
                    nombre for nombre, valor in {
                        "cuenta_comision": cuenta_comision,
                        "cuenta_itf": cuenta_itf,
                        "cuenta_banco": cuenta_banco,
                    }.items()
                    if not str(valor or "").strip()
                ]
                if faltantes:
                    raise ErrorAsientoITFComisiones(
                        f"Faltan cuentas en cuentas_banco.json para "
                        f"empresa={clave_empresa!r}, banco={banco_normalizado!r}, "
                        f"moneda={moneda_normalizada!r}: " + ", ".join(faltantes)
                    )
                return {
                    "empresa": cuenta_comision,        # cuenta de gasto de comisión
                    "empresa_itf": cuenta_itf,         # cuenta de gasto de ITF
                    "empresa_cce": cuenta_cce,         # cuenta de gasto CCE (puede vacío)
                    "banco": cuenta_banco,             # cuenta bancaria (haber)
                    "cc": cc_empresa,                  # Código Centro de Costo
                    "scc": scc_empresa,                # Código Sub Centro de Costo
                }

    # ── 4. Fallback al dict CUENTAS hardcodeado ───────────────────────────
    cuentas_fallback = CUENTAS.get((banco_normalizado, moneda_normalizada))
    if cuentas_fallback is None:
        # Intentar alias SCOTIA → SCOTIABANK
        alias = {"SCOTIABANK": "SCOTIA", "SCOTIA": "SCOTIABANK"}
        alt = alias.get(banco_normalizado)
        if alt:
            cuentas_fallback = CUENTAS.get((alt, moneda_normalizada))
    if cuentas_fallback is None:
        empresa_msg = f"empresa={empresa!r}, " if empresa else ""
        raise ErrorAsientoITFComisiones(
            f"No hay cuentas configuradas para {empresa_msg}"
            f"banco={banco_normalizado!r}, moneda={moneda_normalizada!r}. "
            f"Agregue la empresa/banco al archivo cuentas_banco.json."
        )
    return _validar_y_extraer_cuentas(
        cuentas_fallback, banco_normalizado, moneda_normalizada, cc=cc_empresa, scc=scc_empresa
    )


def _validar_y_extraer_cuentas(
    cuentas: Mapping[str, str],
    banco_normalizado: str,
    moneda_normalizada: str,
    cc: str = "",
    scc: str = "",
) -> Mapping[str, str]:
    """Valida que el mapa de cuentas tenga los campos obligatorios y los normaliza."""
    empresa_itf = cuentas.get("empresa_itf") or cuentas.get("empresa-ITF")
    faltantes = [
        nombre
        for nombre, valor in {
            "empresa": cuentas.get("empresa"),
            "empresa_itf": empresa_itf,
            "banco": cuentas.get("banco"),
        }.items()
        if not str(valor or "").strip()
    ]
    if faltantes:
        raise ErrorAsientoITFComisiones(
            f"Faltan cuentas para {banco_normalizado} {moneda_normalizada}: "
            + ", ".join(faltantes)
            + ". Complete cuentas_banco.json o use cuentas_personalizadas."
        )
    return {
        "empresa": str(cuentas["empresa"]),
        "empresa_itf": str(empresa_itf),
        "empresa_cce": str(cuentas.get("empresa_cce", cuentas.get("empresa", ""))),
        "banco": str(cuentas["banco"]),
        "cc": str(cuentas.get("cc") or cuentas.get("CC") or cc or ""),
        "scc": str(cuentas.get("scc") or cuentas.get("SCC") or scc or ""),
    }



def _resolver_tc_venta(tipo_cambio_venta: TipoCambioVenta, fecha_cierre: date) -> Decimal:
    """Obtiene y valida el TC Venta para la fecha de cierre solicitada."""
    if callable(tipo_cambio_venta):
        valor = tipo_cambio_venta(fecha_cierre)
    elif isinstance(tipo_cambio_venta, Mapping):
        candidatos: tuple[Any, ...] = (
            fecha_cierre,
            datetime.combine(fecha_cierre, datetime.min.time()),
            fecha_cierre.isoformat(),
            fecha_cierre.strftime("%d/%m/%Y"),
        )
        valor = next(
            (tipo_cambio_venta[candidato] for candidato in candidatos if candidato in tipo_cambio_venta),
            None,
        )
        if valor is None:
            raise TipoCambioNoDisponibleError(
                "Falta el TC Venta de Contanet para el "
                f"{fecha_cierre.strftime('%d/%m/%Y')}."
            )
    else:
        valor = tipo_cambio_venta

    tc_venta = _a_decimal(
        valor,
        campo=f"TC Venta de Contanet para {fecha_cierre.strftime('%d/%m/%Y')}",
    )
    if tc_venta <= 0:
        raise TipoCambioNoDisponibleError(
            f"El TC Venta de Contanet debe ser mayor que cero ({fecha_cierre:%d/%m/%Y})."
        )
    return tc_venta


def _importe_redondeado(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVOS, rounding=ROUND_HALF_UP)


def _montos_asiento(
    monto: Decimal,
    moneda: str,
    tc_venta: Decimal,
) -> tuple[dict[str, Decimal], dict[str, Decimal]]:
    """Devuelve importes de empresa (superior) y banco (inferior), equilibrados."""
    importe = abs(monto)
    monto_mn = (
        _importe_redondeado(importe)
        if moneda == "MN"
        else _importe_redondeado(importe * tc_venta)
    )
    monto_me = (
        _importe_redondeado(importe / tc_venta)
        if moneda == "MN"
        else _importe_redondeado(importe)
    )
    banco = {
        "monto_debe": Decimal("0"),
        "monto_haber": Decimal("0"),
        "monto_debe_me": Decimal("0"),
        "monto_haber_me": Decimal("0"),
    }
    empresa = banco.copy()

    # El banco define el sentido: egreso (negativo) al Haber e ingreso
    # (positivo) al Debe. La empresa registra el sentido opuesto.
    lado_banco = "haber" if monto < 0 else "debe"
    lado_empresa = "debe" if monto < 0 else "haber"
    banco[f"monto_{lado_banco}"] = monto_mn
    banco[f"monto_{lado_banco}_me"] = monto_me
    empresa[f"monto_{lado_empresa}"] = monto_mn
    empresa[f"monto_{lado_empresa}_me"] = monto_me
    return empresa, banco


def _escribir_fila(
    hoja: Worksheet,
    fila: int,
    columnas: Mapping[str, int],
    *,
    correlativo: int,
    relacionado: int,
    fecha: date,
    cuenta: str,
    glosa: str,
    montos: Mapping[str, Decimal],
    tc_venta: Decimal,
    cod_centro: str = "",
    cod_sub_centro: str = "",
) -> None:
    """Escribe únicamente los campos variables definidos por la regla contable."""
    valores: dict[str, Any] = {
        "correlativo": correlativo,
        "relacionado": relacionado,
        "ejercicio": fecha.year,
        "periodo": f"{fecha.month:02d}",
        "nro_cuenta": cuenta,
        "glosa": glosa,
        "fecha_emision": fecha,
        "fecha_vencimiento": fecha,
        "fecha_movimiento": fecha,
        "fecha_cbr": fecha,
        "fecha_registro": fecha,
        "monto_debe": montos["monto_debe"],
        "monto_haber": montos["monto_haber"],
        "monto_debe_me": montos["monto_debe_me"],
        "monto_haber_me": montos["monto_haber_me"],
        "cambio_moneda": tc_venta,
    }
    if cod_centro and "cod_centro" in columnas:
        valores["cod_centro"] = cod_centro
    if cod_sub_centro and "cod_sub_centro" in columnas:
        valores["cod_sub_centro"] = cod_sub_centro

    for campo, valor in valores.items():
        hoja.cell(fila, columnas[campo]).value = valor


def generar_asientos_itf_comisiones(
    ruta_conciliacion_final: str | Path,
    ruta_plantilla: str | Path,
    *,
    banco: str,
    moneda: str,
    tipo_cambio_venta: TipoCambioVenta,
    ruta_salida: str | Path | None = None,
    cuentas_personalizadas: Mapping[tuple[str, str], Mapping[str, str]] | None = None,
    empresa: str | None = None,
) -> Path:
    """Genera y guarda la plantilla de importación de ITF y comisiones.

    Cada movimiento crea dos filas consecutivas (empresa y banco) con correlativo
    secuencial (1, 2, 3...). Ambas filas comparten el número de ``Relacionado`` del
    asiento. Los cinco campos de fecha reciben el último día del mes de la
    operación; los montos se generan en MN y ME usando el TC Venta del cierre.
    Si la plantilla incluye ``ImportCONTABILIDAD``, se sincroniza automáticamente.

    Parameters
    ----------
    empresa:
        Abreviatura de la empresa (ej. ``"STN"``, ``"CMT"``). Se usa para
        buscar las cuentas contables en ``cuentas_banco.json``. Si no se
        indica, se intentará detectar desde el nombre del archivo o se
        utilizan las cuentas del diccionario ``CUENTAS`` (fallback).
    """
    ruta_final = Path(ruta_conciliacion_final).expanduser().resolve()
    plantilla = Path(ruta_plantilla).expanduser().resolve()
    if not plantilla.is_file():
        raise FileNotFoundError(f"No existe la plantilla: {plantilla}")

    # Resolver o auto-detectar empresa usando el catálogo de cuentas_banco.json
    datos_json = cargar_cuentas_banco()
    empresas_registradas = list(datos_json.get("empresas", {}).keys())
    if not empresa:
        empresa = _resolver_clave_empresa(ruta_final.stem, empresas_registradas)
    else:
        empresa = _resolver_clave_empresa(empresa, empresas_registradas) or empresa

    movimientos = leer_movimientos_itf_comisiones(ruta_final)
    moneda_normalizada = _normalizar_moneda(moneda)
    cuentas = _cuentas_para(banco, moneda_normalizada, cuentas_personalizadas, empresa=empresa)
    cod_centro = cuentas.get("cc") or ""
    cod_sub_centro = cuentas.get("scc") or ""

    cierres = {ultimo_dia_mes(movimiento.fecha) for movimiento in movimientos}
    if len(cierres) > 1 and not callable(tipo_cambio_venta) and not isinstance(tipo_cambio_venta, Mapping):
        raise TipoCambioNoDisponibleError(
            "La conciliación contiene más de un mes. Indique un TC Venta de "
            "Contanet por cada cierre mensual mediante un diccionario o función."
        )
    tipos_cambio = {
        cierre: _resolver_tc_venta(tipo_cambio_venta, cierre)
        for cierre in cierres
    }

    if ruta_salida is None:
        ruta_salida = ruta_final.with_name(
            f"Asiento_ITF_Comis_{ruta_final.stem}{plantilla.suffix}"
        )
    salida = Path(ruta_salida).expanduser().resolve()
    if plantilla.suffix.lower() == ".xlsm" and salida.suffix.lower() != ".xlsm":
        raise ErrorAsientoITFComisiones(
            "La plantilla tiene macros (.xlsm); la salida también debe terminar en .xlsm."
        )
    if salida in {plantilla, ruta_final}:
        raise ErrorAsientoITFComisiones(
            "La salida debe ser un archivo distinto de la plantilla y de la conciliación final."
        )
    salida.parent.mkdir(parents=True, exist_ok=True)

    libro = load_workbook(plantilla, keep_vba=plantilla.suffix.lower() == ".xlsm")
    try:
        if NOMBRE_HOJA_DESTINO not in libro.sheetnames:
            raise ErrorAsientoITFComisiones(
                f"La plantilla no tiene la hoja '{NOMBRE_HOJA_DESTINO}'."
            )
        hoja = libro[NOMBRE_HOJA_DESTINO]
        fila_encabezado, columnas = _buscar_fila_encabezados_destino(hoja)
        fila_modelo = fila_encabezado + 1

        # Se deja inicialmente la primera fila de datos como modelo y se
        # eliminan las filas de ejemplo restantes. Antes de escribir en ella,
        # se replica su contenido/formato para todos los asientos necesarios.
        _limpiar_importacion(hoja, fila_modelo + 1)

        # Actualizar fila modelo con los centros de costo de la empresa si existen
        if cod_centro and "cod_centro" in columnas:
            hoja.cell(fila_modelo, columnas["cod_centro"]).value = cod_centro
        if cod_sub_centro and "cod_sub_centro" in columnas:
            hoja.cell(fila_modelo, columnas["cod_sub_centro"]).value = cod_sub_centro

        cantidad_filas = len(movimientos) * 2
        for fila in range(fila_modelo + 1, fila_modelo + cantidad_filas):
            _copiar_fila_modelo(hoja, fila_modelo, fila)

        fila_destino = fila_modelo
        correlativo = 0
        for relacionado, movimiento in enumerate(movimientos, start=1):
            fecha_cierre = ultimo_dia_mes(movimiento.fecha)
            tc_venta = tipos_cambio[fecha_cierre]
            # Seleccionar la cuenta de empresa según el tipo de movimiento:
            # ITF → cuenta_itf, TRANFERENCIA CCE → cuenta_cce (si existe), COMISION → empresa
            if movimiento.tipo == "ITF":
                cuenta_empresa = cuentas["empresa_itf"]
            elif movimiento.tipo == "COMISION" and cuentas.get("empresa_cce"):
                cuenta_empresa = cuentas["empresa_cce"]
            else:
                cuenta_empresa = cuentas["empresa"]

            montos_empresa, montos_banco = _montos_asiento(
                movimiento.monto, moneda_normalizada, tc_venta
            )

            correlativo += 1
            _escribir_fila(
                hoja,
                fila_destino,
                columnas,
                correlativo=correlativo,
                relacionado=relacionado,
                fecha=fecha_cierre,
                cuenta=cuenta_empresa,
                glosa=movimiento.descripcion,
                montos=montos_empresa,
                tc_venta=tc_venta,
                cod_centro=cod_centro,
                cod_sub_centro=cod_sub_centro,
            )
            fila_destino += 1

            correlativo += 1
            _escribir_fila(
                hoja,
                fila_destino,
                columnas,
                correlativo=correlativo,
                relacionado=relacionado,
                fecha=fecha_cierre,
                cuenta=cuentas["banco"],
                glosa=movimiento.descripcion,
                montos=montos_banco,
                tc_venta=tc_venta,
                cod_centro=cod_centro,
                cod_sub_centro=cod_sub_centro,
            )
            fila_destino += 1

        # Si la plantilla contiene la hoja técnica ImportCONTABILIDAD (utilizada
        # por Contanet para validar e importar), se sincroniza automáticamente
        # para que el archivo quede listo sin depender de macros de guardado.
        if NOMBRE_HOJA_IMPORT in libro.sheetnames:
            hoja_import = libro[NOMBRE_HOJA_IMPORT]
            _limpiar_importacion(hoja_import, 2)
            for fila_src in range(fila_modelo, fila_destino):
                fila_dst = fila_src - fila_modelo + 2
                for col_dst in range(1, hoja.max_column - 3 + 1):
                    valor_celda = hoja.cell(fila_src, col_dst + 3).value
                    if valor_celda is not None:
                        hoja_import.cell(fila_dst, col_dst).value = valor_celda

        # CONTABILIDAD es la hoja operativa de la plantilla; se deja visible y
        # activa para que el usuario pueda verificar el asiento antes de subirlo.
        hoja.sheet_state = "visible"
        libro.active = libro.index(hoja)
        libro.save(salida)
    finally:
        libro.close()

    return salida


def _parsear_tipos_cambio(argumentos: list[str]) -> TipoCambioVenta:
    """Convierte los argumentos CLI a un TC único o un mapa por fecha."""
    if not argumentos:
        raise TipoCambioNoDisponibleError(
            "Indique el TC Venta consultado en Contanet con --tc-venta."
        )
    if len(argumentos) == 1 and "=" not in argumentos[0]:
        return _a_decimal(argumentos[0], campo="TC Venta de Contanet")

    tipos_cambio: dict[date, Decimal] = {}
    for argumento in argumentos:
        if "=" not in argumento:
            raise TipoCambioNoDisponibleError(
                "Para varios meses use --tc-venta AAAA-MM-DD=VALOR por cada cierre."
            )
        fecha_texto, valor = argumento.split("=", maxsplit=1)
        fecha_cierre = _a_fecha(fecha_texto)
        if fecha_cierre is None:
            raise TipoCambioNoDisponibleError(
                f"Fecha de TC Venta inválida: {fecha_texto!r}. Use AAAA-MM-DD."
            )
        tipos_cambio[fecha_cierre] = _a_decimal(
            valor, campo=f"TC Venta de Contanet para {fecha_texto}"
        )
    return tipos_cambio


def _crear_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Genera el asiento Contanet de ITF y comisiones desde una conciliación final."
    )
    parser.add_argument("conciliacion_final", help="Archivo CBF con la hoja ITF Y COM.")
    parser.add_argument("plantilla", help="Plantilla Contanet .xlsx o .xlsm.")
    parser.add_argument("--banco", required=True, help="Ej.: BCP, Scotia o BCP Miami.")
    parser.add_argument("--moneda", required=True, help="MN/Soles o ME/Dólares.")
    parser.add_argument(
        "--tc-venta",
        required=True,
        action="append",
        metavar="VALOR|FECHA=VALOR",
        help=(
            "TC Venta obtenido de Contanet. Use un valor para un período, o "
            "AAAA-MM-DD=valor por cada cierre mensual."
        ),
    )
    parser.add_argument("--salida", help="Ruta del .xlsx/.xlsm a generar.")
    parser.add_argument(
        "--empresa",
        help="Abreviatura de la empresa (ej. STN, CMT, ITS, THIMBLE).",
    )
    return parser


def main() -> int:
    """Punto de entrada de consola; imprime la ruta lista para cargar."""
    argumentos = _crear_parser().parse_args()
    try:
        salida = generar_asientos_itf_comisiones(
            argumentos.conciliacion_final,
            argumentos.plantilla,
            banco=argumentos.banco,
            moneda=argumentos.moneda,
            tipo_cambio_venta=_parsear_tipos_cambio(argumentos.tc_venta),
            ruta_salida=argumentos.salida,
            empresa=argumentos.empresa,
        )
    except (ErrorAsientoITFComisiones, FileNotFoundError, PermissionError) as exc:
        print(f"Error: {exc}")
        return 1

    print(f"Archivo creado y cerrado: {salida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
