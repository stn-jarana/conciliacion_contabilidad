from __future__ import annotations

from calendar import monthrange
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Optional

from get_TC_at_date import (
    RUTA_HISTORICO_TC,
    cargar_historico_tc,
    normalizar_fecha,
)


def obtener_tipo_cambio(
    fecha: date | str,
    ruta_historico: Optional[Path] = None,
) -> Decimal:
    """Consulta el TC Venta desde el archivo histórico local para la fecha indicada.

    El archivo histórico contiene una entrada por mes (último día natural del
    mes). Si la fecha solicitada no coincide exactamente con ninguna entrada,
    se busca la entrada del mismo año/mes.

    Parameters
    ----------
    fecha:
        ``date`` o cadena ISO ``"YYYY-MM-DD"`` / ``"DD/MM/YYYY"``.
    ruta_historico:
        Ruta alternativa al archivo TXT. Si es ``None`` se usa
        ``RUTA_HISTORICO_TC`` definida en ``get_TC_at_date``.

    Returns
    -------
    Decimal
        Tipo de cambio de venta registrado en el histórico.

    Raises
    ------
    ValueError
        Si no se encuentra un tipo de cambio para ese mes en el histórico.
    FileNotFoundError
        Si el archivo histórico no está disponible en la ruta de red.
    """
    fecha_date = normalizar_fecha(fecha)
    historico = cargar_historico_tc(ruta_historico)

    # Búsqueda exacta
    if fecha_date in historico:
        return historico[fecha_date]

    # Búsqueda por año/mes (el archivo guarda el último día del mes)
    for fecha_reg, tc in historico.items():
        if fecha_reg.year == fecha_date.year and fecha_reg.month == fecha_date.month:
            return tc

    raise ValueError(
        f"No se encontró tipo de cambio en el histórico para "
        f"{fecha_date.year}-{fecha_date.month:02d} ({fecha_date.isoformat()}).\n"
        f"Verifique que el archivo {ruta_historico or RUTA_HISTORICO_TC} "
        f"tenga un registro para ese mes."
    )


def obtener_tc_ultimo_dia_mes(anio: int, mes: int) -> Decimal:
    """Devuelve el TC Venta del último día natural del mes indicado desde el histórico.

    Parameters
    ----------
    anio, mes:
        Año y mes (1-12) del período de conciliación.

    Returns
    -------
    Decimal
        TC Venta del último día del mes según el archivo histórico.
    """
    ultimo_dia = monthrange(anio, mes)[1]
    fecha_cierre = date(anio, mes, ultimo_dia)
    return obtener_tipo_cambio(fecha_cierre)


def import_xlsm_sql():
    from conexion import get_connection
    conexion = get_connection()

    try:
        cursor = conexion.cursor()
        sql = ""
        cursor.execute(sql, )
        pass

    except Exception as e:
        pass
    finally:
        conexion.close()
    pass
