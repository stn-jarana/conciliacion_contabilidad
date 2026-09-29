from __future__ import annotations

from calendar import monthrange
from datetime import date, datetime
from decimal import Decimal

from conexion import get_connection


def obtener_tipo_cambio(fecha: date | str) -> Decimal:
    """Consulta el TC Venta de Contanet para la fecha indicada.

    Parameters
    ----------
    fecha:
        ``date`` o cadena ISO ``"YYYY-MM-DD"`` / ``"DD/MM/YYYY"``.

    Returns
    -------
    Decimal
        Tipo de cambio de venta (Tipo_Venta) registrado en Contanet.

    Raises
    ------
    ValueError
        Si no se encuentra un tipo de cambio para esa fecha.
    """
    # Normalizar fecha a objeto date
    if isinstance(fecha, datetime):
        fecha_date = fecha.date()
    elif isinstance(fecha, date):
        fecha_date = fecha
    elif isinstance(fecha, str):
        texto = fecha.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                fecha_date = datetime.strptime(texto, fmt).date()
                break
            except ValueError:
                continue
        else:
            raise ValueError(f"Formato de fecha no reconocido: {fecha!r}")
    else:
        raise TypeError(f"Se esperaba date o str, se recibió {type(fecha).__name__!r}")

    conexion = get_connection()
    try:
        cursor = conexion.cursor()
        sql = "SELECT Tipo_Venta FROM dbo.CN_TipoCambio WHERE fecha = ?"
        cursor.execute(sql, fecha_date)
        filas = cursor.fetchall()
        if not filas:
            raise ValueError(
                f"No se encontró tipo de cambio en Contanet para la fecha {fecha_date.isoformat()}."
            )
        valor = filas[0][0]
        return Decimal(str(valor))
    finally:
        conexion.close()


def obtener_tc_ultimo_dia_mes(anio: int, mes: int) -> Decimal:
    """Devuelve el TC Venta del último día natural del mes indicado.

    Parameters
    ----------
    anio, mes:
        Año y mes (1-12) del período de conciliación.

    Returns
    -------
    Decimal
        TC Venta Contanet del último día del mes.
    """
    ultimo_dia = monthrange(anio, mes)[1]
    fecha_cierre = date(anio, mes, ultimo_dia)
    return obtener_tipo_cambio(fecha_cierre)


def import_xlsm_sql():
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
