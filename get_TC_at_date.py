"""
Módulo para consultar el Tipo de Cambio (TC) desde el archivo histórico local:
\\\\192.168.30.36\\Sig\\Asistentes Contables 2019\\ARCHIVO CONTABLE DIGITAL\\CONCILIACION BANCARIA\\historico_tc.txt

El archivo tiene una entrada por mes (último día natural del mes) con dos
columnas: Fecha (YYYY-MM-DD) y Tipo de Cambio Venta.

Incluye la regla para el cierre de mes:
La consulta del tipo de cambio del último día del mes se realiza directamente
sobre la entrada del mes en el archivo histórico. Si el mes solicitado no
tiene datos (p. ej. el mes en curso sin cerrar), se lanza un ValueError.
"""

import calendar
import datetime
from decimal import Decimal
from pathlib import Path
from typing import Dict, Optional, Set, Union

# ── Ruta del archivo histórico de tipo de cambio ─────────────────────────────
RUTA_HISTORICO_TC = Path(
    r"\\192.168.30.36\Sig\Asistentes Contables 2019"
    r"\ARCHIVO CONTABLE DIGITAL\CONCILIACION BANCARIA\historico_tc.txt"
)


# ── Helpers de días útiles (se conservan por compatibilidad/utilidad) ─────────

def calcular_pascua(year: int) -> datetime.date:
    """Calcula el Domingo de Pascua (Resurrección) mediante el algoritmo de Butcher/Meeus."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return datetime.date(year, month, day)


def obtener_feriados_peru(year: int) -> Set[datetime.date]:
    """
    Retorna el conjunto de feriados nacionales oficiales de Perú para un año dado.
    Incluye feriados de ley fijos y los feriados móviles de Semana Santa (Jueves y Viernes Santo).
    """
    feriados = {
        datetime.date(year, 1, 1),    # Año Nuevo
        datetime.date(year, 5, 1),    # Día del Trabajo
        datetime.date(year, 6, 7),    # Batalla de Arica y Día de la Bandera
        datetime.date(year, 6, 29),   # San Pedro y San Pablo
        datetime.date(year, 7, 23),   # Día de la Fuerza Aérea del Perú
        datetime.date(year, 7, 28),   # Fiestas Patrias
        datetime.date(year, 7, 29),   # Fiestas Patrias
        datetime.date(year, 8, 6),    # Batalla de Junín
        datetime.date(year, 8, 30),   # Santa Rosa de Lima
        datetime.date(year, 10, 8),   # Combate de Angamos
        datetime.date(year, 11, 1),   # Todos los Santos
        datetime.date(year, 12, 8),   # Inmaculada Concepción
        datetime.date(year, 12, 9),   # Batalla de Ayacucho
        datetime.date(year, 12, 25),  # Navidad
    }

    # Feriados móviles de Semana Santa
    pascua = calcular_pascua(year)
    feriados.add(pascua - datetime.timedelta(days=3))  # Jueves Santo
    feriados.add(pascua - datetime.timedelta(days=2))  # Viernes Santo

    return feriados


def es_dia_util(fecha: datetime.date, feriados_extra: Optional[Set[datetime.date]] = None) -> bool:
    """Determina si una fecha dada es un día útil (no sábado, no domingo, no feriado)."""
    # 5 = Sábado, 6 = Domingo
    if fecha.weekday() >= 5:
        return False

    feriados = obtener_feriados_peru(fecha.year)
    if feriados_extra:
        feriados.update(feriados_extra)

    return fecha not in feriados


def obtener_dia_util_anterior_fin_mes(
    anio: int,
    mes: int,
    feriados_extra: Optional[Set[datetime.date]] = None
) -> datetime.date:
    """
    Obtiene el día útil anterior al último día calendario del mes.

    Regla:
    1. Determina el último día calendario del mes (ej. 30 o 31, o 28/29 en febrero).
    2. Retrocede al día inmediatamente anterior.
    3. Si cae en sábado, domingo o feriado, retrocede sucesivamente hasta hallar
       el día laborable/útil inmediatamente anterior.
    """
    ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
    fecha_cursor = datetime.date(anio, mes, ultimo_dia_mes) - datetime.timedelta(days=1)

    feriados = obtener_feriados_peru(fecha_cursor.year)
    if feriados_extra:
        feriados.update(feriados_extra)

    while fecha_cursor.weekday() >= 5 or fecha_cursor in feriados:
        fecha_cursor -= datetime.timedelta(days=1)
        if fecha_cursor.year != anio:
            feriados.update(obtener_feriados_peru(fecha_cursor.year))

    return fecha_cursor


def normalizar_fecha(fecha: Union[str, datetime.date, datetime.datetime]) -> datetime.date:
    """Convierte una entrada de fecha (str DD/MM/YYYY, YYYY-MM-DD o date/datetime) a datetime.date."""
    if isinstance(fecha, datetime.datetime):
        return fecha.date()
    if isinstance(fecha, datetime.date):
        return fecha
    if isinstance(fecha, str):
        fecha = fecha.strip()
        if "/" in fecha:
            # DD/MM/YYYY
            partes = fecha.split("/")
            return datetime.date(int(partes[2]), int(partes[1]), int(partes[0]))
        elif "-" in fecha:
            # YYYY-MM-DD
            partes = fecha.split("-")
            return datetime.date(int(partes[0]), int(partes[1]), int(partes[2]))
    raise ValueError(f"Formato de fecha no reconocido: {fecha}. Use 'DD/MM/YYYY', 'YYYY-MM-DD' o datetime.date.")


# ── Lectura del archivo histórico ─────────────────────────────────────────────

def cargar_historico_tc(ruta: Optional[Path] = None) -> Dict[datetime.date, Decimal]:
    """
    Lee el archivo histórico de tipo de cambio y devuelve un diccionario
    {fecha (date): tc_venta (Decimal)}.

    Ignora encabezados, separadores y líneas con valor 'Sin datos'.

    Parameters
    ----------
    ruta:
        Ruta al archivo TXT. Si es None se usa RUTA_HISTORICO_TC.

    Returns
    -------
    dict
        Diccionario {datetime.date: Decimal} con los registros disponibles.

    Raises
    ------
    FileNotFoundError
        Si el archivo no se encuentra en la ruta indicada.
    """
    if ruta is None:
        ruta = RUTA_HISTORICO_TC

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo histórico de tipo de cambio:\n{ruta}\n"
            "Verifique que la ruta de red esté disponible."
        )

    historico: Dict[datetime.date, Decimal] = {}

    with open(ruta, encoding="utf-8", errors="replace") as f:
        for linea in f:
            linea = linea.strip()
            if not linea:
                continue
            # Ignorar encabezado y separador
            if linea.startswith("Fecha") or linea.startswith("---"):
                continue

            partes = linea.split()
            if len(partes) < 2:
                continue

            fecha_str = partes[0]
            valor_str = partes[-1]

            # Ignorar filas sin datos
            if valor_str.lower() in ("sin", "datos", "sin datos"):
                continue
            # "Sin datos" puede aparecer como múltiples tokens
            if any(p.lower() in ("sin", "datos") for p in partes[1:]):
                continue

            try:
                fecha = datetime.date.fromisoformat(fecha_str)
                tc_venta = Decimal(valor_str)
                historico[fecha] = tc_venta
            except (ValueError, Exception):
                # Línea con formato inesperado: se omite
                continue

    return historico


def consultar_tipo_cambio_historico(
    fecha: Union[str, datetime.date, datetime.datetime],
    ruta_historico: Optional[Path] = None,
) -> Dict[str, Union[str, Decimal]]:
    """
    Obtiene el tipo de cambio de venta para una fecha específica desde el
    archivo histórico local.

    El archivo contiene una entrada por mes (último día natural del mes).
    Si la fecha solicitada no coincide exactamente con ninguna entrada, se busca
    la entrada correspondiente al mismo año/mes.

    Retorna un diccionario con:
    - 'fecha_consultada': fecha en formato YYYY-MM-DD
    - 'fecha_devuelta':   fecha del registro encontrado (YYYY-MM-DD)
    - 'venta':            tipo de cambio venta (Decimal)

    Raises
    ------
    ValueError
        Si no existe un registro para el mes de la fecha solicitada.
    """
    dt = normalizar_fecha(fecha)
    historico = cargar_historico_tc(ruta_historico)

    # Búsqueda exacta por fecha
    if dt in historico:
        return {
            "fecha_consultada": dt.isoformat(),
            "fecha_devuelta": dt.isoformat(),
            "venta": historico[dt],
        }

    # Búsqueda por año y mes (útil cuando se pasa un día distinto al último del mes)
    for fecha_reg, tc in historico.items():
        if fecha_reg.year == dt.year and fecha_reg.month == dt.month:
            return {
                "fecha_consultada": dt.isoformat(),
                "fecha_devuelta": fecha_reg.isoformat(),
                "venta": tc,
            }

    raise ValueError(
        f"No se encontró tipo de cambio en el histórico para "
        f"{dt.year}-{dt.month:02d} ({dt.isoformat()}).\n"
        f"Verifique que el archivo {ruta_historico or RUTA_HISTORICO_TC} "
        f"tenga un registro para ese mes."
    )


def obtener_tipo_cambio_fin_mes(
    anio: int,
    mes: int,
    feriados_extra: Optional[Set[datetime.date]] = None,
    ruta_historico: Optional[Path] = None,
) -> Dict[str, Union[str, Decimal, datetime.date]]:
    """
    Obtiene el tipo de cambio de venta desde el archivo histórico para el
    cierre de un mes dado.

    El valor devuelto corresponde al registro del mes en el histórico
    (que es el último día natural del mes).

    Returns
    -------
    dict con claves:
        - 'fecha_consultada'  : str ISO de la fecha pedida
        - 'fecha_devuelta'    : str ISO del registro encontrado en el histórico
        - 'venta'             : Decimal con el TC de venta
        - 'ultimo_dia_mes'    : str ISO del último día natural del mes
        - 'dia_util_consultado': str ISO del día útil anterior al fin de mes
                                 (informativo; el valor del histórico es el del mes)
    """
    ultimo_dia_mes = calendar.monthrange(anio, mes)[1]
    fecha_fin_mes = datetime.date(anio, mes, ultimo_dia_mes)
    fecha_util_anterior = obtener_dia_util_anterior_fin_mes(anio, mes, feriados_extra)

    resultado = consultar_tipo_cambio_historico(fecha_fin_mes, ruta_historico)
    resultado["ultimo_dia_mes"] = fecha_fin_mes.isoformat()
    resultado["dia_util_consultado"] = fecha_util_anterior.isoformat()
    return resultado


if __name__ == "__main__":
    import sys

    print("=" * 65)
    print("CONSULTA DE TIPO DE CAMBIO - HISTÓRICO LOCAL")
    print("=" * 65)
    print(f"Archivo: {RUTA_HISTORICO_TC}\n")

    hoy = datetime.date.today()
    anio_consulta = hoy.year
    mes_consulta = hoy.month

    print(f"Consultando tipo de cambio fin de mes para {anio_consulta}-{mes_consulta:02d}...")

    try:
        tc = obtener_tipo_cambio_fin_mes(anio_consulta, mes_consulta)
        print("\nResultado obtenido del histórico:")
        print(f"  Último día del mes   : {tc['ultimo_dia_mes']}")
        print(f"  Fecha del registro   : {tc['fecha_devuelta']}")
        print(f"  TC Venta             : {tc['venta']}")
    except Exception as e:
        print(f"\nAviso al consultar el histórico: {e}")
        sys.exit(1)