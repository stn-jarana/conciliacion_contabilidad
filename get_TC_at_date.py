""" Obtiene el tipo de cambio de la fecha indicada directamente de la SBS. El TC corresponde al día útil anterior al día solicitado """

import requests
from bs4 import BeautifulSoup

URL = "https://www.sbs.gob.pe/app/pp/SISTIP_PORTAL/Paginas/Publicacion/TipoCambioPromedio.aspx"

session = requests.Session()

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0"
    )
}

# 1. Obtener página inicial
r = session.get(URL, headers=headers)
r.raise_for_status()

soup = BeautifulSoup(r.text, "html.parser")

viewstate = soup.find("input", {"name": "__VIEWSTATE"})["value"]
eventvalidation = soup.find("input", {"name": "__EVENTVALIDATION"})["value"]
viewstategenerator = soup.find("input", {"name": "__VIEWSTATEGENERATOR"})["value"]

# Fecha que deseas consultar
fecha_texto = "21/07/2026"
fecha_iso = "2026-07-21"

payload = {
    "ctl00$MainScriptManager":
        "ctl00$cphContent$updConsulta|ctl00$cphContent$btnConsultar",

    "__EVENTTARGET": "",
    "__EVENTARGUMENT": "",
    "__VIEWSTATE": viewstate,
    "__VIEWSTATEGENERATOR": viewstategenerator,
    "__EVENTVALIDATION": eventvalidation,

    "ctl00$cphContent$rdpDate": fecha_iso,
    "ctl00$cphContent$rdpDate$dateInput": fecha_texto,

    "__ASYNCPOST": "true",
    "ctl00$cphContent$btnConsultar": "Consultar",
}

payload["ctl00_cphContent_rdpDate_dateInput_ClientState"] = (
    '{"enabled":true,'
    f'"validationText":"{fecha_iso}-00-00-00",'
    f'"valueAsString":"{fecha_iso}-00-00-00",'
    '"minDateStr":"1000-01-01-00-00-00",'
    '"maxDateStr":"2026-09-23-00-00-00",'
    f'"lastSetTextBoxValue":"{fecha_texto}"'
    '}'
)

yyyy, mm, dd = fecha_iso.split("-")

payload["ctl00_cphContent_rdpDate_calendar_SD"] = (
    f"[[{yyyy},{int(mm)},{int(dd)}]]"
)

payload["ctl00_cphContent_rdpDate_calendar_AD"] = (
    "[[1000,1,1],[2026,9,23],[2026,8,1]]"
)

payload["ctl00_cphContent_rdpDate_ClientState"] = (
    '{"minDateStr":"1000-01-01-00-00-00",'
    '"maxDateStr":"2026-09-23-00-00-00"}'
)

for k, v in payload.items():
    if "rdpDate" in k:
        print(k, "=", v)

headers_post = {
    **headers,
    "X-MicrosoftAjax": "Delta=true",
    "X-Requested-With": "XMLHttpRequest",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Referer": URL,
    "Origin": "https://www.sbs.gob.pe",
}

r2 = session.post(
    URL,
    data=payload,
    headers=headers_post
)

import re

m = re.search(
    r"Tipo de Cambio al (\d{2}/\d{2}/\d{4})",
    r2.text
)



respuesta = r2.text

fecha = re.search(
    r"Tipo de Cambio al (\d{2}/\d{2}/\d{4})",
    respuesta
)

print(fecha.group(1))

print("Fecha enviada:", fecha_texto)
print("Fecha devuelta:", fecha.group(1))

fila = re.search(
    r"Dólar de N\.A\..*?(\d+\.\d+).*?(\d+\.\d+)",
    respuesta,
    re.S
)

compra = fila.group(1)
venta = fila.group(2)

print(
    fecha.group(1),
    compra,
    venta
)

print("Dólar de N.A." in r2.text)
print(r.status_code)
print(r2.status_code)