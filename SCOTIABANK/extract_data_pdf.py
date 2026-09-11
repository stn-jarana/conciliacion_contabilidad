import pymupdf4llm
import re

md = pymupdf4llm.to_markdown('ESTADO DE CUENTA SCOTIABANK STN SOLES 06-2026.pdf')

companies = {
    "20506883301": "CMT", "20514016624": "DYNAMITEX", "20494530865": "DINSURA",
    "20376729126": "STN", "20504334041": "ITS"
}


def get_company():
    doi_match = re.search(r'DOI\|:\s*(\d+)\|', md)

    if doi_match:
        DOI = doi_match.group(1)
        company = companies.get(DOI)

        return (DOI, company)


def cuenta_corriente():
    match = re.search(r"(M\.[NE]\.\s+\w*)\s+No\.\s+([\d\-]+)", md)

    if match:
        moneda = match.group(1)
        cuenta = match.group(2)

        return (moneda, cuenta)
    

def get_movimientos():
    movimientos = []

    for linea in md.splitlines():
        linea = linea.strip()

        if re.match(r"^\|\d{2}/\d{2}\|", linea):
            columnas = [c.strip() for c in linea.strip("|").split("|")]

            cargo = columnas[5]
            cargo = (
                -float(cargo.replace(",", ""))
                if cargo else 0
            )

            movimientos.append({
                "fecha_oper": columnas[0],
                "concepto": columnas[3],
                "operacion": columnas[4],
                "cargo": cargo,
                "abono": columnas[6]
            })


    return movimientos


def get_saldo_final():
    saldo_final = None

    for linea in md.splitlines():
        if "Saldo Final al" in linea:
            columnas = [c.strip() for c in linea.strip("|").split("|")]

            saldo_final = float(columnas[-1].replace(",", ""))

    return saldo_final


print(get_movimientos())
