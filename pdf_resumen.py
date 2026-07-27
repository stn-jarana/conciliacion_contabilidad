"""
pdf_resumen.py
──────────────
Genera un PDF con el resumen de las operaciones del reporte inicial
de conciliación bancaria.

Contenido del PDF:
  - Encabezado con empresa, moneda y fecha de generación
  - Resumen financiero (totales banco vs contabilidad)
  - Tabla de partidas con diferencia (Diferencia != 0)
  - Tabla de partidas conciliadas (Diferencia == 0)
  - Pie de página con nombre del archivo Excel generado

Uso:
    from pdf_resumen import generar_pdf_resumen
    generar_pdf_resumen(bank, conta, anexar1, resumen_df, ruta_pdf, empresa, moneda)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak,
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT


# ══════════════════════════════════════════════════════════════════════
# COLORES
# ══════════════════════════════════════════════════════════════════════
AZUL_OSCURO  = colors.HexColor('#1F4E79')
AZUL_MEDIO   = colors.HexColor('#2E75B6')
VERDE_OSCURO = colors.HexColor('#375623')
NARANJA      = colors.HexColor('#C55A11')
MORADO       = colors.HexColor('#7030A0')
GRIS_CLARO   = colors.HexColor('#F2F2F2')
GRIS_MEDIO   = colors.HexColor('#D9D9D9')
ROJO         = colors.HexColor('#C00000')
VERDE        = colors.HexColor('#375623')
BLANCO       = colors.white
NEGRO        = colors.black


# ══════════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════════

def generar_pdf_resumen(
    bank: pd.DataFrame,
    conta: pd.DataFrame,
    anexar1: pd.DataFrame,
    resumen_df: pd.DataFrame,
    ruta_pdf: Path,
    empresa: str = "Southern Textil",
    moneda: str  = "Dolares (USD)",
    ruta_excel: Path | None = None,
) -> None:
    """Genera el PDF de resumen junto al reporte inicial Excel."""

    doc = SimpleDocTemplate(
        str(ruta_pdf),
        pagesize=landscape(A4),
        leftMargin=1.5*cm,
        rightMargin=1.5*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm,
        title="Resumen Conciliacion Bancaria",
    )

    estilos = _estilos()
    historia = []

    # ── Encabezado ────────────────────────────────────────────────────
    historia += _encabezado(estilos, empresa, moneda, ruta_excel)
    historia.append(Spacer(1, 0.4*cm))
    historia.append(HRFlowable(width="100%", thickness=2, color=AZUL_OSCURO))
    historia.append(Spacer(1, 0.4*cm))

    # ── Resumen financiero ────────────────────────────────────────────
    historia += _seccion_resumen_financiero(estilos, bank, conta)
    historia.append(Spacer(1, 0.5*cm))

    # ── Estadísticas de conciliación ──────────────────────────────────
    historia += _seccion_estadisticas(estilos, resumen_df)
    historia.append(Spacer(1, 0.5*cm))

    # ── Partidas con diferencia ───────────────────────────────────────
    partidas_abiertas = resumen_df[resumen_df['Diferencia'].abs() > 0.01].copy()
    historia += _seccion_tabla(
        estilos,
        titulo="Partidas con diferencia (requieren revision del especialista)",
        df=partidas_abiertas,
        color_cabecera=NARANJA,
        nota=f"Total: {len(partidas_abiertas)} partidas  |  "
             f"Monto pendiente banco: {partidas_abiertas['Monto-Banco'].fillna(0).sum():,.2f}  |  "
             f"Monto pendiente conta: {partidas_abiertas['Monto-Conta'].fillna(0).sum():,.2f}",
    )
    historia.append(Spacer(1, 0.5*cm))

    # ── Partidas ya conciliadas ───────────────────────────────────────
    partidas_ok = resumen_df[resumen_df['Diferencia'].abs() <= 0.01].copy()
    historia += _seccion_tabla(
        estilos,
        titulo="Partidas conciliadas automaticamente (Diferencia = 0)",
        df=partidas_ok,
        color_cabecera=VERDE_OSCURO,
        nota=f"Total: {len(partidas_ok)} partidas  |  "
             f"Monto banco: {partidas_ok['Monto-Banco'].fillna(0).sum():,.2f}  |  "
             f"Monto conta: {partidas_ok['Monto-Conta'].fillna(0).sum():,.2f}",
        max_filas=50,
    )

    # ── Pie de página informativo ─────────────────────────────────────
    historia.append(Spacer(1, 0.5*cm))
    historia.append(HRFlowable(width="100%", thickness=1, color=GRIS_MEDIO))
    historia.append(Spacer(1, 0.2*cm))
    pie = f"Documento generado automaticamente el {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    if ruta_excel:
        pie += f"  |  Archivo Excel: {ruta_excel.name}"
    historia.append(Paragraph(pie, estilos['pie']))

    doc.build(historia)
    print(f"  [OK] Resumen PDF generado   : {ruta_pdf}")


# ══════════════════════════════════════════════════════════════════════
# SECCIONES
# ══════════════════════════════════════════════════════════════════════

def _encabezado(estilos, empresa, moneda, ruta_excel):
    elementos = []
    elementos.append(Paragraph("CONCILIACION BANCARIA", estilos['titulo']))
    elementos.append(Paragraph("Reporte Inicial — Resumen de Operaciones", estilos['subtitulo']))
    elementos.append(Spacer(1, 0.3*cm))

    datos_meta = [
        ["Empresa:",  empresa,  "Moneda:",  moneda],
        ["Fecha:",    datetime.now().strftime('%d/%m/%Y'),
         "Hora:",     datetime.now().strftime('%H:%M:%S')],
    ]
    t = Table(datos_meta, colWidths=[3*cm, 8*cm, 3*cm, 8*cm])
    t.setStyle(TableStyle([
        ('FONTNAME',    (0,0), (-1,-1), 'Helvetica'),
        ('FONTNAME',    (0,0), (0,-1),  'Helvetica-Bold'),
        ('FONTNAME',    (2,0), (2,-1),  'Helvetica-Bold'),
        ('FONTSIZE',    (0,0), (-1,-1), 9),
        ('TEXTCOLOR',   (0,0), (0,-1),  AZUL_OSCURO),
        ('TEXTCOLOR',   (2,0), (2,-1),  AZUL_OSCURO),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('TOPPADDING',    (0,0), (-1,-1), 3),
    ]))
    elementos.append(t)
    return elementos


def _seccion_resumen_financiero(estilos, bank, conta):
    elementos = []
    elementos.append(Paragraph("Resumen Financiero", estilos['seccion']))

    total_b_ing = bank['ingreso'].sum()
    total_b_egr = bank['egreso'].sum()
    total_c_ing = conta['ingreso'].sum()
    total_c_egr = conta['egreso'].sum()
    dif_ing     = total_b_ing - total_c_ing
    dif_egr     = total_b_egr - total_c_egr

    def fmt(v):
        return f"{v:,.2f}"

    def color_dif(v):
        return fmt(v)

    cabecera = [["Concepto", "Banco", "Contabilidad", "Diferencia"]]
    filas = [
        ["Total Ingresos", fmt(total_b_ing), fmt(total_c_ing), color_dif(dif_ing)],
        ["Total Egresos",  fmt(total_b_egr), fmt(total_c_egr), color_dif(dif_egr)],
        ["Movimientos",    str(len(bank)),   str(len(conta)),  ""],
    ]

    t = Table(
        cabecera + filas,
        colWidths=[6*cm, 6*cm, 6*cm, 6*cm],
    )
    t.setStyle(_estilo_tabla_base(AZUL_OSCURO))

    # Colorear diferencias
    for fila_idx, (dif) in enumerate([dif_ing, dif_egr], start=1):
        col_color = ROJO if abs(dif) > 0.01 else VERDE
        t.setStyle(TableStyle([('TEXTCOLOR', (3, fila_idx), (3, fila_idx), col_color)]))

    elementos.append(t)
    return elementos


def _seccion_estadisticas(estilos, resumen_df):
    elementos = []
    elementos.append(Spacer(1, 0.3*cm))
    elementos.append(Paragraph("Estadisticas de Conciliacion", estilos['seccion']))

    total        = len(resumen_df)
    conciliadas  = len(resumen_df[resumen_df['Diferencia'].abs() <= 0.01])
    abiertas     = total - conciliadas
    pct          = (conciliadas / total * 100) if total > 0 else 0

    filas = [
        ["Total de partidas en Resumen",            str(total)],
        ["Partidas conciliadas (Diferencia = 0)",   str(conciliadas)],
        ["Partidas con diferencia (abiertas)",       str(abiertas)],
        ["Porcentaje conciliado",                    f"{pct:.1f}%"],
    ]

    t = Table(filas, colWidths=[10*cm, 4*cm])
    t.setStyle(TableStyle([
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica'),
        ('FONTNAME',      (0,0), (0,-1),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,-1), 9),
        ('BACKGROUND',    (0,0), (-1,0),  GRIS_CLARO),
        ('ROWBACKGROUNDS',(0,0), (-1,-1), [BLANCO, GRIS_CLARO]),
        ('GRID',          (0,0), (-1,-1), 0.5, GRIS_MEDIO),
        ('ALIGN',         (1,0), (1,-1),  'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING',    (0,0), (-1,-1), 4),
        ('TEXTCOLOR',     (1,2), (1,2),   NARANJA if abiertas > 0 else VERDE),
        ('FONTNAME',      (1,2), (1,2),   'Helvetica-Bold'),
    ]))
    elementos.append(t)
    return elementos


def _seccion_tabla(estilos, titulo, df, color_cabecera, nota="", max_filas=None):
    elementos = []
    elementos.append(Spacer(1, 0.3*cm))
    elementos.append(Paragraph(titulo, estilos['seccion']))

    if df.empty:
        elementos.append(Paragraph("No hay registros en esta seccion.", estilos['normal']))
        return elementos

    # Limitar filas si se especifica
    mostrar = df.head(max_filas) if max_filas else df
    truncado = max_filas and len(df) > max_filas

    # Formatear columnas numéricas
    df_fmt = mostrar.copy()
    for col in ['Monto-Conta', 'Monto-Banco', 'Diferencia']:
        if col in df_fmt.columns:
            df_fmt[col] = df_fmt[col].apply(
                lambda v: f"{v:,.2f}" if pd.notna(v) else ""
            )

    cabecera = [list(df_fmt.columns)]
    datos    = df_fmt.values.tolist()

    # Calcular anchos proporcionales
    n_cols   = len(df_fmt.columns)
    ancho_pg = 25.7*cm  # A4 landscape - márgenes
    col_w    = [ancho_pg / n_cols] * n_cols

    t = Table(cabecera + datos, colWidths=col_w, repeatRows=1)
    t.setStyle(_estilo_tabla_base(color_cabecera))
    elementos.append(t)

    if truncado:
        nota += f"  (mostrando primeras {max_filas} de {len(df)} filas)"
    if nota:
        elementos.append(Spacer(1, 0.15*cm))
        elementos.append(Paragraph(nota, estilos['nota']))

    return elementos


# ══════════════════════════════════════════════════════════════════════
# ESTILOS
# ══════════════════════════════════════════════════════════════════════

def _estilos() -> dict:
    base = getSampleStyleSheet()
    return {
        'titulo'   : ParagraphStyle('titulo',    parent=base['Normal'],
                                    fontSize=16, fontName='Helvetica-Bold',
                                    textColor=AZUL_OSCURO, alignment=TA_CENTER,
                                    spaceAfter=4),
        'subtitulo': ParagraphStyle('subtitulo', parent=base['Normal'],
                                    fontSize=11, fontName='Helvetica',
                                    textColor=AZUL_MEDIO, alignment=TA_CENTER,
                                    spaceAfter=6),
        'seccion'  : ParagraphStyle('seccion',   parent=base['Normal'],
                                    fontSize=10, fontName='Helvetica-Bold',
                                    textColor=AZUL_OSCURO, spaceAfter=4),
        'normal'   : ParagraphStyle('normal',    parent=base['Normal'],
                                    fontSize=8,  fontName='Helvetica'),
        'nota'     : ParagraphStyle('nota',      parent=base['Normal'],
                                    fontSize=7,  fontName='Helvetica',
                                    textColor=colors.grey, alignment=TA_RIGHT),
        'pie'      : ParagraphStyle('pie',       parent=base['Normal'],
                                    fontSize=7,  fontName='Helvetica',
                                    textColor=colors.grey, alignment=TA_CENTER),
    }


def _estilo_tabla_base(color_cabecera) -> TableStyle:
    return TableStyle([
        # Cabecera
        ('BACKGROUND',    (0,0), (-1,0),  color_cabecera),
        ('TEXTCOLOR',     (0,0), (-1,0),  BLANCO),
        ('FONTNAME',      (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',      (0,0), (-1,0),  8),
        ('ALIGN',         (0,0), (-1,0),  'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,0),  5),
        ('TOPPADDING',    (0,0), (-1,0),  5),
        # Datos
        ('FONTNAME',      (0,1), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,1), (-1,-1), 7),
        ('ROWBACKGROUNDS',(0,1), (-1,-1), [BLANCO, GRIS_CLARO]),
        ('GRID',          (0,0), (-1,-1), 0.5, GRIS_MEDIO),
        ('ALIGN',         (1,1), (-1,-1), 'RIGHT'),
        ('ALIGN',         (0,1), (0,-1),  'LEFT'),
        ('BOTTOMPADDING', (0,1), (-1,-1), 3),
        ('TOPPADDING',    (0,1), (-1,-1), 3),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
    ])
