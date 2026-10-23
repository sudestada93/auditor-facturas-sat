"""Reporte Excel de una sola hoja: encabezado, indicadores y tabla de facturas."""

from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import datetime

import pandas as pd

from .auditoria import BASES, FUENTES_TC, Resultado

FUENTE = "Arial"
CARBON, GRAFITO, CENIZA = "#181925", "#666666", "#999999"
NIEBLA, LINO, BRUMA = "#E8E8E8", "#FAFAFA", "#F5F5F5"
LAVANDA, CIELO = "#918DF6", "#2C78FC"
MENTA, MENTA_FONDO = "#1E9E46", "#DEF6E4"
ROJO, ROJO_FONDO = "#C62828", "#FDE7E7"
AMBAR, AMBAR_FONDO = "#9A6400", "#FFF2D6"

FILA_TITULO = 1
FILA_INDICADORES = 5     # etiqueta en esta fila, valor en la siguiente
COLUMNA_INICIAL = 1      # la columna A queda como margen


@dataclass
class Columna:
    titulo: str
    ancho: float
    tipo: str = "texto"   # texto | dinero | tc | fecha | liga | estado
    total: bool = False


COLUMNAS_CALCULO = [
    Columna("Estado SAT", 14, "estado"), Columna("¿Cuenta?", 10),
    Columna("Fecha", 11, "fecha"), Columna("Serie y folio", 14),
    Columna("RFC emisor", 15), Columna("Emisor", 30),
    Columna("RFC receptor", 15), Columna("Receptor", 30),
    Columna("Tipo", 12), Columna("Moneda", 9),
    Columna("Subtotal", 14, "dinero"), Columna("Total", 14, "dinero"),
    Columna("Tipo de cambio", 12, "tc"), Columna("Base MXN", 16, "dinero", True),
    Columna("Contraprestación", 17, "dinero", True), Columna("Notas", 42),
    Columna("Concepto", 40), Columna("UUID", 38), Columna("SAT", 9, "liga"),
    Columna("Archivo", 30),
]

COLUMNAS_VERIFICACION = [
    Columna("Estado SAT", 14, "estado"), Columna("EFOS", 9),
    Columna("Fecha", 11, "fecha"), Columna("Serie y folio", 14),
    Columna("RFC emisor", 15), Columna("Emisor", 30),
    Columna("RFC receptor", 15), Columna("Receptor", 30),
    Columna("Tipo", 12), Columna("Moneda", 9),
    Columna("Total", 14, "dinero"), Columna("Estatus cancelación", 20),
    Columna("Notas", 42), Columna("Concepto", 40), Columna("UUID", 38),
    Columna("SAT", 9, "liga"), Columna("Archivo", 30),
]


def tabla_reporte(resultado: Resultado) -> pd.DataFrame:
    """Tabla compacta (la misma que se ve en la app y en el Excel)."""
    d = resultado.detalle
    calcular = resultado.config.calcular
    notas = []
    for motivo, obs in zip(d["Motivo de exclusión"], d["Observaciones"]):
        if not calcular and not motivo.startswith(("Error de lectura", "UUID duplicado")):
            motivo = ""
        notas.append(" · ".join(t for t in (motivo, obs.replace("; ", " · ")) if t))
    serie_folio = [("-".join(p for p in (s, f) if p)) for s, f in zip(d["Serie"], d["Folio"])]
    t = pd.DataFrame({
        "Estado SAT": d["Estado SAT"].where(d["Estado SAT"] != "", "Sin consultar"),
        "¿Cuenta?": d["¿Incluida?"],
        "EFOS": d["Emisor en EFOS"],
        "Fecha": pd.to_datetime(d["Fecha emisión"]),
        "Serie y folio": serie_folio,
        "RFC emisor": d["RFC emisor"], "Emisor": d["Nombre emisor"],
        "RFC receptor": d["RFC receptor"], "Receptor": d["Nombre receptor"],
        "Tipo": d["Tipo"].str.split(" - ").str[-1].str.replace(" (nota de crédito)", "",
                                                                 regex=False),
        "Moneda": d["Moneda"], "Subtotal": d["Subtotal"], "Total": d["Total"],
        "Tipo de cambio": d["Tipo de cambio"],
        # Las que no cuentan se dejan en blanco para que los totales cuadren.
        "Base MXN": d["Base MXN"].where(d["¿Incluida?"] == "Sí"),
        "Contraprestación": d["Contraprestación MXN"].where(d["¿Incluida?"] == "Sí"),
        "Estatus cancelación": d["Estatus cancelación"],
        "Notas": notas, "Concepto": d["Descripción"], "UUID": d["UUID"],
        "SAT": d["Verificar en SAT"], "Archivo": d["Archivo"],
    })
    columnas = COLUMNAS_CALCULO if calcular else COLUMNAS_VERIFICACION
    return t[[c.titulo for c in columnas]]


def generar_excel(resultado: Resultado) -> bytes:
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("Auditoría")
        writer.sheets["Auditoría"] = ws
        f = _formatos(wb)
        columnas = COLUMNAS_CALCULO if resultado.config.calcular else COLUMNAS_VERIFICACION
        tabla = tabla_reporte(resultado)

        ws.hide_gridlines(2)
        ws.set_column(0, 0, 2)
        for j, col in enumerate(columnas, COLUMNA_INICIAL):
            ws.set_column(j, j, col.ancho)

        fila = _encabezado(ws, f, resultado, len(columnas))
        fila = _indicadores(ws, f, resultado, fila)
        _tabla(ws, f, tabla, columnas, fila + 1)
    return salida.getvalue()


def _formatos(wb) -> dict:
    def fmt(**kw):
        return wb.add_format({"font_name": FUENTE, "font_size": 10, "font_color": CARBON,
                              "valign": "vcenter", **kw})

    celda = dict(bottom=1, bottom_color=NIEBLA)
    return {
        "titulo": fmt(font_size=20, bold=True),
        "subtitulo": fmt(font_color=GRAFITO, font_size=11),
        "parametros": fmt(font_color=CENIZA, font_size=9),
        "ind_etiqueta": fmt(font_color=CENIZA, font_size=9, bg_color=LINO, top=1,
                            top_color=NIEBLA, left=1, left_color=NIEBLA, right=1,
                            right_color=NIEBLA, indent=1),
        "ind_valor": fmt(font_size=16, bold=True, bg_color=LINO, bottom=1, bottom_color=NIEBLA,
                         left=1, left_color=NIEBLA, right=1, right_color=NIEBLA, indent=1,
                         align="left"),
        "ind_dinero": fmt(font_size=16, bold=True, bg_color=LINO, bottom=1, bottom_color=NIEBLA,
                          left=1, left_color=NIEBLA, right=1, right_color=NIEBLA, indent=1,
                          align="left", num_format="$#,##0.00"),
        "ind_etiqueta_dest": fmt(font_color=MENTA, font_size=9, bold=True, bg_color=MENTA_FONDO,
                                 top=1, top_color=MENTA_FONDO, indent=1),
        "ind_dinero_dest": fmt(font_size=16, bold=True, bg_color=MENTA_FONDO, indent=1,
                               align="left", num_format="$#,##0.00"),
        "nota": fmt(font_color=GRAFITO, font_size=9),
        "aviso": fmt(font_color=AMBAR, font_size=9),
        "encabezado": fmt(bold=True, bg_color=BRUMA, bottom=1, bottom_color=CARBON,
                          text_wrap=True, indent=1),
        "encabezado_num": fmt(bold=True, bg_color=BRUMA, bottom=1, bottom_color=CARBON,
                              text_wrap=True, align="right"),
        "texto": fmt(**celda, indent=1),
        "dinero": fmt(**celda, num_format="#,##0.00"),
        "tc": fmt(**celda, num_format="#,##0.0000"),
        "fecha": fmt(**celda, num_format="dd/mm/yyyy", align="left", indent=1),
        "liga": fmt(**celda, font_color=CIELO, underline=1, indent=1),
        "estado": fmt(**celda, bold=True, indent=1),
        "total_etiqueta": fmt(bold=True, top=1, top_color=CARBON),
        "total": fmt(bold=True, top=1, top_color=CARBON, num_format="#,##0.00"),
        "vigente": wb.add_format({"font_color": MENTA, "bg_color": MENTA_FONDO}),
        "cancelado": wb.add_format({"font_color": ROJO, "bg_color": ROJO_FONDO}),
        "pendiente": wb.add_format({"font_color": AMBAR, "bg_color": AMBAR_FONDO}),
        "no_cuenta": wb.add_format({"font_color": CENIZA}),
        "efos": wb.add_format({"font_color": ROJO, "bold": True}),
    }


def _encabezado(ws, f, r: Resultado, n_columnas: int) -> int:
    c = r.config
    ws.set_row(0, 10)
    ws.set_row(FILA_TITULO, 32)
    ws.write(FILA_TITULO, COLUMNA_INICIAL, "Auditoría de CFDIs", f["titulo"])
    modo = "Verificación y cálculo de contraprestación" if c.calcular else "Verificación ante el SAT"
    ws.write(FILA_TITULO + 1, COLUMNA_INICIAL,
             f"{modo}  ·  Periodo {c.fecha_inicio:%d/%m/%Y} – {c.fecha_fin:%d/%m/%Y}  ·  "
             f"Generado {datetime.now():%d/%m/%Y %H:%M}", f["subtitulo"])
    if c.calcular:
        tc = FUENTES_TC[c.fuente_tc] if c.fuente_tc != "manual" else f"manual ({c.tc_manual})"
        partes = [f"Tasa {float(c.tasa * 100):g} %", f"Base: {BASES[c.base]}", f"TC USD: {tc}",
                  "Fuera del periodo: " + ("no suman" if c.solo_periodo else "también suman"),
                  "Notas de crédito: " + ("se restan" if c.restar_notas_credito else "se ignoran")]
        if c.rfc_emisor:
            partes.append(f"Emisor: {c.rfc_emisor}")
        ws.write(FILA_TITULO + 2, COLUMNA_INICIAL, "  ·  ".join(partes), f["parametros"])
    return FILA_INDICADORES


def _indicadores(ws, f, r: Resultado, fila: int) -> int:
    total = len(r.detalle)
    if r.config.calcular:
        tarjetas = [
            ("Facturas", total, None),
            ("Cuentan para el cálculo", len(r.incluidas), None),
            ("Base total MXN", r.base_mxn, "dinero"),
            (f"Contraprestación {float(r.config.tasa * 100):g} %", r.contraprestacion, "destacado"),
            ("Con incidencias", len(r.incidencias), None),
        ]
    else:
        tarjetas = [
            ("Facturas", total, None),
            ("Vigentes", r.contar_estado("Vigente"), None),
            ("Canceladas", r.contar_estado("Cancelado"), None),
            ("No encontradas / sin verificar", r.contar_estado("No Encontrado") + r.sin_verificar,
             None),
            ("Emisor en EFOS", r.en_efos, None),
        ]
    ws.set_row(fila, 20)
    ws.set_row(fila + 1, 30)
    for i, (etiqueta, valor, tipo) in enumerate(tarjetas):
        j = COLUMNA_INICIAL + 2 * i
        if tipo == "destacado":
            fe, fv = f["ind_etiqueta_dest"], f["ind_dinero_dest"]
        else:
            fe, fv = f["ind_etiqueta"], f["ind_dinero"] if tipo == "dinero" else f["ind_valor"]
        ws.merge_range(fila, j, fila, j + 1, etiqueta, fe)
        ws.merge_range(fila + 1, j, fila + 1, j + 1, valor, fv)
    fila += 2

    if r.config.calcular and len(r.por_periodo) > 0:
        meses = "   ·   ".join(
            f"{p}: {n} fact. · ${c:,.2f}"
            for p, n, c in r.por_periodo[["Periodo", "Facturas", "Contraprestación MXN"]]
            .itertuples(index=False))
        fila += 1
        ws.write(fila, COLUMNA_INICIAL, f"Contraprestación por mes de emisión  →  {meses}", f["nota"])
    for aviso in r.advertencias:
        fila += 1
        ws.write(fila, COLUMNA_INICIAL, f"⚠  {aviso}", f["aviso"])
    return fila + 1


def _tabla(ws, f, tabla: pd.DataFrame, columnas: list[Columna], fila: int):
    ini = COLUMNA_INICIAL
    fin = ini + len(columnas) - 1
    ws.set_row(fila, 30)
    for j, col in enumerate(columnas, ini):
        numerica = col.tipo in ("dinero", "tc")
        ws.write(fila, j, col.titulo, f["encabezado_num" if numerica else "encabezado"])

    primera = fila + 1
    for i, registro in enumerate(tabla.itertuples(index=False), primera):
        ws.set_row(i, 18)
        for j, (col, valor) in enumerate(zip(columnas, registro), ini):
            if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
                ws.write_blank(i, j, None, f["texto"])
            elif col.tipo == "liga":
                if valor:
                    ws.write_url(i, j, valor, f["liga"], string="Abrir")
                else:
                    ws.write_blank(i, j, None, f["texto"])
            elif col.tipo == "fecha":
                ws.write_datetime(i, j, valor.to_pydatetime(), f["fecha"])
            else:
                ws.write(i, j, valor, f[col.tipo if col.tipo in f else "texto"])
    ultima = primera + len(tabla) - 1

    if len(tabla) == 0:
        ws.write(primera, ini, "No se encontraron facturas en los archivos.", f["nota"])
        return

    # Totales que respetan los filtros de Excel (SUBTOTAL 109 ignora filas ocultas).
    fila_total = ultima + 1
    ws.write(fila_total, ini, "Total visible", f["total_etiqueta"])
    for j, col in enumerate(columnas, ini):
        if j == ini:
            continue
        if col.total:
            letra = _letra(j)
            ws.write_formula(fila_total, j, f"=SUBTOTAL(109,{letra}{primera + 1}:{letra}{ultima + 1})",
                             f["total"], float(tabla[col.titulo].fillna(0).sum()))
        else:
            ws.write_blank(fila_total, j, None, f["total_etiqueta"])

    rango = f"{_letra(ini)}{primera + 1}:{_letra(fin)}{ultima + 1}"
    estado = f"${_letra(ini)}{primera + 1}"
    titulos = [c.titulo for c in columnas]
    col_estado = f"{_letra(ini)}{primera + 1}:{_letra(ini)}{ultima + 1}"
    for valor, formato in (("Vigente", "vigente"), ("Cancelado", "cancelado")):
        ws.conditional_format(col_estado, {"type": "formula", "criteria": f'={estado}="{valor}"',
                                           "format": f[formato]})
    ws.conditional_format(col_estado, {"type": "formula",
                                       "criteria": f'=AND({estado}<>"Vigente",{estado}<>"Cancelado")',
                                       "format": f["pendiente"]})
    if "¿Cuenta?" in titulos:
        cuenta = f"${_letra(ini + titulos.index('¿Cuenta?'))}{primera + 1}"
        ws.conditional_format(rango, {"type": "formula", "criteria": f'={cuenta}="No"',
                                      "format": f["no_cuenta"]})
    if "EFOS" in titulos:
        j = ini + titulos.index("EFOS")
        ws.conditional_format(primera, j, ultima, j, {"type": "cell", "criteria": "==",
                                                       "value": '"Sí"', "format": f["efos"]})

    ws.autofilter(fila, ini, ultima, fin)
    ws.freeze_panes(fila + 1, ini + 1)
    ws.set_landscape()
    ws.set_paper(1)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(fila)
    ws.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)


def _letra(indice: int) -> str:
    letras = ""
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras
