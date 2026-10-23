"""Reporte Excel con hojas Resumen, Detalle e Incidencias."""

from __future__ import annotations

import io

import pandas as pd

from .auditoria import COLUMNAS_MONEDA, COLUMNAS_MXN, Resultado, describir_config

AZUL = "#2E4057"
ANCHOS = {
    "Archivo": 28, "UUID": 38, "Versión": 8, "Tipo": 22, "Serie": 8, "Folio": 10,
    "Fecha emisión": 18, "Periodo": 10, "RFC emisor": 15, "Nombre emisor": 30,
    "RFC receptor": 15, "Nombre receptor": 30, "Moneda": 8, "Tipo de cambio": 12,
    "Fuente TC": 16, "¿Incluida?": 10, "Motivo de exclusión": 34, "Estado SAT": 14,
    "Código SAT": 30, "Estatus cancelación": 18, "Es cancelable": 18, "Emisor en EFOS": 10,
    "Método de pago": 10, "Uso CFDI": 9, "CFDI relacionados": 30, "Descripción": 50,
    "Observaciones": 40, "Verificar en SAT": 14,
}


def generar_excel(resultado: Resultado) -> bytes:
    salida = io.BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        wb = writer.book
        f = _formatos(wb)
        _hoja_resumen(writer, wb, f, resultado)
        _hoja_tabla(writer, wb, f, "Detalle", resultado.detalle, resultado)
        _hoja_tabla(writer, wb, f, "Incidencias", resultado.incidencias, resultado)
    return salida.getvalue()


def _formatos(wb) -> dict:
    base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
    return {
        "titulo": wb.add_format({**base, "bold": True, "font_size": 16, "font_color": AZUL}),
        "sub": wb.add_format({**base, "bold": True, "font_size": 11, "font_color": "white",
                              "bg_color": AZUL}),
        "encabezado": wb.add_format({**base, "bold": True, "font_color": "white",
                                     "bg_color": AZUL, "text_wrap": True}),
        "etiqueta": wb.add_format({**base, "bold": True, "font_color": "#555555"}),
        "texto": wb.add_format(base),
        "dinero": wb.add_format({**base, "num_format": "#,##0.00"}),
        "dinero_b": wb.add_format({**base, "num_format": "$#,##0.00", "bold": True,
                                   "font_size": 12, "bg_color": "#E3F2FD"}),
        "entero": wb.add_format({**base, "num_format": "#,##0"}),
        "tc": wb.add_format({**base, "num_format": "#,##0.0000"}),
        "fecha": wb.add_format({**base, "num_format": "dd/mm/yyyy hh:mm"}),
        "aviso": wb.add_format({**base, "font_color": "#856404", "bg_color": "#FFF3CD",
                                "text_wrap": True}),
        "excluida": wb.add_format({"font_color": "#7F7F7F", "bg_color": "#F2F2F2"}),
        "cancelada": wb.add_format({"font_color": "#9C0006", "bg_color": "#FFC7CE"}),
        "efos": wb.add_format({"font_color": "#9C5700", "bg_color": "#FFEB9C"}),
    }


def _escribir_tabla(ws, fila: int, df: pd.DataFrame, f: dict, formatos_col: dict) -> int:
    for j, col in enumerate(df.columns):
        ws.write(fila, j, col, f["sub"])
    for i, registro in enumerate(df.itertuples(index=False), 1):
        for j, (col, valor) in enumerate(zip(df.columns, registro)):
            ws.write(fila + i, j, valor, formatos_col.get(col, f["texto"]))
    return fila + len(df) + 2


def _hoja_resumen(writer, wb, f, r: Resultado):
    ws = wb.add_worksheet("Resumen")
    writer.sheets["Resumen"] = ws
    ws.set_column(0, 0, 34)
    ws.set_column(1, 3, 22)
    ws.write(0, 0, "Auditoría de CFDIs y contraprestación", f["titulo"])
    ws.fit_to_pages(1, 0)

    fila = 2
    for etiqueta, valor in describir_config(r.config):
        ws.write(fila, 0, etiqueta, f["etiqueta"])
        ws.write(fila, 1, valor, f["texto"])
        fila += 1

    fila += 1
    ws.merge_range(fila, 0, fila, 1, "Resultado", f["sub"])
    total = len(r.detalle)
    cifras = [
        ("Archivos procesados", total, f["entero"]),
        ("Facturas incluidas en la base", len(r.incluidas), f["entero"]),
        ("Facturas excluidas", total - len(r.incluidas), f["entero"]),
        ("Base total (MXN)", r.base_mxn, f["dinero_b"]),
        ("Contraprestación a cobrar (MXN)", r.contraprestacion, f["dinero_b"]),
    ]
    for etiqueta, valor, formato in cifras:
        fila += 1
        ws.write(fila, 0, etiqueta, f["etiqueta"])
        ws.write(fila, 1, valor, formato)

    for aviso in r.advertencias:
        fila += 2
        ws.merge_range(fila, 0, fila, 3, f"⚠ {aviso}", f["aviso"])
        ws.set_row(fila, 30)

    dinero = {c: f["dinero"] for c in COLUMNAS_MONEDA + COLUMNAS_MXN}
    for titulo, df in [("Por mes de emisión (sólo incluidas)", r.por_periodo),
                       ("Por moneda (sólo incluidas)", r.por_moneda),
                       ("Motivos de exclusión", r.exclusiones),
                       ("Estado ante el SAT", r.por_estado)]:
        if df.empty:
            continue
        fila += 2
        ws.write(fila, 0, titulo, f["etiqueta"])
        fila = _escribir_tabla(ws, fila + 1, df, f, {**dinero, "Facturas": f["entero"]}) - 2


def _hoja_tabla(writer, wb, f, nombre: str, df: pd.DataFrame, r: Resultado):
    df = df.copy()
    df["Fecha emisión"] = pd.to_datetime(df["Fecha emisión"])
    ws = wb.add_worksheet(nombre)
    writer.sheets[nombre] = ws
    filas, cols = len(df), len(df.columns)

    if filas == 0:
        ws.write(0, 0, "Sin registros", f["texto"])
        return

    columnas = []
    for col in df.columns:
        spec = {"header": col, "header_format": f["encabezado"]}
        if col in COLUMNAS_MONEDA + COLUMNAS_MXN:
            spec["format"] = f["dinero"]
        elif col == "Tipo de cambio":
            spec["format"] = f["tc"]
        elif col == "Fecha emisión":
            spec["format"] = f["fecha"]
        if col in COLUMNAS_MXN:
            spec["total_function"] = "sum"
        columnas.append(spec)
    columnas[0]["total_string"] = "Total (filtrado)"
    ws.add_table(0, 0, filas + 1, cols - 1, {
        "data": df.astype(object).where(df.notna(), None).values.tolist(),
        "columns": columnas, "total_row": True, "style": "Table Style Light 9",
        "name": f"Tabla{nombre}",
    })

    # Ligas al verificador del SAT como hipervínculos reales.
    j_liga = df.columns.get_loc("Verificar en SAT")
    for i, url in enumerate(df["Verificar en SAT"], 1):
        if url:
            ws.write_url(i, j_liga, url, string="Abrir")

    for j, col in enumerate(df.columns):
        ws.set_column(j, j, ANCHOS.get(col, 14))

    letra = lambda col: _letra(df.columns.get_loc(col))  # noqa: E731
    rango = f"A2:{_letra(cols - 1)}{filas + 1}"
    ws.conditional_format(rango, {"type": "formula", "format": f["cancelada"],
                                  "criteria": f'=${letra("Estado SAT")}2="Cancelado"'})
    ws.conditional_format(rango, {"type": "formula", "format": f["efos"],
                                  "criteria": f'=${letra("Emisor en EFOS")}2="Sí"'})
    ws.conditional_format(rango, {"type": "formula", "format": f["excluida"],
                                  "criteria": f'=${letra("¿Incluida?")}2="No"'})
    ws.freeze_panes(1, 2)
    ws.set_row(0, 30)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0)


def _letra(indice: int) -> str:
    letras = ""
    indice += 1
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras
