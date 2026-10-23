"""Uso por línea de comandos (también sirve dentro de Google Colab):

    python -m auditor_cfdi /ruta/a/facturas --desde 2026-01-01 --hasta 2026-01-31
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from decimal import Decimal

from .auditoria import Configuracion, auditar_archivos
from .lector import leer_carpeta
from .reporte import generar_excel


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="auditor_cfdi", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("carpeta", help="Carpeta con XML y/o ZIP (se busca en subcarpetas)")
    p.add_argument("--desde", required=True, type=date.fromisoformat, help="AAAA-MM-DD")
    p.add_argument("--hasta", required=True, type=date.fromisoformat, help="AAAA-MM-DD")
    p.add_argument("--solo-verificar", action="store_true",
                   help="Sólo verificar ante el SAT, sin calcular contraprestación")
    p.add_argument("--solo-periodo", action="store_true",
                   help="Las facturas emitidas fuera del periodo no suman")
    p.add_argument("--tasa", type=Decimal, default=Decimal("5"), help="Porcentaje (default 5)")
    p.add_argument("--base", choices=["neto", "subtotal"], default="neto")
    p.add_argument("--tc", choices=["cfdi", "banxico", "manual"], default="cfdi",
                   help="Fuente del tipo de cambio USD")
    p.add_argument("--tc-manual", type=Decimal)
    p.add_argument("--token-banxico", default=os.environ.get("BANXICO_TOKEN"))
    p.add_argument("--rfc-emisor", help="Sólo cuentan las facturas de este emisor")
    p.add_argument("--sin-sat", action="store_true", help="No consultar el estado ante el SAT")
    p.add_argument("--ignorar-notas-credito", action="store_true")
    p.add_argument("--incluir-no-verificadas", action="store_true")
    p.add_argument("-o", "--salida", help="Archivo .xlsx de salida")
    a = p.parse_args(argv)

    if a.desde > a.hasta:
        p.error("--desde no puede ser posterior a --hasta")
    if a.solo_verificar and a.sin_sat:
        p.error("--solo-verificar no se puede combinar con --sin-sat")
    if a.tc == "manual" and not a.tc_manual:
        p.error("--tc manual requiere --tc-manual")
    if not os.path.isdir(a.carpeta):
        p.error(f"No existe la carpeta {a.carpeta}")

    config = Configuracion(
        fecha_inicio=a.desde, fecha_fin=a.hasta, calcular=not a.solo_verificar,
        solo_periodo=a.solo_periodo, tasa=a.tasa / 100, base=a.base,
        fuente_tc=a.tc, tc_manual=a.tc_manual, banxico_token=a.token_banxico,
        rfc_emisor=a.rfc_emisor, restar_notas_credito=not a.ignorar_notas_credito,
        verificar_sat=not a.sin_sat, incluir_no_verificadas=a.incluir_no_verificadas)

    def progreso(etapa, actual, total):
        print(f"\r{etapa}: {actual}/{total}   ", end="", file=sys.stderr, flush=True)
        if actual == total:
            print(file=sys.stderr)

    resultado = auditar_archivos(leer_carpeta(a.carpeta), config, progreso)
    salida = a.salida or f"Auditoria_{a.desde}_al_{a.hasta}.xlsx"
    with open(salida, "wb") as f:
        f.write(generar_excel(resultado))

    print(f"Archivos procesados:  {len(resultado.detalle)}")
    if config.calcular:
        print(f"Incluidas en la base: {len(resultado.incluidas)}")
        print(f"Base MXN:             $ {resultado.base_mxn:,.2f}")
        print(f"Contraprestación:     $ {resultado.contraprestacion:,.2f}")
    else:
        print(f"Vigentes:             {resultado.contar_estado('Vigente')}")
        print(f"Canceladas:           {resultado.contar_estado('Cancelado')}")
        print(f"No confirmadas:       "
              f"{resultado.contar_estado('No Encontrado') + resultado.sin_verificar}")
    for aviso in resultado.advertencias:
        print(f"⚠ {aviso}")
    print(f"Reporte: {salida}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
