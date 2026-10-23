import openpyxl

from auditor_cfdi.__main__ import main

from .fabrica import cfdi40


def test_cli_genera_reporte(tmp_path, capsys):
    carpeta = tmp_path / "facturas" / "enero"
    carpeta.mkdir(parents=True)
    (carpeta / "a.xml").write_bytes(cfdi40())
    (carpeta / "notas.txt").write_text("se ignora")
    salida = tmp_path / "reporte.xlsx"

    assert main([str(tmp_path / "facturas"), "--desde", "2026-01-01", "--hasta", "2026-01-31",
                 "--sin-sat", "--tasa", "10", "-o", str(salida)]) == 0
    assert "Contraprestación:     $ 100.00" in capsys.readouterr().out
    assert openpyxl.load_workbook(salida).sheetnames == ["Resumen", "Detalle", "Incidencias"]
