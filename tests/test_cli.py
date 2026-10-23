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
    assert openpyxl.load_workbook(salida).sheetnames == ["Auditoría"]


def test_cli_solo_verificar(tmp_path, capsys, monkeypatch):
    from auditor_cfdi import sat
    monkeypatch.setattr(sat, "consultar", lambda c, *a, **k: sat.EstadoSAT("Vigente"))
    (tmp_path / "a.xml").write_bytes(cfdi40())
    assert main([str(tmp_path), "--desde", "2026-01-01", "--hasta", "2026-01-31",
                 "--solo-verificar", "-o", str(tmp_path / "r.xlsx")]) == 0
    salida = capsys.readouterr().out
    assert "Vigentes:             1" in salida and "Contraprestación" not in salida
