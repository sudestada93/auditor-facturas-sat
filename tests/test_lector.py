import io
import zipfile
from decimal import Decimal

from auditor_cfdi.lector import expandir_archivos, leer_cfdi

from .fabrica import cfdi40


def test_lee_cfdi_40():
    c = leer_cfdi(cfdi40(descuento="100.00", subtotal="1000.00", total="1044.00",
                         conceptos=("Uno\n  con salto", "Dos")), "a.xml")
    assert c.error is None
    assert c.version == "4.0" and c.tipo == "I"
    assert c.uuid == "AAAAAAAA-1111-2222-3333-444444444444"
    assert c.rfc_emisor == "EKU9003173C9" and c.rfc_receptor == "XAXX010101000"
    assert c.subtotal_neto == Decimal("900.00")
    assert c.total_texto == "1044.00"
    assert c.descripcion == "Uno con salto | Dos"
    assert c.fecha.year == 2026 and c.uso_cfdi == "G03"
    assert c.sello_ultimos8 == "345678=="
    assert c.impuestos_trasladados == Decimal("160.00")


def test_lee_cfdi_32_minusculas():
    xml = b"""<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/3" version="3.2"
      tipoDeComprobante="egreso" fecha="2016-05-01T09:00:00" subTotal="50" total="58"
      Moneda="Pesos">
      <cfdi:Emisor rfc="aaa010101aaa" nombre="X"/><cfdi:Receptor rfc="BBB010101BBB"/>
      <cfdi:Complemento><tfd:TimbreFiscalDigital xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital"
        UUID="bbbbbbbb-1111-2222-3333-444444444444"/></cfdi:Complemento>
    </cfdi:Comprobante>"""
    c = leer_cfdi(xml, "v32.xml")
    assert c.error is None
    assert c.tipo == "E" and c.moneda == "MXN" and c.version == "3.2"
    assert c.rfc_emisor == "AAA010101AAA"
    assert c.uuid == "BBBBBBBB-1111-2222-3333-444444444444"


def test_cfdi_relacionados():
    c = leer_cfdi(cfdi40(tipo="E", relacionados=["cccccccc-0000-0000-0000-000000000000"]), "nc.xml")
    assert c.relacionados == ["01:CCCCCCCC-0000-0000-0000-000000000000"]


def test_errores_de_lectura():
    assert "XML inválido" in leer_cfdi(b"<no cerrado", "x.xml").error
    assert "No es un CFDI" in leer_cfdi(b"<Acuse/>", "x.xml").error
    assert "vacío" in leer_cfdi(b"", "x.xml").error
    sin_timbre = cfdi40().replace(b"<cfdi:Complemento>", b"<cfdi:Otro>").replace(
        b"</cfdi:Complemento>", b"</cfdi:Otro>")
    assert "UUID" in leer_cfdi(sin_timbre, "x.xml").error


def test_xml_con_entidades_maliciosas_no_se_procesa():
    bomba = b"""<?xml version="1.0"?><!DOCTYPE l [<!ENTITY a "aaaa"><!ENTITY b "&a;&a;">]>
    <Comprobante>&b;</Comprobante>"""
    assert leer_cfdi(bomba, "bomba.xml").error


def _zip(archivos: dict[str, bytes]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for nombre, datos in archivos.items():
            z.writestr(nombre, datos)
    return buf.getvalue()


def test_expandir_zip_anidado():
    interno = _zip({"b.xml": cfdi40()})
    externo = _zip({"carpeta/a.xml": cfdi40(), "__MACOSX/._a.xml": b"basura",
                    "leeme.txt": b"hola", "interno.zip": interno})
    nombres = [n for n, _ in expandir_archivos([("descarga.zip", externo), ("suelto.xml", b"x"),
                                                ("foto.pdf", b"")])]
    assert nombres == ["descarga.zip/carpeta/a.xml", "descarga.zip/interno.zip/b.xml", "suelto.xml"]


def test_zip_danado():
    [(nombre, datos)] = list(expandir_archivos([("roto.zip", b"no es zip")]))
    assert nombre == "roto.zip" and datos == b""
