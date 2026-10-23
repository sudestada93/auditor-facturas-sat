"""Genera XML de CFDI 4.0 sintéticos para las pruebas."""


def cfdi40(uuid="AAAAAAAA-1111-2222-3333-444444444444", tipo="I", fecha="2026-01-15T10:00:00",
           moneda="MXN", tipo_cambio=None, subtotal="1000.00", descuento=None, total="1160.00",
           rfc_emisor="EKU9003173C9", rfc_receptor="XAXX010101000", metodo="PUE",
           conceptos=("Servicio de maniobras",), relacionados=None) -> bytes:
    attrs = ['Version="4.0"', f'TipoDeComprobante="{tipo}"', f'Fecha="{fecha}"',
             f'Moneda="{moneda}"', f'SubTotal="{subtotal}"', f'Total="{total}"',
             'Serie="A"', 'Folio="100"', 'Sello="abcdefghijklmnopqrstuvwxyz12345678=="',
             f'MetodoPago="{metodo}"']
    if tipo_cambio:
        attrs.append(f'TipoCambio="{tipo_cambio}"')
    if descuento:
        attrs.append(f'Descuento="{descuento}"')
    rel = ""
    if relacionados:
        rel = ('<cfdi:CfdiRelacionados TipoRelacion="01">'
               + "".join(f'<cfdi:CfdiRelacionado UUID="{u}"/>' for u in relacionados)
               + "</cfdi:CfdiRelacionados>")
    conc = "".join(f'<cfdi:Concepto Descripcion="{d}" Importe="1"/>' for d in conceptos)
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4"
    xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital" {' '.join(attrs)}>
  {rel}
  <cfdi:Emisor Rfc="{rfc_emisor}" Nombre="EMISOR SA DE CV"/>
  <cfdi:Receptor Rfc="{rfc_receptor}" Nombre="RECEPTOR" UsoCFDI="G03"/>
  <cfdi:Conceptos>{conc}</cfdi:Conceptos>
  <cfdi:Impuestos TotalImpuestosTrasladados="160.00"/>
  <cfdi:Complemento><tfd:TimbreFiscalDigital Version="1.1" UUID="{uuid}"/></cfdi:Complemento>
</cfdi:Comprobante>"""
    return xml.encode("utf-8")
