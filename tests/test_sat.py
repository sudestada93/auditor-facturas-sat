from unittest import mock

import requests

from auditor_cfdi import sat
from auditor_cfdi.lector import leer_cfdi

from .fabrica import cfdi40

RESPUESTA_VIGENTE = """<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/"><s:Body>
<ConsultaResponse xmlns="http://tempuri.org/"><ConsultaResult
  xmlns:a="http://schemas.datacontract.org/2004/07/Sat.Cfdi.Negocio.ConsultaCfdi.Servicio">
<a:CodigoEstatus>S - Comprobante obtenido satisfactoriamente.</a:CodigoEstatus>
<a:EsCancelable>Cancelable con aceptación</a:EsCancelable>
<a:Estado>Vigente</a:Estado><a:EstatusCancelacion/>
<a:ValidacionEFOS>200</a:ValidacionEFOS>
</ConsultaResult></ConsultaResponse></s:Body></s:Envelope>""".encode()


def test_expresion_y_liga():
    c = leer_cfdi(cfdi40(rfc_emisor="M&amp;M010101AAA"), "a.xml")
    assert sat.expresion_impresa(c) == (
        "?re=M&amp;M010101AAA&rr=XAXX010101000&tt=1160.00&id=AAAAAAAA-1111-2222-3333-444444444444")
    liga = sat.liga_verificador(c)
    assert liga.startswith(sat.URL_VERIFICADOR + "?re=M%26M010101AAA&rr=")
    assert liga.endswith("&fe=345678%3D%3D")


def test_total_formato_32():
    c = leer_cfdi(cfdi40(total="1160.5"), "a.xml")
    c.version = "3.2"
    assert "&tt=0000001160.500000&" in sat.expresion_impresa(c)


def test_interpretar_respuesta():
    e = sat.interpretar_respuesta(RESPUESTA_VIGENTE)
    assert e.estado == "Vigente" and e.verificado and not e.emisor_en_efos
    assert e.es_cancelable == "Cancelable con aceptación"


def test_consultar_envia_soap_y_reintenta():
    c = leer_cfdi(cfdi40(), "a.xml")
    sesion = mock.Mock()
    ok = mock.Mock(status_code=200, content=RESPUESTA_VIGENTE, text=RESPUESTA_VIGENTE.decode())
    sesion.post.side_effect = [requests.ConnectionError("caído"), ok]
    with mock.patch("auditor_cfdi.sat.time.sleep"):
        e = sat.consultar(c, sesion)
    assert e.estado == "Vigente"
    assert sesion.post.call_count == 2
    cuerpo = sesion.post.call_args.kwargs["data"].decode()
    assert "<tem:expresionImpresa>?re=EKU9003173C9&amp;rr=XAXX010101000&amp;tt=1160.00" in cuerpo


def test_consultar_sin_conexion():
    c = leer_cfdi(cfdi40(), "a.xml")
    sesion = mock.Mock()
    sesion.post.side_effect = requests.Timeout("lento")
    with mock.patch("auditor_cfdi.sat.time.sleep"):
        e = sat.consultar(c, sesion, reintentos=2)
    assert e.estado == "Error de conexión" and not e.verificado


def test_consultar_lote_tolera_excepciones():
    cfdis = [leer_cfdi(cfdi40(uuid=f"U{i}"), f"{i}.xml") for i in range(5)]

    def consultor(c):
        if c.uuid == "U3":
            raise RuntimeError("boom")
        return sat.EstadoSAT("Vigente")

    avances = []
    r = sat.consultar_lote(cfdis, hilos=2, consultor=consultor,
                           al_avanzar=lambda a, t: avances.append((a, t)))
    assert r["U3"].estado == "Error de conexión"
    assert sum(e.estado == "Vigente" for e in r.values()) == 4
    assert avances[-1] == (5, 5)
