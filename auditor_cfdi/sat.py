"""Consulta del estado de un CFDI en el servicio público del SAT.

Se habla SOAP directamente con `requests` en lugar de usar zeep: así no se
descarga el WSDL en cada consulta (era lo más lento de la versión anterior)
y las consultas pueden hacerse en paralelo con una sola sesión HTTP.
"""

from __future__ import annotations

import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Callable, Iterable
from xml.sax.saxutils import escape

import requests
from defusedxml import ElementTree as ET

from .lector import CFDI

URL_SERVICIO = "https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc"
SOAP_ACTION = "http://tempuri.org/IConsultaCFDIService/Consulta"
URL_VERIFICADOR = "https://verificacfdi.facturaelectronica.sat.gob.mx/default.aspx"

ESTADOS_VERIFICADOS = {"Vigente", "Cancelado", "No Encontrado"}

_SOBRE = (
    '<?xml version="1.0" encoding="utf-8"?>'
    '<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" '
    'xmlns:tem="http://tempuri.org/"><soapenv:Header/><soapenv:Body>'
    "<tem:Consulta><tem:expresionImpresa>{expresion}</tem:expresionImpresa></tem:Consulta>"
    "</soapenv:Body></soapenv:Envelope>"
)


@dataclass
class EstadoSAT:
    estado: str
    codigo: str = ""
    es_cancelable: str = ""
    estatus_cancelacion: str = ""
    validacion_efos: str = ""

    @property
    def verificado(self) -> bool:
        return self.estado in ESTADOS_VERIFICADOS

    @property
    def emisor_en_efos(self) -> bool:
        # 100 = emisor en lista de EFOS (art. 69-B CFF); 200 = no está.
        return self.validacion_efos.strip() == "100"


def _total_para_expresion(cfdi: CFDI) -> str:
    if cfdi.version and cfdi.version.startswith("3.2"):
        # Anexo 20 v3.2: 10 enteros y 6 decimales con ceros a la izquierda.
        return f"{cfdi.total:017.6f}"
    return cfdi.total_texto or str(cfdi.total)


def expresion_impresa(cfdi: CFDI) -> str:
    """Arma la expresión que se envía al servicio SOAP. Los RFC con '&'
    (p. ej. 'M&M') se escapan como '&amp;', que es como el SAT espera recibirlos."""
    re_ = (cfdi.rfc_emisor or "").replace("&", "&amp;")
    rr = (cfdi.rfc_receptor or "").replace("&", "&amp;")
    return f"?re={re_}&rr={rr}&tt={_total_para_expresion(cfdi)}&id={cfdi.uuid_xml or cfdi.uuid}"


def _valor_url(valor: str) -> str:
    # Sólo se codifica lo que rompería la liga ('&', '+', '#', espacios...); '=' y '/'
    # se dejan tal cual, igual que en la liga del código QR impreso.
    return requests.utils.quote(valor, safe="=/")


def liga_verificador(cfdi: CFDI) -> str:
    """Liga al verificador web del SAT con folio fiscal, RFC emisor, RFC receptor y
    total ya capturados, en el mismo orden que el código QR del CFDI. El folio
    fiscal va tal cual viene en el XML (sin pasarlo a mayúsculas)."""
    if not all([cfdi.rfc_emisor, cfdi.rfc_receptor, cfdi.uuid]):
        return ""
    liga = (f"{URL_VERIFICADOR}?id={cfdi.uuid_xml or cfdi.uuid}&re={_valor_url(cfdi.rfc_emisor)}"
            f"&rr={_valor_url(cfdi.rfc_receptor)}"
            f"&tt={_valor_url(cfdi.total_texto or str(cfdi.total))}")
    if cfdi.sello_ultimos8:
        liga += f"&fe={_valor_url(cfdi.sello_ultimos8)}"
    return liga


def interpretar_respuesta(xml: bytes | str) -> EstadoSAT:
    root = ET.fromstring(xml)
    valores: dict[str, str] = {}
    for elem in root.iter():
        nombre = elem.tag.rsplit("}", 1)[-1]
        if nombre in ("CodigoEstatus", "Estado", "EsCancelable",
                      "EstatusCancelacion", "ValidacionEFOS", "faultstring"):
            valores[nombre] = (elem.text or "").strip()
    if "faultstring" in valores and "Estado" not in valores:
        return EstadoSAT("Error del SAT", valores["faultstring"])
    return EstadoSAT(
        estado=valores.get("Estado", "Sin respuesta"),
        codigo=valores.get("CodigoEstatus", ""),
        es_cancelable=valores.get("EsCancelable", ""),
        estatus_cancelacion=valores.get("EstatusCancelacion", ""),
        validacion_efos=valores.get("ValidacionEFOS", ""),
    )


def consultar(cfdi: CFDI, sesion: requests.Session | None = None,
              timeout: float = 20, reintentos: int = 3) -> EstadoSAT:
    if not all([cfdi.rfc_emisor, cfdi.rfc_receptor, cfdi.uuid, cfdi.total_texto]):
        return EstadoSAT("Datos incompletos", "Faltan RFC, UUID o total en el XML")
    sesion = sesion or requests.Session()
    cuerpo = _SOBRE.format(expresion=escape(expresion_impresa(cfdi)))
    encabezados = {"Content-Type": "text/xml; charset=utf-8", "SOAPAction": f'"{SOAP_ACTION}"'}
    ultimo_error = ""
    for intento in range(reintentos):
        try:
            r = sesion.post(URL_SERVICIO, data=cuerpo.encode("utf-8"),
                            headers=encabezados, timeout=timeout)
            if r.status_code < 500 or "Estado" in r.text:
                return interpretar_respuesta(r.content)
            ultimo_error = f"HTTP {r.status_code}"
        except (requests.RequestException, ET.ParseError) as e:
            ultimo_error = str(e)
        if intento < reintentos - 1:
            time.sleep(1.5 * 2 ** intento)
    return EstadoSAT("Error de conexión", ultimo_error)


def consultar_lote(cfdis: Iterable[CFDI], hilos: int = 6,
                   al_avanzar: Callable[[int, int], None] | None = None,
                   consultor: Callable[[CFDI], EstadoSAT] | None = None) -> dict[str, EstadoSAT]:
    """Consulta varios CFDI en paralelo. Devuelve {uuid: EstadoSAT}."""
    cfdis = list(cfdis)
    resultados: dict[str, EstadoSAT] = {}
    if not cfdis:
        return resultados
    if consultor is None:
        sesion = requests.Session()
        adaptador = requests.adapters.HTTPAdapter(pool_connections=hilos, pool_maxsize=hilos)
        sesion.mount("https://", adaptador)
        consultor = lambda c: consultar(c, sesion)  # noqa: E731
    with ThreadPoolExecutor(max_workers=hilos) as pool:
        futuros = {pool.submit(consultor, c): c for c in cfdis}
        for i, fut in enumerate(as_completed(futuros), 1):
            c = futuros[fut]
            try:
                resultados[c.uuid] = fut.result()
            except Exception as e:  # nunca debe tumbar toda la auditoría
                resultados[c.uuid] = EstadoSAT("Error de conexión", str(e))
            if al_avanzar:
                al_avanzar(i, len(cfdis))
    return resultados

