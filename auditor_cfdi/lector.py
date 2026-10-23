"""Lectura de CFDIs (versiones 3.2, 3.3 y 4.0) a partir de XML o ZIP."""

from __future__ import annotations

import io
import os
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import Iterable, Iterator

from defusedxml import ElementTree as ET  # protege contra XXE / "billion laughs"

# Límites para no reventar la memoria con ZIPs maliciosos o enormes.
MAX_BYTES_XML = 20 * 1024 * 1024
MAX_BYTES_ZIP_DESCOMPRIMIDO = 1024 * 1024 * 1024

# CFDI 3.2 usa palabras completas; 3.3 y 4.0 usan la clave de una letra.
_TIPOS_32 = {"ingreso": "I", "egreso": "E", "traslado": "T"}

_ALIAS_MONEDA = {
    "MXN": "MXN", "MN": "MXN", "M.N.": "MXN", "PESOS": "MXN", "PESO": "MXN",
    "PESOS MEXICANOS": "MXN", "PESO MEXICANO": "MXN", "NACIONAL": "MXN",
    "USD": "USD", "DOLARES": "USD", "DÓLARES": "USD", "DOLAR": "USD", "DÓLAR": "USD",
    "DOLARES AMERICANOS": "USD", "DÓLARES AMERICANOS": "USD", "US": "USD",
}


@dataclass
class Concepto:
    descripcion: str
    clave_prod_serv: str | None = None
    no_identificacion: str | None = None
    cantidad: Decimal | None = None
    clave_unidad: str | None = None
    unidad: str | None = None
    valor_unitario: Decimal | None = None
    importe: Decimal = Decimal("0")
    descuento: Decimal = Decimal("0")


@dataclass
class CFDI:
    archivo: str
    error: str | None = None
    version: str | None = None
    tipo: str | None = None               # I, E, T, N, P
    serie: str | None = None
    folio: str | None = None
    uuid: str | None = None
    fecha: datetime | None = None
    rfc_emisor: str | None = None
    nombre_emisor: str | None = None
    rfc_receptor: str | None = None
    nombre_receptor: str | None = None
    moneda: str | None = None             # normalizada (MXN, USD, EUR...)
    moneda_original: str | None = None    # tal como viene en el XML
    tipo_cambio: Decimal | None = None
    subtotal: Decimal = Decimal("0")
    descuento: Decimal = Decimal("0")
    total: Decimal = Decimal("0")
    total_texto: str | None = None        # se manda tal cual al SAT
    impuestos_trasladados: Decimal | None = None
    impuestos_retenidos: Decimal | None = None
    metodo_pago: str | None = None
    forma_pago: str | None = None
    uso_cfdi: str | None = None
    sello_ultimos8: str = ""
    descripcion: str = ""
    relacionados: list[str] = field(default_factory=list)  # "TipoRelacion:UUID"
    conceptos: list[Concepto] = field(default_factory=list)

    @property
    def subtotal_neto(self) -> Decimal:
        return self.subtotal - self.descuento


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _attr(elem, nombre: str) -> str | None:
    """Atributo sin distinguir mayúsculas (3.2 usa 'total', 3.3+ usa 'Total')."""
    nombre = nombre.lower()
    for k, v in elem.attrib.items():
        if _local(k).lower() == nombre:
            v = v.strip()
            return v or None
    return None


def _hijos(elem, nombre: str):
    return [h for h in elem if _local(h.tag) == nombre]


def _decimal(valor: str | None) -> Decimal | None:
    if valor is None:
        return None
    try:
        return Decimal(valor.replace(",", "").strip())
    except InvalidOperation:
        return None


def _fecha(valor: str | None) -> datetime | None:
    if not valor:
        return None
    try:
        return datetime.fromisoformat(valor.strip())
    except ValueError:
        return None


def normalizar_moneda(valor: str | None) -> str | None:
    if not valor:
        return None
    limpio = valor.strip().upper()
    return _ALIAS_MONEDA.get(limpio, limpio)


def leer_cfdi(contenido: bytes, archivo: str) -> CFDI:
    """Convierte el XML de un CFDI en un objeto CFDI. Nunca lanza excepción:
    si algo falla, el error queda en `cfdi.error`."""
    cfdi = CFDI(archivo=archivo)
    if not contenido:
        cfdi.error = "Archivo vacío, ZIP dañado o demasiado grande"
        return cfdi
    try:
        root = ET.fromstring(contenido)
    except Exception as e:  # XML mal formado o con entidades peligrosas
        cfdi.error = f"XML inválido: {e}"
        return cfdi

    if _local(root.tag) != "Comprobante":
        cfdi.error = f"No es un CFDI (elemento raíz: {_local(root.tag)})"
        return cfdi

    cfdi.version = _attr(root, "Version")
    tipo = _attr(root, "TipoDeComprobante")
    if tipo:
        cfdi.tipo = _TIPOS_32.get(tipo.lower(), tipo.upper())
    cfdi.serie = _attr(root, "Serie")
    cfdi.folio = _attr(root, "Folio")
    cfdi.fecha = _fecha(_attr(root, "Fecha"))
    cfdi.moneda_original = _attr(root, "Moneda")
    cfdi.moneda = normalizar_moneda(cfdi.moneda_original) or "MXN"  # 3.2 sin Moneda = pesos
    cfdi.tipo_cambio = _decimal(_attr(root, "TipoCambio"))
    cfdi.subtotal = _decimal(_attr(root, "SubTotal")) or Decimal("0")
    cfdi.descuento = _decimal(_attr(root, "Descuento")) or Decimal("0")
    cfdi.total_texto = _attr(root, "Total")
    cfdi.total = _decimal(cfdi.total_texto) or Decimal("0")
    cfdi.metodo_pago = _attr(root, "MetodoPago")
    cfdi.forma_pago = _attr(root, "FormaPago")
    cfdi.sello_ultimos8 = (_attr(root, "Sello") or "")[-8:]

    for hijo in root:
        nombre = _local(hijo.tag)
        if nombre == "Emisor":
            cfdi.rfc_emisor = (_attr(hijo, "Rfc") or "").upper() or None
            cfdi.nombre_emisor = _attr(hijo, "Nombre")
        elif nombre == "Receptor":
            cfdi.rfc_receptor = (_attr(hijo, "Rfc") or "").upper() or None
            cfdi.nombre_receptor = _attr(hijo, "Nombre")
            cfdi.uso_cfdi = _attr(hijo, "UsoCFDI")
        elif nombre == "Conceptos":
            for concepto in _hijos(hijo, "Concepto"):
                cfdi.conceptos.append(Concepto(
                    descripcion=" ".join((_attr(concepto, "Descripcion") or "").split()),
                    clave_prod_serv=_attr(concepto, "ClaveProdServ"),
                    no_identificacion=_attr(concepto, "NoIdentificacion"),
                    cantidad=_decimal(_attr(concepto, "Cantidad")),
                    clave_unidad=_attr(concepto, "ClaveUnidad"),
                    unidad=_attr(concepto, "Unidad"),
                    valor_unitario=_decimal(_attr(concepto, "ValorUnitario")),
                    importe=_decimal(_attr(concepto, "Importe")) or Decimal("0"),
                    descuento=_decimal(_attr(concepto, "Descuento")) or Decimal("0"),
                ))
            cfdi.descripcion = " | ".join(c.descripcion for c in cfdi.conceptos if c.descripcion)
        elif nombre == "Impuestos":
            cfdi.impuestos_trasladados = _decimal(_attr(hijo, "TotalImpuestosTrasladados"))
            cfdi.impuestos_retenidos = _decimal(_attr(hijo, "TotalImpuestosRetenidos"))
        elif nombre == "CfdiRelacionados":
            tipo_rel = _attr(hijo, "TipoRelacion") or ""
            for rel in _hijos(hijo, "CfdiRelacionado"):
                u = _attr(rel, "UUID")
                if u:
                    cfdi.relacionados.append(f"{tipo_rel}:{u.upper()}")
        elif nombre == "Complemento":
            for elem in hijo.iter():
                if _local(elem.tag) == "TimbreFiscalDigital" and not cfdi.uuid:
                    u = _attr(elem, "UUID")
                    cfdi.uuid = u.upper() if u else None

    if not cfdi.uuid:
        cfdi.error = "CFDI sin timbre fiscal (no tiene UUID)"
    return cfdi


def expandir_archivos(archivos: Iterable[tuple[str, bytes]]) -> Iterator[tuple[str, bytes]]:
    """Recibe pares (nombre, contenido) de XMLs o ZIPs y devuelve sólo XMLs.
    Los ZIP se abren (incluso anidados), ignorando carpetas de sistema."""
    for nombre, contenido in archivos:
        bajo = nombre.lower()
        if bajo.endswith(".zip"):
            yield from _expandir_zip(nombre, contenido)
        elif bajo.endswith(".xml"):
            yield nombre, contenido


def _expandir_zip(nombre_zip: str, contenido: bytes) -> Iterator[tuple[str, bytes]]:
    try:
        zf = zipfile.ZipFile(io.BytesIO(contenido))
    except zipfile.BadZipFile:
        yield nombre_zip, b""  # se reportará como XML inválido
        return
    acumulado = 0
    with zf:
        for info in zf.infolist():
            ruta = info.filename
            if info.is_dir() or "__MACOSX" in ruta or os.path.basename(ruta).startswith("."):
                continue
            bajo = ruta.lower()
            if not (bajo.endswith(".xml") or bajo.endswith(".zip")):
                continue
            acumulado += info.file_size
            if info.file_size > MAX_BYTES_XML * (50 if bajo.endswith(".zip") else 1) \
                    or acumulado > MAX_BYTES_ZIP_DESCOMPRIMIDO:
                yield f"{nombre_zip}/{ruta}", b""
                continue
            datos = zf.read(info)
            yield from expandir_archivos([(f"{nombre_zip}/{ruta}", datos)])


def leer_carpeta(carpeta: str) -> list[tuple[str, bytes]]:
    """Lee recursivamente los XML y ZIP de una carpeta local (o de Drive en Colab)."""
    archivos = []
    for raiz, _, nombres in os.walk(carpeta):
        for n in sorted(nombres):
            if n.lower().endswith((".xml", ".zip")):
                ruta = os.path.join(raiz, n)
                with open(ruta, "rb") as f:
                    archivos.append((os.path.relpath(ruta, carpeta), f.read()))
    return archivos
