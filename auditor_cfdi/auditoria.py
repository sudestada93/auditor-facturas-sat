"""Reglas de la auditoría y del cálculo de la contraprestación."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from typing import Callable, Iterable

import pandas as pd

from . import sat
from .lector import CFDI, expandir_archivos, leer_cfdi
from .tipo_cambio import ErrorTipoCambio, TablaFIX, descargar_fix

CENTAVO = Decimal("0.01")
MAX_DESCRIPCION = 2000  # Excel no acepta celdas de más de 32,767 caracteres

TIPOS = {"I": "Ingreso", "E": "Egreso (nota de crédito)", "P": "Pago",
         "T": "Traslado", "N": "Nómina"}

FUENTES_TC = {
    "cfdi": "El que declara cada CFDI (atributo TipoCambio)",
    "banxico": "FIX de Banxico publicado en el DOF el día de emisión",
    "manual": "Un tipo de cambio único para todo el periodo",
}

BASES = {
    "neto": "Subtotal − Descuento",
    "subtotal": "Subtotal (sin restar descuentos)",
}

COLUMNAS_MONEDA = ["Subtotal", "Descuento", "Total", "Base moneda original"]
COLUMNAS_MXN = ["Base MXN", "Contraprestación MXN"]


@dataclass
class Configuracion:
    fecha_inicio: date
    fecha_fin: date
    tasa: Decimal = Decimal("0.05")
    base: str = "neto"
    fuente_tc: str = "cfdi"
    tc_manual: Decimal | None = None
    banxico_token: str | None = None
    rfc_emisor: str | None = None          # si se indica, sólo cuentan sus facturas
    restar_notas_credito: bool = True
    verificar_sat: bool = True
    incluir_no_verificadas: bool = False   # si el SAT no respondió, ¿se cuentan?
    hilos_sat: int = 6


@dataclass
class Resultado:
    config: Configuracion
    detalle: pd.DataFrame
    por_periodo: pd.DataFrame
    por_estado: pd.DataFrame
    por_moneda: pd.DataFrame
    exclusiones: pd.DataFrame
    advertencias: list[str] = field(default_factory=list)

    @property
    def incluidas(self) -> pd.DataFrame:
        return self.detalle[self.detalle["¿Incluida?"] == "Sí"]

    @property
    def incidencias(self) -> pd.DataFrame:
        d = self.detalle
        return d[(d["¿Incluida?"] == "No") | (d["Observaciones"] != "")]

    @property
    def base_mxn(self) -> float:
        return float(self.incluidas["Base MXN"].sum())

    @property
    def contraprestacion(self) -> float:
        return float(self.incluidas["Contraprestación MXN"].sum())


Progreso = Callable[[str, int, int], None]


def _redondear(valor: Decimal) -> Decimal:
    return valor.quantize(CENTAVO, rounding=ROUND_HALF_UP)


def auditar_archivos(archivos: Iterable[tuple[str, bytes]], config: Configuracion,
                     al_avanzar: Progreso | None = None,
                     consultor_sat: Callable[[CFDI], sat.EstadoSAT] | None = None) -> Resultado:
    """Punto de entrada: recibe pares (nombre, contenido) de XML o ZIP."""
    avanzar = al_avanzar or (lambda *_: None)
    xmls = list(expandir_archivos(archivos))
    cfdis = []
    for i, (nombre, contenido) in enumerate(xmls, 1):
        cfdis.append(leer_cfdi(contenido, nombre))
        avanzar("Leyendo XML", i, len(xmls))
    return auditar(cfdis, config, avanzar, consultor_sat)


def auditar(cfdis: list[CFDI], config: Configuracion, al_avanzar: Progreso | None = None,
            consultor_sat: Callable[[CFDI], sat.EstadoSAT] | None = None) -> Resultado:
    avanzar = al_avanzar or (lambda *_: None)
    advertencias: list[str] = []

    # Duplicados: el mismo UUID puede venir en varios ZIP o carpetas.
    primero: dict[str, str] = {}
    duplicado_de: dict[int, str] = {}
    for i, c in enumerate(cfdis):
        if c.error or not c.uuid:
            continue
        if c.uuid in primero:
            duplicado_de[i] = primero[c.uuid]
        else:
            primero[c.uuid] = c.archivo

    # Estado ante el SAT: sólo se consultan los que podrían contar.
    estados: dict[str, sat.EstadoSAT] = {}
    if config.verificar_sat:
        por_consultar = [c for i, c in enumerate(cfdis)
                         if not c.error and i not in duplicado_de and _es_relevante(c, config)]
        estados = sat.consultar_lote(por_consultar, hilos=config.hilos_sat,
                                     al_avanzar=lambda a, t: avanzar("Consultando SAT", a, t),
                                     consultor=consultor_sat)
        fallidas = sum(1 for e in estados.values() if not e.verificado)
        if fallidas:
            advertencias.append(
                f"{fallidas} factura(s) no pudieron verificarse ante el SAT "
                f"({'se incluyeron' if config.incluir_no_verificadas else 'se excluyeron'} del cálculo). "
                "Vuelve a correr la auditoría más tarde o revísalas con la liga del verificador.")

    tabla_fix = None
    if config.fuente_tc == "banxico":
        fechas = [c.fecha.date() for c in cfdis if c.fecha and c.moneda == "USD"]
        if fechas:
            if not config.banxico_token:
                advertencias.append("No se indicó token de Banxico: se usó el tipo de cambio de cada CFDI.")
            else:
                try:
                    tabla_fix = descargar_fix(config.banxico_token, min(fechas), max(fechas))
                except ErrorTipoCambio as e:
                    advertencias.append(f"{e} Se usó el tipo de cambio de cada CFDI.")

    filas = []
    for i, c in enumerate(cfdis):
        estado = estados.get(c.uuid) if c.uuid and i not in duplicado_de else None
        filas.append(_fila(c, config, estado, duplicado_de.get(i), tabla_fix))
        avanzar("Calculando", i + 1, len(cfdis))

    detalle = pd.DataFrame(filas, columns=list(_fila_vacia().keys()))
    return _resumir(config, detalle, advertencias)


def _es_relevante(c: CFDI, config: Configuracion) -> bool:
    if c.tipo not in _tipos_que_cuentan(config):
        return False
    if config.rfc_emisor and c.rfc_emisor != config.rfc_emisor.strip().upper():
        return False
    return bool(c.fecha) and config.fecha_inicio <= c.fecha.date() <= config.fecha_fin


def _tipos_que_cuentan(config: Configuracion) -> set[str]:
    return {"I", "E"} if config.restar_notas_credito else {"I"}


def _fila_vacia() -> dict:
    return {
        "Archivo": "", "UUID": "", "Versión": "", "Tipo": "", "Serie": "", "Folio": "",
        "Fecha emisión": None, "Periodo": "",
        "RFC emisor": "", "Nombre emisor": "", "RFC receptor": "", "Nombre receptor": "",
        "Moneda": "", "Subtotal": None, "Descuento": None, "Total": None,
        "Base moneda original": None, "Tipo de cambio": None, "Fuente TC": "",
        "Base MXN": None, "Contraprestación MXN": None,
        "¿Incluida?": "No", "Motivo de exclusión": "",
        "Estado SAT": "", "Código SAT": "", "Estatus cancelación": "", "Es cancelable": "",
        "Emisor en EFOS": "",
        "Método de pago": "", "Uso CFDI": "", "CFDI relacionados": "",
        "Descripción": "", "Observaciones": "", "Verificar en SAT": "",
    }


def _resolver_tc(c: CFDI, config: Configuracion, tabla_fix: TablaFIX | None,
                 observaciones: list[str]) -> tuple[Decimal | None, str]:
    if c.moneda == "MXN":
        return Decimal("1"), "MXN"
    if c.moneda == "USD":
        if config.fuente_tc == "manual" and config.tc_manual:
            return config.tc_manual, "Manual"
        if config.fuente_tc == "banxico" and tabla_fix is not None and c.fecha:
            tc = tabla_fix.para(c.fecha.date())
            if tc is not None:
                return tc, "Banxico FIX (DOF)"
            observaciones.append("Sin FIX de Banxico para la fecha; se usó el TC del CFDI")
    if c.tipo_cambio:
        return c.tipo_cambio, "CFDI"
    return None, ""


def _fila(c: CFDI, config: Configuracion, estado: sat.EstadoSAT | None,
          duplicado_de: str | None, tabla_fix: TablaFIX | None) -> dict:
    fila = _fila_vacia()
    fila["Archivo"] = c.archivo
    if c.error:
        fila["Motivo de exclusión"] = f"Error de lectura: {c.error}"
        return fila

    obs: list[str] = []
    fila.update({
        "UUID": c.uuid, "Versión": c.version or "", "Serie": c.serie or "", "Folio": c.folio or "",
        "Tipo": f"{c.tipo} - {TIPOS.get(c.tipo, 'Desconocido')}" if c.tipo else "",
        "Fecha emisión": c.fecha, "Periodo": c.fecha.strftime("%Y-%m") if c.fecha else "",
        "RFC emisor": c.rfc_emisor or "", "Nombre emisor": c.nombre_emisor or "",
        "RFC receptor": c.rfc_receptor or "", "Nombre receptor": c.nombre_receptor or "",
        "Moneda": c.moneda or "", "Subtotal": float(c.subtotal), "Descuento": float(c.descuento),
        "Total": float(c.total), "Método de pago": c.metodo_pago or "", "Uso CFDI": c.uso_cfdi or "",
        "CFDI relacionados": ", ".join(c.relacionados),
        "Descripción": c.descripcion if len(c.descripcion) <= MAX_DESCRIPCION
                       else c.descripcion[:MAX_DESCRIPCION] + "…",
        "Verificar en SAT": sat.liga_verificador(c),
    })

    if estado is not None:
        fila.update({"Estado SAT": estado.estado, "Código SAT": estado.codigo,
                     "Estatus cancelación": estado.estatus_cancelacion,
                     "Es cancelable": estado.es_cancelable,
                     "Emisor en EFOS": "Sí" if estado.emisor_en_efos else
                                       ("No" if estado.validacion_efos else "")})
        if estado.emisor_en_efos:
            obs.append("Emisor en lista de EFOS (art. 69-B CFF)")
        if estado.estado == "Vigente" and "proceso" in estado.estatus_cancelacion.lower():
            obs.append("Cancelación en proceso")

    # Base y conversión a pesos
    signo = {"I": 1, "E": -1}.get(c.tipo or "", 0)
    if c.tipo == "E" and not config.restar_notas_credito:
        signo = 0
    base = (c.subtotal_neto if config.base == "neto" else c.subtotal) * signo
    tc, fuente = _resolver_tc(c, config, tabla_fix, obs)
    fila["Base moneda original"] = float(base)
    fila["Tipo de cambio"] = float(tc) if tc is not None else None
    fila["Fuente TC"] = fuente
    base_mxn = _redondear(base * tc) if tc is not None else None
    fila["Base MXN"] = float(base_mxn) if base_mxn is not None else None
    if c.moneda not in ("MXN", "USD"):
        obs.append(f"Moneda {c.moneda_original or c.moneda}: revisar tipo de cambio")
    if c.tipo == "I" and c.metodo_pago == "PPD":
        obs.append("Pago en parcialidades o diferido (PPD)")

    # Motivo de exclusión: se reporta el primero que aplique.
    motivo = ""
    fecha_ok = c.fecha and config.fecha_inicio <= c.fecha.date() <= config.fecha_fin
    if duplicado_de:
        motivo = f"UUID duplicado (ya incluido en {duplicado_de})"
    elif c.tipo not in _tipos_que_cuentan(config):
        motivo = f"Tipo {fila['Tipo'] or 'desconocido'}: no forma parte de la base"
    elif config.rfc_emisor and c.rfc_emisor != config.rfc_emisor.strip().upper():
        motivo = "Emitida por un RFC distinto al auditado"
    elif not c.fecha:
        motivo = "Fecha de emisión inválida"
    elif not fecha_ok:
        motivo = "Fuera del periodo auditado"
    elif estado is not None and estado.estado == "Cancelado":
        motivo = "Cancelada ante el SAT"
    elif estado is not None and estado.estado == "No Encontrado":
        motivo = "No encontrada en el SAT"
    elif estado is not None and not estado.verificado and not config.incluir_no_verificadas:
        motivo = f"No se pudo verificar ante el SAT ({estado.estado})"
    elif base_mxn is None:
        motivo = "Sin tipo de cambio para convertir a MXN"

    if estado is None and config.verificar_sat is False and not motivo:
        obs.append("No verificada ante el SAT")
    elif estado is not None and not estado.verificado and not motivo:
        obs.append("Incluida sin verificar ante el SAT")

    fila["Motivo de exclusión"] = motivo
    fila["¿Incluida?"] = "No" if motivo else "Sí"
    fila["Contraprestación MXN"] = (
        float(_redondear(base_mxn * config.tasa)) if not motivo and base_mxn is not None else 0.0)
    fila["Observaciones"] = "; ".join(obs)
    return fila


def _resumir(config: Configuracion, detalle: pd.DataFrame, advertencias: list[str]) -> Resultado:
    inc = detalle[detalle["¿Incluida?"] == "Sí"]

    por_periodo = (inc.groupby("Periodo", as_index=False)
                   .agg(Facturas=("UUID", "count"), **{"Base MXN": ("Base MXN", "sum"),
                        "Contraprestación MXN": ("Contraprestación MXN", "sum")}))
    por_moneda = (inc.groupby("Moneda", as_index=False)
                  .agg(Facturas=("UUID", "count"),
                       **{"Base moneda original": ("Base moneda original", "sum"),
                          "Base MXN": ("Base MXN", "sum")}))
    estados = detalle["Estado SAT"].replace("", "Sin consultar")
    por_estado = (estados.value_counts().rename_axis("Estado SAT").reset_index(name="Facturas"))
    motivos = detalle.loc[detalle["Motivo de exclusión"] != "", "Motivo de exclusión"]
    motivos = motivos.str.replace(r"^Error de lectura:.*", "Error de lectura", regex=True)
    motivos = motivos.str.replace(r"^UUID duplicado.*", "UUID duplicado", regex=True)
    exclusiones = motivos.value_counts().rename_axis("Motivo").reset_index(name="Facturas")

    return Resultado(config, detalle, por_periodo, por_estado, por_moneda, exclusiones, advertencias)


def describir_config(config: Configuracion) -> list[tuple[str, str]]:
    tc = FUENTES_TC[config.fuente_tc]
    if config.fuente_tc == "manual" and config.tc_manual:
        tc += f" ({config.tc_manual})"
    return [
        ("Periodo auditado", f"{config.fecha_inicio:%d/%m/%Y} al {config.fecha_fin:%d/%m/%Y}"),
        ("Tasa de contraprestación", f"{float(config.tasa * 100):g} %"),
        ("Base de cálculo", BASES[config.base]),
        ("Tipo de cambio USD", tc),
        ("RFC emisor auditado", config.rfc_emisor or "Todos"),
        ("Notas de crédito", "Se restan de la base" if config.restar_notas_credito else "Se ignoran"),
        ("Verificación SAT", "Sí" if config.verificar_sat else "No"),
        ("Facturas sin verificar", "Se incluyen" if config.incluir_no_verificadas else "Se excluyen"),
        ("Generado", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
