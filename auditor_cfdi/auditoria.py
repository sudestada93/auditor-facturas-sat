"""Reglas de la auditoría y del cálculo de la contraprestación."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from typing import Callable, Iterable

import pandas as pd

from . import sat
from .lector import CFDI, Concepto, expandir_archivos, leer_cfdi
from .tipo_cambio import ErrorTipoCambio, TablaFIX, descargar_fix

CENTAVO = Decimal("0.01")
MAX_DESCRIPCION = 2000  # Excel no acepta celdas de más de 32,767 caracteres
MAX_LINEAS_CONCEPTOS = 25

TIPOS = {"I": "Ingreso", "E": "Egreso (nota de crédito)", "P": "Pago",
         "T": "Traslado", "N": "Nómina"}

FUENTES_TC = {
    "cfdi": "El de cada CFDI (TipoCambio)",
    "banxico": "FIX de Banxico (DOF) del día",
    "manual": "Uno manual para todo el periodo",
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
    calcular: bool = True                  # False = sólo verificar ante el SAT
    solo_periodo: bool = False             # True = las facturas fuera del periodo no suman
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
    conceptos: pd.DataFrame = field(default_factory=pd.DataFrame)  # una fila por concepto

    @property
    def incluidas(self) -> pd.DataFrame:
        return self.detalle[self.detalle["¿Incluida?"] == "Sí"]

    @property
    def incidencias(self) -> pd.DataFrame:
        d = self.detalle
        if not self.config.calcular:
            return d[(d["Estado SAT"] != "Vigente") | (d["Observaciones"] != "")]
        return d[(d["¿Incluida?"] == "No") | (d["Observaciones"] != "")]

    def contar_estado(self, estado: str) -> int:
        return int((self.detalle["Estado SAT"] == estado).sum())

    @property
    def sin_verificar(self) -> int:
        d = self.detalle
        return int(((d["Estado SAT"] != "") & ~d["Estado SAT"].isin(sat.ESTADOS_VERIFICADOS)).sum())

    @property
    def errores_lectura(self) -> int:
        return int(self.detalle["Motivo de exclusión"].str.startswith("Error de lectura").sum())

    @property
    def en_efos(self) -> int:
        return int((self.detalle["Emisor en EFOS"] == "Sí").sum())

    @property
    def base_mxn(self) -> float:
        return float(self.incluidas["Base MXN"].sum())

    @property
    def total_vigente_mxn(self) -> float:
        d = self.detalle
        return float(d.loc[d["Estado SAT"] == "Vigente", "Total MXN"].sum())

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
            efecto = ""
            if config.calcular:
                efecto = (" (se incluyeron en el cálculo)" if config.incluir_no_verificadas
                          else " (se excluyeron del cálculo)")
            advertencias.append(
                f"{fallidas} factura(s) no pudieron verificarse ante el SAT{efecto}. "
                "Vuelve a correr la auditoría más tarde o revísalas con la liga del verificador.")

    tabla_fix = None
    if config.calcular and config.fuente_tc == "banxico":
        fechas = [c.fecha.date() for c in cfdis if c.fecha and c.moneda == "USD"]
        if fechas:
            if not config.banxico_token:
                advertencias.append("No se indicó token de Banxico: se usó el tipo de cambio de cada CFDI.")
            else:
                try:
                    tabla_fix = descargar_fix(config.banxico_token, min(fechas), max(fechas))
                except ErrorTipoCambio as e:
                    advertencias.append(f"{e} Se usó el tipo de cambio de cada CFDI.")

    filas, conceptos = [], []
    for i, c in enumerate(cfdis):
        estado = estados.get(c.uuid) if c.uuid and i not in duplicado_de else None
        fila = _fila(c, config, estado, duplicado_de.get(i), tabla_fix)
        filas.append(fila)
        if not c.error and i not in duplicado_de:
            conceptos.extend(_filas_conceptos(c, fila))
        avanzar("Calculando", i + 1, len(cfdis))

    detalle = pd.DataFrame(filas, columns=list(_fila_vacia().keys()))
    resultado = _resumir(config, detalle, advertencias)
    resultado.conceptos = pd.DataFrame(conceptos, columns=COLUMNAS_CONCEPTOS)
    return resultado


COLUMNAS_CONCEPTOS = [
    "Estado SAT", "¿Cuenta?", "Fecha", "Serie y folio", "RFC emisor", "Emisor", "Receptor",
    "Clave SAT", "No. identificación", "Descripción", "Cantidad", "Unidad", "Valor unitario",
    "Importe", "Descuento", "Moneda", "Tipo de cambio", "Importe neto MXN", "UUID", "Archivo",
]


def _cantidad(valor: Decimal | None) -> str:
    return format(valor.normalize(), "f") if valor is not None else ""


def _linea_concepto(k: Concepto, moneda: str) -> str:
    unidad = k.unidad or k.clave_unidad or ""
    detalle = f"{_cantidad(k.cantidad)} {unidad}".strip()
    if k.valor_unitario is not None:
        detalle += f" × {k.valor_unitario:,.2f}"
    linea = f"• {k.descripcion or 'Sin descripción'} — {detalle} = {k.importe:,.2f} {moneda}"
    if k.descuento:
        linea += f" (desc. {k.descuento:,.2f})"
    if k.clave_prod_serv:
        linea += f" · clave {k.clave_prod_serv}"
    return linea


def _texto_conceptos(c: CFDI) -> str:
    lineas = [_linea_concepto(k, c.moneda or "") for k in c.conceptos[:MAX_LINEAS_CONCEPTOS]]
    if len(c.conceptos) > MAX_LINEAS_CONCEPTOS:
        lineas.append(f"… y {len(c.conceptos) - MAX_LINEAS_CONCEPTOS} concepto(s) más")
    texto = "\n".join(lineas)
    return texto if len(texto) <= MAX_DESCRIPCION else texto[:MAX_DESCRIPCION] + "…"


def _filas_conceptos(c: CFDI, fila: dict) -> list[dict]:
    tc = fila["Tipo de cambio"]
    signo = -1 if c.tipo == "E" else 1
    serie_folio = "-".join(p for p in (c.serie, c.folio) if p)
    return [{
        "Estado SAT": fila["Estado SAT"], "¿Cuenta?": fila["¿Incluida?"], "Fecha": c.fecha,
        "Serie y folio": serie_folio, "RFC emisor": c.rfc_emisor or "",
        "Emisor": c.nombre_emisor or "", "Receptor": c.nombre_receptor or "",
        "Clave SAT": k.clave_prod_serv or "", "No. identificación": k.no_identificacion or "",
        "Descripción": k.descripcion, "Cantidad": float(k.cantidad) if k.cantidad is not None else None,
        "Unidad": k.unidad or k.clave_unidad or "",
        "Valor unitario": float(k.valor_unitario) if k.valor_unitario is not None else None,
        "Importe": float(k.importe), "Descuento": float(k.descuento), "Moneda": c.moneda or "",
        "Tipo de cambio": tc,
        "Importe neto MXN": (float(_redondear((k.importe - k.descuento) * Decimal(str(tc)) * signo))
                             if tc is not None else None),
        "UUID": c.uuid_xml, "Archivo": c.archivo,
    } for k in c.conceptos]


def _es_relevante(c: CFDI, config: Configuracion) -> bool:
    if not config.calcular:
        return True  # en modo verificación se consulta todo
    if c.tipo not in _tipos_que_cuentan(config):
        return False
    if config.rfc_emisor and c.rfc_emisor != config.rfc_emisor.strip().upper():
        return False
    if config.solo_periodo:
        return _en_periodo(c, config)
    return True


def _en_periodo(c: CFDI, config: Configuracion) -> bool:
    return bool(c.fecha) and config.fecha_inicio <= c.fecha.date() <= config.fecha_fin


def _tipos_que_cuentan(config: Configuracion) -> set[str]:
    return {"I", "E"} if config.restar_notas_credito else {"I"}


def _fila_vacia() -> dict:
    return {
        "Archivo": "", "UUID": "", "Versión": "", "Tipo": "", "Serie": "", "Folio": "",
        "Fecha emisión": None, "Periodo": "", "En periodo": "",
        "RFC emisor": "", "Nombre emisor": "", "RFC receptor": "", "Nombre receptor": "",
        "Moneda": "", "Subtotal": None, "Descuento": None, "Total": None,
        "Base moneda original": None, "Tipo de cambio": None, "Fuente TC": "",
        "Subtotal MXN": None, "Total MXN": None,
        "Base MXN": None, "Contraprestación MXN": None,
        "¿Incluida?": "No", "Motivo de exclusión": "",
        "Estado SAT": "", "Código SAT": "", "Estatus cancelación": "", "Es cancelable": "",
        "Emisor en EFOS": "",
        "Método de pago": "", "Uso CFDI": "", "CFDI relacionados": "",
        "Descripción": "", "Conceptos": "", "Núm. conceptos": 0,
        "Observaciones": "", "Verificar en SAT": "",
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
        "UUID": c.uuid_xml, "Versión": c.version or "", "Serie": c.serie or "", "Folio": c.folio or "",
        "Tipo": f"{c.tipo} - {TIPOS.get(c.tipo, 'Desconocido')}" if c.tipo else "",
        "Fecha emisión": c.fecha, "Periodo": c.fecha.strftime("%Y-%m") if c.fecha else "",
        "En periodo": "Sí" if _en_periodo(c, config) else "No",
        "RFC emisor": c.rfc_emisor or "", "Nombre emisor": c.nombre_emisor or "",
        "RFC receptor": c.rfc_receptor or "", "Nombre receptor": c.nombre_receptor or "",
        "Moneda": c.moneda or "", "Subtotal": float(c.subtotal), "Descuento": float(c.descuento),
        "Total": float(c.total), "Método de pago": c.metodo_pago or "", "Uso CFDI": c.uso_cfdi or "",
        "CFDI relacionados": ", ".join(c.relacionados),
        "Descripción": c.descripcion if len(c.descripcion) <= MAX_DESCRIPCION
                       else c.descripcion[:MAX_DESCRIPCION] + "…",
        "Verificar en SAT": sat.liga_verificador(c),
        "Conceptos": _texto_conceptos(c), "Núm. conceptos": len(c.conceptos),
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
    if tc is not None and not duplicado_de:
        # Importes en pesos (las notas de crédito en negativo) para sumar en ambos modos.
        signo_importe = -1 if c.tipo == "E" else 1
        fila["Subtotal MXN"] = float(_redondear(c.subtotal * tc * signo_importe))
        fila["Total MXN"] = float(_redondear(c.total * tc * signo_importe))
    if config.calcular and signo and c.moneda not in ("MXN", "USD"):
        obs.append(f"Moneda {c.moneda_original or c.moneda}: revisar tipo de cambio")
    if config.calcular and c.tipo == "I" and c.metodo_pago == "PPD":
        obs.append("Pago en parcialidades o diferido (PPD)")

    # Motivo de exclusión: se reporta el primero que aplique.
    motivo = ""
    fecha_ok = _en_periodo(c, config)
    if duplicado_de:
        motivo = f"UUID duplicado (ya incluido en {duplicado_de})"
    elif c.tipo not in _tipos_que_cuentan(config):
        motivo = f"Tipo {fila['Tipo'] or 'desconocido'}: no forma parte de la base"
    elif config.rfc_emisor and c.rfc_emisor != config.rfc_emisor.strip().upper():
        motivo = "Emitida por un RFC distinto al auditado"
    elif config.solo_periodo and not c.fecha:
        motivo = "Fecha de emisión inválida"
    elif config.solo_periodo and not fecha_ok:
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
    elif config.calcular and estado is not None and not estado.verificado and not motivo:
        obs.append("Incluida sin verificar ante el SAT")

    if not c.fecha:
        obs.append("Fecha de emisión inválida")
    elif not fecha_ok and not config.solo_periodo:
        obs.append("Emitida fuera del periodo")

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
    periodo = ("Periodo auditado", f"{config.fecha_inicio:%d/%m/%Y} al {config.fecha_fin:%d/%m/%Y}")
    if not config.calcular:
        return [("Modo", "Sólo verificación ante el SAT"), periodo]
    tc = FUENTES_TC[config.fuente_tc]
    if config.fuente_tc == "manual" and config.tc_manual:
        tc += f" ({config.tc_manual})"
    return [
        ("Modo", "Verificación y cálculo de contraprestación"),
        periodo,
        ("Facturas fuera del periodo", "No suman" if config.solo_periodo else "También suman"),
        ("Tasa de contraprestación", f"{float(config.tasa * 100):g} %"),
        ("Base de cálculo", BASES[config.base]),
        ("Tipo de cambio USD", tc),
        ("RFC emisor auditado", config.rfc_emisor or "Todos"),
        ("Notas de crédito", "Se restan de la base" if config.restar_notas_credito else "Se ignoran"),
        ("Verificación SAT", "Sí" if config.verificar_sat else "No"),
        ("Facturas sin verificar", "Se incluyen" if config.incluir_no_verificadas else "Se excluyen"),
    ]
