"""Tipo de cambio USD→MXN.

Banxico publica el tipo de cambio FIX "para solventar obligaciones
denominadas en moneda extranjera". La serie SF60653 lo indexa por la fecha
de publicación en el DOF, que es la que aplica en una fecha dada. La API SIE
requiere un token gratuito: https://www.banxico.org.mx/SieAPIRest/service/v1/token
"""

from __future__ import annotations

import bisect
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation

import requests

SERIE_FIX_DOF = "SF60653"
URL_SIE = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/{serie}/datos/{ini}/{fin}"
URL_SIE_OPORTUNO = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/{serie}/datos/oportuno"


class ErrorTipoCambio(Exception):
    pass


def _datos_serie(json: dict) -> list[dict]:
    series = json.get("bmx", {}).get("series", [])
    return series[0].get("datos", []) if series else []


def _a_decimal(valor: str) -> Decimal | None:
    try:
        return Decimal(valor.replace(",", ""))
    except (InvalidOperation, AttributeError):
        return None  # Banxico usa "N/E" en días sin dato


class TablaFIX:
    """Tipos de cambio por fecha; para fines de semana y feriados usa el
    último publicado antes de esa fecha."""

    def __init__(self, valores: dict[date, Decimal]):
        self._fechas = sorted(valores)
        self._valores = valores

    def __len__(self) -> int:
        return len(self._fechas)

    def para(self, fecha: date) -> Decimal | None:
        i = bisect.bisect_right(self._fechas, fecha)
        return self._valores[self._fechas[i - 1]] if i else None


def descargar_fix(token: str, desde: date, hasta: date, timeout: float = 20) -> TablaFIX:
    # Se piden días extra hacia atrás por si el periodo empieza en fin de semana.
    url = URL_SIE.format(serie=SERIE_FIX_DOF, ini=(desde - timedelta(days=10)).isoformat(),
                         fin=hasta.isoformat())
    try:
        r = requests.get(url, headers={"Bmx-Token": token.strip()}, timeout=timeout)
    except requests.RequestException as e:
        raise ErrorTipoCambio(f"No se pudo conectar con Banxico: {e}") from e
    if r.status_code in (401, 403):
        raise ErrorTipoCambio("Banxico rechazó el token. Revisa que sea correcto.")
    if r.status_code != 200:
        raise ErrorTipoCambio(f"Banxico respondió HTTP {r.status_code}")
    valores = {}
    for d in _datos_serie(r.json()):
        tc = _a_decimal(d.get("dato", ""))
        if tc is not None:
            valores[datetime.strptime(d["fecha"], "%d/%m/%Y").date()] = tc
    if not valores:
        raise ErrorTipoCambio("Banxico no devolvió datos para el periodo.")
    return TablaFIX(valores)


def tc_referencia(token: str | None = None, timeout: float = 5) -> float | None:
    """Último tipo de cambio conocido, sólo como sugerencia en la interfaz."""
    if token:
        try:
            r = requests.get(URL_SIE_OPORTUNO.format(serie=SERIE_FIX_DOF),
                             headers={"Bmx-Token": token.strip()}, timeout=timeout)
            datos = _datos_serie(r.json())
            if datos and (tc := _a_decimal(datos[-1].get("dato", ""))) is not None:
                return float(tc)
        except Exception:
            pass
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=timeout)
        return float(r.json()["rates"]["MXN"])
    except Exception:
        return None
