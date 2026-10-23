from datetime import date
from decimal import Decimal
from unittest import mock

import pytest

from auditor_cfdi.tipo_cambio import ErrorTipoCambio, TablaFIX, descargar_fix

RESPUESTA = {"bmx": {"series": [{"idSerie": "SF60653", "datos": [
    {"fecha": "02/01/2026", "dato": "17.9000"},
    {"fecha": "05/01/2026", "dato": "N/E"},
    {"fecha": "06/01/2026", "dato": "18.1000"},
]}]}}


def test_tabla_usa_el_ultimo_publicado():
    t = TablaFIX({date(2026, 1, 2): Decimal("17.9"), date(2026, 1, 6): Decimal("18.1")})
    assert t.para(date(2026, 1, 1)) is None
    assert t.para(date(2026, 1, 4)) == Decimal("17.9")
    assert t.para(date(2026, 1, 6)) == Decimal("18.1")


def test_descargar_fix():
    resp = mock.Mock(status_code=200, json=lambda: RESPUESTA)
    with mock.patch("auditor_cfdi.tipo_cambio.requests.get", return_value=resp) as get:
        t = descargar_fix(" token ", date(2026, 1, 5), date(2026, 1, 31))
    url = get.call_args.args[0]
    assert "/SF60653/datos/2025-12-26/2026-01-31" in url
    assert get.call_args.kwargs["headers"] == {"Bmx-Token": "token"}
    assert len(t) == 2 and t.para(date(2026, 1, 5)) == Decimal("17.9000")


def test_token_invalido():
    with mock.patch("auditor_cfdi.tipo_cambio.requests.get",
                    return_value=mock.Mock(status_code=401)):
        with pytest.raises(ErrorTipoCambio, match="token"):
            descargar_fix("x", date(2026, 1, 1), date(2026, 1, 2))
