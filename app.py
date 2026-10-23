"""Interfaz web.  Ejecutar con:  streamlit run app.py"""

from __future__ import annotations

from dataclasses import asdict
from datetime import date, timedelta
from decimal import Decimal

import pandas as pd
import requests
import streamlit as st

from auditor_cfdi import Configuracion, auditar_archivos, generar_excel, sat
from auditor_cfdi.auditoria import BASES, FUENTES_TC, describir_config
from auditor_cfdi.tipo_cambio import tc_referencia

st.set_page_config(page_title="Auditor de CFDIs", page_icon="🧾", layout="wide")


def _secreto(nombre: str) -> str:
    try:
        return st.secrets.get(nombre, "")
    except Exception:  # no hay secrets.toml
        return ""


@st.cache_data(ttl=3600, show_spinner=False)
def _tc_sugerido(token: str) -> float:
    return tc_referencia(token or None) or 18.0


def _dinero(valor: float) -> str:
    return f"$ {valor:,.2f}"


# ── Parámetros ────────────────────────────────────────────────────────────────
hoy = date.today()
fin_mes_pasado = hoy.replace(day=1) - timedelta(days=1)

with st.sidebar:
    st.header("Parámetros")
    fecha_inicio = st.date_input("Desde", fin_mes_pasado.replace(day=1), format="DD/MM/YYYY")
    fecha_fin = st.date_input("Hasta", fin_mes_pasado, format="DD/MM/YYYY")
    tasa = st.number_input("Tasa de contraprestación (%)", 0.0, 100.0, 5.0, 0.5, format="%.2f")
    base = st.selectbox("Base de cálculo", list(BASES), format_func=BASES.get)

    st.subheader("Tipo de cambio USD")
    fuente_tc = st.radio("Fuente", list(FUENTES_TC), format_func=FUENTES_TC.get,
                         label_visibility="collapsed")
    token = _secreto("BANXICO_TOKEN")
    tc_manual = None
    if fuente_tc == "banxico":
        token = st.text_input("Token de Banxico", token, type="password",
                              help="Gratis en https://www.banxico.org.mx/SieAPIRest/service/v1/token")
    elif fuente_tc == "manual":
        tc_manual = st.number_input("Tipo de cambio", min_value=0.0001,
                                    value=_tc_sugerido(token), format="%.4f")

    rfc_emisor = st.text_input("RFC emisor a auditar (opcional)",
                               help="Si lo indicas, sólo cuentan las facturas emitidas por este RFC.")

    with st.expander("Opciones avanzadas"):
        verificar_sat = st.checkbox("Verificar estado ante el SAT", True)
        restar_nc = st.checkbox("Restar notas de crédito (tipo E)", True)
        incluir_nv = st.checkbox(
            "Contar facturas que el SAT no pudo verificar", False,
            help="Si el servicio del SAT falla para alguna factura, por defecto se excluye.")

config = Configuracion(
    fecha_inicio=fecha_inicio, fecha_fin=fecha_fin, tasa=Decimal(str(tasa)) / 100, base=base,
    fuente_tc=fuente_tc, tc_manual=Decimal(str(tc_manual)) if tc_manual else None,
    banxico_token=token or None, rfc_emisor=rfc_emisor.strip().upper() or None,
    restar_notas_credito=restar_nc, verificar_sat=verificar_sat,
    incluir_no_verificadas=incluir_nv)

# ── Carga y ejecución ─────────────────────────────────────────────────────────
st.title("🧾 Auditoría de CFDIs y contraprestación")
st.caption("Verifica las facturas ante el SAT, revisa el periodo y calcula la contraprestación "
           "variable. Los archivos se procesan en memoria y no se guardan.")

subidos = st.file_uploader("Sube los XML o los ZIP de la descarga masiva del SAT",
                           type=["xml", "zip"], accept_multiple_files=True)

if fecha_inicio > fecha_fin:
    st.error("La fecha inicial no puede ser posterior a la final.")

if st.button("▶ Iniciar auditoría", type="primary",
             disabled=not subidos or fecha_inicio > fecha_fin):
    # Las respuestas del SAT se recuerdan durante la sesión para que repetir la
    # auditoría con otros parámetros no vuelva a consultar las mismas facturas.
    memoria = st.session_state.setdefault("estados_sat", {})
    sesion = requests.Session()

    def consultor(c):
        clave = (c.uuid, c.rfc_emisor, c.rfc_receptor, c.total_texto)
        if clave not in memoria:
            estado = sat.consultar(c, sesion)
            if not estado.verificado:
                return estado
            memoria[clave] = estado
        return memoria[clave]

    barra = st.progress(0.0, "Preparando…")

    def avanzar(etapa, actual, total):
        barra.progress(actual / max(total, 1), f"{etapa}: {actual} de {total}")

    resultado = auditar_archivos([(f.name, f.getvalue()) for f in subidos], config,
                                 avanzar, consultor_sat=consultor)
    barra.empty()
    st.session_state["resultado"] = resultado
    st.session_state["excel"] = generar_excel(resultado)

resultado = st.session_state.get("resultado")
if resultado is None:
    st.info("Sube tus archivos, revisa los parámetros de la izquierda y pulsa **Iniciar auditoría**.")
    st.stop()

# ── Resultados ────────────────────────────────────────────────────────────────
if asdict(resultado.config) != asdict(config):
    st.warning("Cambiaste parámetros después de la última auditoría. Pulsa **Iniciar auditoría** "
               "para recalcular (las respuestas del SAT ya obtenidas se reutilizan).")

for aviso in resultado.advertencias:
    st.warning(aviso)

det = resultado.detalle
m1, m2, m3, m4 = st.columns(4)
m1.metric("Archivos procesados", len(det))
m2.metric("Incluidas en la base", f"{len(resultado.incluidas)} de {len(det)}")
m3.metric("Base (MXN)", _dinero(resultado.base_mxn))
m4.metric(f"Contraprestación {float(resultado.config.tasa * 100):g} % (MXN)",
          _dinero(resultado.contraprestacion))

nombre = f"Auditoria_{resultado.config.fecha_inicio}_al_{resultado.config.fecha_fin}"
d1, d2, _ = st.columns([1, 1, 3])
d1.download_button("📥 Descargar Excel", st.session_state["excel"], f"{nombre}.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", use_container_width=True)
d2.download_button("Descargar CSV", det.to_csv(index=False).encode("utf-8-sig"),
                   f"{nombre}.csv", "text/csv", use_container_width=True)

dinero = st.column_config.NumberColumn(format="dollar")
config_columnas = {
    "Fecha emisión": st.column_config.DatetimeColumn(format="DD/MM/YYYY HH:mm"),
    "Tipo de cambio": st.column_config.NumberColumn(format="%.4f"),
    "Verificar en SAT": st.column_config.LinkColumn(display_text="Abrir"),
    **{c: dinero for c in ["Subtotal", "Descuento", "Total", "Base moneda original",
                           "Base MXN", "Contraprestación MXN"]},
}

t_resumen, t_detalle, t_incid = st.tabs(
    ["Resumen", "Detalle", f"Incidencias ({len(resultado.incidencias)})"])

with t_resumen:
    izq, der = st.columns([3, 2])
    with izq:
        st.subheader("Por mes de emisión")
        if resultado.por_periodo.empty:
            st.write("Ninguna factura quedó incluida en la base.")
        else:
            st.dataframe(resultado.por_periodo, hide_index=True, column_config=config_columnas)
            st.bar_chart(resultado.por_periodo, x="Periodo", y="Contraprestación MXN")
            st.subheader("Por moneda")
            st.dataframe(resultado.por_moneda, hide_index=True, column_config=config_columnas)
    with der:
        st.subheader("Estado ante el SAT")
        st.dataframe(resultado.por_estado, hide_index=True)
        if not resultado.exclusiones.empty:
            st.subheader("Motivos de exclusión")
            st.dataframe(resultado.exclusiones, hide_index=True)
        st.subheader("Parámetros usados")
        st.table(pd.DataFrame(describir_config(resultado.config), columns=["", "Valor"])
                 .set_index(""))

with t_detalle:
    f1, f2, f3 = st.columns([2, 1, 1])
    buscar = f1.text_input("Buscar (UUID, RFC, nombre, folio, descripción)")
    incluida = f2.selectbox("¿Incluida?", ["Todas", "Sí", "No"])
    estados = f3.multiselect("Estado SAT", sorted(e for e in det["Estado SAT"].unique() if e),
                           placeholder="Todos")
    vista = det
    if buscar:
        texto = vista[["UUID", "RFC emisor", "Nombre emisor", "RFC receptor", "Nombre receptor",
                       "Serie", "Folio", "Descripción", "Archivo"]].astype(str).agg(" ".join, axis=1)
        vista = vista[texto.str.contains(buscar, case=False, regex=False)]
    if incluida != "Todas":
        vista = vista[vista["¿Incluida?"] == incluida]
    if estados:
        vista = vista[vista["Estado SAT"].isin(estados)]
    st.caption(f"{len(vista)} de {len(det)} registros · contraprestación filtrada: "
               f"{_dinero(vista['Contraprestación MXN'].sum())}")
    st.dataframe(vista, hide_index=True, column_config=config_columnas, height=520)

with t_incid:
    if resultado.incidencias.empty:
        st.success("Sin incidencias: todas las facturas están incluidas y sin observaciones.")
    else:
        st.dataframe(resultado.incidencias[
            ["Archivo", "UUID", "Fecha emisión", "RFC emisor", "Nombre emisor", "Total", "Moneda",
             "Estado SAT", "¿Incluida?", "Motivo de exclusión", "Observaciones", "Verificar en SAT"]],
            hide_index=True, column_config=config_columnas, height=520)
