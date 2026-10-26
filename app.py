"""Interfaz web.  Ejecutar con:  streamlit run app.py"""

from __future__ import annotations

from dataclasses import asdict
from datetime import date, timedelta
from decimal import Decimal
from html import escape

import requests
import streamlit as st

from auditor_cfdi import Configuracion, auditar_archivos, generar_excel, sat
from auditor_cfdi.auditoria import BASES, FUENTES_TC
from auditor_cfdi.reporte import tabla_reporte
from auditor_cfdi.tipo_cambio import tc_referencia

st.set_page_config(page_title="Auditoría de CFDIs", page_icon=":material/fact_check:",
                   layout="wide", initial_sidebar_state="collapsed")

st.html("""
<style>
  :root { --carbon:#181925; --grafito:#666666; --ceniza:#999999; --niebla:#E8E8E8;
          --lino:#FAFAFA; --lavanda:#918DF6; --menta:#1E9E46; --menta-fondo:#DEF6E4; }
  .block-container { max-width: 1120px; padding-top: 2.5rem; padding-bottom: 4rem; }
  [data-testid="stSidebarCollapsedControl"] { display: none; }

  .hero { text-align: center; padding: 1.5rem 0 2.5rem; }
  .chip { display: inline-flex; gap: .5rem; align-items: center; border: 1px solid var(--niebla);
          border-radius: 9999px; padding: .3rem .85rem; font-size: 13px; font-weight: 500;
          color: var(--carbon); background: #fff; }
  .chip b { color: #2C78FC; font-weight: 600; }
  .hero h1 { font-size: 52px; line-height: 1.1; font-weight: 600; letter-spacing: -2px;
             color: var(--carbon); margin: 1.1rem 0 .6rem; padding: 0; }
  .hero p { font-size: 17px; color: var(--grafito); max-width: 640px; margin: 0 auto;
            letter-spacing: -.3px; }

  .paso { display: flex; gap: .9rem; align-items: flex-start; margin-bottom: .4rem; }
  .paso .num { flex: none; width: 32px; height: 32px; border-radius: 9999px; background: var(--lavanda);
               color: #fff; font-weight: 600; font-size: 14px; display: grid; place-items: center; }
  .paso h3 { font-size: 18px; font-weight: 600; letter-spacing: -.3px; margin: 0; padding: 0;
             color: var(--carbon); }
  .paso p { font-size: 14px; color: var(--grafito); margin: .15rem 0 0; }

  [class*="st-key-tarjeta"] { border-radius: 16px !important; padding: 1.5rem 1.75rem !important;
      box-shadow: rgba(0,0,0,.04) 0 1px 3px, rgba(0,0,0,.03) 0 8px 16px; }

  .metricas { display: grid; grid-template-columns: repeat(var(--cols), minmax(0, 1fr));
              gap: 12px; margin: .5rem 0 1rem; }
  @media (max-width: 760px) { .metricas { grid-template-columns: repeat(2, 1fr); } }
  .metrica { border: 1px solid var(--niebla); border-radius: 16px; padding: 14px 16px; background: #fff; }
  .metrica .etq { font-size: 12px; color: var(--ceniza); letter-spacing: -.2px; }
  .metrica .val { font-size: 24px; font-weight: 600; letter-spacing: -.8px; color: var(--carbon);
                  margin-top: 2px; white-space: nowrap; }
  .metrica.destacada { background: var(--menta-fondo); border-color: var(--menta-fondo);
                       grid-column: span 2; }
  .metrica.destacada .etq { color: var(--menta); font-weight: 600; }
  .seccion { font-size: 28px; font-weight: 600; letter-spacing: -.8px; color: var(--carbon);
             margin: 2.5rem 0 .25rem; }
  .sub { color: var(--grafito); font-size: 14px; margin-bottom: .5rem; }
</style>
""")


def _secreto(nombre: str) -> str:
    try:
        return st.secrets.get(nombre, "")
    except Exception:  # no hay secrets.toml
        return ""


@st.cache_data(ttl=3600, show_spinner=False)
def _tc_sugerido(token: str) -> float:
    return tc_referencia(token or None) or 18.0


def _paso(numero: int, titulo: str, texto: str):
    st.html(f'<div class="paso"><div class="num">{numero}</div>'
            f'<div><h3>{titulo}</h3><p>{texto}</p></div></div>')


def _metricas(tarjetas: list[tuple[str, str, bool]]):
    celdas = "".join(
        f'<div class="metrica{" destacada" if dest else ""}"><div class="etq">{escape(etq)}</div>'
        f'<div class="val">{escape(val)}</div></div>' for etq, val, dest in tarjetas)
    columnas = len(tarjetas) + sum(dest for *_, dest in tarjetas)  # la destacada ocupa 2
    st.html(f'<div class="metricas" style="--cols:{columnas}">{celdas}</div>')


def _pesos(valor: float) -> str:
    return f"${valor:,.2f}"


MODOS = {"verificar": "Sólo verificar facturas",
         "calcular": "Verificar y calcular contraprestación"}

# ── Encabezado ────────────────────────────────────────────────────────────────
st.html("""
<div class="hero">
  <span class="chip"><b>SAT</b> CFDI 3.2 · 3.3 · 4.0</span>
  <h1>Auditoría de facturas</h1>
  <p>Confirma ante el SAT que tus CFDI siguen vigentes y, si lo necesitas,
     calcula la contraprestación a cobrar. Tus archivos se procesan en memoria y no se guardan.</p>
</div>
""")

# ── Paso 1: qué hacer ─────────────────────────────────────────────────────────
with st.container(border=True, key="tarjeta-modo"):
    _paso(1, "¿Qué quieres hacer?",
          "Elige si sólo revisas el estado de las facturas o también calculas la contraprestación.")
    modo = st.segmented_control("Modo", list(MODOS), format_func=MODOS.get, default="calcular",
                                required=True, label_visibility="collapsed")
    calcular = modo == "calcular"

# ── Paso 2: periodo y reglas ──────────────────────────────────────────────────
fin_mes_pasado = date.today().replace(day=1) - timedelta(days=1)
token = _secreto("BANXICO_TOKEN")
tc_manual = None
tasa, base, fuente_tc, rfc_emisor = 5.0, "neto", "cfdi", ""
solo_periodo, restar_nc, incluir_nv = False, True, False

with st.container(border=True, key="tarjeta-periodo"):
    if calcular:
        _paso(2, "Periodo y reglas de cálculo",
              "El periodo marca qué facturas son del mes; las reglas definen cómo se calcula.")
    else:
        _paso(2, "Periodo", "Se marcarán las facturas emitidas fuera de este rango.")
    c1, c2, c3 = st.columns(3)
    fecha_inicio = c1.date_input("Desde", fin_mes_pasado.replace(day=1), format="DD/MM/YYYY")
    fecha_fin = c2.date_input("Hasta", fin_mes_pasado, format="DD/MM/YYYY")
    if calcular:
        tasa = c3.number_input("Tasa de contraprestación (%)", 0.0, 100.0, 5.0, 0.5, format="%.2f")
        c1, c2, c3 = st.columns(3)
        base = c1.selectbox("Base de cálculo", list(BASES), format_func=BASES.get)
        fuente_tc = c2.selectbox("Tipo de cambio para USD", list(FUENTES_TC),
                                 format_func=FUENTES_TC.get)
        if fuente_tc == "banxico":
            token = c3.text_input("Token de Banxico", token, type="password",
                                  help="Es gratuito: https://www.banxico.org.mx/SieAPIRest/service/v1/token")
        elif fuente_tc == "manual":
            tc_manual = c3.number_input("Tipo de cambio", min_value=0.0001,
                                        value=_tc_sugerido(token), format="%.4f")
        solo_periodo = not st.toggle(
            "Sumar también las facturas emitidas fuera del periodo", True,
            help="Si lo apagas, sólo cuentan las facturas con fecha de emisión dentro del periodo.")
        with st.expander("Más opciones"):
            rfc_emisor = st.text_input(
                "RFC emisor a auditar (opcional)",
                help="Si lo indicas, sólo cuentan las facturas emitidas por este RFC.")
            restar_nc = st.checkbox("Restar notas de crédito (tipo E)", True)
            incluir_nv = st.checkbox(
                "Contar facturas que el SAT no pudo verificar", False,
                help="Si el servicio del SAT falla para alguna factura, por defecto no se cuenta.")
    if fecha_inicio > fecha_fin:
        st.error("La fecha inicial no puede ser posterior a la final.")

config = Configuracion(
    fecha_inicio=fecha_inicio, fecha_fin=fecha_fin, calcular=calcular, solo_periodo=solo_periodo,
    tasa=Decimal(str(tasa)) / 100, base=base, fuente_tc=fuente_tc,
    tc_manual=Decimal(str(tc_manual)) if tc_manual else None, banxico_token=token or None,
    rfc_emisor=rfc_emisor.strip().upper() or None, restar_notas_credito=restar_nc,
    verificar_sat=True, incluir_no_verificadas=incluir_nv)

# ── Paso 3: archivos ──────────────────────────────────────────────────────────
with st.container(border=True, key="tarjeta-archivos"):
    _paso(3, "Sube tus facturas",
          "Arrastra los XML o los ZIP de la descarga masiva del SAT. Puedes mezclar ambos.")
    subidos = st.file_uploader("Archivos", type=["xml", "zip"], accept_multiple_files=True,
                               label_visibility="collapsed")
    iniciar = st.button("Iniciar auditoría", type="primary", icon=":material/play_arrow:",
                        disabled=not subidos or fecha_inicio > fecha_fin)

if iniciar:
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
    st.stop()

# ── Resultados ────────────────────────────────────────────────────────────────
rc = resultado.config
st.html(f'<div class="seccion">Resultados</div><div class="sub">'
        f'{MODOS["calcular" if rc.calcular else "verificar"]} · periodo '
        f'{rc.fecha_inicio:%d/%m/%Y} – {rc.fecha_fin:%d/%m/%Y}</div>')

if asdict(rc) != asdict(config):
    st.info("Cambiaste opciones después de la última auditoría. Pulsa **Iniciar auditoría** "
            "para actualizar los resultados; no se vuelve a consultar al SAT lo que ya respondió.",
            icon=":material/refresh:")

det = resultado.detalle
if rc.calcular:
    _metricas([
        ("Facturas", f"{len(det):,}", False),
        ("Cuentan para el cálculo", f"{len(resultado.incluidas):,}", False),
        ("Base total MXN", _pesos(resultado.base_mxn), False),
        (f"Contraprestación {float(rc.tasa * 100):g} %", _pesos(resultado.contraprestacion), True),
        ("Con incidencias", f"{len(resultado.incidencias):,}", False),
    ])
else:
    _metricas([
        ("Facturas", f"{len(det):,}", False),
        ("Vigentes", f"{resultado.contar_estado('Vigente'):,}", False),
        ("Canceladas", f"{resultado.contar_estado('Cancelado'):,}", False),
        ("No confirmadas",
         f"{resultado.contar_estado('No Encontrado') + resultado.sin_verificar:,}", False),
        ("Emisor en EFOS", f"{resultado.en_efos:,}", False),
        ("Total vigente MXN", _pesos(resultado.total_vigente_mxn), True),
    ])

for aviso in resultado.advertencias:
    st.warning(aviso, icon=":material/warning:")

nombre = f"Auditoria_{rc.fecha_inicio}_al_{rc.fecha_fin}"
tabla = tabla_reporte(resultado)
conceptos = resultado.conceptos
d1, d2, d3, _ = st.columns([1.2, 1, 1.2, 1.6])
d1.download_button("Descargar Excel", st.session_state["excel"], f"{nombre}.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", icon=":material/download:", width="stretch")
d2.download_button("CSV facturas", tabla.to_csv(index=False).encode("utf-8-sig"),
                   f"{nombre}.csv", "text/csv", width="stretch")
d3.download_button("CSV conceptos", conceptos.to_csv(index=False).encode("utf-8-sig"),
                   f"{nombre}_conceptos.csv", "text/csv", width="stretch")

COLORES_ESTADO = {"Vigente": "color:#1E9E46;background-color:#DEF6E4;font-weight:600",
                  "Cancelado": "color:#C62828;background-color:#FDE7E7;font-weight:600"}
MONTOS = ["Subtotal", "Total", "Subtotal MXN", "Total MXN", "Base MXN", "Contraprestación",
          "Valor unitario", "Importe", "Descuento", "Importe neto MXN", "Importe MXN"]


def _texto(valor, patron: str) -> str:
    return "" if valor is None or valor != valor else patron.format(valor)


def _ordenar(df, primero: list[str], quitar: tuple[str, ...] = ()):
    """Pone primero las columnas clave para el análisis (sin perder las demás)."""
    primero = [c for c in primero if c in df.columns]
    resto = [c for c in df.columns if c not in primero and c not in quitar]
    return df[primero + resto]


ORDEN_FACTURAS = ["Estado SAT", "Verificar en SAT", "¿Cuenta?", "EFOS", "Fecha", "Serie y folio", "Emisor",
                  "Conceptos", "Moneda", "Subtotal", "Total", "Tipo de cambio", "Subtotal MXN",
                  "Total MXN", "Base MXN", "Contraprestación", "Notas"]
ORDEN_CONCEPTOS = ["Estado SAT", "¿Cuenta?", "Fecha", "Serie y folio", "Emisor", "Descripción",
                   "Clave SAT", "Cantidad", "Unidad", "Valor unitario", "Importe", "Descuento",
                   "Moneda", "Importe neto MXN"]
SIN_CUENTA = () if rc.calcular else ("¿Cuenta?",)


def _resumen_conceptos(texto: str) -> str:
    """Primer concepto (sin el detalle) y cuántos más trae la factura."""
    lineas = [l for l in texto.split("\n") if l]
    if not lineas:
        return ""
    primero = lineas[0].removeprefix("• ").split(" — ")[0]
    return primero + (f"  (+{len(lineas) - 1} más)" if len(lineas) > 1 else "")


def _mostrar(df, clave: str | None = None):
    """Tabla con formato. Si se da `clave`, permite seleccionar una fila."""
    df = df.copy()
    numericas = [c for c in MONTOS + ["Tipo de cambio", "Cantidad"] if c in df.columns]
    for col in numericas:
        patron = {"Tipo de cambio": "{:.4f}", "Cantidad": "{:,.6g}"}.get(col, "{:,.2f}")
        df[col] = df[col].map(lambda v, p=patron: _texto(v, p))
    if "Fecha" in df:
        df["Fecha"] = df["Fecha"].map(lambda v: _texto(v, "{:%d/%m/%Y}"))
    if "Conceptos" in df:
        df["Conceptos"] = df["Conceptos"].map(_resumen_conceptos)
    estilo = df.style
    if "Estado SAT" in df:
        estilo = estilo.map(
            lambda v: COLORES_ESTADO.get(v, "color:#9A6400;background-color:#FFF2D6;font-weight:600"),
            subset=["Estado SAT"])
    if "¿Cuenta?" in df.columns:
        estilo = estilo.map(lambda v: "color:#999999" if v == "No" else "", subset=["¿Cuenta?"])
    seleccion = {"on_select": "rerun", "selection_mode": "single-row", "key": clave} if clave else {}
    return st.dataframe(
        estilo, hide_index=True, height=min(540, 35 * (len(df) + 1) + 3),
        column_config={
            **{c: st.column_config.TextColumn(alignment="right") for c in numericas},
            "Conceptos": st.column_config.TextColumn(width="large"),
            "Descripción": st.column_config.TextColumn(width="large"),
            "Notas": st.column_config.TextColumn(width="large"),
            "Verificar en SAT": st.column_config.LinkColumn(display_text="Abrir verificador"),
        }, **seleccion)


def _datos_verificador(factura):
    """Liga y datos listos para copiar en el verificador del SAT."""
    if not factura["Verificar en SAT"]:
        return
    st.markdown("**Verificar en el SAT**")
    c1, c2, c3 = st.columns([2.2, 1, 1])
    c1.caption("Folio fiscal")
    c1.code(factura["UUID"], language=None)
    c2.caption("RFC emisor")
    c2.code(factura["RFC emisor"], language=None)
    c3.caption("RFC receptor")
    c3.code(factura["RFC receptor"], language=None)
    st.caption("Liga con los datos cargados: cópiala con el botón de la derecha y pégala en "
               "la barra de direcciones del navegador; sólo falta escribir el código de la imagen.")
    st.code(factura["Verificar en SAT"], language=None, wrap_lines=True)
    st.link_button("Abrir verificador", factura["Verificar en SAT"],
                   icon=":material/open_in_new:")


def _desglose(vista, evento):
    """Datos para el verificador y conceptos de la factura seleccionada."""
    filas = evento.selection.rows if evento else []
    if not filas:
        st.caption("Haz clic en el recuadro a la izquierda de una factura para ver sus conceptos "
                   "y los datos para el verificador del SAT.")
        return
    factura = vista.iloc[filas[0]]
    propios = conceptos[(conceptos["UUID"] == factura["UUID"]) &
                        (conceptos["Archivo"] == factura["Archivo"])]
    with st.container(border=True, key="tarjeta-desglose"):
        st.markdown(f"#### Factura {factura['Serie y folio'] or factura['UUID']} · "
                    f"{factura['Emisor']} · {factura['Fecha']:%d/%m/%Y}"
                    if factura["Fecha"] == factura["Fecha"] else
                    f"#### Factura {factura['UUID']}")
        _datos_verificador(factura)
        st.markdown("**Conceptos**")
        if propios.empty:
            st.write("Esta factura no tiene conceptos legibles.")
            return
        _mostrar(propios[["Clave SAT", "No. identificación", "Descripción", "Cantidad", "Unidad",
                          "Valor unitario", "Importe", "Descuento", "Moneda", "Importe neto MXN"]])
        st.caption(f"{len(propios)} concepto(s) · importe neto "
                   f"{_pesos(propios['Importe neto MXN'].sum())} MXN")


pestanas = ["Facturas", f"Conceptos ({len(conceptos)})",
            f"Incidencias ({len(resultado.incidencias)})"]
if rc.calcular:
    pestanas.append("Por mes")
t_facturas, t_conceptos, t_incid, *t_mes = st.tabs(pestanas)

with t_facturas:
    f1, f2 = st.columns([2, 1.3])
    buscar = f1.text_input("Buscar", placeholder="UUID, RFC, nombre, folio o concepto",
                           label_visibility="collapsed")
    estados = sorted(e for e in tabla["Estado SAT"].unique())
    filtro = f2.pills("Estado", estados, selection_mode="multi", label_visibility="collapsed")
    vista = tabla
    if buscar:
        texto = vista[["UUID", "RFC emisor", "Emisor", "RFC receptor", "Receptor",
                       "Serie y folio", "Conceptos", "Archivo"]].astype(str).agg(" ".join, axis=1)
        vista = vista[texto.str.contains(buscar, case=False, regex=False)]
    if filtro:
        vista = vista[vista["Estado SAT"].isin(filtro)]
    resumen = f"{len(vista)} de {len(tabla)} facturas"
    if rc.calcular:
        resumen += f" · contraprestación visible {_pesos(vista['Contraprestación'].sum())}"
    else:
        resumen += f" · total visible {_pesos(vista['Total MXN'].sum())} MXN"
    st.caption(resumen)
    vista = _ordenar(vista, ORDEN_FACTURAS)
    _desglose(vista, _mostrar(vista, "tabla-facturas"))

with t_conceptos:
    if conceptos.empty:
        st.write("No se encontraron conceptos en las facturas.")
    else:
        cuentan = (conceptos["¿Cuenta?"] == "Sí") if rc.calcular else \
            (conceptos["Estado SAT"] == "Vigente")
        agrupado = (conceptos[cuentan]
                    .groupby(["Descripción", "Clave SAT"], as_index=False)
                    .agg(Facturas=("UUID", "nunique"), Cantidad=("Cantidad", "sum"),
                         **{"Importe MXN": ("Importe neto MXN", "sum")})
                    .sort_values("Importe MXN", ascending=False))
        st.markdown("**Resumen por concepto**")
        st.caption("Sólo facturas que " + ("cuentan para el cálculo." if rc.calcular
                                            else "están vigentes ante el SAT.")
                   + " Las notas de crédito restan.")
        _mostrar(agrupado)
        st.markdown("**Todos los conceptos**")
        buscar_c = st.text_input("Buscar concepto", placeholder="Descripción, clave SAT, emisor…",
                                 label_visibility="collapsed")
        vista_c = conceptos
        if buscar_c:
            texto = vista_c[["Descripción", "Clave SAT", "No. identificación", "Emisor",
                             "RFC emisor", "Serie y folio", "UUID"]].astype(str).agg(" ".join, axis=1)
            vista_c = vista_c[texto.str.contains(buscar_c, case=False, regex=False)]
        st.caption(f"{len(vista_c)} de {len(conceptos)} conceptos · importe neto visible "
                   f"{_pesos(vista_c['Importe neto MXN'].sum())} MXN")
        _mostrar(_ordenar(vista_c, ORDEN_CONCEPTOS, SIN_CUENTA))

with t_incid:
    incid = tabla.loc[resultado.incidencias.index]
    if incid.empty:
        st.success("Sin incidencias: todo está en orden.", icon=":material/check_circle:")
    else:
        st.caption("Facturas que no cuentan, no están vigentes o tienen alguna observación.")
        _mostrar(_ordenar(incid, ORDEN_FACTURAS))

if t_mes:
    with t_mes[0]:
        if resultado.por_periodo.empty:
            st.write("Ninguna factura cuenta para el cálculo.")
        else:
            st.bar_chart(resultado.por_periodo, x="Periodo", y="Contraprestación MXN",
                         color="#918DF6", height=260)
            st.dataframe(resultado.por_periodo, hide_index=True, column_config={
                "Base MXN": st.column_config.NumberColumn(format="dollar"),
                "Contraprestación MXN": st.column_config.NumberColumn(format="dollar")})
