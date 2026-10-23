# 🧾 Auditor de CFDIs y contraprestación

Verifica facturas (CFDI 3.2, 3.3 y 4.0) ante el SAT, revisa que estén dentro del
periodo auditado y calcula la contraprestación variable (5 % por defecto) en pesos.

Se usa como **aplicación web** (Streamlit), como **comando** en la terminal o
desde **Google Colab**. Todas usan el mismo código, que viene con pruebas automáticas.

## Qué hace

1. Lee los XML sueltos o los **ZIP de la descarga masiva del SAT** (incluso ZIP dentro de ZIP).
2. Detecta UUID duplicados, XML dañados y archivos que no son CFDI.
3. Consulta el estado de cada factura en el servicio del SAT, **en paralelo y con
   reintentos**. Reporta si está vigente, cancelada, no encontrada, si la cancelación
   está en proceso y si el emisor aparece en la **lista de EFOS (art. 69-B CFF)**.
4. Convierte a MXN las facturas en dólares con la fuente de tipo de cambio que elijas.
5. Calcula la base y la contraprestación, y explica por qué se excluyó cada factura.
6. Genera un Excel de una sola hoja: encabezado con los parámetros, indicadores y la
   tabla de facturas con filtros y totales que respetan el filtro.

Antes de empezar eliges el modo: **sólo verificar facturas** (estado SAT, EFOS,
cancelaciones) o **verificar y calcular contraprestación**. En ambos modos cada factura
trae su subtotal y total en pesos y el reporte los suma (las notas de crédito restan).

**Conceptos.** Cada concepto se lee con su clave SAT, cantidad, unidad, valor unitario e
importe. En el Excel aparecen como lista dentro de la celda "Conceptos"; en la página hay
una pestaña con todos los conceptos y un resumen por concepto, y al seleccionar una
factura se muestra su desglose. También se descargan en un CSV aparte.

## Reglas de cálculo

Una factura **entra en la base** sólo si cumple todo esto:

| Regla | Detalle |
|---|---|
| Tipo de comprobante | `I` (ingreso) suma. `E` (nota de crédito) **resta** (se puede desactivar). `P`, `T` y `N` (pago, traslado, nómina) no cuentan. |
| Periodo | Por defecto **también suman** las facturas emitidas fuera del periodo (quedan marcadas). Con el interruptor apagado, sólo cuentan las del rango. |
| Estado SAT | Debe estar **Vigente**. Las canceladas y las no encontradas se excluyen. Si el SAT no responde, la factura se excluye y se avisa (se puede cambiar). |
| No duplicada | Si un UUID aparece dos veces, sólo cuenta la primera. |
| Emisor (opcional) | Si indicas un RFC emisor, sólo cuentan las facturas que él emitió. |

- **Base** = `Subtotal − Descuento` (o sólo `Subtotal`, si así lo eliges) × tipo de cambio.
- **Contraprestación** = Base MXN × tasa, redondeada al centavo en cada factura.
- **Tipo de cambio USD**, a elegir:
  - el que declara cada CFDI (atributo `TipoCambio`), que es la opción por defecto;
  - el **FIX de Banxico publicado en el DOF** el día de emisión. En fin de semana o
    día feriado se usa el último publicado. Requiere el token gratuito de Banxico;
  - un tipo de cambio único que escribes tú.

  Otras monedas, como EUR, usan el `TipoCambio` del CFDI.

## Uso 1: aplicación web

### En tu computadora

Necesitas tener instalado [Python 3.10 o posterior](https://www.python.org/downloads/).

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en el navegador (`http://localhost:8501`). Sigue los tres pasos: elige el modo,
el periodo (y las reglas, si calculas) y sube los XML o ZIP.

### Publicada en internet (gratis), para usarla sin instalar nada

1. Entra a <https://share.streamlit.io> con tu cuenta de GitHub.
2. Elige **Create app**, el repositorio `auditor-facturas-sat`, la rama `main` y el archivo `app.py`.
3. Si vas a usar el tipo de cambio de Banxico, en **Advanced settings → Secrets** escribe:
   ```toml
   BANXICO_TOKEN = "tu-token"
   ```
4. Si el repositorio es privado, en la configuración de la app puedes limitar quién la ve.

Los archivos se procesan en memoria y no se guardan en ningún lado.

## Uso 2: línea de comandos

```bash
pip install -r requirements.txt
python -m auditor_cfdi /ruta/a/facturas --desde 2026-01-01 --hasta 2026-01-31
```

Opciones útiles:

```
--solo-verificar         Sólo estado ante el SAT, sin calcular contraprestación
--solo-periodo           Las facturas fuera del periodo no suman
--tasa 5                 Porcentaje de contraprestación
--base neto|subtotal     Restar o no los descuentos
--tc cfdi|banxico|manual --tc-manual 18.25 --token-banxico XXXX (o variable BANXICO_TOKEN)
--rfc-emisor AAA010101AAA
--sin-sat                No consultar al SAT (más rápido, sin validar estado)
--ignorar-notas-credito  --incluir-no-verificadas
-o reporte.xlsx
```

## Uso 3: Google Colab (con tus carpetas de Drive)

```python
from google.colab import drive, files
drive.mount('/content/drive')
!git clone https://github.com/sudestada93/auditor-facturas-sat.git
%cd auditor-facturas-sat
!pip install -q -r requirements.txt
!python -m auditor_cfdi "/content/drive/MyDrive/Facturas_SAT" --desde 2026-01-01 --hasta 2026-01-31 -o reporte.xlsx
files.download('reporte.xlsx')
```

Si el repositorio es privado, clónalo con un token personal de GitHub
(`https://<token>@github.com/...`).

## Estructura

```
app.py                     Interfaz web (Streamlit)
auditor_cfdi/
  lector.py                Lectura segura de XML/ZIP y normalización de CFDI 3.2/3.3/4.0
  sat.py                   Consulta SOAP al SAT (paralela, con reintentos) y liga al verificador
  tipo_cambio.py           FIX de Banxico (serie SF60653)
  auditoria.py             Reglas de inclusión y cálculo de la contraprestación
  reporte.py               Excel de una hoja (y la tabla que muestra la app)
  __main__.py              Línea de comandos
tests/                     Pruebas automáticas (pytest)
```

## Desarrollo

```bash
pip install -r requirements-dev.txt
python -m pytest
```

Las pruebas no se conectan al SAT ni a Banxico: simulan sus respuestas. GitHub
Actions las corre en cada push.
