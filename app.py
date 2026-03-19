import streamlit as st
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
from zeep import Client
import warnings
import io

# Ignorar advertencias de zeep
warnings.filterwarnings("ignore")

# --- FUNCIONES DE EXTRACCIÓN Y CONSULTA ---
def extraer_datos_cfdi(archivo_xml):
    """Extrae datos de un archivo XML cargado en la web."""
    try:
        # Volver al inicio del archivo por si ya se leyó
        archivo_xml.seek(0)
        tree = ET.parse(archivo_xml)
        root = tree.getroot()
        
        fecha_str = root.get('Fecha') or root.get('fecha')
        total = root.get('Total') or root.get('total')
        subtotal = root.get('SubTotal') or root.get('subTotal')
        serie = root.get('Serie') or root.get('serie')
        folio = root.get('Folio') or root.get('folio')
        moneda = root.get('Moneda') or root.get('moneda')
        
        sello = root.get('Sello') or root.get('sello')
        sello_8 = sello[-8:] if sello else ""
        
        rfc_emisor, nombre_emisor = None, None
        rfc_receptor, nombre_receptor = None, None
        uuid = None
        descripciones = []
        
        for elemento in root.iter():
            etiqueta = elemento.tag.lower()
            if etiqueta.endswith('emisor') and rfc_emisor is None:
                rfc_emisor = elemento.get('Rfc') or elemento.get('rfc')
                nombre_emisor = elemento.get('Nombre') or elemento.get('nombre')
            elif etiqueta.endswith('receptor') and rfc_receptor is None:
                rfc_receptor = elemento.get('Rfc') or elemento.get('rfc')
                nombre_receptor = elemento.get('Nombre') or elemento.get('nombre')
            elif etiqueta.endswith('timbrefiscaldigital') and uuid is None:
                uuid = elemento.get('UUID') or elemento.get('uuid')
            elif etiqueta.endswith('concepto'):
                desc = elemento.get('Descripcion') or elemento.get('descripcion')
                if desc:
                    descripciones.append(desc.replace('\n', ' ').replace('\r', '').strip())
                
        link_sat = ""
        if all([rfc_emisor, rfc_receptor, total, uuid]):
            link_sat = f"https://verificacfdi.facturaelectronica.sat.gob.mx/default.aspx?id={uuid}&re={rfc_emisor}&rr={rfc_receptor}&tt={total}&fe={sello_8}"
        
        descripcion_total = " | ".join(descripciones) if descripciones else "Sin descripción"
        
        return {
            'Archivo': archivo_xml.name,
            'Serie': serie,
            'Folio': folio,
            'UUID': uuid,
            'RFC_Emisor': rfc_emisor,
            'Nombre_Emisor': nombre_emisor,
            'RFC_Receptor': rfc_receptor,
            'Nombre_Receptor': nombre_receptor,
            'Moneda': moneda,
            'Descripcion': descripcion_total,
            'SubTotal': subtotal,
            'Total': total,
            'Fecha_Emision': fecha_str,
            'Link_Validacion_SAT': link_sat,
            'Error_Lectura': None
        }
    except Exception as e:
        return {'Archivo': archivo_xml.name, 'Error_Lectura': str(e)}

def consultar_sat(rfc_emisor, rfc_receptor, total, uuid):
    """Consulta el Web Service del SAT."""
    wsdl = 'https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc?wsdl'
    try:
        client = Client(wsdl)
        expresion = f"?re={rfc_emisor}&rr={rfc_receptor}&tt={total}&id={uuid}"
        respuesta = client.service.Consulta(expresion)
        return respuesta['Estado'], respuesta['CodigoEstatus']
    except Exception as e:
        return 'Error de conexión', str(e)

# --- INTERFAZ DE USUARIO (STREAMLIT) ---
st.set_page_config(page_title="Auditor de Facturas SAT", layout="wide")

st.title("📊 Auditor de Facturas SAT (CFDI)")
st.markdown("Sube tus archivos XML para validar su estatus en el SAT, verificar si están dentro del periodo correcto y calcular la contraprestación.")

# 1. Controles de Fecha
col1, col2 = st.columns(2)
with col1:
    fecha_inicio = st.date_input("Fecha de Inicio del periodo")
with col2:
    fecha_fin = st.date_input("Fecha de Fin del periodo")

# 2. Subida de Archivos
archivos_subidos = st.file_uploader("Arrastra aquí tus facturas en formato XML", type=['xml'], accept_multiple_files=True)

# 3. Botón de Procesamiento
if st.button("🚀 Auditar Facturas", type="primary") and archivos_subidos:
    resultados = []
    
    # Barra de progreso visual
    barra_progreso = st.progress(0)
    texto_estado = st.empty()
    
    total_archivos = len(archivos_subidos)
    
    for i, archivo in enumerate(archivos_subidos):
        texto_estado.text(f"Procesando: {archivo.name} ({i+1}/{total_archivos})")
        
        datos = extraer_datos_cfdi(archivo)
        
        if not datos.get('Error_Lectura'):
            try:
                fecha_dt = datetime.strptime(datos['Fecha_Emision'], '%Y-%m-%dT%H:%M:%S').date()
                if fecha_inicio <= fecha_dt <= fecha_fin:
                    datos['Periodo_Correcto'] = 'Sí'
                else:
                    datos['Periodo_Correcto'] = 'No (Fuera de rango)'
            except (ValueError, TypeError):
                datos['Periodo_Correcto'] = 'Error en formato de fecha'
                
            if all([datos['RFC_Emisor'], datos['RFC_Receptor'], datos['Total'], datos['UUID']]):
                estado, codigo = consultar_sat(datos['RFC_Emisor'], datos['RFC_Receptor'], datos['Total'], datos['UUID'])
                datos['Estado_SAT'] = estado
                datos['Codigo_SAT'] = codigo
            else:
                datos['Estado_SAT'] = 'Datos incompletos'
                datos['Codigo_SAT'] = 'N/A'
                
        resultados.append(datos)
        barra_progreso.progress((i + 1) / total_archivos)
        
    texto_estado.text("¡Auditoría completada!")
    
    # --- PROCESAMIENTO DE DATOS ---
    df = pd.DataFrame(resultados)
    df['SubTotal'] = pd.to_numeric(df['SubTotal'], errors='coerce').fillna(0)
    df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
    df['Contraprestacion_5%'] = df['SubTotal'] * 0.05
    
    columnas_ordenadas = [
        'Archivo', 'Serie', 'Folio', 'UUID', 'RFC_Emisor', 'Nombre_Emisor', 
        'RFC_Receptor', 'Nombre_Receptor', 'Moneda', 'Descripcion', 'SubTotal', 
        'Contraprestacion_5%', 'Total', 'Fecha_Emision', 'Periodo_Correcto', 
        'Estado_SAT', 'Codigo_SAT', 'Link_Validacion_SAT', 'Error_Lectura'
    ]
    df = df[[col for col in columnas_ordenadas if col in df.columns]]
    
    # --- MÉTRICAS VISUALES ---
    facturas_validas = df[(df['Periodo_Correcto'] == 'Sí') & (df['Estado_SAT'] == 'Vigente')]
    
    st.divider()
    st.subheader("Resumen de Facturas Válidas (En rango y Vigentes)")
    
    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Total Facturas", len(facturas_validas))
    col_m2.metric("Suma Subtotal", f"${facturas_validas['SubTotal'].sum():,.2f}")
    col_m3.metric("Contraprestación (5%)", f"${facturas_validas['Contraprestacion_5%'].sum():,.2f}")
    col_m4.metric("Suma Total", f"${facturas_validas['Total'].sum():,.2f}")
    
    # Mostrar vista previa de la tabla en la web
    st.dataframe(df)
    
    # --- EXPORTAR A EXCEL ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Auditoria')
    datos_excel = output.getvalue()
    
    st.download_button(
        label="📥 Descargar Reporte en Excel",
        data=datos_excel,
        file_name=f"Auditoria_{fecha_inicio}_al_{fecha_fin}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

elif st.button("🚀 Auditar Facturas", type="primary") and not archivos_subidos:
    st.warning("⚠️ Por favor, sube al menos un archivo XML antes de procesar.")