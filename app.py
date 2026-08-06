import os
import zipfile
import io
import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import requests
import msal

try:
    from pypdf import PdfWriter
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

CLIENT_ID = "6f6074bd-8f47-4589-a691-7a05cebae707"
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Read"]

PALABRAS_EXCLUIDAS = [
    "estado de cuenta", "resumen de cuenta", 
    "extracto", "boletin", "publicidad", "oferta"
]

PALABRAS_CLAVE_PERMITIDAS = [
    "factura", "comprobante", "electronico", "electronica", 
    "tiquete", "nota de credito", "documento electronico", "fe-"
]

# ---------------------------------------------------------
# 1. AUTENTICACIÓN
# ---------------------------------------------------------
def get_graph_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:
        try:
            result = app.acquire_token_interactive(scopes=SCOPES)
        except Exception as e:
            st.error(f"Error de autenticación: {e}")
            return None

    if result and "access_token" in result:
        return result["access_token"]
    else:
        st.error("No se pudo obtener el token de acceso.")
        return None

def get_quarter(month):
    if month in [1, 2, 3]: return "Q1 (Ene-Mar)"
    elif month in [4, 5, 6]: return "Q2 (Abr-Jun)"
    elif month in [7, 8, 9]: return "Q3 (Jul-Sep)"
    else: return "Q4 (Oct-Dic)"

def parse_xml_invoice(xml_bytes):
    try:
        root = ET.fromstring(xml_bytes)
        def find_text(tag_name):
            for elem in root.iter():
                if elem.tag.endswith(tag_name):
                    return elem.text
            return "0"

        emisor = find_text("Nombre") or "Proveedor Desconocido"
        subtotal = float(find_text("TotalComprobante") or find_text("TotalVentaNeto") or 0)
        iva = float(find_text("TotalImpuesto") or 0)
        total = float(find_text("TotalComprobante") or 0)

        if subtotal == total and iva > 0:
            subtotal = total - iva

        return {"Proveedor": emisor, "Subtotal": subtotal, "IVA": iva, "Total": total}
    except Exception:
        return None

# ---------------------------------------------------------
# 2. PROCESAMIENTO EN MEMORIA RAM (SÚPER RÁPIDO)
# ---------------------------------------------------------
def download_invoices_in_memory(access_token, fecha_inicio, fecha_fin):
    headers = {'Authorization': f'Bearer {access_token}'}
    start_iso = fecha_inicio.strftime('%Y-%m-%dT00:00:00Z')
    end_iso = fecha_fin.strftime('%Y-%m-%dT23:59:59Z')

    endpoint = (
        "https://graph.microsoft.com/v1.0/me/messages"
        f"?$filter=hasAttachments eq true and receivedDateTime ge {start_iso} and receivedDateTime le {end_iso}"
        "&$select=id,subject,from,receivedDateTime"
        "&$top=100"
    )
    
    response = requests.get(endpoint, headers=headers)
    if response.status_code != 200:
        st.error(f"Error consultando Microsoft Graph: {response.status_code}")
        return [], {}

    messages = response.json().get('value', [])
    records = []
    files_in_memory = {}  # Guarda los bytes en memoria en lugar de disco

    for msg in messages:
        subject = msg.get('subject') or "Sin Asunto"
        subject_lower = subject.lower()

        if any(excl in subject_lower for excl in PALABRAS_EXCLUIDAS):
            continue

        msg_id = msg['id']
        sender_info = msg.get('from', {}).get('emailAddress', {}) if msg.get('from') else {}
        sender = f"{sender_info.get('name', '')} <{sender_info.get('address', '')}>"
        
        raw_date = msg.get('receivedDateTime')
        msg_date = datetime.datetime.fromisoformat(raw_date.replace('Z', '+00:00')) if raw_date else datetime.datetime.now()

        year = str(msg_date.year)
        quarter = get_quarter(msg_date.month)

        attach_endpoint = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/attachments?$select=id,name,contentBytes"
        attach_res = requests.get(attach_endpoint, headers=headers)
        
        if attach_res.status_code == 200:
            attachments = attach_res.json().get('value', [])
            xml_data = None
            pdf_filename = ""
            pdf_bytes = None

            for att in attachments:
                name = att.get('name', '')
                name_lower = name.lower()

                es_pdf = name_lower.endswith('.pdf')
                es_xml = name_lower.endswith('.xml')
                es_factura = any(p in subject_lower for p in PALABRAS_CLAVE_PERMITIDAS) or \
                             any(p in name_lower for p in PALABRAS_CLAVE_PERMITIDAS)

                if (es_pdf or es_xml) and es_factura and 'contentBytes' in att:
                    import base64
                    file_bytes = base64.b64decode(att['contentBytes'])
                    safe_filename = f"{msg_date.strftime('%Y%m%d')}_{name}"

                    if es_pdf:
                        pdf_filename = safe_filename
                        pdf_bytes = file_bytes
                        files_in_memory[safe_filename] = file_bytes
                    elif es_xml and not xml_data:
                        xml_data = parse_xml_invoice(file_bytes)

            if pdf_filename:
                records.append({
                    "Fecha": msg_date.strftime('%Y-%m-%d'),
                    "Año": year,
                    "Trimestre": quarter,
                    "Proveedor": xml_data["Proveedor"] if xml_data else sender,
                    "Subtotal": xml_data["Subtotal"] if xml_data else 0.0,
                    "IVA": xml_data["IVA"] if xml_data else 0.0,
                    "Total": xml_data["Total"] if xml_data else 0.0,
                    "Asunto": subject,
                    "Archivo": pdf_filename
                })

    return records, files_in_memory

def merge_pdfs_from_memory(files_dict, filenames):
    if not HAS_PYPDF: return None
    merger = PdfWriter()
    for name in filenames:
        if name in files_dict:
            try:
                merger.append(io.BytesIO(files_dict[name]))
            except Exception:
                pass
    output_pdf = io.BytesIO()
    merger.write(output_pdf)
    merger.close()
    return output_pdf.getvalue()

# ---------------------------------------------------------
# 3. INTERFAZ STREAMLIT
# ---------------------------------------------------------
st.set_page_config(page_title="Control de Facturas - Outlook", page_icon="📩", layout="wide")
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

st.title("📩 Control y Gestor de Facturas desde Outlook")
st.markdown("Busca, consolida datos financieros de facturas y unifica documentos para impresión.")

st.sidebar.header("🎯 Rango de Fechas a Consultar")
fecha_inicio = st.sidebar.date_input("Fecha Inicio", datetime.date(2026, 1, 1))
fecha_fin = st.sidebar.date_input("Fecha Fin", datetime.date(2026, 3, 31))

st.sidebar.divider()

if st.sidebar.button("🔄 Buscar y Descargar Facturas", use_container_width=True):
    if CLIENT_ID == "TU_CLIENT_ID_COPIADO_DE_AZURE":
        st.sidebar.error("Por favor pega tu Client ID de Azure en la variable CLIENT_ID de app.py.")
    else:
        with st.spinner("Procesando facturas directamente en memoria..."):
            token = get_graph_token()
            if token:
                records, files_dict = download_invoices_in_memory(token, fecha_inicio, fecha_fin)
                df_invoices = pd.DataFrame(records)
                
                if not df_invoices.empty:
                    df_invoices = df_invoices.sort_values(by="Fecha", ascending=True)
                    st.session_state['df_invoices_graph'] = df_invoices
                    st.session_state['files_in_memory'] = files_dict
                    st.success(f"¡Listo! Se procesaron {len(df_invoices)} facturas.")
                else:
                    st.info("No se encontraron facturas en el rango de fechas seleccionado.")

if 'df_invoices_graph' not in st.session_state:
    st.session_state['df_invoices_graph'] = pd.DataFrame()
    st.session_state['files_in_memory'] = {}

df = st.session_state['df_invoices_graph']
files_dict = st.session_state['files_in_memory']

if not df.empty:
    st.divider()
    st.subheader("📊 Consolidado de Datos Financieros del Trimestre")
    
    total_subtotal = df['Subtotal'].sum()
    total_iva = df['IVA'].sum()
    total_general = df['Total'].sum()

    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
    col_m1.metric("Facturas Procesadas", len(df))
    col_m2.metric("Subtotal Acumulado", f"₡{total_subtotal:,.2f}")
    col_m3.metric("IVA Acumulado", f"₡{total_iva:,.2f}")
    col_m4.metric("Total Acumulado", f"₡{total_general:,.2f}")

    st.write("### Listado Ordenado por Fecha de Emisión")
    st.dataframe(
        df[['Fecha', 'Trimestre', 'Proveedor', 'Subtotal', 'IVA', 'Total', 'Asunto', 'Archivo']], 
        use_container_width=True
    )

    st.divider()
    st.subheader("⚡ Opciones de Descarga y Consolidación")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### 🖨️ PDF Unificado (Impresión)")
        st.caption("Junta todos los PDFs en un solo documento listo para imprimir.")
        
        if HAS_PYPDF:
            merged_pdf_bytes = merge_pdfs_from_memory(files_dict, df['Archivo'].tolist())
            if merged_pdf_bytes:
                st.download_button(
                    label="📄 Descargar PDF Consolidado",
                    data=merged_pdf_bytes,
                    file_name=f"Facturas_Consolidadas_{fecha_inicio}_al_{fecha_fin}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

    with col2:
        st.markdown("### 📈 Reporte Resumen XML")
        st.caption("Excel consolidado con Subtotal, IVA, Total acumulado y proveedores.")
        
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df[['Fecha', 'Trimestre', 'Proveedor', 'Subtotal', 'IVA', 'Total', 'Asunto', 'Archivo']].to_excel(
                writer, index=False, sheet_name='Detalle_Por_Fecha'
            )
            df_prov = df.groupby('Proveedor')[['Subtotal', 'IVA', 'Total']].sum().reset_index()
            df_prov.to_excel(writer, index=False, sheet_name='Desglose_Por_Proveedor')

            df_resumen = pd.DataFrame([{
                "Rango Fechas": f"{fecha_inicio} al {fecha_fin}",
                "Cant. Facturas": len(df),
                "Subtotal Acumulado": total_subtotal,
                "IVA Acumulado": total_iva,
                "Total Acumulado": total_general
            }])
            df_resumen.to_excel(writer, index=False, sheet_name='Resumen_Trimestre')

        st.download_button(
            label="📊 Descargar Reporte Resumen (XML)",
            data=output_excel.getvalue(),
            file_name=f"Reporte_Resumen_Consolidado_{fecha_inicio}_al_{fecha_fin}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col3:
        st.markdown("### 📦 Paquete Completo (.ZIP)")
        st.caption("Descarga todos los PDFs en un solo archivo comprimido.")
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for fname, fbytes in files_dict.items():
                zip_file.writestr(fname, fbytes)

        st.download_button(
            label="📁 Descargar Paquete ZIP",
            data=zip_buffer.getvalue(),
            file_name=f"Facturas_Archivos_{fecha_inicio}_al_{fecha_fin}.zip",
            mime="application/zip",
            use_container_width=True
        )
else:
    st.info("Selecciona el rango de fechas en el panel izquierdo y haz clic en **'Buscar y Descargar Facturas'**.")
