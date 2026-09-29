import os
import zipfile
import io
import datetime
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import requests
import msal
import unicodedata

try:
    from pypdf import PdfWriter
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

st.set_page_config(page_title="Control de Facturas - Outlook", page_icon="📩", layout="wide")
st.markdown('<meta name="google" content="notranslate">', unsafe_allow_html=True)

CLIENT_ID = st.secrets.get("CLIENT_ID", "TU_CLIENT_ID_COPIADO_DE_AZURE")
REFRESH_TOKEN = st.secrets.get("REFRESH_TOKEN", None)
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.Read"]

# Palabras estrictas para DESCARTAR (Estados de cuenta, bancos, tiquetes, etc.)
PALABRAS_EXCLUIDAS = [
    "estado de cuenta", "resumen de cuenta", "extracto", 
    "boletin", "publicidad", "oferta", "newsletter", "promocion",
    "notificacion", "comercio afiliado", "tiquete", "banco nacional", "bncr"
]

# Palabras obligatorias para ACEPTAR (Solo facturas y documentos válidos)
PALABRAS_CLAVE_PERMITIDAS = [
    "factura", "comprobante", "electronico", "electronica", 
    "nota de credito", "documento electronico", "fe-", "fe_"
]

def limpiar_texto(texto):
    if not texto: return ""
    texto_norm = unicodedata.normalize('NFD', texto)
    return "".join(c for c in texto_norm if unicodedata.category(c) != 'Mn').lower()

# ---------------------------------------------------------
# 1. AUTENTICACIÓN
# ---------------------------------------------------------
def get_access_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    if REFRESH_TOKEN:
        result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=SCOPES)
        if "access_token" in result:
            return result["access_token"]
        else:
            st.error(f"Error con el token permanente: {result.get('error_description')}")
            return None

    st.warning("⚠️ **Generador de Token Permanente (Paso Único)**")
    if "flow" not in st.session_state or st.session_state["flow"] is None:
        st.session_state["flow"] = app.initiate_device_flow(scopes=SCOPES)
    flow = st.session_state["flow"]

    st.markdown(f"""
        1. Entra a: **[{flow['verification_uri']}]({flow['verification_uri']})**
        2. Escribe este código: **`{flow['user_code']}`**
    """)
    if st.button("🔑 Generar Refresh Token", use_container_width=True):
        result = app.acquire_token_by_device_flow(flow)
        if "refresh_token" in result:
            st.success("Copia esto en Secrets de Streamlit:")
            st.code(f'REFRESH_TOKEN = "{result["refresh_token"]}"', language="toml")
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
# 2. BÚSQUEDA Y FILTRADO INTELIGENTE
# ---------------------------------------------------------
def download_invoices_in_memory(access_token, fecha_inicio, fecha_fin):
    headers = {'Authorization': f'Bearer {access_token}'}
    
    start_iso = fecha_inicio.strftime('%Y-%m-%dT00:00:00Z')
    end_iso = (fecha_fin + datetime.timedelta(days=1)).strftime('%Y-%m-%dT00:00:00Z')

    endpoint = (
        "https://graph.microsoft.com/v1.0/me/messages"
        f"?$filter=receivedDateTime ge {start_iso} and receivedDateTime le {end_iso}"
        "&$select=id,subject,from,receivedDateTime,hasAttachments"
        "&$top=100"
    )
    
    records = []
    files_in_memory = {}
    logs = []
    
    url = endpoint
    messages = []
    
    while url and len(messages) < 500:
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            data = res.json()
            messages.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
        else:
            break

    for msg in messages:
        if not msg.get('hasAttachments'):
            continue

        subject_raw = msg.get('subject') or "Sin Asunto"
        subject_clean = limpiar_texto(subject_raw)
        
        sender_info = msg.get('from', {}).get('emailAddress', {}) if msg.get('from') else {}
        sender_name = limpiar_texto(sender_info.get('name', ''))
        sender_email = limpiar_texto(sender_info.get('address', ''))
        sender_full = f"{sender_name} {sender_email}"

        # 1. FILTRAR EXCLUSIONES (Si el asunto o emisor contiene palabras prohibidas, se descarta)
        if any(excl in subject_clean or excl in sender_full for excl in PALABRAS_EXCLUIDAS):
            logs.append(f"❌ Descartado por filtro: '{subject_raw}' (Emisor: {sender_email})")
            continue

        # 2. FILTRAR PERMISOS (Debe contener alguna palabra clave de factura)
        es_factura_valida = any(p in subject_clean for p in PALABRAS_CLAVE_PERMITIDAS)

        msg_id = msg['id']
        raw_date = msg.get('receivedDateTime')
        msg_date = datetime.datetime.fromisoformat(raw_date.replace('Z', '+00:00')) if raw_date else datetime.datetime.now()

        attach_endpoint = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}/attachments"
        attach_res = requests.get(attach_endpoint, headers=headers)
        
        if attach_res.status_code == 200:
            attachments = attach_res.json().get('value', [])
            xml_data = None
            pdf_filename = ""

            for att in attachments:
                name_raw = att.get('name', '')
                name_clean = limpiar_texto(name_raw)

                es_pdf = name_clean.endswith('.pdf')
                es_xml = name_clean.endswith('.xml')
                
                # Verificar si el nombre del archivo también indica factura
                if any(p in name_clean for p in PALABRAS_CLAVE_PERMITIDAS):
                    es_factura_valida = True

                if (es_pdf or es_xml) and 'contentBytes' in att:
                    import base64
                    file_bytes = base64.b64decode(att['contentBytes'])
                    safe_filename = f"{msg_date.strftime('%Y%m%d')}_{name_raw}"

                    if es_pdf:
                        pdf_filename = safe_filename
                        files_in_memory[safe_filename] = file_bytes
                    elif es_xml and not xml_data:
                        try:
                            xml_data = parse_xml_invoice(file_bytes)
                        except:
                            pass

            # Solo agregar si pasó el filtro de factura válida y tiene monto mayor a 0 o proveedor detectado
            if pdf_filename and es_factura_valida:
                total_monto = xml_data["Total"] if xml_data else 0.0
                
                # Descartar si el total es 0 (evita falsos positivos sin monto)
                if total_monto > 0:
                    records.append({
                        "Fecha": msg_date.strftime('%Y-%m-%d'),
                        "Trimestre": get_quarter(msg_date.month),
                        "Proveedor": xml_data["Proveedor"] if xml_data else sender_info.get('name', 'Proveedor'),
                        "Subtotal": xml_data["Subtotal"] if xml_data else 0.0,
                        "IVA": xml_data["IVA"] if xml_data else 0.0,
                        "Total": total_monto,
                        "Asunto": subject_raw,
                        "Archivo": pdf_filename
                    })
                    logs.append(f"✅ Factura aceptada: '{subject_raw}' - Monto: ₡{total_monto}")
                else:
                    logs.append(f"⚠️ Descartado (Total en 0): '{subject_raw}'")

    return records, files_in_memory, logs

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
st.title("📩 Control y Gestor de Facturas desde Outlook")

access_token = get_access_token()

if access_token:
    st.sidebar.success("🟢 Conexión a Outlook activa")
    fecha_inicio = st.sidebar.date_input("Fecha Inicio", datetime.date(2026, 9, 1))
    fecha_fin = st.sidebar.date_input("Fecha Fin", datetime.date(2026, 9, 28))

    if st.sidebar.button("🔄 Buscar y Descargar Facturas", use_container_width=True):
        with st.spinner("Filtrando facturas válidas..."):
            records, files_dict, logs = download_invoices_in_memory(access_token, fecha_inicio, fecha_fin)
            df_invoices = pd.DataFrame(records)
            
            st.session_state['df_invoices_graph'] = df_invoices
            st.session_state['files_in_memory'] = files_dict
            st.session_state['logs'] = logs
            
            if not df_invoices.empty:
                st.success(f"¡Listo! Se procesaron {len(df_invoices)} facturas válidas.")
            else:
                st.warning("No se encontraron facturas con los nuevos filtros.")

if 'logs' in st.session_state and st.session_state['logs']:
    with st.expander("🛠️ Ver registro de filtros (Diagnóstico)"):
        for linea in st.session_state['logs']:
            st.text(linea)

if 'df_invoices_graph' in st.session_state and not st.session_state['df_invoices_graph'].empty:
    df = st.session_state['df_invoices_graph']
    files_dict = st.session_state['files_in_memory']

    st.divider()
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Facturas Válidas", len(df))
    col_m2.metric("Subtotal Acumulado", f"₡{df['Subtotal'].sum():,.2f}")
    col_m3.metric("Total Acumulado", f"₡{df['Total'].sum():,.2f}")

    st.dataframe(df[['Fecha', 'Proveedor', 'Subtotal', 'IVA', 'Total', 'Asunto']], use_container_width=True)

    st.subheader("⚡ Descargas")
    col1, col2, col3 = st.columns(3)

    with col1:
        if HAS_PYPDF:
            merged_pdf_bytes = merge_pdfs_from_memory(files_dict, df['Archivo'].tolist())
            if merged_pdf_bytes:
                st.download_button("📄 Descargar PDF Consolidado", data=merged_pdf_bytes, file_name="Facturas_Consolidadas.pdf", mime="application/pdf", use_container_width=True)

    with col2:
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Detalle')
        st.download_button("📊 Descargar Excel Resumen", data=output_excel.getvalue(), file_name="Reporte_Facturas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with col3:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for fname, fbytes in files_dict.items():
                zip_file.writestr(fname, fbytes)
        st.download_button("📦 Descargar Paquete ZIP", data=zip_buffer.getvalue(), file_name="Facturas_Zip.zip", mime="application/zip", use_container_width=True)
