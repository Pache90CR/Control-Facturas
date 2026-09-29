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

# ---------------------------------------------------------
# 2. BÚSQUEDA SIMPLIFICADA (EVITA COLAPSO DE MICROSOFT)
# ---------------------------------------------------------
def download_invoices_in_memory(access_token, fecha_inicio, fecha_fin):
    headers = {'Authorization': f'Bearer {access_token}'}
    
    start_iso = fecha_inicio.strftime('%Y-%m-%dT00:00:00Z')
    end_iso = (fecha_fin + datetime.timedelta(days=1)).strftime('%Y-%m-%dT00:00:00Z')

    # Consulta súper simple a Microsoft (Solo fechas, sin filtros complejos ni ordenamiento)
    endpoint = (
        "https://graph.microsoft.com/v1.0/me/messages"
        f"?$filter=receivedDateTime ge {start_iso} and receivedDateTime le {end_iso}"
        "&$select=id,subject,from,receivedDateTime,hasAttachments"
        "&$top=100"
    )
    
    records = []
    files_in_memory = {}
    logs = []
    
    logs.append(f"Consultando fechas desde {start_iso} hasta {end_iso}...")
    
    url = endpoint
    messages = []
    
    # Extraer correos (hasta 500)
    while url and len(messages) < 500:
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            data = res.json()
            messages.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
        else:
            logs.append(f"Error de Microsoft Graph: {res.status_code} - {res.text}")
            break

    logs.append(f"Se descargó la lista de {len(messages)} correos. Filtrando adjuntos en Python...")

    for msg in messages:
        # Filtrado de Python (Más eficiente y sin errores)
        if not msg.get('hasAttachments'):
            continue

        subject_raw = msg.get('subject') or "Sin Asunto"
        msg_id = msg['id']
        sender_info = msg.get('from', {}).get('emailAddress', {}) if msg.get('from') else {}
        sender = f"{sender_info.get('name', '')} <{sender_info.get('address', '')}>"
        
        raw_date = msg.get('receivedDateTime')
        msg_date = datetime.datetime.fromisoformat(raw_date.replace('Z', '+00:00')) if raw_date else datetime.datetime.now()
        
        logs.append(f"📥 Revisando: '{subject_raw}' (De: {sender})")

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
                
                if es_pdf or es_xml:
                    if 'contentBytes' in att:
                        import base64
                        file_bytes = base64.b64decode(att['contentBytes'])
                        safe_filename = f"{msg_date.strftime('%Y%m%d')}_{name_raw}"

                        if es_pdf:
                            pdf_filename = safe_filename
                            files_in_memory[safe_filename] = file_bytes
                            logs.append(f"   ✅ PDF guardado: {name_raw}")
                        elif es_xml and not xml_data:
                            try:
                                root = ET.fromstring(file_bytes)
                                def find_t(tag):
                                    for el in root.iter():
                                        if el.tag.endswith(tag): return el.text
                                    return "0"
                                emisor = find_t("Nombre") or "Proveedor"
                                total = float(find_t("TotalComprobante") or 0)
                                subtotal = float(find_t("TotalVentaNeto") or total)
                                xml_data = {"Prov": emisor, "Sub": subtotal, "Tot": total}
                                logs.append(f"   ✅ XML procesado. Total: ₡{total}")
                            except:
                                xml_data = {"Prov": sender, "Sub": 0.0, "Tot": 0.0}
                                logs.append(f"   ⚠️ XML leído pero sin formato estándar de factura.")
                else:
                    logs.append(f"   ⏭️ Ignorado: {name_raw} (No es PDF/XML)")

            if pdf_filename:
                records.append({
                    "Fecha": msg_date.strftime('%Y-%m-%d'),
                    "Proveedor": xml_data["Prov"] if xml_data else sender,
                    "Total": xml_data["Tot"] if xml_data else 0.0,
                    "Asunto": subject_raw,
                    "Archivo": pdf_filename
                })

    return records, files_in_memory, logs

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
        with st.spinner("Buscando facturas (Puede tardar de 10 a 30 segundos)..."):
            records, files_dict, logs = download_invoices_in_memory(access_token, fecha_inicio, fecha_fin)
            df_invoices = pd.DataFrame(records)
            
            st.session_state['df_invoices_graph'] = df_invoices
            st.session_state['files_in_memory'] = files_dict
            st.session_state['logs'] = logs
            
            if not df_invoices.empty:
                st.success(f"¡Listo! Se procesaron {len(df_invoices)} facturas.")
            else:
                st.error("Búsqueda completada, pero no se encontraron PDFs ni XMLs en las fechas.")

if 'logs' in st.session_state and st.session_state['logs']:
    with st.expander("🛠️ MODO DIAGNÓSTICO: Haz clic aquí para ver qué ocurrió"):
        for linea in st.session_state['logs']:
            st.text(linea)

if 'df_invoices_graph' in st.session_state and not st.session_state['df_invoices_graph'].empty:
    df = st.session_state['df_invoices_graph']
    files_dict = st.session_state['files_in_memory']

    st.divider()
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Facturas Procesadas", len(df))
    col_m2.metric("Total Acumulado", f"₡{df['Total'].sum():,.2f}")

    st.dataframe(df, use_container_width=True)

    st.subheader("⚡ Descargas")
    col1, col2 = st.columns(2)

    with col1:
        if HAS_PYPDF:
            merger = PdfWriter()
            for name in df['Archivo'].tolist():
                if name in files_dict: merger.append(io.BytesIO(files_dict[name]))
            output_pdf = io.BytesIO()
            merger.write(output_pdf)
            merger.close()
            st.download_button("📄 Descargar PDF Consolidado", data=output_pdf.getvalue(), file_name="Consolidado.pdf", mime="application/pdf", use_container_width=True)

    with col2:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for fname, fbytes in files_dict.items():
                zip_file.writestr(fname, fbytes)
        st.download_button("📦 Descargar Paquete ZIP (Todos los PDFs)", data=zip_buffer.getvalue(), file_name="Archivos.zip", mime="application/zip", use_container_width=True)
