import streamlit as st
import sqlite3
import imaplib
import email
from email.header import decode_header
import xml.etree.ElementTree as ET
import os
import pandas as pd

# ==========================================
# 1. GESTIÓN DE LA BASE DE DATOS (SQLite)
# ==========================================
DB_NAME = "contabilidad.db"

def inicializar_db():
    """Crea las tablas necesarias en SQLite si aún no existen."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Tabla de Proveedores
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS proveedores (
            id_proveedor INTEGER PRIMARY KEY AUTOINCREMENT,
            razon_social TEXT NOT NULL,
            identificacion TEXT UNIQUE NOT NULL
        )
    """)
    
    # Tabla de Facturas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS facturas (
            id_factura INTEGER PRIMARY KEY AUTOINCREMENT,
            clave TEXT UNIQUE,
            id_proveedor INTEGER,
            fecha_emision TEXT,
            trimestre TEXT,
            subtotal REAL,
            monto_impuesto REAL,
            monto_total REAL,
            categoria TEXT DEFAULT 'Por clasificar',
            estado_pago TEXT DEFAULT 'Pagado',
            archivo_origen TEXT,
            FOREIGN KEY (id_proveedor) REFERENCES proveedores (id_proveedor)
        )
    """)
    
    conn.commit()
    conn.close()

def guardar_facturas_en_db(lista_facturas):
    """Inserta las facturas y proveedores en la base de datos sin duplicados."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    registradas = 0
    duplicadas = 0

    for f in lista_facturas:
        # 1. Insertar o recuperar Proveedor
        cursor.execute("SELECT id_proveedor FROM proveedores WHERE identificacion = ?", (f["emisor_id"],))
        res = cursor.fetchone()
        
        if res:
            id_proveedor = res[0]
        else:
            cursor.execute(
                "INSERT INTO proveedores (razon_social, identificacion) VALUES (?, ?)",
                (f["emisor_nombre"], f["emisor_id"])
            )
            id_proveedor = cursor.lastrowid

        # 2. Insertar Factura (ignora si la clave ya existe)
        try:
            cursor.execute("""
                INSERT INTO facturas (clave, id_proveedor, fecha_emision, trimestre, subtotal, monto_impuesto, monto_total, archivo_origen)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                f["clave"],
                id_proveedor,
                f["fecha"],
                f["trimestre"],
                f["subtotal"],
                f["impuesto"],
                f["total"],
                f["archivo_origen"]
            ))
            registradas += 1
        except sqlite3.IntegrityError:
            duplicadas += 1 # Ya estaba registrada en la base de datos

    conn.commit()
    conn.close()
    return registradas, duplicadas

# ==========================================
# 2. LECTOR DE CORREO OUTLOOK & PARSER XML
# ==========================================
def obtener_facturas_outlook(usuario, password):
    """Se conecta a Outlook por IMAP y descarga archivos XML adjuntos de correos no leídos."""
    IMAP_SERVER = "outlook.office365.com"
    archivos_descargados = []
    
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    mail.login(usuario, password)
    mail.select("inbox")

    # Buscar correos NO LEÍDOS
    status, mensajes = mail.search(None, 'UNSEEN')
    ids_mensajes = mensajes[0].split()

    os.makedirs("temp_facturas", exist_ok=True)

    for mail_id in ids_mensajes:
        _, msg_data = mail.fetch(mail_id, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                        continue
                    
                    filename = part.get_filename()
                    if filename and filename.lower().endswith('.xml'):
                        filename_decoded, encoding = decode_header(filename)[0]
                        if isinstance(filename_decoded, bytes):
                            filename = filename_decoded.decode(encoding or 'utf-8')
                        
                        filepath = os.path.join("temp_facturas", filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        
                        archivos_descargados.append(filepath)

    mail.logout()
    return archivos_descargados

def parsear_factura_xml(ruta_xml):
    """Extrae montos, emisor y fecha del archivo XML de factura electrónica."""
    tree = ET.parse(ruta_xml)
    root = tree.getroot()

    # Eliminar namespaces para simplificar la lectura
    for elem in root.iter():
        if '}' in elem.tag:
            elem.tag = elem.tag.split('}', 1)[1]

    def buscar_texto(tag_name, default=""):
        node = root.find(f".//{tag_name}")
        return node.text.strip() if node is not None and node.text else default

    clave = buscar_texto("Clave") or buscar_texto("NumeroConsecutivo") or buscar_texto("NumeroFactura")
    emisor_nombre = buscar_texto("Nombre")
    emisor_id = buscar_texto("Numero")
    fecha = buscar_texto("FechaEmision")[:10] if buscar_texto("FechaEmision") else ""
    
    subtotal = float(buscar_texto("TotalVenta") or buscar_texto("TotalBaseImponible") or 0.0)
    impuesto = float(buscar_texto("TotalImpuesto") or 0.0)
    total = float(buscar_texto("TotalComprobante") or buscar_texto("TotalFactura") or 0.0)

    # Determinar el trimestre fiscal (Q1 a Q4)
    trimestre = ""
    if fecha:
        mes = int(fecha.split("-")[1])
        anio = fecha.split("-")[0]
        q = (mes - 1) // 3 + 1
        trimestre = f"Q{q}-{anio}"

    return {
        "clave": clave,
        "emisor_nombre": emisor_nombre,
        "emisor_id": emisor_id,
        "fecha": fecha,
        "trimestre": trimestre,
        "subtotal": subtotal,
        "impuesto": impuesto,
        "total": total,
        "archivo_origen": os.path.basename(ruta_xml)
    }

# ==========================================
# 3. INTERFAZ DE USUARIO EN STREAMLIT
# ==========================================
st.set_page_config(page_title="Control de Facturas", layout="wide")
inicializar_db()

st.title("Sistema de Gestión de Facturas y Trimestres")

# Barra lateral para configurar credenciales de Outlook
st.sidebar.header("Conexión a Outlook")
user_email = st.sidebar.text_input("Correo de Outlook", placeholder="ejemplo@outlook.com")
user_password = st.sidebar.text_input("Contraseña / App Password", type="password")

st.markdown("### 1. Sincronización desde Correo")

if st.button("Buscar Facturas Nuevas en Outlook", type="primary"):
    if not user_email or not user_password:
        st.warning("Por favor ingresa tu correo y contraseña en el menú lateral.")
    else:
        with st.spinner("Conectando a Outlook y escaneando adjuntos XML..."):
            try:
                archivos = obtener_facturas_outlook(user_email, user_password)
                if not archivos:
                    st.info("No se encontraron facturas XML nuevas no leídas.")
                else:
                    facturas_leidas = [parsear_factura_xml(f) for f in archivos]
                    st.session_state["facturas_pendientes"] = facturas_leidas
                    st.success(f"Se encontraron {len(facturas_leidas)} facturas en tu bandeja de entrada.")
            except Exception as e:
                st.error(f"Error al conectar con Outlook: {e}")

# Previsualización y guardado
if "facturas_pendientes" in st.session_state and st.session_state["facturas_pendientes"]:
    st.markdown("#### Facturas encontradas pendientes de guardar:")
    df_pendientes = pd.DataFrame(st.session_state["facturas_pendientes"])
    st.dataframe(df_pendientes, use_container_width=True)
    
    if st.button("Guardar todas en la Base de Datos"):
        registradas, duplicadas = guardar_facturas_en_db(st.session_state["facturas_pendientes"])
        st.success(f"Proceso completado: **{registradas}** facturas nuevas guardadas. ({duplicadas} ya existían en el sistema).")
        del st.session_state["facturas_pendientes"]
        st.rerun()

st.divider()

# ==========================================
# 4. VISUALIZACIÓN DE BASE DE DATOS Y EXPORTACIÓN
# ==========================================
st.markdown("### 2. Facturas Registradas en el Sistema")

conn = sqlite3.connect(DB_NAME)
query = """
    SELECT f.id_factura, f.trimestre, f.fecha_emision, p.razon_social AS proveedor, 
           p.identificacion, f.subtotal, f.monto_impuesto, f.monto_total, f.categoria
    FROM facturas f
    LEFT JOIN proveedores p ON f.id_proveedor = p.id_proveedor
    ORDER BY f.fecha_emision DESC
"""
df_facturas = pd.read_sql_query(query, conn)
conn.close()

if not df_facturas.empty:
    # Filtro por trimestre
    trimestres_disponibles = ["Todos"] + list(df_facturas["trimestre"].unique())
    trimestre_sel = st.selectbox("Filtrar por Trimestre:", trimestres_disponibles)
    
    if trimestre_sel != "Todos":
        df_filtrado = df_facturas[df_facturas["trimestre"] == trimestre_sel]
    else:
        df_filtrado = df_facturas

    st.dataframe(df_filtrado, use_container_width=True)
    
    # Resumen contable rápido
    col1, col2, col3 = st.columns(3)
    col1.metric("Subtotal Base", f"${df_filtrado['subtotal'].sum():,.2f}")
    col2.metric("Total IVA / Impuestos", f"${df_filtrado['monto_impuesto'].sum():,.2f}")
    col3.metric("Total Gastos", f"${df_filtrado['monto_total'].sum():,.2f}")

    # Exportar a Excel para el contador
    excel_data = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Descargar Reporte en CSV/Excel para el Contador",
        data=excel_data,
        file_name=f"Reporte_Facturas_{trimestre_sel}.csv",
        mime="text/csv",
    )
else:
    st.info("Aún no hay facturas guardadas en la base de datos.")
