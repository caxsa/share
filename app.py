import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext


st.set_page_config(page_title="📦 Solicitud de Materiales", layout="centered")

st.markdown("""
    <style>
        .main { background-color: #f5f7fa; }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            border-radius: 8px;
            padding: 0.5em 1.5em;
        }
        .stDownloadButton>button {
            background-color: #2196F3;
            color: white;
            border-radius: 8px;
        }
    </style>
""", unsafe_allow_html=True)

st.title("🧰 Registro de Solicitudes de Materiales")

archivo_local = "solicitudes.xlsx"

# Cargar datos desde archivo local si existe
if os.path.exists(archivo_local):
    df_historico = pd.read_excel(archivo_local)
    st.session_state.solicitudes = df_historico.to_dict(orient="records")
else:
    st.session_state.solicitudes = []

# Formulario de solicitud
with st.form("form_solicitud"):
    tecnico = st.text_input("👨‍🔧 Técnico", max_chars=50)
    proyecto = st.text_input("🏗️ Proyecto", max_chars=100)
    material = st.text_input("🧱 Material", max_chars=100)
    unidades = st.number_input("🔢 Unidades", min_value=1, step=1)
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    enviar = st.form_submit_button("➕ Generar solicitud")

    if enviar:
        solicitud = {
            "Técnico": tecnico,
            "Proyecto": proyecto,
            "Material": material,
            "Unidades": int(unidades),
            "Fecha": fecha
        }
        st.session_state.solicitudes.append(solicitud)
        st.success("✅ Solicitud registrada correctamente.")

# Mostrar registros
df = pd.DataFrame(st.session_state.solicitudes)
if not df.empty:
    st.subheader("📋 Solicitudes registradas")
    st.dataframe(df, use_container_width=True)

    st.subheader("📦 Total de materiales solicitados")
    total = df.groupby(["Material"])["Unidades"].sum().reset_index()
    st.dataframe(total, use_container_width=True)

    # Guardar archivo permanente local
    df.to_excel(archivo_local, index=False)

    # Subida a SharePoint
    st.subheader("☁️ Sincronización con SharePoint")

    # DATOS A CAMBIAR
    sharepoint_url = "https://tusitio.sharepoint.com/sites/TuSitio"
    usuario = "tucorreo@empresa.com"
    contrasena = "tu_contraseña"
    carpeta_destino = "Documentos compartidos/Solicitudes"

    try:
        ctx_auth = AuthenticationContext(sharepoint_url)
        if ctx_auth.acquire_token_for_user(usuario, contrasena):
            ctx = ClientContext(sharepoint_url, ctx_auth)
            carpeta = ctx.web.get_folder_by_server_relative_url(carpeta_destino)
            with open(archivo_local, "rb") as archivo:
                nombre_archivo = os.path.basename(archivo_local)
                carpeta.upload_file(nombre_archivo, archivo.read()).execute_query()
            st.success("✅ Archivo subido a SharePoint correctamente.")
        else:
            st.error("❌ Error autenticando con SharePoint.")
    except Exception as e:
        st.error(f"Error subiendo a SharePoint: {e}")

    # Descargar Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Solicitudes")
        total.to_excel(writer, index=False, sheet_name="Totales")
    output.seek(0)

    fecha_excel = datetime.now().strftime("%Y-%m-%d")
    st.download_button(
        label="📄 Descargar Excel",
        data=output,
        file_name=f"solicitudes_materiales_{fecha_excel}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("No hay solicitudes registradas todavía.")
