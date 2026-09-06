# -*- coding: utf-8 -*-
"""
Plataforma Integral de Gestión Aduanera — Zona Franca
================================================================================
Módulos: 
1. Procesador Masivo de Declaraciones de Importación (DIM - Formulario 500)
2. Procesador Masivo de Actas de Inventario e Inconsistencias (Tránsito Aduanero)
"""

import io
import re
import zipfile
import os
import base64
from datetime import datetime

import fitz  # PyMuPDF para DIMs
import pdfplumber  # Para Actas de Tránsito
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuración general y Estilos Corporativos con Fondo Personalizado
# --------------------------------------------------------------------------

st.set_page_config(
    page_title="Plataforma Aduanera | Zona Franca",
    page_icon="🏢",
    layout="wide",
)

fondo_css = ""
if os.path.exists("Fondo ZFC.png"):
    with open("Fondo ZFC.png", "rb") as f:
        fondo_bytes = f.read()
    fondo_base64 = base64.b64encode(fondo_bytes).decode()
    fondo_css = f"""
    .stApp {{
        background-image: linear-gradient(rgba(244, 247, 246, 0.9), rgba(244, 247, 246, 0.9)), url("data:image/png;base64,{fondo_base64}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
    }}
    """
else:
    fondo_css = """
    .stApp {
        background-color: #F4F7F6;
    }
    """

st.markdown(f"""
<style>
    :root {{
        --zf-green-dark: #1B4D3E;
        --zf-green-medium: #2C6B56;
        --zf-olive: #8A9A28;
        --zf-card-bg: #FFFFFF;
        --zf-text-main: #2C3E50;
    }}

    {fondo_css}

    h1, h2, h3 {{
        color: var(--zf-green-dark) !important;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }}

    div[data-testid="stVerticalBlock"] > div[style*="border"] {{
        background-color: var(--zf-card-bg);
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        border: 1px solid #E1E8E5 !important;
        padding: 20px;
    }}

    .stButton>button {{
        background-color: var(--zf-green-dark);
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: 600;
        padding: 0.5rem 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }}

    .stButton>button:hover {{
        background-color: var(--zf-green-medium);
        color: white;
        border: none;
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }}

    div[data-testid="stMetricValue"] {{
        color: var(--zf-green-dark);
        font-weight: 700;
    }}
    div[data-testid="stMetricLabel"] {{
        color: #556B2F;
        font-weight: 600;
    }}

    .stDataFrame {{
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #E1E8E5;
    }}
</style>
""", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# Menú de Navegación Lateral (Selector de Módulos)
# --------------------------------------------------------------------------

st.sidebar.markdown("### 🧭 Navegación")
st.sidebar.caption("Seleccione el Módulo Operativo:")

modulo_seleccionado = st.sidebar.radio(
    "Portal:",
    [
        "📑 Procesador de DIM",
        "📋 Actas de Tránsito / PICIZ"
    ],
    label_visibility="collapsed"
)

st.sidebar.divider()

# Carga de Logo en Sidebar o Cabecera
logo_path = None
for filename in ["LOGO ZFS-ZFC.jpeg", "logo.jpeg", "logo.jpg", "logo.png"]:
    if os.path.exists(filename):
        logo_path = filename
        break

if logo_path:
    st.sidebar.image(logo_path, width=200)

st.sidebar.markdown("---")
st.sidebar.markdown("**Zona Franca de Cúcuta**  \nOperada por Zona Franca Santander")


# ==========================================================================
# MÓDULO 1: PROCESADOR DE DIM (Declaraciones de Importación)
# ==========================================================================

if modulo_seleccionado == "📑 Procesador de DIM":
    
    COLUMNAS_DIM = [
        "Número de formulario",
        "NIT Importador",
        "Razón Social Importador",
        "Factura",
        "Manifiesto de carga",
        "Documento de transporte",
        "Cod. País Procedencia",
        "Cod. Modo Transporte",
        "Código de Bandera",
        "Tasa de Cambio",
        "Subpartida Arancelaria",
        "Cod. País Origen",
        "Cod. País Compra",
        "Peso Bruto (Kgs)",
        "Peso Neto (Kgs)",
        "Código de Embalaje",
        "Cod. Unidad Comercial (76)",
        "Cantidad (77)",
        "No. Bultos",
        "Valor FOB (USD)",
        "Sumatoria Fletes/Seguros/Otros (USD)",
        "Acta de Inspección No.",
        "Levante No.",
        "Fecha del Levante",
        "Archivo",
        "Campos_no_encontrados",
    ]
    
    MONTO_DIM = r"[\d\.,]+"
    ENTERO_MILES_DIM = r"[\d\.,]+"

    def limpiar_monto_dim(val_str):
        if not val_str:
            return 0.0
        val_str = str(val_str).strip()
        val_str = re.sub(r'[^\d.,]', '', val_str)
        if not val_str:
            return 0.0
        if ',' in val_str:
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            partes = val_str.split('.')
            if len(partes) == 2:
                if len(partes[1]) == 3 and len(partes[0]) <= 3:
                    val_str = val_str.replace('.', '')
            elif len(partes) > 2:
                if len(partes[-1]) == 2:
                    val_str = "".join(partes[:-1]) + "." + partes[-1]
                else:
                    val_str = val_str.replace('.', '')
        try:
            return float(val_str)
        except ValueError:
            return 0.0

    def _buscar_dim(patron, texto, flags=re.IGNORECASE, grupo=1):
        m = re.search(patron, texto, flags)
        if m:
            try:
                return m.group(grupo).strip()
            except IndexError:
                return None
        return None

    def dividir_dims(texto: str):
        partes = re.split(r"(?=Declaraci[oó]n de Importaci[oó]n)", texto, flags=re.IGNORECASE)
        return [p for p in partes if re.search(r"N[uú]mero de formulario", p, re.IGNORECASE)]

    def extraer_campos_dim(chunk_texto: str, texto_completo: str, nombre_archivo: str) -> dict:
        faltantes = []

        def campo(nombre, patron, grupo=1, default=""):
            valor = _buscar_dim(patron, chunk_texto, grupo=grupo)
            if not valor:
                valor = _buscar_dim(patron, texto_completo, grupo=grupo)
            if not valor:
                faltantes.append(nombre)
                return default
            return " ".join(valor.split())

        numero_formulario = campo("Número de formulario", r"4\s*\.\s*N[uú]mero de formulario\s*\n?\s*(\S+)")
        nit_importador = campo("NIT Importador", r"5\s*\.\s*N[uú]mero de Identificaci[oó]n Tributaria \(NIT\)\s*(\d{9,10})")
        razon_social = campo("Razón Social Importador", r"11\s*\.\s*Apellidos y nombres o Raz[oó]n Social\s*([^\n]+)")
        factura = campo("Factura", r"51\s*\.\s*No\.\s*de\s*factura\s*\n\s*(\S+)")
        
        manifiesto_carga = campo("Manifiesto de carga", r"42\s*\.?\s*Manifiesto\s+de\s+carga\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")
        if not manifiesto_carga:
            manifiesto_carga = campo("Manifiesto de carga", r"42\s*\.?\s*Manifiesto\s+de\s+carga[^\n]*\n\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")

        documento_transporte = campo("Documento de transporte", r"44\s*\.?\s*Documento\s+de\s+transporte\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")
        if not documento_transporte:
            documento_transporte = campo("Documento de transporte", r"44\s*\.?\s*Documento\s+de\s+transporte[^\n]*\n\s*(?:No\.?\s*)?([A-Za-z0-9\-]+)")

        cod_pais_procedencia = campo("Cod. País Procedencia", r"53\s*\.\s*(?:C[oó]d\.?\s*)?pa[ií]s\s+(?:de\s+)?procedencia\s*([A-Za-z0-9]{2,3})")
        cod_modo_transporte = campo("Cod. Modo Transporte", r"54\s*\.\s*Cod\.\s*Modo\s*Transporte\s*(\d)")
        codigo_bandera = campo("Código de Bandera", r"55\s*\.\s*C[oó]digo\s+(?:de\s+)?bandera\s*([A-Za-z0-9]{2,3})")
        tasa_cambio = campo("Tasa de Cambio", r"Tasa de cambio\s*\$?\s*cvs\.?\s*\n?\s*([\d.,]+)")
        subpartida = campo("Subpartida Arancelaria", r"59\s*\.\s*Subpartida arancelaria\s*(\d{10})")
        cod_pais_origen = campo("Cod. País Origen", r"66\s*\.\s*(?:C[oó]d\.?\s*)?pa[ií]s\s+(?:de\s+)?origen\s*([A-Za-z0-9]{2,3})")
        cod_pais_compra = campo("Cod. País Compra", r"70\s*\.\s*Cod\s*\.\s*pa[ií]s\s*\n?\s*compra\s*(\d{2,3})")
        codigo_embalaje = campo("Código de Embalaje", r"73\s*\.\s*C[oó]digo\s*\n?\s*embalaje\s*([A-Za-z0-9]{1,4})")
        
        peso_bruto = limpiar_monto_dim(campo("Peso Bruto (Kgs)", r"71\s*\.\s*Peso bruto kgs\.\s*dcms\.\s*(" + MONTO_DIM + ")"))
        peso_neto = limpiar_monto_dim(campo("Peso Neto (Kgs)", r"72\s*\.\s*Peso neto kgs\.\s*dcms\.\s*(" + MONTO_DIM + ")"))
        valor_fob = limpiar_monto_dim(campo("Valor FOB (USD)", r"78\s*\.\s*Valor FOB USD\s*(" + MONTO_DIM + ")"))
        sumatoria_fletes = limpiar_monto_dim(campo("Sumatoria Fletes/Seguros/Otros (USD)", r"82\s*\.\s*Sumatoria de fletes,?\s*seguros\s*\n?\s*y otros gastos USD\s*(" + MONTO_DIM + ")"))
        
        cod_unidad_comercial = campo("Cod. Unidad Comercial (76)", r"76\s*\.\s*Cod\.?\s*unidad\s*comercial\s*[\r\n\s]+([A-Za-z]{1,4})\b")
        if not cod_unidad_comercial:
            cod_unidad_comercial = campo("Cod. Unidad Comercial (76)", r"76\s*\.\s*Cod\.?\s*unidad\s*comercial\s+([A-Za-z]{1,4})\b")

        cantidad_str = campo("Cantidad (77)", r"77\s*\.\s*Cantidad\s*(?:dcms\.?)?\s*[\r\n\s]+(" + MONTO_DIM + ")")
        if not cantidad_str:
            cantidad_str = campo("Cantidad (77)", r"77\s*\.\s*Cantidad\s*(?:dcms\.?)?\s*(" + MONTO_DIM + ")")
        cantidad_comercial = limpiar_monto_dim(cantidad_str)

        n_bultos_str = campo("No. Bultos", r"74\s*\.\s*No\.\s*bultos\s*(" + ENTERO_MILES_DIM + ")")
        try:
            no_bultos = int(re.sub(r'[^\d]', '', n_bultos_str)) if n_bultos_str else 0
        except ValueError:
            no_bultos = 0

        acta_inspeccion = ""
        m_acta = re.search(r"ACTA\s+DE\s+INSPECCI[OÓ]N\s*(?:No\.?|Número)?\s*[:\.]?\s*([0-9]{8,15})", texto_completo, re.IGNORECASE)
        if m_acta:
            acta_inspeccion = m_acta.group(1).strip()

        levante_no = ""
        m_lev_box = re.search(r"134\.?\s*Levante\s+No\.?\s*([0-9]{8,15})", chunk_texto, re.IGNORECASE)
        if m_lev_box:
            levante_no = m_lev_box.group(1).strip()
        if not levante_no:
            m_lev_gen = re.search(r"(?:Levante|Auto(?:rización)?)\s*(?:No\.?|Número)?\s*[:\.]?\s*([0-9]{8,15})", texto_completo, re.IGNORECASE)
            if m_lev_gen:
                levante_no = m_lev_gen.group(1).strip()
        if not levante_no:
            faltantes.append("Levante No.")

        fecha_levante = campo("Fecha del Levante", r"135\.?\s*Fecha[^\d\n]*(\d{4}\s*[-/\.]\s*\d{2}\s*[-/\.]\s*\d{2})")
        if not fecha_levante:
            m_fec = re.search(r"\b(20\d{2}[-/\.](?:0[1-9]|1[0-2])[-/\.](?:0[1-9]|[12]\d|3[01]))\b", texto_completo)
            if m_fec:
                fecha_levante = m_fec.group(1)
                if "Fecha del Levante" in faltantes:
                    faltantes.remove("Fecha del Levante")
        if fecha_levante:
            fecha_levante = re.sub(r"\s+", "", fecha_levante)
            fecha_levante = re.sub(r"[/.]", "-", fecha_levante)

        return {
            "Número de formulario": numero_formulario,
            "NIT Importador": nit_importador,
            "Razón Social Importador": razon_social,
            "Factura": factura,
            "Manifiesto de carga": manifiesto_carga,
            "Documento de transporte": documento_transporte,
            "Cod. País Procedencia": cod_pais_procedencia,
            "Cod. Modo Transporte": cod_modo_transporte,
            "Código de Bandera": codigo_bandera,
            "Tasa de Cambio": tasa_cambio,
            "Subpartida Arancelaria": subpartida,
            "Cod. País Origen": cod_pais_origen,
            "Cod. País Compra": cod_pais_compra,
            "Peso Bruto (Kgs)": peso_bruto,
            "Peso Neto (Kgs)": peso_neto,
            "Código de Embalaje": codigo_embalaje,
            "Cod. Unidad Comercial (76)": cod_unidad_comercial,
            "Cantidad (77)": cantidad_comercial,
            "No. Bultos": no_bultos,
            "Valor FOB (USD)": valor_fob,
            "Sumatoria Fletes/Seguros/Otros (USD)": sumatoria_fletes,
            "Acta de Inspección No.": acta_inspeccion,
            "Levante No.": levante_no,
            "Fecha del Levante": fecha_levante,
            "Archivo": nombre_archivo,
            "Campos_no_encontrados": ", ".join(faltantes) if faltantes else "",
        }

    st.markdown("### 📑 Módulo de Declaraciones de Importación (DIM)")
    st.markdown("**Procesador Masivo — Formulario 500 DIAN**")
    st.divider()

    if "df_resultado_dim" not in st.session_state:
        st.session_state.df_resultado_dim = pd.DataFrame(columns=COLUMNAS_DIM)
    if "uploader_key_dim" not in st.session_state:
        st.session_state.uploader_key_dim = 0

    with st.container():
        st.subheader("1. Carga de Documentación DIM")
        uploaded_files_dim = st.file_uploader(
            "Arrastra archivos PDF o paquetes .ZIP con DIMs",
            type=["pdf", "zip"],
            accept_multiple_files=True,
            key=f"dim_uploader_{st.session_state.uploader_key_dim}",
        )

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            procesar_dim_btn = st.button("🚀 Procesar DIMs", type="primary", use_container_width=True)
        with col_b2:
            limpiar_dim_btn = st.button("🧹 Limpiar Panel DIM", use_container_width=True)

    if limpiar_dim_btn:
        st.session_state.df_resultado_dim = pd.DataFrame(columns=COLUMNAS_DIM)
        st.session_state.uploader_key_dim += 1
        st.rerun()

    if procesar_dim_btn:
        if not uploaded_files_dim:
            st.warning("Por favor carga al menos un archivo PDF o ZIP.")
        else:
            progreso = st.progress(0.0, text="Iniciando motor de extracción DIM...")
            tareas = []
            for uf in uploaded_files_dim:
                contenido = uf.read()
                if uf.name.lower().endswith(".zip"):
                    with zipfile.ZipFile(io.BytesIO(contenido)) as zf:
                        for info in zf.infolist():
                            if info.filename.lower().endswith(".pdf") and not info.is_dir():
                                tareas.append((info.filename.split("/")[-1], zf.read(info.filename)))
                else:
                    tareas.append((uf.name, contenido))

            filas = []
            total = max(len(tareas), 1)
            for i, (nombre_pdf, data) in enumerate(tareas, start=1):
                try:
                    with fitz.open(stream=data, filetype="pdf") as doc:
                        texto_completo = "\n".join(page.get_text() for page in doc)
                    dim_chunks = dividir_dims(texto_completo)
                    if not dim_chunks:
                        dim_chunks = [texto_completo]
                    for chunk in dim_chunks:
                        filas.append(extraer_campos_dim(chunk, texto_completo, nombre_pdf))
                except Exception as exc:
                    fila = {c: "" for c in COLUMNAS_DIM}
                    fila["Archivo"] = nombre_pdf
                    fila["Campos_no_encontrados"] = f"ERROR: {exc}"
                    filas.append(fila)
                progreso.progress(i / total, text=f"Procesando: {nombre_pdf}")

            progreso.empty()
            df_res = pd.DataFrame(filas, columns=COLUMNAS_DIM)
            if "Número de formulario" in df_res.columns:
                df_res["Número de formulario"] = df_res["Número de formulario"].astype(str).str.strip()
                df_res = df_res[df_res["Número de formulario"].notna() & (df_res["Número de formulario"] != "") & (df_res["Número de formulario"].str.lower() != "nan")]
                df_res = df_res.drop_duplicates(subset=["Número de formulario"], keep="last").reset_index(drop=True)
            if "Levante No." in df_res.columns:
                df_res["Levante No."] = df_res["Levante No."].astype(str).str.strip()

            st.session_state.df_resultado_dim = df_res
            st.success(f"✅ Extracción exitosa. {len(df_res)} declaración(es) consolidada(s).")

    df_dim = st.session_state.df_resultado_dim
    if not df_dim.empty:
        st.markdown("---")
        st.subheader("2. Analítica y Consolidado DIM")
        cols_m = st.columns(4)
        cols_m[0].metric("Total Valor FOB (USD)", f"${df_dim['Valor FOB (USD)'].sum():,.2f}")
        cols_m[1].metric("Total Fletes/Seguros", f"${df_dim['Sumatoria Fletes/Seguros/Otros (USD)'].sum():,.2f}")
        cols_m[2].metric("Peso Bruto (Kgs)", f"{df_dim['Peso Bruto (Kgs)'].sum():,.2f}")
        cols_m[3].metric("Peso Neto (Kgs)", f"{df_dim['Peso Neto (Kgs)'].sum():,.2f}")

        st.markdown("<br>", unsafe_allow_html=True)
        busq_dim = st.text_input("🔍 Buscar en DIMs (Formulario, NIT, Importador...)", "")
        df_vista_dim = df_dim.copy()
        if busq_dim:
            mask = df_vista_dim.apply(lambda f: f.astype(str).str.contains(busq_dim, case=False, na=False).any(), axis=1)
            df_vista_dim = df_vista_dim[mask]

        st.dataframe(df_vista_dim, use_container_width=True, height=400)

        st.markdown("---")
        st.subheader("3. Exportación de DIMs")
        df_exp_dim = df_dim.drop(columns=["Campos_no_encontrados"])
        
        # Excel generator para DIM
        def generar_excel_dim(dframe):
            buf = io.BytesIO()
            df_ex = dframe.copy()
            if not df_ex.empty:
                t_row = {c: "" for c in df_ex.columns}
                t_row["Número de formulario"] = "TOTALES CONSOLIDADOS"
                t_row["Valor FOB (USD)"] = df_ex["Valor FOB (USD)"].sum()
                t_row["Sumatoria Fletes/Seguros/Otros (USD)"] = df_ex["Sumatoria Fletes/Seguros/Otros (USD)"].sum()
                t_row["Peso Bruto (Kgs)"] = df_ex["Peso Bruto (Kgs)"].sum()
                t_row["Peso Neto (Kgs)"] = df_ex["Peso Neto (Kgs)"].sum()
                t_row["No. Bultos"] = df_ex["No. Bultos"].sum()
                if "Cantidad (77)" in df_ex.columns:
                    t_row["Cantidad (77)"] = df_ex["Cantidad (77)"].sum()
                df_ex = pd.concat([df_ex, pd.DataFrame([t_row])], ignore_index=True)

            with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                df_ex.to_excel(wr, index=False, sheet_name="DIM")
                ws = wr.sheets["DIM"]
                header_fill = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                thin_b = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
                for col_idx in range(1, df_ex.shape[1] + 1):
                    c = ws.cell(row=1, column=col_idx)
                    c.fill = header_fill
                    c.font = header_font
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.border = thin_b
                for r in range(2, df_ex.shape[0] + 2):
                    for col_idx in range(1, df_ex.shape[1] + 1):
                        c = ws.cell(row=r, column=col_idx)
                        c.border = thin_b
                        c.alignment = Alignment(vertical="top", wrap_text=True)
                        if r == df_ex.shape[0] + 1:
                            c.fill = PatternFill(start_color="D9E1D9", end_color="D9E1D9", fill_type="solid")
                            c.font = Font(bold=True)
                for col_idx, col_name in enumerate(df_ex.columns, start=1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 22
            return buf.getvalue()

        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button("⬇️ Descargar Excel DIM", data=generar_excel_dim(df_exp_dim), file_name=f"dim_zona_franca_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_d2:
            st.download_button("⬇️ Descargar CSV DIM", data=df_exp_dim.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"), file_name=f"dim_zona_franca_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv", use_container_width=True)


# ==========================================================================
# MÓDULO 2: ACTAS DE TRÁNSITO / PICIZ
# ==========================================================================

elif modulo_seleccionado == "📋 Actas de Tránsito / PICIZ":

    def extraer_datos_acta(pdf_file, nombre_archivo):
        with pdfplumber.open(pdf_file) as pdf:
            texto = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        usuario = re.search(r"consignados?\s+al\s*\n?\s*([A-Z0-9\.\-\s]+?)(?=\s+y\s+amparado|\s+DECLARACION|\n\s*DECLARACION|$)", texto, re.IGNORECASE)
        val_usuario = " ".join(usuario.group(1).strip().split()) if usuario else "N/A"

        doc_form = re.search(r"DOCUMENTO\s+FORMULARIO\s+MERCANC[ÍI]A[\s\S]*?\n\s*([A-Z0-9\.\-_]+)\s+(\d+)", texto, re.IGNORECASE)
        if not doc_form:
            doc_form = re.search(r"\b([A-Z0-9\.\-_]{5,})\s+(\d{7,10})\b", texto)

        transito = re.search(r"DECLARACION\s+DE\s+TRANSITO\s+ADUANERO\s*\n?\s*(?:Número|N[úu]mero)?\s*[:\.]?\s*(\d+)", texto, re.IGNORECASE)
        fecha_ingreso = re.search(r"ACTA\s+DE\s+DESPRECINTAJE[\s\S]*?(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
        fecha_auto = re.search(r"Fecha\s+de\s+la\s+autorizaci[oó]n\s+de\s+la\s+operaci[oó]n[^\d]*(\d{4}[-/]\d{2}[-/]\d{2})", texto, re.IGNORECASE)
        fecha_limite = re.search(r"Fecha\s+l[ií]mite\s+para\s+finalizar\s+el\s+r[eé]gimen[^\d]*(\d{4}[-/]\d{2}[-/]\d{2})", texto, re.IGNORECASE)
        acta_n = re.search(r"Acta\s+N\.?\s*(\d+)", texto, re.IGNORECASE)
        
        fecha_acta_match = re.search(r"FECHA\s+GENERACI[OÓ]N\s+DEL\s+ACTA:\s*(\d{2}/\d{2}/\d{4})", texto, re.IGNORECASE)
        fecha_acta = fecha_acta_match.group(1).strip() if fecha_acta_match else "N/A"

        peso_match = re.search(r"TOTALES\s*:\s*[\d\.,]+\s+([\d\.,]+)", texto, re.IGNORECASE) or re.search(r"TOTALES\s*:\s*([\d\.,]+)", texto, re.IGNORECASE)
        peso_bascula = peso_match.group(1).strip() if peso_match else "N/A"

        obs_match = re.search(r"Observaciones[\s\S]*?\n([\s\S]*?)(?=\n\s*(?:DOCUMENTO|TOTALES|USUARIO\s+OPERADOR|FECHA\s+GENERACI|\Z))", texto, re.IGNORECASE)
        if obs_match:
            lineas = obs_match.group(1).split("\n")
            lineas_limpias = [l.strip() for l in lineas if l.strip() and not re.match(r"^(Descripción\s*N/A|Bultos|Estado|Términos|Otra)\b", l.strip(), re.IGNORECASE)]
            observaciones = " ".join(lineas_limpias) if lineas_limpias else "N/A"
        else:
            observaciones = "N/A"

        return {
            "Usuario": val_usuario,
            "Documento de transporte": doc_form.group(1).strip() if doc_form else "N/A",
            "Transito N°": transito.group(1).strip() if transito else "N/A",
            "FMM N°": doc_form.group(2).strip() if doc_form else "N/A",
            "FECHA INGRESO ÚLTIMO VEHÍCULO": fecha_ingreso.group(1).strip() if fecha_ingreso else "N/A",
            "Fecha de autorización": fecha_auto.group(1).strip() if fecha_auto else "N/A",
            "Tránsito Fecha Maxima Finalización": fecha_limite.group(1).strip() if fecha_limite else "N/A",
            "Acta de Inventario e Inconsistencias PICIZ": acta_n.group(1).strip() if acta_n else "N/A",
            "Fecha acta de inventario e inconsistencias": fecha_acta,
            "No. Planilla de Recepción (FECHA)": fecha_acta,
            "Peso Báscula ZFC": peso_bascula,
            "OBSERVACIONES/ INCONSISTENCIAS": observaciones,
            "Archivo": nombre_archivo,
        }

    st.markdown("### 📋 Módulo de Actas de Tránsito / PICIZ")
    st.markdown("**Extractor Automático de Actas de Inventario e Inconsistencias**")
    st.divider()

    if "df_resultado_actas" not in st.session_state:
        st.session_state.df_resultado_actas = pd.DataFrame()
    if "uploader_key_actas" not in st.session_state:
        st.session_state.uploader_key_actas = 0

    with st.container():
        st.subheader("1. Carga de Documentación de Actas")
        uploaded_files_actas = st.file_uploader(
            "Carga archivos PDF de actas de inventario",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"actas_uploader_{st.session_state.uploader_key_actas}",
        )

        col_a1, col_a2 = st.columns(2)
        with col_a1:
            procesar_actas_btn = st.button("🚀 Procesar Actas", type="primary", use_container_width=True)
        with col_a2:
            limpiar_actas_btn = st.button("🧹 Limpiar Panel Actas", use_container_width=True)

    if limpiar_actas_btn:
        st.session_state.df_resultado_actas = pd.DataFrame()
        st.session_state.uploader_key_actas += 1
        st.rerun()

    if procesar_actas_btn:
        if not uploaded_files_actas:
            st.warning("Por favor carga al menos un archivo PDF de actas.")
        else:
            with st.spinner("Extrayendo campos clave de actas de tránsito..."):
                datos = [extraer_datos_acta(f, f.name) for f in uploaded_files_actas]
                st.session_state.df_resultado_actas = pd.DataFrame(datos)
            st.success(f"✅ Extracción completada para {len(uploaded_files_actas)} acta(s).")

    df_actas = st.session_state.df_resultado_actas
    if not df_actas.empty:
        st.markdown("---")
        st.subheader("2. Resultados Consolidados de Actas")
        busq_actas = st.text_input("🔍 Buscar en Actas (Usuario, Tránsito, Acta PICIZ...)", "")
        df_vista_actas = df_actas.copy()
        if busq_actas:
            mask = df_vista_actas.apply(lambda f: f.astype(str).str.contains(busq_actas, case=False, na=False).any(), axis=1)
            df_vista_actas = df_vista_actas[mask]

        st.dataframe(df_vista_actas, use_container_width=True, height=430)

        st.markdown("---")
        st.subheader("3. Exportación de Actas")
        
        def generar_excel_actas(dframe):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as wr:
                dframe.to_excel(wr, index=False, sheet_name="Actas")
                ws = wr.sheets["Actas"]
                header_fill = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                thin_b = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))
                for col_idx in range(1, dframe.shape[1] + 1):
                    c = ws.cell(row=1, column=col_idx)
                    c.fill = header_fill
                    c.font = header_font
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.border = thin_b
                for r in range(2, dframe.shape[0] + 2):
                    for col_idx in range(1, dframe.shape[1] + 1):
                        c = ws.cell(row=r, column=col_idx)
                        c.border = thin_b
                        c.alignment = Alignment(vertical="top", wrap_text=True)
                for col_idx in range(1, dframe.shape[1] + 1):
                    ws.column_dimensions[get_column_letter(col_idx)].width = 25
            return buf.getvalue()

        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.download_button("⬇️ Descargar Excel Actas", data=generar_excel_actas(df_actas), file_name=f"actas_zona_franca_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_e2:
            st.download_button("⬇️ Descargar CSV Actas", data=df_actas.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"), file_name=f"actas_zona_franca_{datetime.now().strftime('%Y%m%d_%H%M')}.csv", mime="text/csv", use_container_width=True)