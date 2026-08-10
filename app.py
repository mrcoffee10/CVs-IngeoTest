import os
import json
import re
import streamlit as st
from google import genai
from docxtpl import DocxTemplate
from pypdf import PdfReader
from json_repair import repair_json

# Configuración de página con colores e imagen corporativa
st.set_page_config(page_title="Generador de CVs — INGEOTEST S.A.C.", page_icon="⚡", layout="wide")

# Estilos CSS Corporativos
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { background-color: #1B365D; color: white; font-weight: bold; width: 100%; border-radius: 8px; height: 3em; }
    .stButton>button:hover { background-color: #0B1F3A; color: white; }
    .header-box { background: linear-gradient(135deg, #1B365D 0%, #0B1F3A 100%); color: white; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 25px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="header-box">
        <h2 style="color: white; margin:0;">⚡ GENERADOR CORPORATIVO DE CVs</h2>
        <p style="color: #B0C4DE; margin:5px 0 0 0;">INGEOTEST INGENIEROS S.A.C. — Estándar Maestro de Presentación Técnica</p>
    </div>
""", unsafe_allow_html=True)

# Obtener API Key desde Secrets de Streamlit
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY"))
Nombre_Plantilla_Word = "PLANTILLA_MAESTRA_INGEOTEST.docx"

client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

def extraer_texto_pdf(pdf_file):
    reader = PdfReader(pdf_file)
    texto = ""
    for page in reader.pages:
        text_page = page.extract_text()
        if text_page:
            texto += text_page + "\n"
    return texto

def procesar_cv_con_ia(texto_cv_original, cargo_objetivo):
    prompt = f"""
    Eres el Auditor Principal de Selección Técnica y Control de Calidad Documental de INGEOTEST INGENIEROS S.A.C.
    Tu objetivo es reestructurar la información del CV del candidato al ESTÁNDAR CORPORATIVO INGEOTEST, reorientando su perfil profesional al cargo objetivo de: "{cargo_objetivo}".

    =============================================================================
    PROTOCOLOS ESTRICTOS DE VERACIDAD Y AUDITORÍA DE CALIDAD:
    =============================================================================
    1. INDICADORES EN FACTORES DE 5 Y FORMATO CIP:
       - 'ANOS_EXP': Calcula la experiencia real en múltiplos de 5 con '+' (ejemplo: 8-11 años -> "+10"; 12-14 años -> "+10").
       - 'NUM_PROYECTOS': Cuenta los proyectos reales en múltiplos de 5 con '+' (ejemplo: 6 a 9 proyectos -> "+5"; 10 a 14 -> "+10"; 30+ -> "+30").
       - 'CIP_NUMERO': OBLIGATORIAMENTE con el prefijo 'CIP' seguido del número (ejemplo: "CIP 166577" o "CIP 116099").

    2. SECCIÓN HERRAMIENTAS Y SOFTWARE (OBLIGATORIA DE 2 A 3 FILAS):
       Extrae las herramientas informáticas Y de gestión operacional reales del candidato:
       - SOFT_1_CAT y SOFT_1_PROG: Herramientas de Gestión / SSOMA / Operativas (ej: "Gestión y SSOMA" -> "Sistema Integrado de Gestión (SIG), Matriz IPERC, PETAR, ATS").
       - SOFT_2_CAT y SOFT_2_PROG: Planificación / Software Técnico / Productividad (ej: "Planificación y Productividad" -> "AutoCAD, Control de Avances, MS Office").
       - SOFT_3_CAT y SOFT_3_PROG: Si existe software especializado adicional (ej. ArcGIS, Slide, SAP), colócalo aquí. De lo contrario deja SOFT_3 en blanco.

    3. CERTIFICADOS, REGISTROS Y CAPACITACIONES (ORDEN STRICTO DE VIÑETAS):
       La lista 'capacitaciones' debe construirse con las siguientes viñetas exactas extraídas del CV:
       - Viñeta 1: "Colegiatura CIP N.º [Número] – Habilitado."
       - Viñeta 2: "Licencia de conducir: [Categoría y Número]." (Si figura en el CV).
       - Viñetas siguientes: Los cursos, diplomados y certificaciones reales mencionados en el expediente.

    4. SIMETRÍA EN FECHAS Y RAZÓN SOCIAL DE EMPRESAS:
       - Periodos idénticos en resumen y detalle (ej. "2024 – 2025").
       - 'EMPRESA_1' y 'EMPRESA_2': ÚNICAMENTE la razón social oficial contratante (ej. "BRAYAM S.A.C."). Prohibido fusionar con clientes.

    5. RESUMEN EJECUTIVO (80 a 110 palabras):
       Redacción técnica en 3ra persona formal: "[Profesión Base] colegiado y habilitado con más de [X] años de trayectoria profesional en... Como [Cargo Objetivo] en INGEOTEST INGENIEROS S.A.C., lidera... Destaca por...".

    =============================================================================
    CV ORIGINAL DEL CANDIDATO:
    =============================================================================
    {texto_cv_original}

    =============================================================================
    ESTRUCTURA JSON DE SALIDA (DEVOLVER ÚNICAMENTE UN JSON VÁLIDO):
    =============================================================================
    {{
        "NOMBRE_COMPLETO": "",
        "TITULO_PROFESIONAL": "",
        "CARGO_OBJETIVO": "{cargo_objetivo}",
        "CIP_HABILITACION": "",
        "TELEFONO": "",
        "CORREO": "",
        "UBICACION": "",
        "RESUMEN_EJECUTIVO": "",
        "PROPUESTA_1": "", "PROPUESTA_2": "", "PROPUESTA_3": "", "PROPUESTA_4": "",
        "propuesta_valor": ["", "", "", ""],
        "ANOS_EXP": "", "NUM_PROYECTOS": "", "CIP_NUMERO": "", "NIVEL_SENIORITY": "", "AREA_ENFOQUE": "",
        "COMP_1": "", "COMP_2": "", "COMP_3": "", "COMP_4": "", "COMP_5": "",
        "COMP_6": "", "COMP_7": "", "COMP_8": "", "COMP_9": "", "COMP_10": "",
        "EMPRESA_RESUMEN_1": "", "CARGO_RESUMEN_1": "", "PERIODO_RESUMEN_1": "",
        "EMPRESA_RESUMEN_2": "", "CARGO_RESUMEN_2": "", "PERIODO_RESUMEN_2": "",
        "EMPRESA_1": "", "CARGO_1": "", "PERIODO_1": "", "ALCANCE_1": "",
        "RESP_1_1": "", "RESP_1_2": "", "RESP_1_3": "", "RESP_1_4": "",
        "EMPRESA_2": "", "CARGO_2": "", "PERIODO_2": "", "ALCANCE_2": "",
        "RESP_2_1": "", "RESP_2_2": "", "RESP_2_3": "", "RESP_2_4": "",
        "EXP_TRANSVERSAL_1": "", "EXP_TRANSVERSAL_2": "", "EXP_TRANSVERSAL_3": "",
        "EXP_TRANSVERSAL_4": "", "EXP_TRANSVERSAL_5": "", "EXP_TRANSVERSAL_6": "", "EXP_TRANSVERSAL_7": "",
        "SECTOR_1": "", "SECTOR_2": "", "SECTOR_3": "", "SECTOR_4": "", "SECTOR_5": "", "SECTOR_6": "",
        "PROY_1_NOMBRE": "", "PROY_1_EMPRESA": "", "PROY_1_UBICACION": "", "PROY_1_CARGO": "", "PROY_1_ANO": "",
        "PROY_2_NOMBRE": "", "PROY_2_EMPRESA": "", "PROY_2_UBICACION": "", "PROY_2_CARGO": "", "PROY_2_ANO": "",
        "PROY_3_NOMBRE": "", "PROY_3_EMPRESA": "", "PROY_3_UBICACION": "", "PROY_3_CARGO": "", "PROY_3_ANO": "",
        "PROY_4_NOMBRE": "", "PROY_4_EMPRESA": "", "PROY_4_UBICACION": "", "PROY_4_CARGO": "", "PROY_4_ANO": "",
        "TITULO_ACADEMICO": "", "UNIVERSIDAD": "", "FECHA_TITULO": "",
        "SOFT_1_CAT": "", "SOFT_1_PROG": "",
        "SOFT_2_CAT": "", "SOFT_2_PROG": "",
        "SOFT_3_CAT": "", "SOFT_3_PROG": "",
        "IDIOMAS": "", "MEMBRESIAS": "",
        "capacitaciones": []
    }}
    """

    modelos = ['gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-2.0-flash']
    for mod in modelos:
        try:
            response = client.models.generate_content(
                model=mod,
                contents=prompt,
                config={"response_mime_type": "application/json", "temperature": 0.1}
            )
            texto_puro = response.text.strip()
            if texto_puro.startswith("```"):
                texto_puro = re.sub(r'^```json\s*|```$', '', texto_puro, flags=re.MULTILINE).strip()
            
            json_reparado = repair_json(texto_puro)
            datos = json.loads(json_reparado)
            
            if "capacitaciones" in datos and isinstance(datos["capacitaciones"], list):
                for idx, cap in enumerate(datos["capacitaciones"][:6]):
                    datos[f"CAPACITACION_{idx+1}"] = cap
            elif "CAPACITACION_1" in datos:
                datos["capacitaciones"] = [
                    datos.get(f"CAPACITACION_{i}", "") for i in range(1, 7) if datos.get(f"CAPACITACION_{i}")
                ]
                
            return datos
        except Exception as e:
            if "503" in str(e) or "UNAVAILABLE" in str(e) or "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
                continue
            else:
                continue

    raise Exception("Los modelos de IA están ocupados temporalmente. Por favor, intenta de nuevo en 30 segundos.")

# Interfaz en la columna principal
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Cargar Documento del Candidato")
    uploaded_file = st.file_uploader("Sube el CV en formato PDF", type=["pdf"])

    st.subheader("2. Definir Cargo Objetivo")
    cargo_drop = st.selectbox(
        "Selecciona el cargo:",
        [
            "Ingeniero de Seguridad", 
            "Ingeniero de Proyectos", 
            "Ingeniero Geólogo", 
            "Especialista en Geotecnia e Hidrogeología", 
            "Jefe de Laboratorio Geotécnico", 
            "Escribir cargo personalizado..."
        ]
    )
    
    cargo_custom = ""
    if cargo_drop == "Escribir cargo personalizado...":
        cargo_custom = st.text_input("Escribe el cargo personalizado:")

    btn_generar = st.button("🚀 Generar CV Estandarizado")

with col2:
    st.subheader("3. Resultado y Descarga")
    if btn_generar:
        if uploaded_file is None:
            st.error("⚠️ Por favor sube un archivo PDF antes de generar.")
        else:
            with st.spinner("Procesando y acondicionando expediente al Estándar INGEOTEST..."):
                try:
                    cargo_final = cargo_custom.strip() if cargo_drop == "Escribir cargo personalizado..." and cargo_custom.strip() else cargo_drop
                    
                    texto_cv = extraer_texto_pdf(uploaded_file)
                    datos_json = procesar_cv_con_ia(texto_cv, cargo_final)
                    
                    if "NOMBRE_COMPLETO" in datos_json:
                        datos_json["NOMBRE_COMPLETO"] = str(datos_json["NOMBRE_COMPLETO"]).upper()
                    if "TITULO_PROFESIONAL" in datos_json:
                        datos_json["TITULO_PROFESIONAL"] = str(datos_json["TITULO_PROFESIONAL"]).upper()
                        
                    doc = DocxTemplate(Nombre_Plantilla_Word)
                    doc.render(datos_json)
                    
                    for table in doc.docx.tables:
                        rows_to_remove = []
                        for row in table.rows:
                            texto_fila = "".join([cell.text.strip() for cell in row.cells])
                            if texto_fila == "":
                                rows_to_remove.append(row)
                        for row in rows_to_remove:
                            tr = row._tr
                            tr.getparent().remove(tr)
                            
                    nombre_salida = f"CV_INGEOTEST_{datos_json.get('NOMBRE_COMPLETO', 'GENERADO').replace(' ', '_')}.docx"
                    doc.save(nombre_salida)
                    
                    with open(nombre_salida, "rb") as file:
                        btn_download = st.download_button(
                            label="📥 Descargar Documento Word (.docx)",
                            data=file,
                            file_name=nombre_salida,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    st.success("✅ ¡Documento estandarizado con éxito!")
                except Exception as e:
                    st.error(f"❌ Error al procesar: {str(e)}")
