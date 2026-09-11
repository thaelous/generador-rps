import io
import os
import re
import json
import time
import base64
import requests
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
import fitz  # PyMuPDF
import streamlit as st

st.set_page_config(page_title="Generador RPS AAM", page_icon="📄", layout="centered")
st.title("Generador Automático de RPS con Evidencias")
st.write("Sube la Factura, la Orden de Compra (OC) y tus fotos de evidencia para generar el RPS completo.")

api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

if not api_key:
    api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")

# Insumos
uploaded_factura = st.file_uploader("1. Factura en PDF (CFDI)", type=["pdf"])
uploaded_oc = st.file_uploader("2. Orden de Compra en PDF (OC)", type=["pdf"])
uploaded_fotos = st.file_uploader("3. Fotos de Evidencia ('DESPUÉS' - hasta 3 imágenes)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

linea_po_manual = st.text_input(
    "Número de línea en PO (opcional):", 
    placeholder="Ej. 3-1 o déjalo vacío si no aplica"
)

SYSTEM_PROMPT = """
Eres un asistente contable y fiscal. Extrae la información de la factura en PDF para completar el formato RPS corporativo.
Devuelve EXCLUSIVAMENTE un objeto JSON válido con la siguiente estructura:
{
  "orden_compra": string,
  "nombre_proveedor": string,
  "folio_factura": string,
  "tipo_servicio": string,
  "solicitante": string,
  "moneda": string,
  "subtotal": number,
  "lineas": [
    {
      "cantidad": number,
      "unidad": string,
      "descripcion": string,
      "monto": number
    }
  ]
}
Reglas estrictas:
- En 'orden_compra' coloca solo números o el código limpio (ej. si dice 'OC 786794', extrae '786794').
- En 'solicitante', extrae el nombre de la persona que solicita el servicio (revisa la adenda o el cuerpo de la descripción).
- En 'lineas.descripcion', coloca ÚNICAMENTE el concepto o servicio brindado. ELIMINA por completo cualquier mención al nombre del solicitante y al número de OC/PO (ej. no incluir 'Solicitante: Juan Perez', ni 'OC 786794').
- Devuelve únicamente el JSON sin comentarios ni bloques adicionales.
"""

def extraer_datos(pdf_bytes, raw_key):
    clean_key = raw_key.strip().strip("'").strip('"')
    pdf_b64 = base64.b64encode(pdf_bytes).decode('utf-8')
    
    payload = {
        "system_instruction": {"parts": [{"text": SYSTEM_PROMPT}]},
        "contents": [{
            "parts": [
                {"inline_data": {"mime_type": "application/pdf", "data": pdf_b64}},
                {"text": "Extrae los datos de esta factura para llenar el RPS."}
            ]
        }],
        "generationConfig": {
            "response_mime_type": "application/json",
            "temperature": 0.1
        }
    }
    
    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": clean_key
    }
    
    modelos = [
        "gemini-2.5-flash-lite",
        "gemini-2.5-flash",
        "gemini-3.6-flash"
    ]
    ultimo_error = None
    
    for mod in modelos:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent".strip()
        for intento in range(3):
            try:
                response = requests.post(url, headers=headers, json=payload, timeout=60)
                if response.status_code == 200:
                    data = response.json()
                    texto = data["candidates"][0]["content"]["parts"][0]["text"]
                    return json.loads(texto)
                elif response.status_code in (429, 503):
                    ultimo_error = f"Código {response.status_code}: Servidor ocupado ({mod}). Reintentando..."
                    time.sleep(2 * (intento + 1))
                    continue
                else:
                    ultimo_error = f"Código {response.status_code}: {response.text}"
                    break
            except Exception as e:
                ultimo_error = str(e)
                time.sleep(2)
                
    raise RuntimeError(ultimo_error)

def limpiar_descripcion(desc, solicitante, oc):
    texto = desc
    if solicitante and solicitante.strip():
        patron_sol = re.compile(re.escape(solicitante.strip()), re.IGNORECASE)
        texto = patron_sol.sub("", texto)
    texto = re.sub(r'(?i)(solicita(?:nte)?|atenci[oó]n|contacto)\s*[:\-]?\s*', '', texto)
    if oc and str(oc).strip():
        patron_oc = re.compile(rf'(?i)(?:oc|orden\s*(?:de)?\s*compra|po)\s*[:#\-]?\s*{re.escape(str(oc).strip())}')
        texto = patron_oc.sub("", texto)
        texto = re.sub(rf'\b{re.escape(str(oc).strip())}\b', '', texto)
    texto = re.sub(r'(?i)\b(?:oc|po)\s*[:#\-]?\b', '', texto)
    texto = re.sub(r'\s+', ' ', texto)
    texto = re.sub(r'^[\s,\.\-_/]+|[\s,\.\-_/]+$', '', texto)
    return texto.strip()

def extraer_pagina_partidas_oc(oc_bytes):
    """Busca la página de la OC con partidas y la convierte en imagen"""
    doc = fitz.open(stream=oc_bytes, filetype="pdf")
    target_page_idx = 0
    
    # Buscar la página que contenga las partidas/líneas
    for i, page in enumerate(doc):
        texto = page.get_text()
        if "Partida/Descripción" in texto or "Precio ampliado" in texto or "Precio unitario" in texto:
            target_page_idx = i
            break
            
    page = doc[target_page_idx]
    pix = page.get_pixmap(dpi=150)
    img = PILImage.open(io.BytesIO(pix.tobytes("png")))
    
    # Redimensionar para que calce bien en el marco "ANTES"
    img.thumbnail((380, 520), PILImage.Resampling.LANCZOS)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    return img_byte_arr

def llenar_plantilla(datos, linea_po_usuario="", oc_bytes=None, fotos_bytes=[], plantilla_path="plantilla_RPS.xlsx"):
    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb["3151"] if "3151" in wb.sheetnames else wb.active
    
    folio = str(datos.get("folio_factura", "RPS"))
    oc_val = datos.get("orden_compra", "")
    solicitante_val = datos.get("solicitante", "")
    ws.title = folio
    
    # 1. Orden de Compra (Fila 7, Columna E)
    if oc_val:
        ws.cell(row=7, column=5, value=int(oc_val) if str(oc_val).isdigit() else str(oc_val))
    
    # 2. Proveedor (Fila 9, Columna E)
    ws.cell(row=9, column=5, value=datos.get("nombre_proveedor", ""))
    
    # 3. Folio Factura y Tipo de Servicio
    ws.cell(row=11, column=5, value=int(folio) if folio.isdigit() else str(folio))
    ws.cell(row=11, column=10, value=datos.get("tipo_servicio", "Entrenamiento"))
    
    # 4. Detalle de Concepto (Fila 15)
    lineas = datos.get("lineas", [])
    if lineas:
        l = lineas[0]
        ws.cell(row=15, column=1, value=linea_po_usuario.strip() if linea_po_usuario else "")
        ws.cell(row=15, column=2, value=l.get("cantidad", 1))
        
        unidad = str(l.get("unidad", "LOT")).strip()
        if "E48" in unidad.upper() or not unidad:
            unidad = "LOT"
        else:
            unidad = unidad.replace("E48", "").replace("-", "").strip()
        ws.cell(row=15, column=4, value=unidad)
        
        ws.cell(row=15, column=5, value=l.get("monto", datos.get("subtotal", 0)))
        ws.cell(row=15, column=6, value=datos.get("moneda", "MXN"))
        
        desc_original = l.get("descripcion", "")
        desc_limpia = limpiar_descripcion(desc_original, solicitante_val, oc_val)
        ws.cell(row=15, column=7, value=desc_limpia)
        ws.cell(row=15, column=16, value="YES")
        
    # 5. Monto Total a Recibir (Fila 25)
    subtotal = datos.get("subtotal", 0)
    ws.cell(row=25, column=5, value=subtotal)
    ws.cell(row=25, column=6, value=subtotal)
    
    # 6. Solicitante (Fila 27)
    if solicitante_val:
        ws.cell(row=27, column=3, value=solicitante_val)
        
    # Limpiar imágenes existentes en la plantilla para no duplicar
    ws._images.clear()

    # 7. Insertar Imagen de Orden de Compra ("ANTES") en S8
    if oc_bytes:
        try:
            img_oc_bytes = extraer_pagina_partidas_oc(oc_bytes)
            img_oc = OpenpyxlImage(img_oc_bytes)
            ws.add_image(img_oc, "S8")
        except Exception as e:
            st.warning(f"No se pudo insertar la imagen de la OC: {e}")
            
    # 8. Insertar Fotos de Evidencia ("DESPUÉS") en AC8, AC15, AC22
    celdas_despues = ["AC8", "AC15", "AC22"]
    for i, f_bytes in enumerate(fotos_bytes[:3]):
        try:
            p_img = PILImage.open(io.BytesIO(f_bytes))
            p_img.thumbnail((320, 150), PILImage.Resampling.LANCZOS)
            b_arr = io.BytesIO()
            p_img.save(b_arr, format='PNG')
            b_arr.seek(0)
            
            excel_img = OpenpyxlImage(b_arr)
            ws.add_image(excel_img, celdas_despues[i])
        except Exception as e:
            st.warning(f"No se pudo insertar la foto {i+1}: {e}")
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output, folio, (oc_val if oc_val else "RPS")

if uploaded_factura and api_key:
    if st.button("Procesar Factura y Generar RPS"):
        with st.spinner("Procesando documentos e integrando imágenes..."):
            try:
                datos = extraer_datos(uploaded_factura.getvalue(), api_key)
                
                oc_bytes = uploaded_oc.getvalue() if uploaded_oc else None
                fotos_bytes = [f.getvalue() for f in uploaded_fotos] if uploaded_fotos else []
                
                excel_salida, folio, oc = llenar_plantilla(
                    datos, 
                    linea_po_usuario=linea_po_manual,
                    oc_bytes=oc_bytes,
                    fotos_bytes=fotos_bytes
                )
                
                st.success("¡RPS con evidencias generado exitosamente!")
                
                with st.expander("Ver datos extraídos"):
                    st.json(datos)
                
                nombre_descarga = f"RPS {oc} {folio}.xlsx"
                st.download_button(
                    label=f"📥 Descargar {nombre_descarga}",
                    data=excel_salida,
                    file_name=nombre_descarga,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error al procesar: {e}")
