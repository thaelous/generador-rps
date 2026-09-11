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

# ReportLab para la generación del PDF oficial
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.pdfgen import canvas

st.set_page_config(page_title="Generador RPS AAM", page_icon="📄", layout="centered")
st.title("Generador Automático de RPS con Evidencias")
st.write("Sube la Factura, la Orden de Compra y tus evidencias para obtener tu Excel y PDF listos.")

api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

if not api_key:
    api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")

# Insumos
uploaded_factura = st.file_uploader("1. Factura en PDF (CFDI)", type=["pdf"])
uploaded_oc = st.file_uploader("2. Orden de Compra en PDF (OC)", type=["pdf"])
uploaded_fotos = st.file_uploader("3. Fotos de Evidencia ('DESPUÉS' - hasta 3 imágenes)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

linea_po_manual = st.text_input(
    "Número de línea en PO (Opcional - La IA lo detectará automáticamente si lo dejas vacío):", 
    placeholder="Ej. 3-1 o déjalo vacío"
)

SYSTEM_PROMPT = """
Eres un auditor contable corporativo. Se te proporcionan dos documentos:
1. Factura (CFDI en PDF).
2. Orden de Compra (OC en PDF).

Tu tarea:
1. Extraer los datos fiscales del CFDI para el formato RPS.
2. Comparar el concepto y monto facturado contra la tabla de partidas de la Orden de Compra.
3. Identificar el número de línea que corresponde al servicio facturado en la OC y formatearlo estrictamente como 'X-1' (por ejemplo: '1-1', '2-1', '3-1', etc.).

Devuelve EXCLUSIVAMENTE un objeto JSON válido con la siguiente estructura:
{
  "orden_compra": string,
  "nombre_proveedor": string,
  "folio_factura": string,
  "tipo_servicio": string,
  "solicitante": string,
  "moneda": string,
  "subtotal": number,
  "linea_po_detectada": string,
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
- En 'linea_po_detectada', coloca el código de línea correspondiente en la OC en formato 'X-1'. Si no lo encuentras con certeza, usa '1-1'.
- En 'orden_compra' coloca solo números o el código limpio.
- En 'solicitante', extrae el nombre de la persona que solicita el servicio.
- En 'lineas.descripcion', coloca ÚNICAMENTE el concepto o servicio. ELIMINA menciones al solicitante y a la OC/PO.
- Devuelve únicamente el JSON sin comentarios.
"""

def extraer_datos_con_oc(factura_bytes, oc_bytes, raw_key):
    clean_key = raw_key.strip().strip("'").strip('"')
    fac_b64 = base64.b64encode(factura_bytes).decode('utf-8')
    
    parts = [
        {"inline_data": {"mime_type": "application/pdf", "data": fac_b64}},
        {"text": "Documento 1: Factura emitida (CFDI)."}
    ]
    
    if oc_bytes:
        oc_b64 = base64.b64encode(oc_bytes).decode('utf-8')
        parts.extend([
            {"inline_data": {"mime_type": "application/pdf", "data": oc_b64}},
            {"text": "Documento 2: Orden de Compra oficial (OC). Cruza el concepto facturado con la OC para identificar la línea de PO en formato 'X-1'."}
        ])
    else:
        parts.append({"text": "No se adjuntó OC. Extrae únicamente los datos de la factura."})

    payload = {
        "system_instruction": {"parts": [{"text": SYSTEM_PROMPT}]},
        "contents": [{"parts": parts}],
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
                response = requests.post(url, headers=headers, json=payload, timeout=75)
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
    doc = fitz.open(stream=oc_bytes, filetype="pdf")
    target_page_idx = 0
    for i, page in enumerate(doc):
        texto = page.get_text()
        if "Partida/Descripción" in texto or "Precio ampliado" in texto or "Precio unitario" in texto:
            target_page_idx = i
            break
    page = doc[target_page_idx]
    pix = page.get_pixmap(dpi=150)
    img = PILImage.open(io.BytesIO(pix.tobytes("png")))
    img.thumbnail((380, 520), PILImage.Resampling.LANCZOS)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')
    img_byte_arr.seek(0)
    return img_byte_arr

def llenar_plantilla_excel(datos, linea_po_final, oc_bytes=None, fotos_bytes=[], plantilla_path="plantilla_RPS.xlsx"):
    wb = openpyxl.load_workbook(plantilla_path)
    ws = wb["3151"] if "3151" in wb.sheetnames else wb.active
    
    folio = str(datos.get("folio_factura", "RPS"))
    oc_val = datos.get("orden_compra", "")
    solicitante_val = datos.get("solicitante", "")
    ws.title = folio
    
    if oc_val:
        ws.cell(row=7, column=5, value=int(oc_val) if str(oc_val).isdigit() else str(oc_val))
    
    ws.cell(row=9, column=5, value=datos.get("nombre_proveedor", ""))
    ws.cell(row=11, column=5, value=int(folio) if folio.isdigit() else str(folio))
    ws.cell(row=11, column=10, value=datos.get("tipo_servicio", "Entrenamiento"))
    
    lineas = datos.get("lineas", [])
    if lineas:
        l = lineas[0]
        ws.cell(row=15, column=1, value=linea_po_final)
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
        
    subtotal = datos.get("subtotal", 0)
    ws.cell(row=25, column=5, value=subtotal)
    ws.cell(row=25, column=6, value=subtotal)
    
    if solicitante_val:
        ws.cell(row=27, column=3, value=solicitante_val)
        
    ws._images.clear()

    if oc_bytes:
        try:
            img_oc_bytes = extraer_pagina_partidas_oc(oc_bytes)
            img_oc = OpenpyxlImage(img_oc_bytes)
            ws.add_image(img_oc, "S8")
        except Exception:
            pass
            
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
        except Exception:
            pass
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def generar_pdf_oficial(datos, linea_po_final, oc_bytes=None, fotos_bytes=[]):
    """Genera exactamente el PDF de 2 páginas del RPS con ReportLab"""
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    ancho, alto = letter
    
    # ------------------ PÁGINA 1: FORMATO RPS ------------------
    # Encabezado rojo / logo
    p.setFillColor(colors.HexColor("#C00000"))
    p.setFont("Helvetica-Bold", 16)
    p.drawString(40, alto - 50, "AAM")
    p.setFillColor(colors.black)
    p.setFont("Helvetica-Bold", 13)
    p.drawString(100, alto - 50, "REPLACEMENT PACKING SLIP")
    
    p.setFont("Helvetica-Bold", 7)
    p.setFillColor(colors.HexColor("#C00000"))
    p.drawString(40, alto - 65, "INSTRUCCIONES: TODOS LOS CAMPOS SOMBREADOS DEBEN SER LLENADOS PARA PODER PROCEDER CON EL RECIBO.")
    p.drawString(40, alto - 75, "ES NECESARIO ADJUNTAR LA FACTURA A ESTE DOCUMENTO")
    
    # Cajas de datos principales
    p.setFont("Helvetica-Bold", 8)
    p.setFillColor(colors.black)
    p.drawString(40, alto - 100, "ORDEN DE COMPRA #:")
    p.drawString(40, alto - 120, "NOMBRE DE PROVEEDOR:")
    p.drawString(40, alto - 140, "FACTURA #:")
    
    p.setFont("Helvetica", 8)
    oc_val = str(datos.get("orden_compra", ""))
    folio_val = str(datos.get("folio_factura", ""))
    proveedor_val = str(datos.get("nombre_proveedor", ""))
    solicitante_val = str(datos.get("solicitante", ""))
    tipo_servicio_val = str(datos.get("tipo_servicio", "Entrenamiento"))
    
    # Valores sombreados
    p.drawString(170, alto - 100, oc_val)
    p.drawString(170, alto - 120, proveedor_val)
    p.drawString(170, alto - 140, folio_val)
    
    p.setFont("Helvetica-Bold", 8)
    p.drawString(380, alto - 140, "TIPO DE SERVICIO A RECIBIR:")
    p.setFont("Helvetica", 8)
    p.drawString(510, alto - 140, tipo_servicio_val)
    
    # Tabla de Concepto
    y_tabla = alto - 170
    p.setFillColor(colors.HexColor("#FFF2CC"))
    p.rect(40, y_tabla - 60, ancho - 80, 60, fill=1, stroke=1)
    
    p.setFillColor(colors.black)
    p.setFont("Helvetica-Bold", 7)
    p.drawString(45, y_tabla - 12, "# DE LINEA PO")
    p.drawString(105, y_tabla - 12, "CANT.")
    p.drawString(135, y_tabla - 12, "UOM")
    p.drawString(165, y_tabla - 12, "MONTO A RECIBIR")
    p.drawString(250, y_tabla - 12, "MONEDA")
    p.drawString(300, y_tabla - 12, "DESCRIPCIÓN DE LA LINEA A RECIBIR")
    p.drawString(520, y_tabla - 12, "COMPLETA?")
    
    lineas = datos.get("lineas", [])
    subtotal = datos.get("subtotal", 0)
    moneda = datos.get("moneda", "MXN")
    desc = ""
    if lineas:
        l = lineas[0]
        desc = limpiar_descripcion(l.get("descripcion", ""), solicitante_val, oc_val)
        
    p.setFont("Helvetica", 7.5)
    p.drawString(45, y_tabla - 35, str(linea_po_final))
    p.drawString(110, y_tabla - 35, "1")
    p.drawString(135, y_tabla - 35, "LOT")
    p.drawString(165, y_tabla - 35, f"${subtotal:,.2f}")
    p.drawString(255, y_tabla - 35, moneda)
    p.drawString(300, y_tabla - 35, desc[:55])
    p.drawString(530, y_tabla - 35, "YES")
    
    # Totales y Solicitante
    p.setFont("Helvetica-Bold", 8)
    p.drawString(40, y_tabla - 85, "MONTO TOTAL A RECIBIR:")
    p.drawString(165, y_tabla - 85, f"${subtotal:,.2f} {moneda}")
    
    p.drawString(40, y_tabla - 120, "SOLICITANTE:")
    p.setFont("Helvetica", 8)
    p.drawString(120, y_tabla - 120, solicitante_val)
    
    p.setFont("Helvetica-Bold", 8)
    p.drawString(40, y_tabla - 145, "APROBADOR:")
    p.setFont("Helvetica", 8)
    p.drawString(120, y_tabla - 145, "Laura Maciel Hernández")
    
    p.showPage()  # Siguiente página
    
    # ------------------ PÁGINA 2: EVIDENCIAS (ANTES Y DESPUÉS) ------------------
    p.setFillColor(colors.HexColor("#C00000"))
    p.setFont("Helvetica-Bold", 16)
    p.drawString(40, alto - 50, "AAM")
    p.setFillColor(colors.black)
    p.setFont("Helvetica-Bold", 13)
    p.drawString(100, alto - 50, "REPLACEMENT PACKING SLIP EVIDENCE")
    
    p.setFont("Helvetica-Bold", 7)
    p.setFillColor(colors.HexColor("#C00000"))
    p.drawString(40, alto - 65, "INSTRUCCIONES: LAS FOTOS DEBEN ESTAR DEL TAMAÑO DEL RECUADRO MARCADO Y RESOLUCIÓN DE CALIDAD")
    p.drawString(40, alto - 75, "NOTA: SOLO APLICA PARA SERVICIOS")
    
    # Títulos de las cajas
    p.setFont("Helvetica-Bold", 9)
    p.setFillColor(colors.black)
    p.drawString(130, alto - 100, "FOTOS DEL ANTES")
    p.drawString(390, alto - 100, "FOTOS DEL DESPUES")
    
    # Recuadro ANTES (Izquierda)
    w_box = 245
    h_box = 480
    y_box = alto - 595
    p.setStrokeColor(colors.gray)
    p.rect(40, y_box, w_box, h_box, fill=0, stroke=1)
    
    # Recuadro DESPUÉS (Derecha)
    p.rect(320, y_box, w_box, h_box, fill=0, stroke=1)
    
    # Pegar imagen de la OC en el ANTES
    if oc_bytes:
        try:
            img_oc_io = extraer_pagina_partidas_oc(oc_bytes)
            pil_oc = PILImage.open(img_oc_io)
            temp_oc_path = f"/tmp/oc_{int(time.time())}.png"
            pil_oc.save(temp_oc_path)
            p.drawImage(temp_oc_path, 45, y_box + 10, width=w_box - 10, height=h_box - 20, preserveAspectRatio=True)
            if os.path.exists(temp_oc_path):
                os.remove(temp_oc_path)
        except Exception:
            pass
            
    # Pegar fotos en el DESPUÉS
    if fotos_bytes:
        n_fotos = min(len(fotos_bytes), 3)
        h_foto = (h_box - 20) / n_fotos
        for i, fb in enumerate(fotos_bytes[:3]):
            try:
                p_foto = PILImage.open(io.BytesIO(fb))
                temp_foto_path = f"/tmp/foto_{i}_{int(time.time())}.png"
                p_foto.save(temp_foto_path)
                y_pos = (y_box + h_box - 10) - (i + 1) * h_foto
                p.drawImage(temp_foto_path, 325, y_pos + 5, width=w_box - 10, height=h_foto - 10, preserveAspectRatio=True)
                if os.path.exists(temp_foto_path):
                    os.remove(temp_foto_path)
            except Exception:
                pass
                
    p.save()
    buffer.seek(0)
    return buffer

# ----------------- EJECUCIÓN STREAMLIT -----------------
if uploaded_factura and api_key:
    if st.button("Procesar Factura y Generar Documentos"):
        with st.spinner("Analizando factura y OC con IA, extrayendo línea y generando archivos..."):
            try:
                oc_bytes = uploaded_oc.getvalue() if uploaded_oc else None
                fotos_bytes = [f.getvalue() for f in uploaded_fotos] if uploaded_fotos else []
                
                # 1. Extracción y cruce inteligente
                datos = extraer_datos_con_oc(uploaded_factura.getvalue(), oc_bytes, api_key)
                
                # Si el usuario escribió manualmente la línea se respeta; si no, toma la detectada por la IA
                linea_po_final = linea_po_manual.strip() if linea_po_manual.strip() else datos.get("linea_po_detectada", "1-1")
                
                # 2. Generar Excel
                excel_salida = llenar_plantilla_excel(
                    datos, 
                    linea_po_final=linea_po_final,
                    oc_bytes=oc_bytes,
                    fotos_bytes=fotos_bytes
                )
                
                # 3. Generar PDF Oficial
                pdf_salida = generar_pdf_oficial(
                    datos,
                    linea_po_final=linea_po_final,
                    oc_bytes=oc_bytes,
                    fotos_bytes=fotos_bytes
                )
                
                st.success(f"¡Documentos generados exitosamente! (Línea de PO asignada: {linea_po_final})")
                
                with st.expander("Ver datos extraídos y validados por la IA"):
                    st.json(datos)
                
                folio = str(datos.get("folio_factura", "RPS"))
                oc = str(datos.get("orden_compra", "OC"))
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label=f"📊 Descargar RPS {oc} {folio}.xlsx",
                        data=excel_salida,
                        file_name=f"RPS {oc} {folio}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    st.download_button(
                        label=f"📄 Descargar RPS {oc} {folio}.pdf",
                        data=pdf_salida,
                        file_name=f"RPS {oc} {folio}.pdf",
                        mime="application/pdf"
                    )
            except Exception as e:
                st.error(f"Error al procesar: {e}")
