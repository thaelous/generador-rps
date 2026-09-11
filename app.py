import base64
import io
import json
import os
import re
import time
import openpyxl
import requests
import streamlit as st

st.set_page_config(
    page_title="Generador RPS AAM", page_icon="📄", layout="centered"
)
st.title("Generador Automático de RPS")
st.write(
    "Sube la factura en PDF para obtener el Excel idéntico y prellenado con"
    " logos y formato oficial."
)

api_key = st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))

if not api_key:
  api_key = st.text_input("Ingresa tu Gemini API Key:", type="password")

uploaded_pdf = st.file_uploader(
    "Selecciona la Factura en PDF (CFDI)", type=["pdf"]
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
      "linea_po": string,
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
  pdf_b64 = base64.b64encode(pdf_bytes).decode("utf-8")

  payload = {
      "system_instruction": {"parts": [{"text": SYSTEM_PROMPT}]},
      "contents": [{
          "parts": [
              {"inline_data": {"mime_type": "application/pdf", "data": pdf_b64}},
              {"text": "Extrae los datos de esta factura para llenar el RPS."},
          ]
      }],
      "generationConfig": {
          "response_mime_type": "application/json",
          "temperature": 0.1,
      },
  }

  headers = {
      "Content-Type": "application/json",
      "x-goog-api-key": clean_key,
  }

  # Modelos Flash disponibles con cuota libre
  modelos = [
      "gemini-2.5-flash",
      "gemini-2.5-flash-lite",
      "gemini-flash-latest",
  ]
  ultimo_error = None

  for mod in modelos:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{mod}:generateContent".strip()
    try:
      response = requests.post(url, headers=headers, json=payload, timeout=60)
      if response.status_code == 200:
        data = response.json()
        texto = data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(texto)
      else:
        ultimo_error = f"Código {response.status_code}: {response.text}"
        time.sleep(1)
    except Exception as e:
      ultimo_error = str(e)

  raise RuntimeError(ultimo_error)


def limpiar_descripcion(desc, solicitante, oc):
  texto = desc

  if solicitante and solicitante.strip():
    patron_sol = re.compile(re.escape(solicitante.strip()), re.IGNORECASE)
    texto = patron_sol.sub("", texto)

  texto = re.sub(
      r"(?i)(solicita(?:nte)?|atenci[oó]n|contacto)\s*[:\-]?\s*", "", texto
  )

  if oc and str(oc).strip():
    patron_oc = re.compile(
        rf"(?i)(?:oc|orden\s*(?:de)?\s*compra|po)\s*[:#\-]?\s*{re.escape(str(oc).strip())}"
    )
    texto = patron_oc.sub("", texto)
    texto = re.sub(rf"\b{re.escape(str(oc).strip())}\b", "", texto)

  texto = re.sub(r"(?i)\b(?:oc|po)\s*[:#\-]?\b", "", texto)
  texto = re.sub(r"\s+", " ", texto)
  texto = re.sub(r"^[\s,\.\-_/]+|[\s,\.\-_/]+$", "", texto)

  return texto.strip()


def llenar_plantilla(datos, plantilla_path="plantilla_RPS.xlsx"):
  wb = openpyxl.load_workbook(plantilla_path)
  ws = wb["3151"] if "3151" in wb.sheetnames else wb.active

  folio = str(datos.get("folio_factura", "RPS"))
  oc_val = datos.get("orden_compra", "")
  solicitante_val = datos.get("solicitante", "")
  ws.title = folio

  # 1. Orden de Compra (Fila 7, Columna E)
  if oc_val:
    ws.cell(
        row=7,
        column=5,
        value=int(oc_val) if str(oc_val).isdigit() else str(oc_val),
    )

  # 2. Proveedor (Fila 9, Columna E)
  ws.cell(row=9, column=5, value=datos.get("nombre_proveedor", ""))

  # 3. Folio Factura (Fila 11, Columna E) y Tipo de Servicio (Fila 11, Columna J)
  ws.cell(row=11, column=5, value=int(folio) if folio.isdigit() else str(folio))
  ws.cell(row=11, column=10, value=datos.get("tipo_servicio", "Entrenamiento"))

  # 4. Detalle de Concepto (Fila 15)
  lineas = datos.get("lineas", [])
  if lineas:
    l = lineas[0]

    # Vacío en '# de línea en PO'
    ws.cell(row=15, column=1, value="")

    # Cantidad
    ws.cell(row=15, column=2, value=l.get("cantidad", 1))

    # Unidad limpia sin E48
    unidad = str(l.get("unidad", "LOT")).strip()
    if "E48" in unidad.upper() or not unidad:
      unidad = "LOT"
    else:
      unidad = unidad.replace("E48", "").replace("-", "").strip()
    ws.cell(row=15, column=4, value=unidad)

    # Monto y Moneda
    ws.cell(row=15, column=5, value=l.get("monto", datos.get("subtotal", 0)))
    ws.cell(row=15, column=6, value=datos.get("moneda", "MXN"))

    # Descripción limpia (sin solicitante ni OC)
    desc_original = l.get("descripcion", "")
    desc_limpia = limpiar_descripcion(desc_original, solicitante_val, oc_val)
    ws.cell(row=15, column=7, value=desc_limpia)

    ws.cell(row=15, column=16, value="YES")

  # 5. Monto Total a Recibir (Fila 25, Columnas E y F)
  subtotal = datos.get("subtotal", 0)
  ws.cell(row=25, column=5, value=subtotal)
  ws.cell(row=25, column=6, value=subtotal)

  # 6. Solicitante (Fila 27, Columna C)
  if solicitante_val:
    ws.cell(row=27, column=3, value=solicitante_val)

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output, folio, oc_val if oc_val else "RPS"


if uploaded_pdf and api_key:
  if st.button("Procesar Factura y Generar RPS"):
    with st.spinner("Leyendo factura con Gemini y llenando formato..."):
      try:
        datos = extraer_datos(uploaded_pdf.getvalue(), api_key)
        excel_salida, folio, oc = llenar_plantilla(datos)

        st.success("¡RPS generado exitosamente!")

        with st.expander("Ver datos extraídos"):
          st.json(datos)

        nombre_descarga = f"RPS {oc} {folio}.xlsx"
        st.download_button(
            label=f"📥 Descargar {nombre_descarga}",
            data=excel_salida,
            file_name=nombre_descarga,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      except Exception as e:
        st.error(f"Error al procesar: {e}")
