import io
import json
import os
import time
import google.generativeai as genai
import openpyxl
import streamlit as st

st.set_page_config(
    page_title="Generador RPS AAM", page_icon="📄", layout="centered"
)
st.title("Generador Automático de RPS")
st.write(
    "Sube la factura en PDF para obtener el Excel idéntico y prellenado con"
    " logos y formato oficial."
)

# Obtener la API key de los Secrets de Streamlit
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
- En 'solicitante', revisa tanto los campos de adenda como el texto dentro de la descripción del concepto.
- Devuelve únicamente el JSON sin comentarios adicionales.
"""


def extraer_datos(pdf_bytes, key):
  genai.configure(api_key=key)

  modelos = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
  ultimo_error = None

  for mod_name in modelos:
    try:
      model = genai.GenerativeModel(
          model_name=mod_name,
          generation_config={"response_mime_type": "application/json"},
          system_instruction=SYSTEM_PROMPT,
      )

      pdf_part = {"mime_type": "application/pdf", "data": pdf_bytes}

      response = model.generate_content([
          pdf_part,
          "Extrae los datos de esta factura para llenar el RPS.",
      ])
      return json.loads(response.text)
    except Exception as err:
      ultimo_error = err
      time.sleep(1)
      continue

  raise ultimo_error


def llenar_plantilla(datos, plantilla_path="plantilla_RPS.xlsx"):
  wb = openpyxl.load_workbook(plantilla_path)
  ws = wb["3151"] if "3151" in wb.sheetnames else wb.active

  folio = str(datos.get("folio_factura", "RPS"))
  ws.title = folio

  # 1. Orden de Compra (Fila 7, Columna E)
  if datos.get("orden_compra"):
    oc = datos["orden_compra"]
    ws.cell(
        row=7, column=5, value=int(oc) if str(oc).isdigit() else str(oc)
    )

  # 2. Proveedor (Fila 9, Columna E)
  ws.cell(row=9, column=5, value=datos.get("nombre_proveedor", ""))

  # 3. Folio Factura (Fila 11, Columna E) y Tipo de Servicio (Fila 11, Columna J)
  ws.cell(
      row=11, column=5, value=int(folio) if folio.isdigit() else str(folio)
  )
  ws.cell(row=11, column=10, value=datos.get("tipo_servicio", "Entrenamiento"))

  # 4. Detalle de Concepto (Fila 15)
  lineas = datos.get("lineas", [])
  if lineas:
    l = lineas[0]
    ws.cell(row=15, column=1, value=l.get("linea_po", "3-1"))
    ws.cell(row=15, column=2, value=l.get("cantidad", 1))
    ws.cell(row=15, column=4, value=l.get("unidad", "LOT"))
    ws.cell(row=15, column=5, value=l.get("monto", datos.get("subtotal", 0)))
    ws.cell(row=15, column=6, value=datos.get("moneda", "MXN"))
    ws.cell(row=15, column=7, value=l.get("descripcion", ""))
    ws.cell(row=15, column=16, value="YES")

  # 5. Monto Total a Recibir (Fila 25, Columnas E y F)
  subtotal = datos.get("subtotal", 0)
  ws.cell(row=25, column=5, value=subtotal)
  ws.cell(row=25, column=6, value=subtotal)

  # 6. Solicitante (Fila 27, Columna C)
  if datos.get("solicitante"):
    ws.cell(row=27, column=3, value=datos["solicitante"])

  output = io.BytesIO()
  wb.save(output)
  output.seek(0)
  return output, folio, datos.get("orden_compra", "RPS")


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
