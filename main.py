import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os

try:
    import pypandoc
    pandoc_available = True
except (ImportError, OSError):
    pandoc_available = False

st.title("Generador de Facturas")

csv_file = st.file_uploader("Sube el archivo CSV", type=["csv"])
docx_template = st.file_uploader("Sube la plantilla Word", type=["docx"])
output_format = st.radio("Formato de salida", options=["DOCX", "PDF"])

if st.button("Generar facturas") and csv_file and docx_template:
    if output_format == "PDF" and not pandoc_available:
        st.error("La conversión a PDF requiere Pandoc instalado localmente. No es compatible con Streamlit Cloud.")
    else:
        df = pd.read_csv(csv_file, sep=',', decimal='.', thousands=',')
        zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name

        with zipfile.ZipFile(zip_path, "w") as zipf:
            for i, row in df.iterrows():
                doc = DocxTemplate(docx_template)
                subtotal = float(row["pago_anual"])
                iva_pct = float(row["pct_iva"])
                total = subtotal + iva_pct * subtotal

                context = {
                    "nombre": row["Nombre"],
                    "CIF": row["CIF"],
                    "tipo_cuota": row["tipo_cuota"],
                    "pago_anual": "{:.2f}".format(subtotal).replace('.', ','),
                    "direccion": row["direccion"],
                    "codigo_postal": row["codigo_postal"],
                    "municipio": row["municipio"],
                    "tipo_socio": row["tipo_socio"],
                    "pct_iva": "{:.2f}".format(iva_pct).replace('.', ','),
                    "valor_iva": "{:.2f}".format(subtotal * iva_pct).replace('.', ','),
                    "num_factura": f"{datetime.now(UTC).year}{(i + 1):03d}",
                    "fecha": datetime.now(UTC).strftime("%d/%m/%Y"),
                    "total": "{:.2f}".format(total).replace('.', ',')
                }

                doc.render(context)

                # Save as .docx
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                    doc.save(tmp_docx.name)

                    if output_format == "DOCX":
                        output_name = f"FACTURA {context['num_factura']} {context['nombre']}.docx"
                        zipf.write(tmp_docx.name, output_name)

                    elif output_format == "PDF":
                        tmp_pdf = tmp_docx.name.replace(".docx", ".pdf")
                        try:
                            pypandoc.convert_file(tmp_docx.name, 'pdf', outputfile=tmp_pdf)
                            output_name = f"FACTURA {context['num_factura']} {context['nombre']}.pdf"
                            zipf.write(tmp_pdf, output_name)
                            os.remove(tmp_pdf)
                        except Exception as e:
                            st.error(f"No se pudo convertir a PDF: {e}")
                    os.remove(tmp_docx.name)

        with open(zip_path, "rb") as f:
            st.download_button("Descargar todas las facturas (.zip)", f, file_name="facturas.zip")
