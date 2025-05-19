import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os
import subprocess

st.title("Generador de Facturas")

csv_file = st.file_uploader("Sube el archivo CSV", type=["csv"])
docx_template = st.file_uploader("Sube la plantilla Word", type=["docx"])
output_format = st.radio("Formato de salida", options=["DOCX", "PDF"])

def convert_to_pdf(docx_path, output_dir):
    """Converts a DOCX file to PDF using LibreOffice."""
    try:
        subprocess.run([
            "soffice",  # LibreOffice CLI
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            docx_path
        ], check=True)
        return os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    except subprocess.CalledProcessError as e:
        st.error(f"Error al convertir a PDF: {e}")
        return None

if st.button("Generar facturas") and csv_file and docx_template:
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

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                doc.save(tmp_docx.name)

                if output_format == "DOCX":
                    output_name = f"FACTURA {context['num_factura']} {context['nombre']}.docx"
                    zipf.write(tmp_docx.name, output_name)

                elif output_format == "PDF":
                    pdf_path = convert_to_pdf(tmp_docx.name, os.path.dirname(tmp_docx.name))
                    if pdf_path:
                        output_name = f"FACTURA {context['num_factura']} {context['nombre']}.pdf"
                        zipf.write(pdf_path, output_name)
                        os.remove(pdf_path)

                os.remove(tmp_docx.name)

    with open(zip_path, "rb") as f:
        st.download_button("Descargar todas las facturas (.zip)", f, file_name="facturas.zip")
