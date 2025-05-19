import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os


# Only import docx2pdf on Windows/macOS
try:
    from docx2pdf import convert
except ImportError:
    convert = None

st.title("Generador de Facturas")

csv_file = st.file_uploader("Sube el archivo CSV", type=["csv"])
docx_template = st.file_uploader("Sube la plantilla Word", type=["docx"])
output_format = st.selectbox("Formato de descarga", ["DOCX", "PDF"])

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
            factura_name = f"FACTURA {context['num_factura']} {context['nombre']}"
            docx_path = os.path.join(tempfile.gettempdir(), f"{factura_name}.docx")
            doc.save(docx_path)

            # Convert to PDF if selected
            if output_format == "PDF":
                pdf_path = os.path.join(tempfile.gettempdir(), f"{factura_name}.pdf")
                if convert:
                    convert(docx_path, pdf_path)
                    zipf.write(pdf_path, f"{factura_name}.pdf")
                else:
                    st.warning("La conversión a PDF solo está disponible en macOS o Windows.")
                    zipf.write(docx_path, f"{factura_name}.docx")
            else:
                zipf.write(docx_path, f"{factura_name}.docx")

    with open(zip_path, "rb") as f:
        st.download_button(
            label=f"Descargar todas las facturas ({output_format}) (.zip)",
            data=f,
            file_name=f"facturas_{output_format.lower()}.zip"
        )
