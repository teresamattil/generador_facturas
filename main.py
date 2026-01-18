import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os

st.title("Generador de Facturas")

excel_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])
docx_template = st.file_uploader("Sube la plantilla Word", type=["docx"])

if st.button("Generar facturas") and excel_file and docx_template:
    df = pd.read_excel(excel_file)
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

            output_name = f"FACTURA {context['num_factura']} {context['nombre']}.docx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                zipf.write(tmp.name, output_name)
                os.remove(tmp.name)

    with open(zip_path, "rb") as f:
        st.download_button(
            "Descargar todas las facturas (.zip)",
            f,
            file_name="facturas.zip"
        )
