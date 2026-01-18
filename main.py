import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os

st.title("Generador de Facturas")

# Tipo_extra	cuota_extra	pct_iva_extra
excel_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])
docx_template = st.file_uploader("Sube la plantilla Word", type=["docx"])

if st.button("Generar facturas") and excel_file and docx_template:
    df = pd.read_excel(excel_file)
    zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name

    with zipfile.ZipFile(zip_path, "w") as zipf:
        for i, row in df.iterrows():
            doc = DocxTemplate(docx_template)
            subtotal = float(row["cuota_anual"])
            iva_pct = float(row["pct_iva"])
            total = subtotal + iva_pct * subtotal

            #Extras
            subtotal_extra = float(row("cuota_extra"))
            iva_pct_extra = float(row("pct_iva_extra"))
            total_extra = subtotal_extra + iva_pct_extra * subtotal_extra

            total_final = total + total_extra

            context = {
                "nombre": row["Nombre"],
                "direccion": row["direccion"],
                "codigo_postal": row["codigo_postal"],
                "municipio": row["municipio"],
                "CIF": row["CIF"],
                "num_factura": f"{datetime.now(UTC).year}{(i + 1):03d}",
                "fecha": datetime.now(UTC).strftime("%d/%m/%Y"),

                "tipo_socio": row["tipo_socio"],
                "cuota_anual": "{:.2f}".format(subtotal).replace('.', ','),
                "pct_iva_socio": "{:.2f}".format(iva_pct).replace('.', ','),
                "valor_iva_socio": "{:.2f}".format(subtotal * iva_pct).replace('.', ','),

                "tipo_extra": row["tipo_extra"],
                "cuota_extra": row["cuota_extra"],
                "pct_iva_extra": row["pct_iva_extra"],
                "valor_iva_extra": "{:.2f}".format(total_extra).replace('.', ','),


                "total": "{:.2f}".format(total).replace('.', ','),

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
