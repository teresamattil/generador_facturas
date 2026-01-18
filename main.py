import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os
import math

st.title("Generador de Facturas")

EXPECTED_COLUMNS = {
    "Nombre",
    "direccion",
    "codigo_postal",
    "municipio",
    "CIF",
    "tipo_socio",
    "cuota_anual",
    "pct_iva",
    "tipo_extra",
    "cuota_extra",
    "pct_iva_extra",
}

excel_file = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])
docx_template = st.file_uploader("Sube la plantilla Word", type=["docx"])

if st.button("Generar facturas") and excel_file and docx_template:
    df = pd.read_excel(excel_file)

    given_columns = set(df.columns)
    missing_columns = EXPECTED_COLUMNS - given_columns
    unused_columns = given_columns - EXPECTED_COLUMNS

    if missing_columns:
        st.error("❌ Error en las columnas del Excel")
        st.write("**Columnas esperadas:**", sorted(EXPECTED_COLUMNS))
        st.write("**Columnas recibidas:**", sorted(given_columns))
        st.write("**Columnas faltantes:**", sorted(missing_columns))
        st.write("**Columnas no usadas:**", sorted(unused_columns))
        st.stop()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tpl:
        tpl.write(docx_template.read())
        template_path = tpl.name

    zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name

    with zipfile.ZipFile(zip_path, "w") as zipf:
        for i, row in df.iterrows():
            doc = DocxTemplate(template_path)

            subtotal = float(row["cuota_anual"])
            iva_pct = float(row["pct_iva"])
            total = subtotal + subtotal * iva_pct

            tiene_extra = (
                pd.notna(row["tipo_extra"])
                and pd.notna(row["cuota_extra"])
                and float(row["cuota_extra"]) > 0
            )

            if tiene_extra:
                subtotal_extra = float(row["cuota_extra"])
                iva_pct_extra = float(row["pct_iva_extra"])
                total_extra = subtotal_extra + subtotal_extra * iva_pct_extra
            else:
                subtotal_extra = 0
                iva_pct_extra = 0
                total_extra = 0

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
                "pct_iva_socio": f"{iva_pct:.2f}".replace('.', ','),
                "valor_iva_socio": f"{subtotal * iva_pct:.2f}".replace('.', ','),

                "tipo_extra": row["tipo_extra"] if tiene_extra else None,
                "cuota_extra": f"{subtotal_extra:.2f}".replace('.', ',') if tiene_extra else "",
                "pct_iva_extra": f"{iva_pct_extra:.2f}".replace('.', ',') if tiene_extra else "",
                "valor_iva_extra": f"{subtotal_extra * iva_pct_extra:.2f}".replace('.', ',') if tiene_extra else "",

                "total": f"{total_final:.2f}".replace('.', ','),
            }

            doc.render(context)

            output_name = f"FACTURA {context['num_factura']} {context['nombre']}.docx"
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                zipf.write(tmp.name, output_name)
                os.remove(tmp.name)

    os.remove(template_path)

    with open(zip_path, "rb") as f:
        st.download_button(
            "Descargar todas las facturas (.zip)",
            f,
            file_name="facturas.zip"
        )
