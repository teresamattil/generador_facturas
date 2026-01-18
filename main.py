import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os

st.title("Generador de Facturas")

EXPECTED_COLUMNS = {
    "Nombre",
    "direccion",
    "codigo_postal",
    "municipio",
    "CIF",
    "tipo_socio",
    "cuota_anual",
    "pct_iva_socio",
    "Tipo_extra",
    "cuota_extra",
    "pct_iva_extra",
}

uploaded_excel = st.file_uploader("Sube el archivo Excel", type=["xlsx", "xls"])
uploaded_template = st.file_uploader("Sube la plantilla Word", type=["docx"])

if st.button("Generar facturas") and uploaded_excel and uploaded_template:
    try:
        df = pd.read_excel(uploaded_excel)
    except Exception as e:
        st.error("❌ No se pudo leer el archivo Excel")
        st.stop()

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
        tpl.write(uploaded_template.read())
        template_path = tpl.name

    zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name

    with zipfile.ZipFile(zip_path, "w") as zipf:
        for i, row in df.iterrows():
            doc = DocxTemplate(template_path)

            subtotal = float(row.get("cuota_anual", 0))
            iva_pct = float(row.get("pct_iva_socio", 0))
            subtotal_extra = float(row.get("cuota_extra", 0))
            iva_pct_extra = float(row.get("pct_iva_extra", 0))

            total = subtotal * (1 + iva_pct)
            total_extra = subtotal_extra * (1 + iva_pct_extra)
            total_final = total + total_extra

            context = {
                "nombre": row.get("Nombre", ""),
                "direccion": row.get("direccion", ""),
                "codigo_postal": row.get("codigo_postal", ""),
                "municipio": row.get("municipio", ""),
                "CIF": row.get("CIF", ""),
                "num_factura": f"{datetime.now(UTC).year}{(i + 1):03d}",
                "fecha": datetime.now(UTC).strftime("%d/%m/%Y"),

                "tipo_socio": row.get("tipo_socio", ""),
                "cuota_anual": f"{subtotal:.2f}".replace('.', ','),
                "pct_iva_socio": f"{iva_pct:.2f}".replace('.', ','),
                "valor_iva_socio": f"{subtotal * iva_pct:.2f}".replace('.', ','),

                "Tipo_extra": row.get("Tipo_extra", ""),
                "cuota_extra": f"{subtotal_extra:.2f}".replace('.', ','),
                "pct_iva_extra": f"{iva_pct_extra:.2f}".replace('.', ','),
                "valor_iva_extra": f"{subtotal_extra * iva_pct_extra:.2f}".replace('.', ','),

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
