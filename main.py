import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC
import tempfile
import zipfile
import os
import subprocess

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


def docx_to_pdf(docx_path, output_dir):
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            output_dir,
            docx_path,
        ],
        check=True
    )


if st.button("Generar facturas") and uploaded_excel and uploaded_template:
    df = pd.read_excel(uploaded_excel)

    given_columns = set(df.columns)
    missing_columns = EXPECTED_COLUMNS - given_columns
    unused_columns = given_columns - EXPECTED_COLUMNS

    if missing_columns:
        st.error("❌ Error en las columnas del Excel")
        st.write("**Esperadas:**", sorted(EXPECTED_COLUMNS))
        st.write("**Recibidas:**", sorted(given_columns))
        st.write("**Faltantes:**", sorted(missing_columns))
        st.write("**No usadas:**", sorted(unused_columns))
        st.stop()

    st.success("✔️ Generación iniciada")

    progress_bar = st.progress(0)
    status_text = st.empty()
    total_rows = len(df)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tpl:
        tpl.write(uploaded_template.read())
        template_path = tpl.name

    zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name

    total_gente_extra = 0

    with zipfile.ZipFile(zip_path, "w") as zipf:
        for i, row in df.iterrows():
            status_text.text(f"Generando factura {i + 1} de {total_rows} — {row['Nombre']}")
            progress_bar.progress((i + 1) / total_rows)

            doc = DocxTemplate(template_path)

            subtotal = float(row["cuota_anual"])
            iva_pct = float(row["pct_iva_socio"]/100)
            total = subtotal * (1 + iva_pct)

            tiene_extra = (
                pd.notna(row["Tipo_extra"])
                and pd.notna(row["cuota_extra"])
                and float(row["cuota_extra"]) > 0
            )

            if tiene_extra:
                total_gente_extra += 1
                subtotal_extra = float(row["cuota_extra"])
                iva_pct_extra = float(row["pct_iva_extra"]/100)
                total_extra = subtotal_extra * (1 + iva_pct_extra)
            else:
                Tipo_extra = "\u200b"  # zero-width space (invisible)
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
                "cuota_anual": f"{subtotal:.2f}".replace('.', ','),
                "pct_iva_socio": f"{iva_pct:.2f}".replace('.', ','),
                "valor_iva_socio": f"{subtotal * iva_pct:.2f}".replace('.', ','),

                "Tipo_extra": row["Tipo_extra"] if tiene_extra else None,
                "cuota_extra": f"{subtotal_extra:.2f}".replace('.', ',') if tiene_extra else "",
                "pct_iva_extra": f"{iva_pct_extra*100:.2f}".replace('.', ',') if tiene_extra else "",
                "valor_iva_extra": f"{subtotal_extra * iva_pct_extra:.2f}".replace('.', ',') if tiene_extra else "",

                "total": f"{total_final:.2f}".replace('.', ','),
            }

            doc.render(context)

            output_name = f"FACTURA {context['num_factura']} {context['nombre']}.docx"
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)

                # Rutas finales dentro del ZIP
                docx_zip_path = f"docx/{output_name}"
                pdf_zip_path = f"pdf/{output_name.replace('.docx', '.pdf')}"

                # Añadir DOCX
                zipf.write(tmp.name, docx_zip_path)

                # Convertir a PDF
                docx_to_pdf(tmp.name, tempfile.gettempdir())
                pdf_path = tmp.name.replace(".docx", ".pdf")

                # Añadir PDF
                zipf.write(pdf_path, pdf_zip_path)

                # Limpieza
                os.remove(tmp.name)
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)


    os.remove(template_path)

    progress_bar.progress(1.0)
    status_text.text("✔️ Facturas generadas correctamente")

    with open(zip_path, "rb") as f:
        st.download_button(
            "Descargar todas las facturas (.zip)",
            f,
            file_name="facturas.zip"
        )

