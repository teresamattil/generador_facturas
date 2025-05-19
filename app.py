import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime, UTC

# Carga el Excel
df = pd.read_csv('/Users/teresamattil/Desktop/proyecto mama/info_socios_maestro.csv', sep=',', decimal='.', thousands=',')

# Carga la plantilla
template = DocxTemplate('/Users/teresamattil/Desktop/proyecto mama/FACTURA plantilla socios.docx')

for i, row in df.iterrows():
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
        "num_factura": f"{datetime.now(UTC).year}{(i+1):03d}",
        "fecha": datetime.now(UTC).strftime("%d/%m/%Y"),
        "total": "{:.2f}".format(total).replace('.', ',')
    }
    print(f"iva: {context['pct_iva']} pago_anual {context['pago_anual']} valor iva {context['valor_iva']}")
    template.render(context)
    output_filename = f"output/FACTURA {context['num_factura']} {context['nombre']}.docx"
    template.save(output_filename)

