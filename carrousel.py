import pandas as pd
import os
import sys
import textwrap
import base64
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from io import BytesIO
from PIL import Image as PILImage

def load_image_from_base64(base64_file):
    with open(base64_file, 'r') as f:
        base64_data = f.read()
    image_data = base64.b64decode(base64_data)
    image = PILImage.open(BytesIO(image_data))
    return image


def truncate_text(text, max_length):
    """Convertir a text si es un número."""
    if isinstance(text, (int, float)):
        text = str(text)
    """Truncar el texto si excede max_length, agregando '...' al final."""
    if len(text) > max_length:
        return textwrap.shorten(text, width=max_length, placeholder="...")
    return text

def generate_reports(input_file, output_folder=None, logo_path=None):
    try:
        if output_folder is None:
            home = os.path.expanduser("~")
            desktop = os.path.join(home, "Desktop")
            output_folder = os.path.join(desktop, "reportes")

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        df = pd.read_excel(input_file, sheet_name=0)

        df.iloc[:, 11] = pd.to_numeric(df.iloc[:, 11], errors='coerce')
        df.iloc[:, 12] = pd.to_numeric(df.iloc[:, 12], errors='coerce')

        df = df.dropna(subset=[df.columns[11], df.columns[12]])

        families = df.groupby(df.iloc[:, 0])

        for name, group in families:
            if group.empty:
                continue

            report_df = group.iloc[:, [4, 5, 11, 12]].copy()

            total_items_sold = report_df.iloc[:, 2].sum()
            total_income = report_df.iloc[:, 3].sum()
            amount_to_collect = total_income * 0.6

            report_path = os.path.join(output_folder, f'Reporte_{name}.pdf')
            create_pdf_report_with_logo(report_path, name, report_df, total_items_sold, total_income, amount_to_collect, logo_path)

            print(f'Reporte generado: {report_path}')

    except Exception as e:
        print(f"Error al generar los reportes: {e}")

def format_quantity(value):
    """Formatear la cantidad como un número entero."""
    return int(value)

def format_value(value):
    """Formatear el valor, quitando .0 si es redondeado."""
    return f'{value:.2f}'.rstrip('0').rstrip('.')  # Eliminar .0 si no es necesario

def create_pdf_report_with_logo(report_path, family_name, report_df, total_items_sold, total_income, amount_to_collect, base64_file):
    doc = SimpleDocTemplate(report_path, pagesize=letter)
    elements = []

    # Load and decode the base64 image
    logo_image = load_image_from_base64(base64_file)
    logo_buffer = BytesIO()
    logo_image.save(logo_buffer, format="PNG")
    logo_data = logo_buffer.getvalue()

    # Header with logo (puedes omitir esto si no quieres un logo)
    # Header with logo
    logo = Image(BytesIO(logo_data), width=7.5*inch, height=2.4*inch)
    logo.hAlign = 'CENTER'
    elements.append(logo)
    elements.append(Spacer(1, 12))

    elements.append(Spacer(1, 24))

    # Sales table
    data = [['Proveedora', 'Artículo', 'Nombre de Artículo', 'Cantidad', 'Valor']]

    # Separar el código de artículo y el nombre del artículo, y formatear la cantidad y el valor
    for index, row in report_df.iterrows():
        data.append([
            truncate_text(family_name, 20),  # Proveedora (columna A)
            truncate_text(row.iloc[0], 10),  # Código del artículo (columna E)
            truncate_text(row.iloc[1], 30),  # Nombre del artículo (columna F)
            format_quantity(row.iloc[2]),  # Cantidad (columna L)
            format_value(row.iloc[3])  # Valor (columna M)
        ])

    # Ajuste de los anchos de las columnas
    table = Table(data, colWidths=[1.8 * inch, .8 * inch, 3.0 * inch, 0.8 * inch, 1.0 * inch])

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F2D3B2")),  # Fondo del encabezado
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#974806")),  # Color del texto del encabezado
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Alineación centrada en el encabezado
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Fuente y negrita en el encabezado
        ('FONTSIZE', (0, 0), (-1, 0), 10),  # Tamaño de la fuente del encabezado
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#FCEFE3")),  # Fondo de las filas
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),  # Color del texto de las filas
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),  # Alineación centrada en las filas
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Fuente en las filas
        ('FONTSIZE', (0, 1), (-1, -1), 8),  # Tamaño de la fuente en las filas
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#976D48")),  # Bordes de la tabla
       # ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor("#0000FF")),  # Línea azul debajo del encabezado
        #('LINEBEFORE', (0, 0), (0, -1), 1, colors.HexColor("#0000FF")),  # Línea azul a la izquierda de la primera columna
    ]))

    elements.append(table)

    elements.append(Spacer(1, 24))

    # Summary table (ajustada dinámicamente)
    summary_data = [
        ['Artículos Vendidos:', format_value(total_items_sold)],
        ['Cantidades Vendidas:', format_value(total_items_sold)],
        ['Ingresos Totales:', f'${format_value(total_income)}'],
        ['A cobrar: 60% del monto total', f'${format_value(amount_to_collect)}'],
        ['', ''],
        ['Crédito adicional del 15% disponible si decidís reinvertir en productos de Carrousel', ''],
        ['(Este crédito es válido indefinidamente y puede ser utilizado en cualquier momento que encuentres productos de su interés)', '']
    ]

    # Ajuste dinámico basado en el tamaño disponible en la página
    available_width = 7.4 * inch  # Ajustado para que se ajuste mejor al tamaño del documento
    summary_table = Table(summary_data, colWidths=[available_width * 0.7, available_width * 0.3])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 3), colors.HexColor("#F2D3B2")),
        ('TEXTCOLOR', (0, 0), (-1, 3), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 3), 12),
        ('GRID', (0, 0), (-1, 3), 1, colors.black),
        ('SPAN', (0, 5), (-1, 5)),
        ('SPAN', (0, 6), (-1, 6)),
        ('TEXTCOLOR', (0, 4), (-1, -1), colors.black),
        ('FONTNAME', (0, 4), (-1, -1), 'Helvetica'),
    ]))
    elements.append(summary_table)

    doc.build(elements)

def get_logo_path():
    if hasattr(sys, '_MEIPASS'):
        # Esto ocurre cuando el ejecutable corre desde un archivo empaquetado
        return os.path.join(sys._MEIPASS, 'logo_base64.txt')
    else:
        # Esto ocurre cuando corres el script sin empaquetar
        return 'logo_base64.txt'

if __name__ == "__main__":
    if len(sys.argv) < 2:
        input_file = input("Por favor, ingresa la ruta completa del archivo Excel: ")
    else:
        input_file = sys.argv[1]

    if not input_file:
        print("No se seleccionó ningún archivo. Saliendo...")
        sys.exit(1)

    output_folder = None
    base64_file = get_logo_path()  # Archivo con el logo en base64

    generate_reports(input_file, output_folder, base64_file)
