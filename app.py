from flask import Flask, render_template, request, send_file
import openpyxl
from datetime import datetime
import io
import zipfile

# Librerías 100% nativas de Python para generar PDF en Windows
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)

def generar_pdf_fiel(datos, fecha_hoy):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=40,
        rightMargin=40,
        topMargin=50,
        bottomMargin=40
    )

    # Anchos de columnas (suman exactamente los 532pt disponibles de la página)
    col_widths = [50, 216, 56, 50, 80, 80]

    # Formateo del valor FOB
    try:
        fob_num = float(datos['FOB'])
        fob_str = f"${fob_num:,.2f}"
    except (ValueError, TypeError):
        fob_str = "$0.00"

    # Estilos tipográficos
    title_style = ParagraphStyle('Title', fontName='Helvetica-Bold', fontSize=10.5, alignment=1, spaceAfter=18)
    p_h_left = ParagraphStyle('HLeft', fontName='Helvetica', fontSize=9, leading=11, alignment=0)
    p_h_center = ParagraphStyle('HCenter', fontName='Helvetica', fontSize=9, leading=11, alignment=1)
    p_val_center = ParagraphStyle('VCenter', fontName='Helvetica', fontSize=9, leading=11, alignment=1)
    p_shipper = ParagraphStyle('Shipper', fontName='Helvetica-Bold', fontSize=9.5, leading=12, alignment=1)
    p_importer = ParagraphStyle('Importer', fontName='Helvetica-Bold', fontSize=7.5, leading=9.5, alignment=1)
    p_origin = ParagraphStyle('Origin', fontName='Helvetica', fontSize=9, leading=15, alignment=0)
    p_th = ParagraphStyle('TH', fontName='Helvetica', fontSize=7.5, leading=9.5, alignment=1)
    p_desc = ParagraphStyle('Desc', fontName='Helvetica', fontSize=7, leading=9.5, alignment=1)
    p_total_lbl = ParagraphStyle('TotLbl', fontName='Helvetica', fontSize=9, leading=11, alignment=2)
    p_total_val = ParagraphStyle('TotVal', fontName='Helvetica', fontSize=8.5, leading=10.5, alignment=1)
    p_legal = ParagraphStyle('Legal', fontName='Helvetica', fontSize=8, leading=11, alignment=0)
    p_legal_sub = ParagraphStyle('LegalSub', fontName='Helvetica', fontSize=8, leading=11, alignment=0, spaceAfter=8)

    story = []
    story.append(Paragraph("COMMERCIAL INVOICE", title_style))

    # Construcción de la cuadrícula idéntica a tu formato
    peso_texto = f"{datos['Kg']}kg" if datos['Kg'] is not None else ""
    cant_texto = str(int(datos['Item'])) if datos['Item'] is not None else "1"

    t_data = [
        # Fila 0: Encabezados Fecha y AWB
        [Paragraph("EXPORT DATE (Fecha Exportación)", p_h_left), "", Paragraph("AWB SKY BILL", p_h_center), "", "", ""],
        # Fila 1: Valores Fecha y AWB
        [Paragraph(fecha_hoy, p_val_center), "", Paragraph(str(datos['SH']), p_val_center), "", "", ""],
        # Fila 2: Encabezados Shipper e Importer
        [Paragraph("SHIPPER/REMITENTE", p_h_left), "", Paragraph("IMPORTER (If Other than consignee)", p_h_left), "", "", ""],
        # Fila 3: Valores Shipper e Importer
        [Paragraph("PLUSCOURIER S.A.S", p_shipper), "", Paragraph(str(datos['Destinatario']), p_importer), "", "", ""],
        # Fila 4: Origen y Destino | Cuadro vacío contiguo
        [Paragraph("COUNTRY ORIGIN OF GOODS<br/><br/>(País de origen)<br/><br/>ESTADOS UNIDOS<br/><br/>COUNTRY DESTINATION OF GOODS<br/><br/>(País de destino) Ecuador", p_origin), "", "", "", "", ""],
        # Fila 5: Encabezados de ítems
        [
            Paragraph("No Of PKGS<br/>(Número de Envíos)", p_th),
            Paragraph("COMPLETE DESCRIPTION OF GOODS<br/>(Número de Envíos)", p_th),
            Paragraph("WEIGH T<br/>(Peso)", p_th),
            Paragraph("QUANTI TY<br/>(Cantidad)", p_th),
            Paragraph("UNIT VALUES<br/>(Valor unitario)", p_th),
            Paragraph("TOTAL VALUES<br/>(Valor total)", p_th)
        ],
        # Fila 6: Detalle del paquete
        [
            Paragraph("1", p_val_center),
            Paragraph(str(datos['Description'] or ""), p_desc),
            Paragraph(peso_texto, p_val_center),
            Paragraph(cant_texto, p_val_center),
            Paragraph(fob_str, p_val_center),
            Paragraph(fob_str, p_val_center)
        ],
        # Fila 7: Fila de TOTAL
        ["", "", "", "", Paragraph("TOTAL", p_total_lbl), Paragraph(fob_str, p_total_val)]
    ]

    t_style = TableStyle([
        # Borde exterior continuo y cerrado
        ('BOX', (0,0), (-1,-1), 1.3, colors.black),
        
        # Combinaciones de celdas horizontales superiores
        ('SPAN', (0,0), (1,0)),
        ('SPAN', (2,0), (5,0)),
        ('SPAN', (0,1), (1,1)),
        ('SPAN', (2,1), (5,1)),
        ('SPAN', (0,2), (1,2)),
        ('SPAN', (2,2), (5,2)),
        ('SPAN', (0,3), (1,3)),
        ('SPAN', (2,3), (5,3)),
        ('SPAN', (0,4), (1,4)),
        ('SPAN', (2,4), (5,4)),
        ('SPAN', (0,7), (3,7)),

        # Líneas horizontales divisorias
        ('LINEBELOW', (0,1), (-1,1), 1.3, colors.black),
        ('LINEBELOW', (0,3), (-1,3), 1.3, colors.black),
        ('LINEBELOW', (0,4), (-1,4), 1.3, colors.black),
        ('LINEBELOW', (0,5), (-1,5), 1.3, colors.black),
        ('LINEBELOW', (0,6), (-1,6), 1.3, colors.black),
        
        # Línea vertical central (divide las secciones superiores en 50/50)
        ('LINEAFTER', (1,0), (1,4), 1.3, colors.black),
        
        # Cuadrícula interna de las columnas de ítems
        ('INNERGRID', (0,5), (-1,6), 1.3, colors.black),
        
        # Líneas para la fila TOTAL
        ('LINEBEFORE', (4,7), (4,7), 1.3, colors.black),
        ('LINEBEFORE', (5,7), (5,7), 1.3, colors.black),

        # Alineación y espaciados
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('VALIGN', (0,0), (1,0), 'TOP'),
        ('VALIGN', (2,0), (5,0), 'TOP'),
        ('VALIGN', (0,2), (1,2), 'TOP'),
        ('VALIGN', (2,2), (5,2), 'TOP'),
        ('VALIGN', (0,4), (1,4), 'TOP'),

        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        
        ('BOTTOMPADDING', (0,1), (-1,1), 12),
        ('BOTTOMPADDING', (0,3), (-1,3), 10),
        ('TOPPADDING', (0,4), (-1,4), 8),
        ('BOTTOMPADDING', (0,4), (-1,4), 12),
        ('TOPPADDING', (0,6), (-1,6), 15),
        ('BOTTOMPADDING', (0,6), (-1,6), 15),
        ('TOPPADDING', (4,7), (5,7), 7),
        ('BOTTOMPADDING', (4,7), (5,7), 7),
    ])

    t = Table(t_data, colWidths=col_widths)
    t.setStyle(t_style)
    story.append(t)
    story.append(Spacer(1, 16))

    # Textos legales inferiores
    story.append(Paragraph("THESE COMMODITIES ARE LICENSED FOR THE ULTIMATE DESTINATION SHOWN DIVERSION CONTRARY TO THE UNITED STATES LAW IS PROHIBITED.", p_legal))
    story.append(Paragraph("(Estos envíos están autorizados para el destino indicado.Cambios contrarios a las leyes Americanas están Prohibidos).", p_legal_sub))
    story.append(Paragraph("I DECLARE THAT ALL THE INFORMATION IN THIS INVOICE IS TRUE AND CORRECT.", p_legal))
    story.append(Paragraph("(Declaro que toda la información contenida es verdad.)", p_legal_sub))

    doc.build(story)
    return buf.getvalue()

def leer_manifiesto(stream_excel):
    wb = openpyxl.load_workbook(stream_excel, data_only=True)
    ws = wb.active

    header_row = None
    col = {}
    for r in range(1, min(10, ws.max_row + 1)):
        row_map = {}
        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val is not None:
                row_map[str(val).strip().upper()] = c
        if 'SH' in row_map:
            header_row = r
            col = row_map
            break

    if not header_row:
        raise ValueError("No se encontró la columna 'SH' en el manifiesto.")

    def get_c(*names):
        for n in names:
            if n.upper() in col:
                return col[n.upper()]
        return None

    c_sh = get_c('SH')
    c_dest = get_c('DESTINATARIO', 'CONSIGNEE')
    c_desc = get_c('DESCRIPTION', 'DESCRIPCION')
    c_kg = get_c('KG', 'KILOS')
    c_item = get_c('ITEM', 'ITEMS')
    c_fob = get_c('US$ FOB', 'FOB')

    registros = []
    for r in range(header_row + 1, ws.max_row + 1):
        sh_val = ws.cell(r, c_sh).value if c_sh else None
        if not sh_val or str(sh_val).strip().upper() == 'SH':
            continue
        registros.append({
            'SH': str(sh_val).strip(),
            'Destinatario': ws.cell(r, c_dest).value if c_dest else '',
            'Description': ws.cell(r, c_desc).value if c_desc else '',
            'Kg': ws.cell(r, c_kg).value if c_kg else None,
            'Item': ws.cell(r, c_item).value if c_item else 1,
            'FOB': ws.cell(r, c_fob).value if c_fob else 0.0
        })
    return registros

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/generar', methods=['POST'])
def generar():
    try:
        file_manifiesto = request.files.get('manifiesto')
        if not file_manifiesto or file_manifiesto.filename == '':
            return "Por favor sube el archivo de manifiesto.", 400

        registros = leer_manifiesto(file_manifiesto)
        hoy = datetime.now().strftime("%d/%m/%Y")
        hoy_slug = datetime.now().strftime("%Y%m%d")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for idx, reg in enumerate(registros, start=1):
                pdf_data = generar_pdf_fiel(reg, hoy)
                clean_sh = reg['SH'].replace('/', '_').replace('\\', '_')
                filename = f"{idx:03d}_{clean_sh}.pdf"
                zf.writestr(filename, pdf_data)

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=f"FACTURAS_PDF_{hoy_slug}.zip",
            mimetype="application/zip"
        )
    except Exception as e:
        return f"Error al generar: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
