from flask import Flask, render_template, request, send_file
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate-quotation', methods=['POST'])
def generate_quotation():
    sections = {
        'entrance': 'Entrance AREA',
        'tv': 'TV Unit Wall Paneling',
        'crockery': 'Crockery and Puja',
        'kitchen': 'KITCHEN',
        'master_bedroom': 'MASTER BEDROOM',
        'children_bedroom': 'CHILDREN\'s BEDROOM',
        'others': 'Others'
    }
    
    # Retrieve dynamic fields from the form
    name = request.form.get('name')
    type_ = request.form.get('type')
    address = request.form.get('address')
    
    data = [['S.No', 'Item', 'Material', 'Length', 'Height', 'Qty(sqt)', 'Price/sft', 'Qty', 'Total']]
    
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    heading_style = styles['Heading2']
    
    grand_total_style = ParagraphStyle(
        'GrandTotalStyle',
        parent=styles['Normal'],
        fontSize=14,
        fontName='Helvetica-Bold',
        alignment=1,  # Center alignment
        spaceAfter=12
    )

    section_colors = [colors.lightblue, colors.lightgreen, colors.lightpink, colors.lightgoldenrodyellow,
                      colors.lightcoral, colors.lightcyan, colors.HexColor('#D3D3D3')]  

    serial_number = 1
    grand_total = 0

    # Add the header paragraph
    header = [
        Paragraph(f'To<br/>{name}<br/>{type_}<br/>{address}<br/><br/>', normal_style),
        Paragraph('Dear Sir/Mam,<br/><br/>This is with reference to our discussions with considering the floor plan of your flat, herewith submitting the Block quote for the interior works as follows:', normal_style)
    ]
    
    for i, (section_key, section_name) in enumerate(sections.items()):
        data.append([Paragraph(section_name, heading_style), '', '', '', '', '', '', '', ''])
        section_color = section_colors[i % len(section_colors)] 

        item_descriptions = request.form.getlist(f'item_description_{section_key}[]')
        material_specifications = request.form.getlist(f'material_specification_{section_key}[]')
        lengths = request.form.getlist(f'length_{section_key}[]')
        heights = request.form.getlist(f'height_{section_key}[]')
        quantities_sqft = request.form.getlist(f'qty_{section_key}[]')
        unit_prices = request.form.getlist(f'unit_price_{section_key}[]')
        quantities = request.form.getlist(f'quantity_{section_key}[]')
        totals = request.form.getlist(f'total_{section_key}[]')

        for j in range(len(item_descriptions)):
            total_value = float(totals[j]) if totals[j] else 0
            grand_total += total_value

            data.append([
                str(serial_number),
                Paragraph(item_descriptions[j], normal_style),
                Paragraph(material_specifications[j], normal_style),
                Paragraph(lengths[j], normal_style),
                Paragraph(heights[j], normal_style),
                Paragraph(quantities_sqft[j], normal_style),
                Paragraph(unit_prices[j], normal_style),
                Paragraph(quantities[j], normal_style),
                Paragraph(f'{total_value:.2f}', normal_style)
            ])
            serial_number += 1

    buffer = BytesIO()
    
    left_margin = 26
    right_margin = 26
    top_margin = 10
    bottom_margin = 10
    doc = SimpleDocTemplate(buffer, pagesize=letter,
                            leftMargin=left_margin, rightMargin=right_margin,
                            topMargin=top_margin, bottomMargin=bottom_margin)

    col_widths = [doc.width * 0.05,
                  doc.width * 0.25,
                  doc.width * 0.2,
                  doc.width * 0.1,
                  doc.width * 0.1,
                  doc.width * 0.1,
                  doc.width * 0.1,
                  doc.width * 0.05,
                  doc.width * 0.1]   

    table = Table(data, colWidths=col_widths, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    for i, row in enumerate(data):
        if isinstance(row[0], Paragraph):
            if row[0].getPlainText() == 'Grand Total':
                style.add('SPAN', (0, i), (8, i))
                style.add('ALIGN', (0, i), (8, i), 'CENTER')
                style.add('VALIGN', (0, i), (8, i), 'MIDDLE')
                style.add('BACKGROUND', (0, i), (8, i), colors.lightgrey)
                style.add('TEXTCOLOR', (0, i), (8, i), colors.black)
                style.add('FONTNAME', (0, i), (8, i), 'Helvetica-Bold')
                style.add('FONTSIZE', (0, i), (8, i), 12)
                style.add('BOTTOMPADDING', (0, i), (8, i), 12)
                style.add('GRID', (0, i), (8, i), 0, colors.transparent)
            else:
                style.add('SPAN', (0, i), (8, i))
                style.add('ALIGN', (0, i), (8, i), 'CENTER')
                style.add('VALIGN', (0, i), (8, i), 'MIDDLE')
                style.add('BACKGROUND', (0, i), (8, i), section_colors[i % len(section_colors)])
                style.add('TEXTCOLOR', (0, i), (8, i), colors.black)

    table.setStyle(style)

    elements = header + [Spacer(1, 12)] + [table]

    grand_total_paragraph = Paragraph(f'<b>Grand Total:</b> {grand_total:.2f}/-', grand_total_style)
    elements.append(Spacer(1, 12))
    elements.append(grand_total_paragraph)

    def draw_page_border(canvas, doc):
        canvas.saveState()
        width, height = doc.pagesize
        canvas.setStrokeColor(colors.black)
        canvas.setLineWidth(2)
        canvas.rect(10, 10, width - 20, height - 20)
        canvas.restoreState()

    doc.build(elements, onFirstPage=draw_page_border, onLaterPages=draw_page_border)

    buffer.seek(0)
    return send_file(buffer, mimetype='application/pdf', as_attachment=True, download_name='quotation.pdf')
    
if __name__ == '__main__':
    app.run(debug=True)
