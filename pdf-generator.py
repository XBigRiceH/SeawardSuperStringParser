from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.pdfencrypt import padding
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
from record_types import *

f = open('./TestResults.sss', 'rb')
test_results: List[TestResult] = []
machine_info = None

# parsing
while True:
    start_byte = f.read(1)
    assert start_byte == b'\x55'

    record_length = f.read(2)
    length_val = struct.unpack('<H', record_length)[0]

    checksum = f.read(2)
    checksum_val = struct.unpack('<H', checksum)[0]

    empty_byte = f.read(2)
    assert empty_byte == b'\x00\x00'

    # I don't know why but sometimes the length will short a little bit
    record_content = f.read(length_val)
    calculated_checksum = (sum(record_content)) & 0xffff
    if calculated_checksum == checksum_val or calculated_checksum == checksum_val + 1:
        pass
    else:
        length_val += 1
        record_content += f.read(1)
        calculated_checksum = (sum(record_content)) & 0xffff
        assert calculated_checksum == checksum_val or calculated_checksum == checksum_val + 1

    record_type = record_content[0]
    if record_type == 0xaa and record_content[1] == 0xff:
        break

    instance = record_type_class_defs[record_type](record_content[1:])
    if type(instance) is MachineInfo:
        machine_info = instance
    else:
        test_results.append(instance)


def to_para(content, style):
    result = []
    for i in content:
        result.append(Paragraph(i, style))
    return result


def create_pat_report():
    # 1. 设置文档
    filename = "pat_report_template.pdf"
    doc = SimpleDocTemplate(filename, pagesize=landscape(A4),
                            rightMargin=1 * cm, leftMargin=1 * cm,
                            topMargin=1 * cm, bottomMargin=1 * cm)
    elements = []
    styles = getSampleStyleSheet()

    style_section_header = ParagraphStyle(name='SectionHeader', parent=styles['Normal'], fontSize=9,
                                          textColor=colors.white, fontName='Helvetica-Bold')
    style_normal = ParagraphStyle(name='NormalText', parent=styles['Normal'], fontSize=8, fontName='Helvetica')
    highlight_size = ParagraphStyle(name='normal_size_bold', parent=styles['Normal'], fontSize=10, fontName='Helvetica')

    dark_blue = colors.Color(0.1, 0.2, 0.4)
    light_grey = colors.Color(0.95, 0.95, 0.95)

    header_row_style = [
        ('BACKGROUND', (0, 0), (-1, 0), dark_blue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ]

    elements.append(Paragraph("Portable Appliance Test (PAT) Report", ParagraphStyle(
        name='HeaderTitle',
        parent=styles['Normal'],
        fontSize=16,
        fontName='Helvetica-Bold',
        leftIndent=-0.25 * cm
    )))
    elements.append(Spacer(1, 1 * cm))

    # %% tec info
    tec_info_table_content = [
        [Paragraph("TESTING CARRIED OUT BY", style_section_header)],
        [Paragraph("<b>TEC PA & Lighting</b>", highlight_size)],
        [Paragraph("""<b>Email:</b> info@nottinghamtec.co.uk<br/>
        <b>Website:</b> www.nottinghamtec.co.uk<br/>
        <b>Tel:</b> 0115 84 68720<br/>
        <b>Address:</b><br/>Portland Building<br/>University Park<br/>Nottingham<br/>NG7 2RD<br/>""", style_normal)]
    ]
    tec_info_table = Table(tec_info_table_content, colWidths=[doc.width / 2 - 1 * cm], )
    tec_info_table.setStyle(TableStyle(header_row_style + [
        ('BOX', (0, 0), (-1, -1), 1, colors.grey),
    ]))

    tec_logo = Image("imgs/teclogo.jpg")
    w, h = tec_info_table.wrap(doc.width / 2 - 1 * cm, doc.height)
    target_height = h * 0.8
    aspect_ratio = tec_logo.imageWidth / tec_logo.imageHeight
    tec_logo.drawHeight = target_height
    tec_logo.drawWidth = target_height * aspect_ratio

    tec_info = Table(
        [[tec_info_table, tec_logo]],
        colWidths=[doc.width / 2, doc.width / 2],
        hAlign='CENTER'
    )
    tec_info.setStyle(TableStyle([
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('VALIGN', (0, 0), (0, 0), 'TOP'),
        ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
    ]))

    elements.append(tec_info)
    elements.append(Spacer(1, 0.5 * cm))

    # %% tester info
    tester_info_content = [
        [Paragraph("PAT TESTER INFO", style_section_header), ""],
        ["Serial Number", "Make and Model"],
        ["44G-0452", "Seaward - Apollo 600"],
    ]
    tester_info_table = Table(tester_info_content, colWidths=[doc.width / 2, doc.width / 2])
    tester_info_table.setStyle(TableStyle(
        header_row_style +
        [
            ('SPAN', (0, 0), (1, 0))
        ] +
        [
            ('GRID', (0, 1), (-1, -1), 0.5, colors.lightgrey),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BACKGROUND', (0, 1), (-1, 1), light_grey),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
        ] + [
            ('BOX', (0, 0), (-1, -1), 1, colors.grey),
        ]
    ))
    elements.append(tester_info_table)
    elements.append(Spacer(1, 0.5 * cm))

    # %% test result table header
    result_header_content = [
        [Paragraph("APPLIANCE DETAILS AND TEST RESULTS", style_section_header), "", "", "", "", "", '', ''],
        [Paragraph("<b>Key</b><br/>P = PASS / F = FAIL / I = INFO / S = SKIP (OR N/A)", style_normal), "", "",
         "", "", "", "", "", "", '', ''],
        ['Appliance ID', 'Appliance Description', 'Test Date', 'Operator', 'Program', 'Test Items', '', "", "",
         "Overall Status", 'Comments'],
        ['', '', '', '', '', 'Test Type', 'Result', "Unit", "Status", "", ''],
    ]
    result_col_widths = [
        3 * cm,
        6 * cm,
        2 * cm,
        2 * cm,
        2.5 * cm,
        3 * cm, 1.5 * cm, 1.25 * cm, 1.5 * cm,
        2.5 * cm,
        2.45 * cm
    ]
    result_header_table = Table(result_header_content, colWidths=result_col_widths)
    result_header_table.setStyle(TableStyle(
        header_row_style +
        [
            ('BACKGROUND', (0, 1), (-1, 3), light_grey),
            ('SPAN', (0, 0), (-1, 0)),
            ('SPAN', (0, 1), (-1, 1)),
            ('FONTSIZE', (0, 2), (-1, -1), 8),
            ('GRID', (0, 1), (-1, -1), 0.5, colors.lightgrey),
            ('SPAN', (5, 2), (8, 2)),

            ('SPAN', (0, 2), (0, 3)),
            ('SPAN', (1, 2), (1, 3)),
            ('SPAN', (2, 2), (2, 3)),
            ('SPAN', (3, 2), (3, 3)),
            ('SPAN', (4, 2), (4, 3)),
            ('SPAN', (9, 2), (9, 3)),
            ('SPAN', (10, 2), (10, 3)),
            ('VALIGN', (0, 2), (-1, 3), 'MIDDLE'),
            ('FONTNAME', (0, 2), (-1, 3), 'Helvetica-Bold'),
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.grey),
            ('LINEBEFORE', (0, 0), (0, -1), 1, colors.grey),
            ('LINEAFTER', (10, 0), (10, -1), 1, colors.grey),
        ]
    ))
    elements.append(result_header_table)

    # %% data
    result_content = []
    result_style = []

    for record in test_results:
        formatted_results = []
        for visual_test_result in record.visual_test_results:
            formatted_results.append([
                visual_test_result.name,
                visual_test_result.result if visual_test_result.unit else "",
                visual_test_result.unit,
                visual_test_result.flags[0]
            ])
        for physical_test_result in record.physical_test_results:
            if isinstance(physical_test_result, EarthResistanceTestResult):
                formatted_results.append([
                    "Earth Continuity",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, IECLeadContinuityTestResult):
                formatted_results.append([
                    "IEC Lead Continuity",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, PointToPointTestResult):
                formatted_results.append([
                    "Point To Point Resistance",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, InsulationTestResult):
                formatted_results.append([
                    "Insulation",
                    physical_test_result.get_value(),
                    physical_test_result.resistance.unit,
                    physical_test_result.get_status()
                ])
                formatted_results.append([
                    "Insulation Voltage",
                    physical_test_result.voltage.value,
                    physical_test_result.voltage.unit,
                    "INFO"
                ])
            elif isinstance(physical_test_result, SubstituteLeakageTestResult):
                formatted_results.append([
                    "Substitute Leakage Current",
                    physical_test_result.get_value(),
                    physical_test_result.current.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, PolarityTestResult):
                formatted_results.append([
                    "IEC Lead Polarity",
                    physical_test_result.get_value(),
                    "",
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, MainVoltageTestResult):
                formatted_results.append([
                    "Main Voltage",
                    physical_test_result.get_value(),
                    physical_test_result.voltage.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, TouchOrLeakageCurrentTestResult):
                formatted_results.append([
                    "Touch Or Leakage Test Load Current",
                    physical_test_result.load_current.value,
                    physical_test_result.load_current.unit,
                    physical_test_result.get_status()
                ])
                formatted_results.append([
                    "Touch Or Leakage Test Leakage Current",
                    physical_test_result.leakage_current.value,
                    physical_test_result.leakage_current.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, RCDTestResult):
                formatted_results.append([
                    "RCD Test Current",
                    physical_test_result.test_current.value,
                    physical_test_result.test_current.unit,
                    "INFO"
                ])
                formatted_results.append([
                    "RCD Test Circle Angle",
                    physical_test_result.circle_angle.value,
                    physical_test_result.circle_angle.unit,
                    "INFO"
                ])
                formatted_results.append([
                    "RCD Test Trip time",
                    physical_test_result.get_value(),
                    physical_test_result.trip_time.unit,
                    physical_test_result.get_status()
                ])
            elif isinstance(physical_test_result, StringComment):
                formatted_results.append([
                    physical_test_result.string_value,
                    "",
                    "",
                    physical_test_result.get_status()
                ])

        for i, tr in enumerate(formatted_results):
            tr[0] = str(tr[0])
            tr[1] = str(tr[1])
            tr[2] = str(tr[2])
            tr[3] = str(tr[3])

            row = []
            if i == 0:
                row.extend([
                    record.asset_id,
                    '',
                    record.test_time.strftime("%d/%m/%Y"),
                    record.test_operator,
                    record.program,
                ])
                row.extend(tr)
                row.extend([
                    record.get_status(),
                    record.comments
                ])
            else:
                row.extend([""] * 5)
                row.extend(tr)
                row.extend([""] * 2)
            result_content.append(to_para(row, style_normal))

    result_table = Table(result_content, colWidths=result_col_widths)
    result_table.setStyle(TableStyle(
        header_row_style +
        [
            ('SPAN', (0, 0), (-1, 0)),
            ('GRID', (0, 1), (-1, -1), 0.5, colors.lightgrey),
            ('LINEBELOW', (0, -1), (-1, -1), 1, colors.grey),
            ('LINEBEFORE', (0, 0), (0, -1), 1, colors.grey),
            ('LINEAFTER', (10, 0), (10, -1), 1, colors.grey),
        ]
    ))
    elements.append(result_table)

    doc.build(elements)
    print(f"PDF Generated successfully: {filename}")


if __name__ == "__main__":
    create_pat_report()
