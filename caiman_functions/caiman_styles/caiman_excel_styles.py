# Hay que meter las fórmulas en inglés

import openpyxl
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties, RegularTextRun
from openpyxl.styles import Alignment, Border, Font, Side


def automatic_column_widths(sheet):
    # automatic_column_widths function calculates de length of every value in a column and shet the witdh column to the maximun lenght.
    dims = {}
    for row in sheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max(
                    (dims.get(cell.column_letter, 0), len(str(cell.value))))
    for col, value in dims.items():
        sheet.column_dimensions[col].width = value


def automatic_no_borders_cell(wb, sheet):
    # automatic_no_borders_cell function remove the borders of all the cells in a worksheet.
    no_fill = openpyxl.styles.PatternFill(fill_type=None)
    side = openpyxl.styles.Side(border_style=None)
    no_border = openpyxl.styles.borders.Border(
        left=side,
        right=side,
        top=side,
        bottom=side,
    )

    # Loop through all cells in all worksheets
    for sheet in wb.worksheets:
        for row in sheet:
            for cell in row:
                # Apply colorless and borderless styles
                cell.fill = no_fill
                cell.border = no_border


def style_oscillation_calculated_parameters_header(cell, sheet):
    font_title = Font(color="FF0000", bold=True)  # Red
    sheet[cell].font = font_title
    sheet[cell].number_format = '0.0000'
    header_position = sheet[cell]
    header_position.alignment = Alignment(
        horizontal="right", vertical="center", wrapText=False, indent=1)


def style_average_values(cell, sheet):
    font_title = Font(color="FF0000", bold=True)  # Red
    sheet[cell].font = font_title
    sheet[cell].number_format = '0.0000'
    header_position = sheet[cell]
    header_position.alignment = Alignment(
        horizontal="center", vertical="center", wrapText=True)


def style_headers(cell, sheet):
    font_title = Font(bold=True)
    sheet[cell].font = font_title
    header_position = sheet[cell]
    header_position.alignment = Alignment(
        horizontal="center", vertical="center", wrapText=True)


def style_time(cell, sheet):
    font_title = Font(bold=False)
    sheet[cell].font = font_title
    header_position = sheet[cell]
    header_position.alignment = Alignment(
        horizontal="center", vertical="center", wrapText=True)


def style_number(cell, sheet):
    font_title = Font(bold=False)
    sheet[cell].font = font_title
    sheet[cell].number_format = '0.0000'
    header_position = sheet[cell]
    header_position.alignment = Alignment(
        horizontal="center", vertical="center")


def cell_border_style(cell, sheet):
    thin_border = Border(left=Side(style=None),
                         right=Side(style=None),
                         top=Side(style=None),
                         bottom=Side(style=None))
    font_title = Font(bold=False)
    sheet[cell].font = font_title
    sheet[cell].number_format = '0.0000'
    sheet[cell].border = thin_border
    header_position = sheet[cell]
    header_position.alignment = Alignment(
        horizontal="center", vertical="center")


def style_chart(text, chart):
    # Style chart title
    font = openpyxl.drawing.text.Font(typeface='Calibri')
    size = 1100  # 14 point size

    # chart.title = Title()
    # paraprops = ParagraphProperties()
    # paraprops.defRPr = CharacterProperties(latin=font, sz=size, b=False)
    # paras = Paragraph(pPr=paraprops, endParaRPr=paraprops.defRPr)

    # X and Y axes numbers
    cp = CharacterProperties(latin=font, sz=size, b=False)  # Not bold
    pp = ParagraphProperties(defRPr=cp)
    rtp = RichText(p=[Paragraph(pPr=pp, endParaRPr=cp)])
    chart.x_axis.txPr = rtp        # Works!
    chart.y_axis.txPr = rtp        # Works!

    # X and Y axes titles
    chart.x_axis.title.tx.rich.p[0].pPr = pp       # Works!
    chart.y_axis.title.tx.rich.p[0].pPr = pp       # Works!
    # chart.title.tx.rich.paragraphs = paras
    # cp = openpyxl.drawing.text.CharacterProperties(latin=font_test, sz=1200)
    # chart.x_axis.txPr = RichText(
    # p=[openpyxl.drawing.text.Paragraph(pPr=openpyxl.drawing.text.ParagraphProperties(defRPr=cp), endParaRPr=cp)])
