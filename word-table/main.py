from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT


def generate_table():
    context = {
        "col_labels": ["fruit", "vegetable", "stone", "thing", "hhe"],
        "tbl_contents": [
            {"label": "yellow", "cols": ["banana", "capsicum", ["10", "pyrite"], "taxi", "zhongli"]},
            {"label": "red", "cols": ["apple", "tomato", ["10", "cinnabar"], "doubledecker", "jinjan"]},
            {"label": "green", "cols": ["apple", "tomato", ["10", "cinnabar"], "doubledecker", "jinjan"]},
        ],
    }
    doc = Document()
    style = doc.styles['Normal']
    font_style = doc.styles['Normal'].font
    font_style.size = Pt(13)
    font_style.name = 'Times New Roman'
    style.paragraph_format.line_spacing = 1.2  #WD_LINE_SPACING.SINGLE
    # style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    table = doc.add_table(
        rows=2 + len(context["tbl_contents"]),
        cols=1 + len(context["col_labels"])
    )
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    hdr_cells_2 = table.rows[1].cells
    hdr_cells[0].text = "Color of think"
    hdr_cells[0].width = Inches(1.5)
    hdr_cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    # table.rows[1].cells[0].width = Inches(1.5)
    hdr_cells[0].merge(hdr_cells_2[0])


    for i in range(0, len(context["col_labels"])):
        hdr_cells[1].merge(hdr_cells[i+1])
        hdr_cells_2[i+1].text = context["col_labels"][i]
        hdr_cells_2[i+1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    hdr_cells[1].text = "Type of thing"
    hdr_cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # doc.render(context)
    doc.save("generated_doc.docx")


if __name__ == '__main__':
    generate_table()
