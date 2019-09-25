#Column Widths
from docx.shared import Cm, Inches

def set_col_widths(table):
    widths = [Inches(0.61), Inches(1.6), Inches(.71), Inches(3.34)]
    for row in table.rows:
        row.height = Inches(0.1)
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
