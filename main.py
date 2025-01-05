# = Automation program
from pathlib import Path
import pandas as pd
import glob
from fpdf import FPDF


filepaths = glob.glob('invoices/*.xlsx')
for filepath in filepaths:
    df = pd.read_excel(filepath)
    # print(df)

    filename = Path(filepath).stem
    invoice_nr, invoice_date  = filename.split('-')

    # Create PDF Output
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Header
    pdf.image('pille.png', w=30, x=170, y=10)
    pdf.set_font(family='Times', style='B', size=18)
    pdf.cell(w=0, h=12, text=f'Invoice nr. {invoice_nr}', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(w=0, h=12, text=f'Date {invoice_date}', align='L', new_x='LMARGIN', new_y='NEXT')

    # Body
    pdf.set_font(family='Times', size=12)
    pdf.set_y(65)
    total = 0
    with pdf.table(col_widths=(30,70,20,20,20), text_align=("LEFT", "LEFT", "LEFT", "LEFT", "RIGHT")) as table:
        headings = list(df.columns)
        head_row = table.row()
        for head in headings:
            header = str(head).replace('_', ' ')
            head_row.cell(header.title())

        for index, row in df.iterrows():
            new_row = table.row()
            total = total + int(row['total_price'])
            for entry in row:
                new_row.cell(str(entry))

        total_row = table.row()
        total_row.cell("", colspan=4)
        pdf.set_font(family='Times', size=12, style="B")
        total_row.cell(str(total))

    # Footer
    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=0, h=10, text='', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(w=0, h=10, text=f'The total due amount is {total} euro.', align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(w=0, h=10, text=f'PayPill Company', align='L')

    pdf.output(f"output/INV-{invoice_nr}.pdf")









