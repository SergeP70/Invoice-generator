# = Automation program
from pathlib import Path
import pandas as pd
import glob
from fpdf import FPDF


filepaths = glob.glob('invoices/*.xlsx')
for filepath in filepaths:
    df = pd.read_excel(filepath)

    filename = Path(filepath).stem
    invoice_nr = filename.split('-')[0]
    invoice_date = filename.split('-')[1]
    print(invoice_nr, invoice_date)
    """
    for index, line in df.iterrows():
        product_id = str(df['product_id'])
        product_name = str(['product_name'])
        amount = df['amount_purchased']
        product_price = df['price_per_unit']
        total_price = df['total_price']
        print(product_id, " ", product_name) # + " " + amount + " x " + product_price + "=" + total_price) """



    # Create PDF Output
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Header
    pdf.set_font(family='Times', style='B', size=18)
    pdf.cell(w=0, h=12, text='Invoice nr. ' + invoice_nr, align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.cell(w=0, h=12, text='Date ' + invoice_date, align='L', new_x='LMARGIN', new_y='NEXT')
    pdf.line(10, 40, 200, 40)

    # Body
    # Footer

    # print(invoice_nr, " ", invoice_date)
    pdf.output(f"output/INV-{invoice_nr}.pdf")









