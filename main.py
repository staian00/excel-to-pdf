import pandas as pd
from glob import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob('Invoices/*xlsx')

for filepath in filepaths:
    # Resetting instances, setting variables
    pdf = FPDF(format='A4', orientation='L', unit='mm', )
    x1 = 10
    y1 = 30
    dx = 50
    df = pd.read_excel(filepath)
    # Creating dictionary for heading
    l1, l2 = df.columns, [i.replace('_', ' ').title() for i in df.columns]
    d = dict(zip(l1, l2))
    # Creating page, creating filename
    pdf.add_page()
    filename = Path(filepath).stem
    pdf.set_font(family='Times', style='B', size=20)
    pdf.cell(w=60, h=8, txt=f'Invoice nr.{filename}', ln=1)
    for col in df.columns:
        pdf.set_xy(x1, y1)
        pdf.set_font(family='Times', style='B', size=11)
        pdf.cell(w=dx, h=8, txt=str(d[str(col)]), ln=1, border=1, align='R')
        pdf.set_font(family='Times', style='', size=11)
        for i in range(df.shape[0]):
            pdf.set_x(x1)
            pdf.cell(w=dx, h=8, txt=str(df.iloc[i][col]), ln=1, border=1, align='R')
        if col == 'total_price':
            s = sum(df[:][col])
            pdf.set_x(x1)
            pdf.set_font(family='Times', style='B', size=11)
            pdf.cell(w=dx, h=8, txt=f'SUM = {s}', ln=1, border=1, align='R')

        x1 += dx

    pdf.output(f'PDF/{filename}.pdf')
