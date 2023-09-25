from fpdf import FPDF, XPos, YPos
import pandas as pd

import glob
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for file_path in filepaths:
    df= pd.read_excel(file_path, sheet_name='Sheet 1')
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times', size= 14, style='B')
    file_name = Path(file_path).stem
    #getting invoice no. and date from the name of file
    invoice_no, date = file_name.split('-')
    pdf.cell(w=0, h=10, txt= 'Invoice No : '+ invoice_no , align='L',new_x=XPos.LMARGIN, new_y=YPos.NEXT )
    pdf.cell(w=0, h=10, txt= 'DATE: '+ date , align='L',new_x=XPos.LMARGIN, new_y=YPos.NEXT )
    #adding header
    df_columns = df.columns
    # removing '_' from header names and cpitalizing same
    df_columns= [item.replace('_', ' ').title() for item in df_columns ]
    pdf.set_font(family='Times', size= 10, style='B')
    pdf.cell(w=30, h=10, txt= df_columns[0], align='L', border=1 )
    pdf.cell(w=40, h=10, txt= df_columns[1], align='L', border=1)
    pdf.cell(w=30, h=10, txt= df_columns[2], align='L', border=1 )
    pdf.cell(w=25, h=10, txt= df_columns[3], align='L', border=1)
    pdf.cell(w=25, h=10, txt= df_columns[4], align='L', border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT )
    #ading rows with details
    for index, row in df.iterrows():
            pdf.set_font(family='Times', size= 10)
            pdf.cell(w=30, h=10, txt= str(row['product_id']), align='L', border=1 )
            pdf.cell(w=40, h=10, txt= row['product_name'], align='L', border=1)
            pdf.cell(w=30, h=10, txt= str(row['amount_purchased']), align='R', border=1 )
            pdf.cell(w=25, h=10, txt= str(row['price_per_unit']), align='R', border=1)
        # using new_x=XPos.LMARGIN, new_y=YPos.NEXT for breaking into next line
            pdf.cell(w=25, h=10, txt= str(row['total_price']), align='R', border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    #adding total price  
    total_price = df['total_price'].sum()
    pdf.set_font(family='Times', size= 10)
    pdf.cell(w=30, h=10, txt= '', align='L', border=1 )
    pdf.cell(w=40, h=10, txt= '', align='L', border=1)
    pdf.cell(w=30, h=10, txt= '', align='R', border=1 )
    pdf.cell(w=25, h=10, txt= '', align='R', border=1)
    pdf.cell(w=25, h=10, txt= str(total_price), align='R', border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # adding bottom details with total price and company logo
    pdf.set_font(family='Times', size= 10, style='B')

    pdf.cell(w=30, h=10, txt= f'Total Price payable : {total_price}', align='L', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.image('pythonhow.png', w=10)
    pdf.output(f'PDF_invoices/{file_name}.pdf')
    


    
