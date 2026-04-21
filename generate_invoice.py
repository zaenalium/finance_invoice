from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import subprocess
from docx.shared import Pt
import pandas as pd
import numpy as np
import os
import shutil
from tqdm import tqdm
import glob
from zipfile import ZipFile 
import concurrent
pd.options.mode.chained_assignment = None 
import json
from docx.shared import RGBColor


from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tqdm import tqdm

def set_cell_text(cell, text, font_size=9, font_name='Arial', font_color=None, bold=False, align=None):
    """Set cell text with font formatting."""
    cell.text = (text or '').strip()
    for paragraph in cell.paragraphs:
        pPr = paragraph._p.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:after'), '0')
        spacing.set(qn('w:before'), '0')
        pPr.append(spacing)

        if align:
            paragraph.alignment = align  # 👈

        for run in paragraph.runs:
            run.font.size = Pt(font_size)
            run.font.name = font_name
            run.font.bold = bold
            if font_color:
                run.font.color.rgb = font_color


def generate_from_excel(file_path):
    df = pd.read_excel(file_path)
    df['price_qty'] = df['price_qty'].round(0)
    df['vat']= df['price_qty'].round(0)
    df['amount']= df['price_qty'].round(0)
    df['invoice_date'] = pd.to_datetime(df['invoice_date']).dt.strftime('%d/%m/%Y')

    for i in tqdm(df.invoice_no.unique(), total = df.invoice_no.nunique()):
        data = df[df['invoice_no'] == i].fillna('').to_dict(orient = 'records')
        f = open('Finance Invoice template.docx', 'rb')
        
        
                        
        doc = Document(f)
        
        # Dynamic rows — line items
        subtotal = sum([x.get('amount', 0) for x in data])
        vat = sum([x.get('vat', 0) for x in data])
        total = f"Rp. {(round(subtotal + vat)):,}".replace('.0', '')
        
        subtotal = f"Rp. {round(subtotal):,}".replace('.0', '')
        vat =f"Rp. {round(vat):,}".replace('.0', '')
        
            
        inv_no = data[0].get('invoice_no')
        # Static fields
        static_cells = [
            (1, 4, data[0].get('invoice_date')),
            (2, 4, inv_no),
            (3, 4, data[0].get('po')),
            (14, 0, data[0].get('company_name')),
            (15, 0, data[0].get('address_1')),
            (16, 0, data[0].get('address_2')),
            (17, 0, data[0].get('address_3')),
            (18, 0, data[0].get('address_4')),
            (19, 0, data[0].get('address_5')),
            # (27, 4, str(subtotal).replace('.0', '')),
            # (28, 4, str(vat).replace('.0', '')),
            # (29, 4, str(total).replace('.0', '')),  # fixed: was overwriting row 28 twice
            (32, 0, data[0].get('payment_notes')),
            (36, 0, 'Account Name           : ' + str(data[0].get('bank_account_name', '')).replace('.0', '')),
            (37, 0, 'Bank                          : ' + str(data[0].get('bank', ''))),
            (38, 0, 'Account Number        : ' + str(data[0].get('bank_account_number', ''))),
            (39, 0, 'Branch                       : ' + str(data[0].get('bank_branch', ''))),
            (40, 0, 'Swift Code                 : ' + str(data[0].get('swift_code', ''))),  # fixed: was overwriting row 39
        ]
        
        for row, col, value in static_cells:
            set_cell_text(doc.tables[0].cell(row, col), value)
        
        # Bold cells — subtotal, vat, total
        set_cell_text(doc.tables[0].cell(27, 4), str(subtotal).replace('.0', ''), bold=True)
        set_cell_text(doc.tables[0].cell(28, 4), str(vat).replace('.0', ''),      bold=True)
        set_cell_text(doc.tables[0].cell(29, 4), str(total).replace('.0', ''),    bold=True)
        
        set_cell_text(doc.tables[0].cell(27, 3), "Subtotal", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        set_cell_text(doc.tables[0].cell(28, 3), "VAT Total",      bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        set_cell_text(doc.tables[0].cell(29, 3), "Total",    bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
        
        for i, item in enumerate(data):
            set_cell_text(doc.tables[0].cell(22 + i, 0), item.get('description'), align=WD_ALIGN_PARAGRAPH.LEFT)
            set_cell_text(doc.tables[0].cell(22 + i, 2), f"{item.get('price_qty', ''):,}".replace('.0', ''), align=WD_ALIGN_PARAGRAPH.RIGHT)
            set_cell_text(doc.tables[0].cell(22 + i, 3), str(item.get('qty', '')).replace('.0', ''), align=WD_ALIGN_PARAGRAPH.RIGHT)
            set_cell_text(doc.tables[0].cell(22 + i, 4), f"{float(item.get('amount', '')):,}".replace('.0', ''), align=WD_ALIGN_PARAGRAPH.RIGHT)
        
        
        if not os.path.exists(f'output'):
            os.makedirs(f'output')
        
        tmp_name = f'output/Invoice_{inv_no}.docx'
        doc.save(tmp_name)
        f.close()
        subprocess.run(['libreoffice', '--convert-to', 'pdf' ,
                            tmp_name, '--outdir', 'output']
                        ,stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL
                        )
            
        os.remove(tmp_name)