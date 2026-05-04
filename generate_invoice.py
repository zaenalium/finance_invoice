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

    log_success = []
    for inv in tqdm(df.invoice_no.unique(), total = df.invoice_no.nunique()):
        try:
            data = df[df['invoice_no'] == inv].fillna('').to_dict(orient = 'records')
            
            dt_inv = pd.to_datetime(data[0]['invoice_date']).strftime('%d/%m/%Y')

            template_path = os.path.join(os.path.dirname(__file__), 'Finance Invoice template.docx')
            f = open(template_path, 'rb')
            
            doc = Document(f)
            
            # Dynamic rows — line items
            subtotal = round(sum([int(x.get('amount', 0)) for x in data]))   
            vat_ori = round(sum([int(x.get('vat', 0)) for x in data]))
            total = f"Rp. {(round(subtotal + vat_ori)):,}".replace('.0', '')
            
            subtotal = f"Rp. {round(subtotal):,}".replace('.0', '')
            vat =f"Rp. {round(vat_ori):,}".replace('.0', '')
                
            inv_no = data[0].get('invoice_no')
            # Static fields

            attention = data[0].get('attention', None)

            if not attention:
                addess_company = [(14, 0, data[0].get('company_name')),
                    (15, 0, data[0].get('address_1')),
                    (16, 0, data[0].get('address_2')),
                    (17, 0, data[0].get('address_3')),
                    (18, 0, data[0].get('address_4')),
                    (19, 0, data[0].get('address_5'))]
            else:
                addess_company = [(14, 0, data[0].get('attention')),
                    (15, 0, data[0].get('company_name')),
                    (16, 0, data[0].get('address_1')),
                    (17, 0, data[0].get('address_2')),
                    (18, 0, data[0].get('address_3')),
                    (19, 0, data[0].get('address_4'))]

            static_cells = [
                (1, 4, dt_inv),
                (2, 4, inv_no),
                (3, 4, data[0].get('po')),
                # (27, 4, str(subtotal).replace('.0', '')),
                # (28, 4, str(vat).replace('.0', '')),
                # (29, 4, str(total).replace('.0', '')),  # fixed: was overwriting row 28 twice
                (32, 0, data[0].get('payment_notes')),
                (36, 0, 'Account Name           : ' + str(data[0].get('bank_account_name', '')).replace('.0', '')),
                (37, 0, 'Bank                          : ' + str(data[0].get('bank', ''))),
                (38, 0, 'Account Number        : ' + str(data[0].get('bank_account_number', ''))),
                (39, 0, 'Branch                       : ' + str(data[0].get('bank_branch', ''))),
                (40, 0, 'Swift Code                 : ' + str(data[0].get('swift_code', ''))),  # fixed: was overwriting row 39
            ] + addess_company
            
            for row, col, value in static_cells:
                set_cell_text(doc.tables[0].cell(row, col), value)
            
            # Bold cells — subtotal, vat, total
            if vat_ori == 0:
                set_cell_text(doc.tables[0].cell(27, 4), str(round(total)).replace('.0', ''),    bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(27, 3), "Total",    bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
            else:
                set_cell_text(doc.tables[0].cell(27, 4), str(round(subtotal)).replace('.0', ''), bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(28, 4), str(round(vat)).replace('.0', ''),      bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(29, 4), str(round(total)).replace('.0', ''),    bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)

                set_cell_text(doc.tables[0].cell(27, 3), "Subtotal", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(28, 3), "VAT Total",      bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(29, 3), "Total",    bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
                
            for j, item in enumerate(data):
                set_cell_text(doc.tables[0].cell(22 + j, 0), item.get('description'), align=WD_ALIGN_PARAGRAPH.LEFT)
                set_cell_text(doc.tables[0].cell(22 + j, 2), f"{round(float(item.get('price_qty', ''))):,}".replace('.0', ''), align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(22 + j, 3), str(round(float(item.get('qty', '')))).replace('.0', ''), align=WD_ALIGN_PARAGRAPH.RIGHT)
                set_cell_text(doc.tables[0].cell(22 + j, 4), f"{round(float(item.get('amount', ''))):,}".replace('.0', ''), align=WD_ALIGN_PARAGRAPH.RIGHT)
            
            
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
            log_success.append({'invoice_no': inv, 'status': 'Success'})
        except Exception as e:
            log_success.append({'invoice_no': inv, 'status': 'Failed', 'error': str(e).replace('\n', '  ')})
    df_log = pd.DataFrame(log_success)
    df_log.to_excel('output/log.xlsx', index=False)