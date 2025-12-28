import docx
from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
import os

doc_path = r"d:\Projects\Dify\word_service\test_output_en.docx"

if os.path.exists(doc_path):
    doc = Document(doc_path)
    
    print("--- Document Structure ---")
    for i, element in enumerate(doc.element.body):
        if isinstance(element, CT_P):
            p = Paragraph(element, doc)
            print(f"[{i}] Paragraph: {p.text[:100]}")
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            print(f"[{i}] Table: {len(table.rows)} rows, {len(table.columns)} columns")
            if len(table.rows) > 0:
                print(f"    First row: {[c.text.strip().replace('\n', ' ') for c in table.rows[0].cells]}")
else:
    print(f"File not found: {doc_path}")
