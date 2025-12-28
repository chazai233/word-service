from docx import Document
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.paragraph.paragraph import Paragraph
import os

doc_path = r"d:\Projects\Dify\word_service\test_output_en.docx"

if os.path.exists(doc_path):
    doc = Document(doc_path)
    
    print("--- Document Structure ---")
    body = doc._body
    for i, element in enumerate(body):
        if isinstance(element, Paragraph):
            print(f"[{i}] Paragraph: {element.text[:50]}")
        elif isinstance(element, CT_Tbl):
            table = Table(element, doc)
            print(f"[{i}] Table: {len(table.rows)} rows, {len(table.columns)} columns")
            if len(table.rows) > 0:
                print(f"    First row: {[c.text.strip() for c in table.rows[0].cells]}")
else:
    print(f"File not found: {doc_path}")
