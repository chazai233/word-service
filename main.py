"""
Dify/n8n Word Document Service - Ultimate Version
功能：模板填充、智能缩进、日期天气更新、人员统计、附表填报
"""
import base64
import io
import json
import re
import os
from datetime import datetime, timedelta
from typing import List, Optional, Dict, Any
from enum import Enum
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

app = FastAPI(title="Word Service Ultimate", version="4.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- 辅助函数 ---

def set_cell_text(cell, text):
    """设置单元格字体：中文宋体，英文Times New Roman，五号"""
    cell.text = ""
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = p.add_run(str(text))
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)

def apply_smart_indentation(doc: Document):
    """
    智能首行缩进（增强版）
    忽略 Markdown 符号(*)、序号(1.) 和空格，精准匹配关键词
    """
    keywords = ["人员投入", "设备投入", "累计工程量", "人员：", "设备："]
    indent_size = Pt(24)
    
    def process_paragraph(p):
        # 清洗文本：去除 *、空格、制表符
        clean_text = re.sub(r"[\*\s\t]", "", p.text)
        for kw in keywords:
            if clean_text.startswith(kw):
                p.paragraph_format.first_line_indent = indent_size
                break

    for p in doc.paragraphs:
        process_paragraph(p)
        
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_paragraph(p)

def find_target_table(doc: Document, index: int) -> Optional[Any]:
    if 0 <= index < len(doc.tables):
        return doc.tables[index]
    return None

def update_table_row(table, row_name: str, today: str, total: str):
    """通用表格行更新逻辑"""
    if not table.rows: return
    
    # 简单的列推断：假设第2列是项目名，倒数第3列今日，倒数第2列累计
    # 你可以根据实际模板调整这些索引！
    name_col = 1
    today_col = len(table.rows[0].cells) - 3
    total_col = len(table.rows[0].cells) - 2
    
    for row in table.rows:
        if len(row.cells) <= max(name_col, today_col, total_col): continue
        cell_text = row.cells[name_col].text.strip()
        if row_name in cell_text: # 模糊匹配
            if today and today != "-":
                set_cell_text(row.cells[today_col], today)
            if total and total != "-":
                set_cell_text(row.cells[total_col], total)
            return

# --- 请求模型 ---

class FillTemplateRequest(BaseModel):
    template_base64: str
    content: str
    table_index: int = 0
    row_index: int = 4
    col_index: int = 2
    # 兼容字段
    update_date_weather: bool = False
    upload_to_feishu: bool = False
    feishu_token: Optional[str] = None

class UpdateDateWeatherRequest(BaseModel):
    document_base64: str
    feishu_token: Optional[str] = None

class UpdatePersonnelRequest(BaseModel):
    document_base64: str
    personnel_text: str # "人员投入详情...\n共计XX人"
    feishu_token: Optional[str] = None

class AppendixTableData(BaseModel):
    table_index: int
    row_name: str
    today_qty: str
    total_qty: str

class UpdateAppendixRequest(BaseModel):
    document_base64: str
    data: List[AppendixTableData]
    feishu_token: Optional[str] = None

# --- 接口定义 ---

@app.post("/fill-template")
async def fill_template(req: FillTemplateRequest):
    try:
        file_bytes = base64.b64decode(req.template_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        # 填充主内容
        if doc.tables and len(doc.tables) > req.table_index:
            table = doc.tables[req.table_index]
            if len(table.rows) > req.row_index:
                cell = table.cell(req.row_index, req.col_index)
                cell.text = req.content
                apply_smart_indentation(doc)
        
        out = io.BytesIO()
        doc.save(out)
        return {"success": True, "document_base64": base64.b64encode(out.getvalue()).decode()}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.post("/update-date-weather")
async def update_date_weather(req: UpdateDateWeatherRequest):
    try:
        file_bytes = base64.b64decode(req.document_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        now = datetime.now()
        date_str = f"{now.year}年{now.month}月{now.day}日"
        weather_str = "天气：晴                气温：20℃~30℃" # 示例，可接真实API
        
        if doc.tables:
            table = doc.tables[0]
            if len(table.rows) > 0:
                cells = table.rows[0].cells
                set_cell_text(cells[0], date_str) # 左上角日期
                set_cell_text(cells[-1], weather_str) # 右上角天气
        
        out = io.BytesIO()
        doc.save(out)
        return {"success": True, "document_base64": base64.b64encode(out.getvalue()).decode()}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.post("/update-personnel-stats")
async def update_personnel_stats(req: UpdatePersonnelRequest):
    try:
        file_bytes = base64.b64decode(req.document_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        # 寻找“人员投入详情”应该填入的位置
        # 逻辑：在主表格下方寻找包含“人员投入”字样的单元格进行追加或替换
        # 这里假设它填入主表格(Index 0)的最后一行，或者根据你的模板定制
        # 临时方案：直接在文档末尾追加一段，或者你可以指定格子
        doc.add_paragraph("\n【自动统计】\n" + req.personnel_text)
        
        out = io.BytesIO()
        doc.save(out)
        return {"success": True, "document_base64": base64.b64encode(out.getvalue()).decode()}
    except Exception as e:
        return {"success": False, "message": str(e)}

@app.post("/update-appendix-tables")
async def update_appendix_tables(req: UpdateAppendixRequest):
    try:
        file_bytes = base64.b64decode(req.document_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        for item in req.data:
            target_table = find_target_table(doc, item.table_index)
            if target_table:
                update_table_row(target_table, item.row_name, item.today_qty, item.total_qty)
        
        out = io.BytesIO()
        doc.save(out)
        return {"success": True, "document_base64": base64.b64encode(out.getvalue()).decode()}
    except Exception as e:
        return {"success": False, "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
