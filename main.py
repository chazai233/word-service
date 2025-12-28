"""
Word Document Generation Service for Dify (Ultimate Version - Format Fixed)
修复：
1. "人员投入"、"累计工程量"、"设备投入" 强制加粗。
2. 上述行 + 以 (1)/(2) 开头的具体内容行，强制首行缩进 2 字符。
"""

import base64
import io
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
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = FastAPI(title="Word Service Ultimate", version="7.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- 辅助函数：格式核心 ----------------

def set_cell_text(cell, text, bold=False):
    """
    设置单元格文本，保留基础字体格式
    """
    cell.text = "" 
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    
    run = p.add_run(str(text))
    
    # 字体设置
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)
    
    if bold:
        run.bold = True

def apply_smart_indentation(doc: Document):
    """
    智能首行缩进逻辑（精准匹配用户要求）
    """
    # 2字符缩进 (小四/五号字通常 24pt 左右视作2字符)
    indent_size = Pt(24) 
    
    # 需要缩进的特定关键词
    target_keywords = ["人员投入", "设备投入", "累计工程量", "人员：", "设备：", "累计："]
    
    def process_paragraph(p):
        text = p.text.strip()
        if not text: return
        
        should_indent = False
        
        # 规则1：包含特定关键词的行 -> 缩进
        for kw in target_keywords:
            if kw in text:
                should_indent = True
                break
        
        # 规则2：以 (数字) 或 （数字） 开头的行 -> 缩进
        # 匹配 (1), (2), （1）, （2）
        if re.match(r"^[\(（]\d+[\)）]", text):
            should_indent = True

        # 执行缩进
        if should_indent:
            p.paragraph_format.first_line_indent = indent_size

    # 1. 遍历表格内容 (日报主要内容在表格)
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    process_paragraph(p)
                    
    # 2. 遍历文档正文 (防止有人写在表格外)
    for p in doc.paragraphs:
        process_paragraph(p)

def find_target_table(doc: Document, index: int) -> Optional[Any]:
    if 0 <= index < len(doc.tables):
        return doc.tables[index]
    return None

def update_table_row(table, row_name: str, today: str, total: str):
    """表格行更新逻辑"""
    if not table.rows: return
    
    name_col = 1
    cols_count = len(table.rows[0].cells)
    today_col = 4 if cols_count > 4 else cols_count - 2
    total_col = 5 if cols_count > 5 else cols_count - 1
    
    for row in table.rows:
        if len(row.cells) <= max(name_col, today_col, total_col): continue
        
        cell_text = row.cells[name_col].text.strip()
        
        if row_name in cell_text: 
            if today and today != "-":
                set_cell_text(row.cells[today_col], today)
            if total and total != "-":
                set_cell_text(row.cells[total_col], total)
            return

# ---------------- 模型定义 ----------------

class FillTemplateRequest(BaseModel):
    template_base64: str
    content: str
    table_index: int = 0
    row_index: int = 4
    col_index: int = 2
    update_date_weather: bool = False
    upload_to_feishu: bool = False
    feishu_token: Optional[str] = None

class UpdateDateWeatherRequest(BaseModel):
    document_base64: str
    feishu_token: Optional[str] = None

class UpdatePersonnelRequest(BaseModel):
    document_base64: str
    personnel_text: str 
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

# ---------------- API 接口实现 ----------------

@app.post("/fill-template")
async def fill_template(req: FillTemplateRequest):
    try:
        file_bytes = base64.b64decode(req.template_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        if doc.tables and len(doc.tables) > req.table_index:
            table = doc.tables[req.table_index]
            if len(table.rows) > req.row_index:
                cell = table.cell(req.row_index, req.col_index)
                
                # --- 重写写入逻辑：逐行控制格式 ---
                cell.text = "" 
                p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                
                lines = req.content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line: continue
                    
                    # 1. 判定是否需要【加粗】
                    # 条件A: 大标题（包含“施工第”或以冒号结尾）
                    is_main_title = ("：" in line or ":" in line) and ("施工第" in line)
                    # 条件B: 关键统计项（包含“人员投入”、“累计工程量”、“设备投入”）
                    is_sub_title = any(kw in line for kw in ["人员投入", "累计工程量", "设备投入"])
                    
                    should_bold = is_main_title or is_sub_title
                    
                    # 写入文本
                    run = p.add_run(line + "\n")
                    
                    # 统一字体设置
                    run.font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    run.font.size = Pt(10.5)
                    
                    # 应用加粗
                    if should_bold:
                        run.bold = True
                
                # --- 应用缩进 (调用增强版函数) ---
                apply_smart_indentation(doc)
        
        out = io.BytesIO()
        doc.save(out)
        return {"success": True, "document_base64": base64.b64encode(out.getvalue()).decode()}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "message": str(e)}

@app.post("/update-date-weather")
async def update_date_weather(req: UpdateDateWeatherRequest):
    try:
        file_bytes = base64.b64decode(req.document_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        now = datetime.now()
        date_str = f"{now.year}年{now.month}月{now.day}日"
        weather_str = "天气：晴                气温：20℃~30℃"
        
        if doc.tables:
            table = doc.tables[0]
            if len(table.rows) > 0:
                cells = table.rows[0].cells
                if len(cells) > 0: set_cell_text(cells[0], date_str)
                if len(cells) > 1: set_cell_text(cells[-1], weather_str)
        
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
        
        # 统计信息追加在最后，也应用字体规范
        p = doc.add_paragraph()
        run = p.add_run("\n" + req.personnel_text)
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(10.5)
        # 统计信息一般不需要首行缩进，或者你可以选择缩进
        # p.paragraph_format.first_line_indent = Pt(24) 
        
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
