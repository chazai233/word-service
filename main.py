"""
Word Document Generation Service - Precision Formatting Version (v8.0)
修复重点：
1. 消除大标题(1、)的错误缩进。
2. 实现"人员投入"等关键词的【局部加粗】（而非整行）。
3. 严格控制子标题((1))和统计项的首行缩进。
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

app = FastAPI(title="Word Service Precision", version="8.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- 核心排版逻辑 ----------------

def format_run_font(run, size=10.5, bold=False):
    """统一设置字体格式"""
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(size)
    run.bold = bold

def process_and_add_line(cell, line_text):
    """
    智能处理每一行的格式：缩进、加粗、分割
    """
    line_text = line_text.strip()
    if not line_text: return

    # 创建新段落（注意：不使用 add_run("\n") 而是 add_paragraph 以便单独控制每一行的缩进）
    # 如果是单元格刚清空后的第一个段落，直接使用，否则新建
    if len(cell.paragraphs) == 1 and not cell.paragraphs[0].text:
        p = cell.paragraphs[0]
    else:
        p = cell.add_paragraph()

    # --- 1. 规则匹配 ---
    
    # 规则A：大标题 (例如 "1、右岸施工营地")
    # 特征：数字开头 + 顿号或点
    if re.match(r"^\d+[、\.]", line_text):
        p.paragraph_format.first_line_indent = Pt(0) # 【关键】强制不缩进
        run = p.add_run(line_text)
        format_run_font(run, bold=True) # 大标题整行加粗
        return

    # 规则B：统计项 (例如 "人员投入：...")
    # 特征：包含特定关键词
    keywords = ["人员投入", "设备投入", "累计工程量"]
    hit_keyword = None
    for kw in keywords:
        if kw in line_text:
            hit_keyword = kw
            break
    
    if hit_keyword:
        p.paragraph_format.first_line_indent = Pt(24) # 【关键】强制缩进 2 字符
        
        # 【局部加粗逻辑】
        # 将文本切分为两部分：关键词前缀(加粗) + 剩余内容(不加粗)
        # 例如 "人员投入：张三" -> "人员投入：" (粗) + " 张三" (细)
        
        # 尝试找到冒号的位置
        split_index = -1
        if "：" in line_text:
            split_index = line_text.index("：") + 1
        elif ":" in line_text:
            split_index = line_text.index(":") + 1
        else:
            # 如果没有冒号，就只加粗关键词本身
            split_index = line_text.index(hit_keyword) + len(hit_keyword)
            
        prefix = line_text[:split_index]
        content = line_text[split_index:]
        
        # 写入前缀（加粗）
        run1 = p.add_run(prefix)
        format_run_font(run1, bold=True)
        
        # 写入内容（正常）
        run2 = p.add_run(content)
        format_run_font(run2, bold=False)
        return

    # 规则C：子标题 / 具体内容 (例如 "(1) 场地精平")
    # 特征：以 (数字) 或 （数字） 开头
    if re.match(r"^[\(（]\d+[\)）]", line_text):
        p.paragraph_format.first_line_indent = Pt(24) # 【关键】强制缩进 2 字符
        run = p.add_run(line_text)
        format_run_font(run, bold=False) # 内容不加粗
        return

    # 规则D：其他普通文本
    # 默认缩进2字符（因为通常是正文延续），或者0？
    # 根据你的截图，如果不符合上述规则，通常是正文描述，建议缩进2字符对齐
    p.paragraph_format.first_line_indent = Pt(24)
    run = p.add_run(line_text)
    format_run_font(run, bold=False)

# ---------------- 辅助函数 ----------------

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
            # 填入数字时也应用字体规范
            if today and today != "-":
                cell = row.cells[today_col]
                cell.text = ""
                run = cell.paragraphs[0].add_run(str(today))
                format_run_font(run)
            if total and total != "-":
                cell = row.cells[total_col]
                cell.text = ""
                run = cell.paragraphs[0].add_run(str(total))
                format_run_font(run)
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
                
                # 清空单元格
                cell.text = "" 
                
                # 逐行处理，精确控制格式
                lines = req.content.split('\n')
                for line in lines:
                    process_and_add_line(cell, line)
        
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
                if len(cells) > 0: 
                    cells[0].text = ""
                    run = cells[0].paragraphs[0].add_run(date_str)
                    format_run_font(run)
                if len(cells) > 1: 
                    cells[-1].text = ""
                    run = cells[-1].paragraphs[0].add_run(weather_str)
                    format_run_font(run)
        
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
        
        # 统计信息在文末追加，默认不需要特殊缩进，但需要字体规范
        p = doc.add_paragraph()
        run = p.add_run("\n" + req.personnel_text)
        format_run_font(run)
        
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
