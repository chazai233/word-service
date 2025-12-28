"""
Word Document Generation Service for Dify (Ultimate Version)
适配：多区域人员统计、多区域累计工程量、复杂格式排版
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

app = FastAPI(title="Word Service Ultimate", version="6.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- 辅助函数：排版核心 ----------------

def set_cell_text(cell, text, bold=False):
    """
    设置单元格文本，保留格式：
    1. 中文：宋体
    2. 英文/数字：Times New Roman
    3. 字号：五号 (10.5pt)
    4. bold: 是否加粗
    """
    cell.text = "" # 清空原有内容
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    
    run = p.add_run(str(text))
    
    # 字体设置
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)
    
    # 加粗设置
    if bold:
        run.bold = True

def apply_smart_indentation(doc: Document):
    """
    智能首行缩进
    逻辑：只要这一行看起来像是正文（不是极短的标题），且前面有序号，就缩进。
    """
    indent_size = Pt(24) # 2字符
    
    def process_paragraph(p):
        text = p.text.strip()
        # 排除空行
        if not text: return
        
        # 排除明显的短标题 (比如 "1、右岸施工营地")，通常标题不缩进，或者已经加粗了
        # 这里策略是：只要包含特定的关键词，或者看起来像描述性段落，就缩进
        
        # 关键词匹配
        keywords = ["人员投入", "设备投入", "累计工程量", "人员：", "设备：", "累计："]
        should_indent = False
        
        # 1. 命中关键词
        for kw in keywords:
            if kw in text:
                should_indent = True
                break
        
        # 2. 或者以 (1), (2) 开头的长段落
        if re.match(r"^\(\d+\)", text) and len(text) > 20:
             should_indent = True

        if should_indent:
            p.paragraph_format.first_line_indent = indent_size

    # 遍历表格内容 (日报的主要内容都在表格里)
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
    """表格行更新逻辑"""
    if not table.rows: return
    
    # 列索引推断
    name_col = 1
    cols_count = len(table.rows[0].cells)
    today_col = 4 if cols_count > 4 else cols_count - 2
    total_col = 5 if cols_count > 5 else cols_count - 1
    
    for row in table.rows:
        if len(row.cells) <= max(name_col, today_col, total_col): continue
        
        cell_text = row.cells[name_col].text.strip()
        
        # 模糊匹配：只要 AI 提取的名字(row_name) 在表格这一行(cell_text) 里出现即可
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

# ---------------- API 接口 ----------------

@app.post("/fill-template")
async def fill_template(req: FillTemplateRequest):
    try:
        file_bytes = base64.b64decode(req.template_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        if doc.tables and len(doc.tables) > req.table_index:
            table = doc.tables[req.table_index]
            if len(table.rows) > req.row_index:
                cell = table.cell(req.row_index, req.col_index)
                
                # --- 写入主内容 ---
                cell.text = "" 
                p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
                
                lines = req.content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line: continue
                    
                    # 识别标题：以冒号结尾，或包含"施工第"
                    # 你的截图显示：1、xxx (施工第xx天)：
                    is_title = ("：" in line or ":" in line) and ("施工第" in line)
                    
                    run = p.add_run(line + "\n")
                    
                    # 字体
                    run.font.name = 'Times New Roman'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
                    run.font.size = Pt(10.5)
                    
                    if is_title:
                        run.bold = True
                
                # --- 应用缩进 ---
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
        
        # 追加人员统计信息到文档末尾（或者你可以指定填入主表格的某一行）
        # 这里演示为：追加在正文下方
        if doc.tables:
             # 尝试追加到第一个表格的最后一个单元格（通常是备注或者正文格）
             # 如果你的正文格是 row=4, col=2，我们也可以追加到那里
             table = doc.tables[0]
             # 这里简单起见，追加到文档最后
             p = doc.add_paragraph()
             run = p.add_run("\n" + req.personnel_text)
             run.font.name = 'Times New Roman'
             run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
             run.font.size = Pt(10.5)
        
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
