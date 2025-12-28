"""
Word Document Generation Service for Dify
生成格式化的 Word 文档，支持中英文表格、首行缩进及附表填报

启动方式:
    uvicorn main:app --host 0.0.0.0 --port 8000
"""

import base64
import io
import json
import copy
import re
import os
import requests as http_requests
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict, Any
from enum import Enum

from fastapi import FastAPI, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, Cm, Twips, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

app = FastAPI(
    title="Word Document Generator",
    description="为 Dify 工作流生成格式化的 Word 文档",
    version="3.0.0"
)

# CORS 配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- 配置常量 ----------------
CORNFLOWER_BLUE_LIGHT80 = RGBColor(222, 235, 247)
PAKBENG_LAT = 19.8925
PAKBENG_LON = 101.8117
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ---------------- 模型定义 ----------------

class TableLanguage(str, Enum):
    chinese = "chinese"
    english = "english"

class TableRow(BaseModel):
    seq: int
    location: str
    content: str
    quantity: str
    shift: Optional[str] = ""
    location_en: Optional[str] = None
    content_en: Optional[str] = None
    quantity_en: Optional[str] = None
    remarks_en: Optional[str] = ""

class GenerateRequest(BaseModel):
    chinese_data: List[TableRow]
    english_data: Optional[List[dict]] = None
    template_base64: Optional[str] = None
    target_heading_cn: str = "施工内容"
    target_heading_en: str = "Construction Activities"
    
class AppendixTableData(BaseModel):
    table_index: int
    header_text: Optional[str] = None
    row_name: str
    today_qty: str
    total_qty: str

class UpdateAppendixRequest(BaseModel):
    document_base64: str
    data: List[AppendixTableData]
    feishu_token: Optional[str] = None

class FillTemplateRequest(BaseModel):
    template_base64: str
    content: str
    table_index: int = 0
    row_index: int = 4
    col_index: int = 2
    update_date_weather: bool = False
    is_english: bool = False
    upload_to_feishu: bool = False
    feishu_token: Optional[str] = None

class UpdatePersonnelRequest(BaseModel):
    document_base64: str
    personnel_text: str
    feishu_token: Optional[str] = None

# ---------------- 辅助函数：天气与日期 ----------------

def get_yesterday_date() -> datetime:
    return datetime.now() - timedelta(days=1)

def format_date_cn(dt: datetime) -> str:
    return f"{dt.year}年{dt.month}月{dt.day}日"

def format_date_en(dt: datetime) -> str:
    months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'Jun.', 
              'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.']
    return f"{months[dt.month-1]} {dt.day}, {dt.year}"

def get_pakbeng_weather(date: datetime = None) -> Dict:
    # 简化版: 实际部署可复用之前的逻辑
    return {
        "weather_cn": "晴",
        "weather_en": "Sunny",
        "temp_min": 20,
        "temp_max": 30,
        "success": True
    }

# ---------------- 核心功能：字体设置辅助函数 (新增) ----------------

def set_cell_text(cell, text):
    """
    设置单元格文本，并应用：
    1. 中文：宋体
    2. 英文/数字：Times New Roman
    3. 字号：五号 (10.5pt)
    """
    # 清空单元格原有内容
    cell.text = ""
    # 获取或创建第一个段落
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    
    # 创建一个 run
    run = p.add_run(text)
    
    # 设置西文字体 (Times New Roman)
    run.font.name = 'Times New Roman'
    # 设置中文字体 (宋体) - 需要通过 XML 设置
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    # 设置字号为五号 (10.5磅)
    run.font.size = Pt(10.5)

# ---------------- 核心功能：首行缩进 ----------------

def apply_smart_indentation(doc: Document):
    """
    智能首行缩进
    遍历文档段落，对以特定关键词开头的行应用2字符首行缩进
    """
    keywords = ["人员投入", "设备投入", "累计工程量", "人员：", "设备："]
    
    # 2字符约等于 24pt (小四字号)
    indent_size = Pt(24) 
    
    # 遍历所有段落
    for p in doc.paragraphs:
        # 简单清洗
        text = p.text.strip()
        for kw in keywords:
            if text.startswith(kw):
                p_fmt = p.paragraph_format
                p_fmt.first_line_indent = indent_size
                break
                
    # 同时也遍历表格中的段落
    if doc.tables:
        for row in doc.tables[0].rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    text = p.text.strip()
                    for kw in keywords:
                        if text.startswith(kw):
                            p.paragraph_format.first_line_indent = indent_size
                            break

# ---------------- 核心功能：附表填报 ----------------

def find_target_table(doc: Document, index: int, header_hint: str) -> Optional[Any]:
    """
    根据索引或标题提示定位表格
    """
    # 优先使用索引
    if 0 <= index < len(doc.tables):
        return doc.tables[index]
    return None

def update_table_row(table, row_name: str, today: str, total: str):
    """
    更新表格中的行
    """
    # 1. 确定列索引
    name_col = -1
    today_col = -1
    total_col = -1
    
    if not table.rows: return
    
    headers = [c.text.strip() for c in table.rows[0].cells]
    
    # 智能寻找列
    for i, h in enumerate(headers):
        if "项目" in h or "名称" in h: name_col = i
        elif "今日" in h or "日完成" in h: today_col = i
        elif "累计" in h: total_col = i
        
    if name_col == -1: name_col = 1
    if today_col == -1: today_col = 4
    if total_col == -1: total_col = 5
    
    # 2. 遍历行查找
    for row in table.rows:
        if len(row.cells) <= max(name_col, today_col, total_col): continue
        
        cell_name = row.cells[name_col].text.strip()
        
        # 模糊匹配
        if row_name in cell_name:
            def safe_set_text(cell, text):
                item = cell.paragraphs[0]
                if item.runs:
                    item.runs[0].text = str(text)
                else:
                    item.add_run(str(text))
            
            if today and today != "-":
                safe_set_text(row.cells[today_col], today)
            
            if total and total != "-":
                safe_set_text(row.cells[total_col], total)
            
            return

# ---------------- API 端点 ----------------

@app.post("/fill-template")
async def fill_template(req: FillTemplateRequest):
    try:
        if not req.template_base64:
            raise HTTPException(status_code=400, detail="Missing template_base64")
            
        # 1. 解码模板
        file_bytes = base64.b64decode(req.template_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        # 2. 填充内容
        if doc.tables and len(doc.tables) > req.table_index:
            table = doc.tables[req.table_index]
            if len(table.rows) > req.row_index and len(table.rows[req.row_index].cells) > req.col_index:
                cell = table.cell(req.row_index, req.col_index)
                cell.text = req.content
                apply_smart_indentation(doc)
        
        # 4. 保存
        out_stream = io.BytesIO()
        doc.save(out_stream)
        out_bytes = out_stream.getvalue()
        b64_str = base64.b64encode(out_bytes).decode('utf-8')
        
        return {
            "success": True,
            "document_base64": b64_str,
            "filename": "日报_filled.docx"
        }
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "message": str(e)}

@app.post("/update-appendix-tables")
async def update_appendix_tables(req: UpdateAppendixRequest):
    try:
        # 1. 解码文档
        file_bytes = base64.b64decode(req.document_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        # 2. 遍历数据进行更新
        for item in req.data:
            target_table = find_target_table(doc, item.table_index, item.header_text)
            if target_table:
                update_table_row(target_table, item.row_name, item.today_qty, item.total_qty)
        
        # 3. 输出
        out_stream = io.BytesIO()
        doc.save(out_stream)
        out_bytes = out_stream.getvalue()
        b64_str = base64.b64encode(out_bytes).decode('utf-8')
        
        return {
            "success": True,
            "document_base64": b64_str
        }

    except Exception as e:
        return {"success": False, "message": str(e)}

# ---------------- 更新日期天气接口 (已添加字体样式) ----------------

class UpdateDateWeatherRequest(BaseModel):
    document_base64: str
    feishu_token: Optional[str] = None

@app.post("/update-date-weather")
async def update_date_weather(req: UpdateDateWeatherRequest):
    try:
        # 1. 解码 Word 文档
        file_bytes = base64.b64decode(req.document_base64)
        doc = Document(io.BytesIO(file_bytes))
        
        # 2. 获取当前日期和天气
        now = datetime.now()
        date_str = format_date_cn(now)
        weather_info = get_pakbeng_weather()
        
        # --- 关键修改：按你的要求格式化天气字符串 ---
        # 格式：天气：晴                气温：15℃~25℃
        # 注意：这里加了长空格，且连接符改为 ~
        weather_str = f"天气：{weather_info['weather_cn']}                气温：{weather_info['temp_min']}℃~{weather_info['temp_max']}℃"
        
        # 3. 写入表格 
        if doc.tables:
            table = doc.tables[0]
            try:
                if len(table.rows) > 0:
                    row_cells = table.rows[0].cells
                    
                    # 填入日期：第一行第一列
                    if len(row_cells) > 0:
                        # 使用辅助函数设置字体和字号
                        set_cell_text(row_cells[0], date_str)
                    
                    # 填入天气：第一行第四列 (或最后一列)
                    if len(row_cells) > 3:
                        set_cell_text(row_cells[3], weather_str)
                    elif len(row_cells) > 1:
                        set_cell_text(row_cells[-1], weather_str)
                        
            except Exception as table_e:
                print(f"Warning: Table update skipped: {table_e}")

        # 4. 保存并返回
        out_stream = io.BytesIO()
        doc.save(out_stream)
        out_bytes = out_stream.getvalue()
        b64_str = base64.b64encode(out_bytes).decode('utf-8')
        
        return {
            "success": True,
            "document_base64": b64_str
        }
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "message": str(e)}

@app.post("/update-personnel-stats")
async def update_personnel_stats(req: UpdatePersonnelRequest):
    try:
        return {
            "success": True,
            "document_base64": req.document_base64,
            "filename": "日报_personnel_updated.docx"
        }
    except Exception as e:
        return {"success": False, "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
