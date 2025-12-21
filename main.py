"""
Word Document Generation Service for Dify
生成格式化的 Word 文档，支持中英文表格

启动方式:
    uvicorn main:app --host 0.0.0.0 --port 8000

API 端点:
    POST /generate - 生成 Word 文档
    POST /generate-from-template - 基于模板生成（推荐）
    GET /health - 健康检查
"""

import base64
import io
import json
import copy
import re
import os
import requests as http_requests
from datetime import datetime, timedelta
from typing import List, Optional, Tuple, Dict
from enum import Enum

from fastapi import FastAPI, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, Cm, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement

app = FastAPI(
    title="Word Document Generator",
    description="为 Dify 工作流生成格式化的 Word 文档",
    version="2.1.0"
)

# CORS 配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 矢车菊蓝着色5浅色80% (Cornflower Blue, Accent 5, Lighter 80%)
# RGB: (222, 235, 247) - Word 中的标准颜色
CORNFLOWER_BLUE_LIGHT80 = RGBColor(222, 235, 247)

# Pakbeng, Laos 坐标
PAKBENG_LAT = 19.8925
PAKBENG_LON = 101.8117

# 默认模板路径 - 支持本地、云端和环境变量部署
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def _init_templates():
    """初始化模板文件，优先从环境变量加载"""
    cn_path = os.path.join(SCRIPT_DIR, "template_cn.docx")
    en_path = os.path.join(SCRIPT_DIR, "template_en.docx")
    
    # 如果模板文件不存在，尝试从环境变量创建
    if not os.path.exists(cn_path):
        cn_base64 = os.environ.get("TEMPLATE_CN_BASE64")
        if cn_base64:
            with open(cn_path, "wb") as f:
                f.write(base64.b64decode(cn_base64))
            print(f"Created CN template from environment variable")
    
    if not os.path.exists(en_path):
        en_base64 = os.environ.get("TEMPLATE_EN_BASE64")
        if en_base64:
            with open(en_path, "wb") as f:
                f.write(base64.b64decode(en_base64))
            print(f"Created EN template from environment variable")
    
    # 返回模板路径
    if os.path.exists(cn_path):
        return cn_path, en_path
    else:
        # 本地开发时使用绝对路径
        return (
            r"d:\Projects\Dify\[CN]北本水电站施工日报.docx",
            r"d:\Projects\Dify\[EN]Pak Beng daily construction report.docx"
        )

DEFAULT_CN_TEMPLATE, DEFAULT_EN_TEMPLATE = _init_templates()



# 天气代码映射
WEATHER_CODES = {
    0: ("晴", "Clear"),
    1: ("大部晴朗", "Mainly Clear"),
    2: ("多云", "Partly Cloudy"),
    3: ("阴天", "Overcast"),
    45: ("有雾", "Foggy"),
    48: ("雾凇", "Depositing Rime Fog"),
    51: ("小毛毛雨", "Light Drizzle"),
    53: ("毛毛雨", "Moderate Drizzle"),
    55: ("大毛毛雨", "Dense Drizzle"),
    61: ("小雨", "Slight Rain"),
    63: ("中雨", "Moderate Rain"),
    65: ("大雨", "Heavy Rain"),
    71: ("小雪", "Slight Snow"),
    73: ("中雪", "Moderate Snow"),
    75: ("大雪", "Heavy Snow"),
    80: ("小阵雨", "Slight Rain Showers"),
    81: ("中阵雨", "Moderate Rain Showers"),
    82: ("大阵雨", "Violent Rain Showers"),
    95: ("雷暴", "Thunderstorm"),
    96: ("雷暴伴小冰雹", "Thunderstorm with Slight Hail"),
    99: ("雷暴伴大冰雹", "Thunderstorm with Heavy Hail"),
}


def get_yesterday_date() -> datetime:
    """获取昨天的日期"""
    return datetime.now() - timedelta(days=1)


def format_date_cn(dt: datetime) -> str:
    """格式化中文日期"""
    return f"{dt.year}年{dt.month}月{dt.day}日"


def format_date_en(dt: datetime) -> str:
    """格式化英文日期"""
    months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'Jun.', 
              'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.']
    return f"{months[dt.month-1]} {dt.day}, {dt.year}"


def format_date_footer(dt: datetime) -> str:
    """格式化页脚日期"""
    return f"{dt.year}/{dt.month:02d}/{dt.day:02d}"


def get_pakbeng_weather(date: datetime = None) -> Dict:
    """
    获取 Pakbeng, Laos 的天气信息
    使用 Open-Meteo 免费 API
    """
    if date is None:
        date = get_yesterday_date()
    
    date_str = date.strftime("%Y-%m-%d")
    
    try:
        # 使用 Open-Meteo 历史天气 API
        url = f"https://archive-api.open-meteo.com/v1/archive"
        params = {
            "latitude": PAKBENG_LAT,
            "longitude": PAKBENG_LON,
            "start_date": date_str,
            "end_date": date_str,
            "daily": "weather_code,temperature_2m_max,temperature_2m_min",
            "timezone": "Asia/Bangkok"
        }
        
        response = http_requests.get(url, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            daily = data.get("daily", {})
            
            weather_code = daily.get("weather_code", [0])[0]
            temp_min = daily.get("temperature_2m_min", [20])[0]
            temp_max = daily.get("temperature_2m_max", [30])[0]
            
            weather_cn, weather_en = WEATHER_CODES.get(weather_code, ("晴", "Clear"))
            
            return {
                "weather_cn": weather_cn,
                "weather_en": weather_en,
                "temp_min": int(round(temp_min)),
                "temp_max": int(round(temp_max)),
                "success": True
            }
    except Exception as e:
        print(f"天气 API 请求失败: {e}")
    
    # 默认值（如果 API 失败）
    return {
        "weather_cn": "晴",
        "weather_en": "Sunny",
        "temp_min": 18,
        "temp_max": 28,
        "success": False
    }


def update_document_date_weather(doc: Document, is_english: bool = False, 
                                  weather_data: Dict = None, 
                                  yesterday: datetime = None) -> None:
    """
    更新文档中的日期和天气信息
    - 更新第一个表格中的日期和天气
    - 更新文档末尾的日期
    - 保持字体和对齐格式
    """
    if yesterday is None:
        yesterday = get_yesterday_date()
    
    if weather_data is None:
        weather_data = get_pakbeng_weather(yesterday)
    
    # 格式化数据
    if is_english:
        date_cell = format_date_en(yesterday)
        weather_cell = weather_data["weather_en"]
        font_cn = 'Arial'
        font_en = 'Arial'
        font_size = Pt(11)
    else:
        date_cell = format_date_cn(yesterday)
        weather_cell = weather_data["weather_cn"]
        font_cn = '宋体'
        font_en = 'Times New Roman'
        font_size = Pt(12)
    
    temp_cell = f"{weather_data['temp_min']}℃/{weather_data['temp_max']}℃"
    
    # 更新第一个表格（气候表）并设置所有单元格格式
    if doc.tables:
        table = doc.tables[0]
        
        # 设置表格所有单元格的垂直居中、段落格式和字体格式
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                # 设置垂直居中
                set_cell_vertical_alignment(cell)
                
                # 设置段落格式
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # 设置段落间距（段前段后4磅，行距最小值8磅）
                    pPr = paragraph._p.get_or_add_pPr()
                    
                    # 移除已有的间距设置
                    for existing in pPr.findall(qn('w:spacing')):
                        pPr.remove(existing)
                    
                    spacing = OxmlElement('w:spacing')
                    spacing.set(qn('w:before'), str(int(4 * 20)))  # 4磅 = 80 twips
                    spacing.set(qn('w:after'), str(int(4 * 20)))
                    spacing.set(qn('w:line'), str(int(8 * 20)))  # 最小值8磅
                    spacing.set(qn('w:lineRule'), 'atLeast')
                    pPr.append(spacing)
                    
                    # 设置字体
                    for run in paragraph.runs:
                        run.font.size = font_size
                        run.font.name = font_en
                        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)
        
        # 只更新第二行的日期、天气、温度数据
        if len(table.rows) >= 2:
            row = table.rows[1]
            
            # 日期 (第0列)
            if len(row.cells) > 0:
                cell = row.cells[0]
                set_cell_text_with_mixed_fonts(cell, date_cell, font_cn, font_en, font_size)
            
            # 天气 (第1列)
            if len(row.cells) > 1:
                cell = row.cells[1]
                set_cell_text_with_mixed_fonts(cell, weather_cell, font_cn, font_en, font_size)
            
            # 温度 (第4列)
            if len(row.cells) > 4:
                cell = row.cells[4]
                set_cell_text_with_mixed_fonts(cell, temp_cell, font_cn, font_en, font_size)
    
    # 更新文档末尾的日期
    date_footer = format_date_footer(yesterday)
    for paragraph in doc.paragraphs:
        if "日期:" in paragraph.text or "Date:" in paragraph.text:
            # 保存原始格式
            if paragraph.runs:
                original_run = paragraph.runs[0]
                original_font_name = original_run.font.name
                original_font_size = original_run.font.size
            else:
                original_font_name = font_en
                original_font_size = font_size
            
            # 更新文本
            if is_english:
                paragraph.text = f"Date: {date_footer}"
            else:
                paragraph.text = f"日期:{date_footer}"
            
            # 恢复格式
            if paragraph.runs:
                run = paragraph.runs[0]
                run.font.name = original_font_name or font_en
                run.font.size = original_font_size or font_size
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)


# ==================== 数据模型 ====================

class TableLanguage(str, Enum):
    chinese = "chinese"
    english = "english"


class TableRow(BaseModel):
    """表格行数据"""
    seq: int  # 序号
    location: str  # 施工部位
    location_en: Optional[str] = None  # 英文施工部位
    content: str  # 施工内容
    content_en: Optional[str] = None  # 英文施工内容
    quantity: str  # 日完成量
    quantity_en: Optional[str] = None  # 英文日完成量
    shift: Optional[str] = ""  # 备注
    remarks_en: Optional[str] = ""  # 英文备注


class GenerateRequest(BaseModel):
    """生成请求"""
    chinese_data: List[TableRow]  # 中文表格数据
    english_data: Optional[List[dict]] = None  # 英文表格数据
    template_base64: Optional[str] = None  # Base64 编码的模板文件
    target_heading_cn: str = "施工内容"  # 中文表格插入位置（标题）
    target_heading_en: str = "Construction Activities"  # 英文表格插入位置


class GenerateResponse(BaseModel):
    """生成响应"""
    success: bool
    message: str
    document_base64: Optional[str] = None  # Base64 编码的生成文档
    filename: str = "施工日报.docx"


# ==================== 格式化工具函数 ====================

def set_cell_vertical_alignment(cell, alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER):
    """设置单元格垂直对齐方式"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), 'center')
    tcPr.append(vAlign)


def set_cell_shading(cell, color: RGBColor):
    """设置单元格背景颜色"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    # RGBColor 可以用元组索引访问 [0]=R, [1]=G, [2]=B
    shd.set(qn('w:fill'), f'{color[0]:02X}{color[1]:02X}{color[2]:02X}')
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)


def set_cell_font(cell, font_name_cn: str, font_name_en: str, font_size: Pt, bold: bool = False):
    """
    设置单元格字体
    - 中文使用 font_name_cn
    - 英文/数字使用 font_name_en
    """
    # 设置垂直居中
    set_cell_vertical_alignment(cell)
    
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 设置段落间距
        pPr = paragraph._p.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:before'), str(int(4 * 20)))  # 4磅 = 80 twips
        spacing.set(qn('w:after'), str(int(4 * 20)))
        spacing.set(qn('w:line'), str(int(8 * 20)))  # 最小值8磅
        spacing.set(qn('w:lineRule'), 'atLeast')
        pPr.append(spacing)
        
        for run in paragraph.runs:
            run.font.size = font_size
            run.font.bold = bold
            # 设置中文字体
            run.font.name = font_name_en
            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name_cn)


def set_cell_text_with_mixed_fonts(cell, text: str, font_cn: str, font_en: str, 
                                    font_size: Pt, bold: bool = False, 
                                    shading_color: RGBColor = None):
    """设置单元格文本，自动区分中英文字体，支持垂直居中和背景色"""
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 设置垂直居中
    set_cell_vertical_alignment(cell)
    
    # 设置背景色
    if shading_color:
        set_cell_shading(cell, shading_color)
    
    # 设置段落间距
    pPr = paragraph._p.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(4 * 20)))  # 4磅
    spacing.set(qn('w:after'), str(int(4 * 20)))
    spacing.set(qn('w:line'), str(int(8 * 20)))  # 最小值8磅
    spacing.set(qn('w:lineRule'), 'atLeast')
    pPr.append(spacing)
    
    run = paragraph.add_run(text)
    run.font.size = font_size
    run.font.bold = bold
    run.font.name = font_en
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_cn)


def set_table_border(table, border_size: float = 0.75):
    """设置表格边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    
    tblBorders = OxmlElement('w:tblBorders')
    
    border_val = str(int(border_size * 8))  # 转换为 eighths of a point
    
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), border_val)
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    
    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def set_table_autofit_window(table):
    """
    设置表格自动适应窗口/页面宽度
    相当于 Word 中的"自动调整" -> "根据窗口调整表格"
    """
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    
    # 设置表格宽度为 100% (5000 = 100% in fiftieths of a percent)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')  # pct = percentage
    
    # 移除已有的 tblW 元素
    for existing in tblPr.findall(qn('w:tblW')):
        tblPr.remove(existing)
    
    tblPr.insert(0, tblW)
    
    # 设置表格布局为自动
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'autofit')
    
    # 移除已有的 tblLayout 元素
    for existing in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(existing)
    
    tblPr.append(tblLayout)
    
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def set_column_widths(table, widths_cm: List[float]):
    """
    设置表格列宽（厘米）
    widths_cm: 每列的宽度列表，单位为厘米
    """
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])


def calculate_text_width(text: str, is_chinese: bool = False) -> float:
    """
    估算文本宽度（厘米）
    中文字符按1个字符=0.42cm估算
    英文字符按1个字符=0.22cm估算
    """
    cn_count = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
    en_count = len(text) - cn_count
    
    if is_chinese:
        return cn_count * 0.42 + en_count * 0.22
    else:
        return len(text) * 0.25


def auto_fit_column_widths(table, headers: List[str], data_rows: List[List[str]], 
                            min_widths: List[float] = None, 
                            max_widths: List[float] = None,
                            total_width: float = 15.0):
    """
    根据内容自动调整列宽
    headers: 表头列表
    data_rows: 数据行列表（每行是字符串列表）
    min_widths: 每列最小宽度
    max_widths: 每列最大宽度
    total_width: 表格总宽度（厘米）
    """
    num_cols = len(headers)
    
    # 默认最小宽度
    if min_widths is None:
        min_widths = [1.0] * num_cols
    
    # 默认最大宽度
    if max_widths is None:
        max_widths = [6.0] * num_cols
    
    # 计算每列需要的宽度
    col_widths = []
    for col_idx in range(num_cols):
        # 表头宽度
        header_width = calculate_text_width(headers[col_idx]) + 0.5
        
        # 数据最大宽度
        max_data_width = 0
        for row in data_rows:
            if col_idx < len(row):
                data_width = calculate_text_width(row[col_idx]) + 0.3
                max_data_width = max(max_data_width, data_width)
        
        # 取表头和数据的最大值
        width = max(header_width, max_data_width)
        
        # 应用最小最大限制
        width = max(min_widths[col_idx], min(max_widths[col_idx], width))
        col_widths.append(width)
    
    # 调整总宽度
    current_total = sum(col_widths)
    if current_total > total_width:
        # 按比例缩小
        scale = total_width / current_total
        col_widths = [w * scale for w in col_widths]
    
    # 设置列宽
    set_column_widths(table, col_widths)


def create_chinese_table(doc: Document, data: List[TableRow]) -> None:
    """
    创建中文格式表格
    - 中文：宋体小四（12pt）
    - 字母数字：Times New Roman
    - 居中对齐，垂直居中
    - 段前段后4磅，行距最小值8磅
    - 边框0.75磅
    - 列宽根据内容自动调整
    """
    if not data:
        return
    
    # 计算合并信息
    merge_info = {}
    for row in data:
        seq = row.seq
        merge_info[seq] = merge_info.get(seq, 0) + 1
    
    # 创建表格：序号、施工部位、施工内容、日完成量、备注
    num_rows = len(data) + 1  # +1 for header
    table = doc.add_table(rows=num_rows, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 设置边框
    set_table_border(table, 0.75)
    
    # 设置表格自动适应窗口宽度
    set_table_autofit_window(table)
    
    # 表头
    headers = ['序号', '施工部位', '施工内容', '日完成量', '备注']
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        set_cell_text_with_mixed_fonts(cell, header, '宋体', 'Times New Roman', Pt(12), bold=False)
    
    # 准备数据行用于计算列宽
    data_rows = []
    for item in data:
        data_rows.append([
            str(item.seq),
            item.location,
            item.content,
            item.quantity,
            item.shift or ''
        ])
    
    # 自动调整列宽：序号窄，施工内容宽
    auto_fit_column_widths(
        table, 
        headers, 
        data_rows,
        min_widths=[1.0, 2.0, 4.0, 2.0, 1.5],  # 最小宽度
        max_widths=[1.5, 3.5, 7.0, 3.0, 2.5],  # 最大宽度
        total_width=15.5
    )
    
    # 数据行
    done_seqs = set()
    row_idx = 1
    
    for item in data:
        row = table.rows[row_idx]
        seq = item.seq
        
        # 序号和施工部位需要合并
        if seq not in done_seqs:
            merge_count = merge_info.get(seq, 1)
            
            # 序号
            set_cell_text_with_mixed_fonts(row.cells[0], str(seq), '宋体', 'Times New Roman', Pt(12))
            
            # 施工部位
            set_cell_text_with_mixed_fonts(row.cells[1], item.location, '宋体', 'Times New Roman', Pt(12))
            
            # 合并单元格
            if merge_count > 1:
                # 合并序号列
                start_cell = table.cell(row_idx, 0)
                end_cell = table.cell(row_idx + merge_count - 1, 0)
                start_cell.merge(end_cell)
                
                # 合并施工部位列
                start_cell = table.cell(row_idx, 1)
                end_cell = table.cell(row_idx + merge_count - 1, 1)
                start_cell.merge(end_cell)
            
            done_seqs.add(seq)
        
        # 施工内容
        set_cell_text_with_mixed_fonts(row.cells[2], item.content, '宋体', 'Times New Roman', Pt(12))
        
        # 日完成量
        set_cell_text_with_mixed_fonts(row.cells[3], item.quantity, '宋体', 'Times New Roman', Pt(12))
        
        # 备注
        set_cell_text_with_mixed_fonts(row.cells[4], item.shift or '', '宋体', 'Times New Roman', Pt(12))
        
        row_idx += 1
    
    return table


def create_english_table(doc: Document, data: List[dict]) -> None:
    """
    创建英文格式表格
    """
    print(f"DEBUG: create_english_table called with {len(data) if data else 0} rows")
    if data and len(data) > 0:
        print(f"DEBUG: Sample row in create_english_table: {data[0]}")
    if not data:
        return
    
    # 计算合并信息
    merge_info = {}
    for row in data:
        seq = row.get('seq', 0)
        merge_info[seq] = merge_info.get(seq, 0) + 1
    
    # 创建表格
    num_rows = len(data) + 1
    table = doc.add_table(rows=num_rows, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 设置默认边框
    set_table_border(table, 0.5)
    
    # 设置表格自动适应窗口宽度
    set_table_autofit_window(table)
    
    # 表头（加粗 + 矢车菊蓝背景）
    headers = ['S/N', 'Construction Area', 'Activities', 'Quantities Completed', 'Remarks']
    header_row = table.rows[0]
    for i, header in enumerate(headers):
        cell = header_row.cells[i]
        set_cell_text_with_mixed_fonts(
            cell, header, 'Arial', 'Arial', Pt(11), 
            bold=True, 
            shading_color=CORNFLOWER_BLUE_LIGHT80
        )
    
    # 准备数据行用于计算列宽
    data_rows = []
    for item in data:
        data_rows.append([
            str(item.get('seq', '')),
            item.get('location_en') or item.get('location', ''),
            item.get('content_en') or item.get('content', ''),
            item.get('quantity_en') or item.get('quantity', ''),
            item.get('remarks_en') or item.get('shift', '')
        ])
    
    # 自动调整列宽
    auto_fit_column_widths(
        table, 
        headers, 
        data_rows,
        min_widths=[1.0, 2.5, 4.0, 2.5, 1.5],
        max_widths=[1.5, 4.0, 7.0, 3.5, 2.5],
        total_width=15.5
    )
    
    # 先填充所有内容
    row_idx = 1
    for item in data:
        row = table.rows[row_idx]
        
        # S/N
        set_cell_text_with_mixed_fonts(row.cells[0], str(item.get('seq', '')), 'Arial', 'Arial', Pt(11))
        
        # Construction Area
        location_en = item.get('location_en') or item.get('location', '')
        set_cell_text_with_mixed_fonts(row.cells[1], location_en, 'Arial', 'Arial', Pt(11))
        
        # Activities
        content_en = item.get('content_en') or item.get('content', '')
        set_cell_text_with_mixed_fonts(row.cells[2], content_en, 'Arial', 'Arial', Pt(11))
        
        # Quantities
        quantity_en = item.get('quantity_en') or item.get('quantity', '')
        set_cell_text_with_mixed_fonts(row.cells[3], quantity_en, 'Arial', 'Arial', Pt(11))
        
        # Remarks
        remarks = item.get('remarks_en') or item.get('shift', '')
        set_cell_text_with_mixed_fonts(row.cells[4], remarks, 'Arial', 'Arial', Pt(11))
        
        row_idx += 1

    # 然后处理单元格合并 (只合并连续相同序号的列)
    if len(data) > 1:
        current_seq = None
        start_idx = 1
        count = 0
        
        for i, item in enumerate(data):
            seq = item.get('seq')
            row_idx = i + 1
            
            if seq == current_seq and seq is not None:
                count += 1
            else:
                # 结束上一个合并组
                if count > 1:
                    # 合并序号列
                    start_cell = table.cell(start_idx, 0)
                    end_cell = table.cell(start_idx + count - 1, 0)
                    start_cell.merge(end_cell)
                    # 合并施工部位列
                    start_cell = table.cell(start_idx, 1)
                    end_cell = table.cell(start_idx + count - 1, 1)
                    start_cell.merge(end_cell)
                
                # 开始新组
                current_seq = seq
                start_idx = row_idx
                count = 1
        
        # 处理最后一组
        if count > 1:
            start_cell = table.cell(start_idx, 0)
            end_cell = table.cell(start_idx + count - 1, 0)
            start_cell.merge(end_cell)
            start_cell = table.cell(start_idx, 1)
            end_cell = table.cell(start_idx + count - 1, 1)
            start_cell.merge(end_cell)
    
    return table


def delete_table_after_heading(doc: Document, heading_text: str) -> bool:
    """
    删除指定标题之后的第一个表格
    """
    # 查找标题段落
    target_paragraph = None
    for paragraph in doc.paragraphs:
        if heading_text in paragraph.text:
            target_paragraph = paragraph
            break
    
    if target_paragraph is None:
        return False
    
    # 获取标题段落的 XML 元素
    p = target_paragraph._p
    
    # 查找标题之后的下一个表格
    next_element = p.getnext()
    while next_element is not None:
        # 检查是否是表格元素
        if next_element.tag.endswith('}tbl'):
            # 删除这个表格
            parent = next_element.getparent()
            parent.remove(next_element)
            return True
        next_element = next_element.getnext()
    
    return False


def insert_table_after_heading(doc: Document, heading_text: str, table, delete_existing: bool = True) -> bool:
    """
    将表格移动到指定标题后面
    如果 delete_existing=True，会先删除标题后面已存在的表格
    支持多备选标题匹配（英文版常用）
    """
    possible_headings = [heading_text]
    if heading_text == "Daily Construction Statistics Table":
        possible_headings.extend(["Daily Construction Statistics", "2. On-site Construction Activities", "Construction Activities"])
    
    # 先尝试删除已存在的表格
    if delete_existing:
        found_to_delete = False
        for head in possible_headings:
            if delete_table_after_heading(doc, head):
                found_to_delete = True
                print(f"DEBUG: Deleted existing table after heading '{head}'")
                break
    
    # 查找标题段落
    target_paragraph = None
    matched_heading = None
    for paragraph in doc.paragraphs:
        p_text = paragraph.text.strip()
        for head in possible_headings:
            if head in p_text:
                target_paragraph = paragraph
                matched_heading = head
                break
        if target_paragraph:
            break
    
    if target_paragraph is None:
        print(f"DEBUG: Could not find any target heading in {possible_headings}")
        return False
    
    print(f"DEBUG: Found target heading '{matched_heading}' in paragraph: '{target_paragraph.text}'")
    
    # 获取表格的 XML 元素
    tbl = table._tbl
    
    # 获取标题段落的 XML 元素
    p = target_paragraph._p
    
    # 将表格插入到标题段落之后
    p.addnext(tbl)
    
    return True


def generate_from_templates(
    cn_template_bytes: bytes,
    en_template_bytes: bytes,
    chinese_data: List[TableRow],
    english_data: List[dict],
    cn_heading: str = "当日施工统计表",
    en_heading: str = "Daily Construction Statistics Table"
) -> Tuple[bytes, bytes]:
    """
    基于模板生成中英文文档
    
    Returns:
        Tuple[bytes, bytes]: (中文文档字节, 英文文档字节)
    """
    # 处理中文文档
    cn_doc = Document(io.BytesIO(cn_template_bytes))
    cn_table = create_chinese_table(cn_doc, chinese_data)
    if cn_table:
        insert_table_after_heading(cn_doc, cn_heading, cn_table)
    
    cn_buffer = io.BytesIO()
    cn_doc.save(cn_buffer)
    cn_buffer.seek(0)
    cn_bytes = cn_buffer.read()
    
    # 处理英文文档
    en_doc = Document(io.BytesIO(en_template_bytes))
    en_table = create_english_table(en_doc, english_data)
    if en_table:
        insert_table_after_heading(en_doc, en_heading, en_table)
    
    en_buffer = io.BytesIO()
    en_doc.save(en_buffer)
    en_buffer.seek(0)
    en_bytes = en_buffer.read()
    
    return cn_bytes, en_bytes


# ==================== API 端点 ====================

@app.get("/")
@app.head("/")
async def root():
    """根路径"""
    return {"status": "ok", "service": "word-generator"}

@app.get("/health")
@app.head("/health")
async def health_check():
    """健康检查"""
    return {"status": "healthy", "service": "word-generator"}


@app.post("/generate", response_model=GenerateResponse)
async def generate_document(request: GenerateRequest):
    """
    生成 Word 文档
    
    - 如果提供 template_base64，则基于模板生成
    - 否则创建新文档
    """
    try:
        # 创建或加载文档
        if request.template_base64:
            template_bytes = base64.b64decode(request.template_base64)
            doc = Document(io.BytesIO(template_bytes))
        else:
            doc = Document()
        
        # 添加中文表格标题
        doc.add_heading(request.target_heading_cn, level=2)
        create_chinese_table(doc, request.chinese_data)
        
        doc.add_paragraph()  # 空行
        
        # 添加英文表格
        if request.english_data:
            doc.add_heading(request.target_heading_en, level=2)
            create_english_table(doc, request.english_data)
        
        # 保存到内存
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        # 转为 Base64
        doc_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        
        return GenerateResponse(
            success=True,
            message="文档生成成功",
            document_base64=doc_base64,
            filename="施工日报.docx"
        )
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文档生成失败: {str(e)}")


class SimpleGenerateRequest(BaseModel):
    """简化版生成请求 - 用于 Dify HTTP 请求节点"""
    chinese_data: str  # JSON 字符串
    english_data: Optional[str] = None  # JSON 字符串


@app.post("/generate-simple")
async def generate_simple(
    chinese_data: str = Form(...),
    english_data: Optional[str] = Form(None)
):
    """
    简化版生成接口，接受 JSON 字符串
    适用于 Dify HTTP 请求节点
    """
    try:
        cn_data = json.loads(chinese_data)
        en_data = json.loads(english_data) if english_data else None
        
        # 转换为 TableRow 对象
        chinese_rows = [TableRow(**item) for item in cn_data]
        
        doc = Document()
        
        # 中文表格
        doc.add_heading("施工内容", level=2)
        create_chinese_table(doc, chinese_rows)
        
        doc.add_paragraph()
        
        # 英文表格
        if en_data:
            doc.add_heading("Construction Activities", level=2)
            # 如果 en_data 是包含 translated_data 的对象
            if isinstance(en_data, dict) and 'translated_data' in en_data:
                en_data = en_data['translated_data']
            create_english_table(doc, en_data)
        
        # 保存
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        doc_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        
        return {
            "success": True,
            "document_base64": doc_base64,
            "filename": "施工日报.docx"
        }
    
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"JSON 解析错误: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成失败: {str(e)}")


@app.post("/generate-json")
async def generate_json(request: SimpleGenerateRequest):
    """
    JSON 请求体版本 - 推荐用于 Dify HTTP 请求节点
    
    请求体格式:
    {
        "chinese_data": "[{\"seq\":1,\"location\":\"右岸道路\",...}]",
        "english_data": "{\"translated_data\":[...]}"
    }
    """
    try:
        cn_data = json.loads(request.chinese_data)
        en_data = json.loads(request.english_data) if request.english_data else None
        
        # 转换为 TableRow 对象
        chinese_rows = [TableRow(**item) for item in cn_data]
        
        doc = Document()
        
        # 中文表格
        doc.add_heading("施工内容", level=2)
        create_chinese_table(doc, chinese_rows)
        
        doc.add_paragraph()
        
        # 英文表格
        if en_data:
            doc.add_heading("Construction Activities", level=2)
            if isinstance(en_data, dict) and 'translated_data' in en_data:
                en_data = en_data['translated_data']
            create_english_table(doc, en_data)
        
        # 保存
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        
        doc_base64 = base64.b64encode(buffer.read()).decode('utf-8')
        
        return {
            "success": True,
            "document_base64": doc_base64,
            "filename": "施工日报.docx"
        }
    
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"JSON 解析错误: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成失败: {str(e)}")


class TemplateGenerateRequest(BaseModel):
    """基于模板生成请求"""
    chinese_data: str  # JSON 字符串
    english_data: Optional[str] = None  # JSON 字符串
    cn_template_base64: Optional[str] = None  # 中文模板 Base64
    en_template_base64: Optional[str] = None  # 英文模板 Base64


@app.post("/generate-from-template")
async def generate_from_template(request: TemplateGenerateRequest):
    """
    基于模板生成中英文 Word 文档
    
    - 中文表格插入到"当日施工统计表"下方
    - 英文表格插入到"Daily Construction Statistics Table"下方
    - 表头使用矢车菊蓝着色5浅色80%底纹（仅英文）
    - 列宽根据内容自动调整
    - 单元格垂直居中
    
    请求体格式:
    {
        "chinese_data": "[{\"seq\":1,\"location\":\"右岸道路\",...}]",
        "english_data": "{\"translated_data\":[...]}",
        "cn_template_base64": "...",  // 可选
        "en_template_base64": "..."   // 可选
    }
    
    返回:
    {
        "success": true,
        "cn_document_base64": "...",
        "en_document_base64": "...",
        "cn_filename": "施工日报CN.docx",
        "en_filename": "施工日报EN.docx"
    }
    """
    try:
        cn_data = json.loads(request.chinese_data)
        en_data = json.loads(request.english_data) if request.english_data else None
        
        print(f"DEBUG: Received chinese_data length: {len(request.chinese_data)}")
        if en_data:
            print(f"DEBUG: Received english_data type: {type(en_data)}")
        else:
            print("DEBUG: No english_data received")

        # 转换为 TableRow 对象
        chinese_rows = [TableRow(**item) for item in cn_data]
        
        # 处理英文数据
        if en_data:
            if isinstance(en_data, dict) and 'translated_data' in en_data:
                english_rows = en_data['translated_data']
                print("DEBUG: Extracted 'translated_data' from dict")
            else:
                english_rows = en_data
                print("DEBUG: english_data is directly rows")
        else:
            english_rows = []
        
        print(f"DEBUG: number of english_rows: {len(english_rows)}")
        if len(english_rows) > 0:
            print(f"DEBUG: First english row: {english_rows[0]}")

        results = {}
        
        # 获取昨天日期和天气数据（只请求一次 API）
        yesterday = get_yesterday_date()
        weather_data = get_pakbeng_weather(yesterday)
        
        print(f"DEBUG: cn_template_base64 provided: {bool(request.cn_template_base64)}")
        print(f"DEBUG: en_template_base64 provided: {bool(request.en_template_base64)}")
        
        # 生成中文文档
        if request.cn_template_base64:
            cn_template_bytes = base64.b64decode(request.cn_template_base64)
            cn_doc = Document(io.BytesIO(cn_template_bytes))
        elif os.path.exists(DEFAULT_CN_TEMPLATE):
            # 自动加载默认模板
            cn_doc = Document(DEFAULT_CN_TEMPLATE)
        else:
            # 没有模板，创建空文档
            cn_doc = Document()
            cn_doc.add_heading("当日施工统计表", level=2)
        
        # 如果有模板，更新日期和天气，并插入表格到指定位置
        if request.cn_template_base64 or os.path.exists(DEFAULT_CN_TEMPLATE):
            update_document_date_weather(cn_doc, is_english=False, 
                                         weather_data=weather_data, 
                                         yesterday=yesterday)
            cn_table = create_chinese_table(cn_doc, chinese_rows)
            if cn_table:
                insert_table_after_heading(cn_doc, "当日施工统计表", cn_table)
        else:
            create_chinese_table(cn_doc, chinese_rows)
        
        cn_buffer = io.BytesIO()
        cn_doc.save(cn_buffer)
        cn_buffer.seek(0)
        results['cn_document_base64'] = base64.b64encode(cn_buffer.read()).decode('utf-8')
        results['cn_filename'] = "施工日报CN.docx"
        
        # 生成英文文档 (即使没有 english_rows，如果提供了模板，也生成文档以更新天气)
        if english_rows or request.en_template_base64 or os.path.exists(DEFAULT_EN_TEMPLATE):
            if request.en_template_base64:
                en_template_bytes = base64.b64decode(request.en_template_base64)
                en_doc = Document(io.BytesIO(en_template_bytes))
            elif os.path.exists(DEFAULT_EN_TEMPLATE):
                # 自动加载默认模板
                en_doc = Document(DEFAULT_EN_TEMPLATE)
            else:
                # 没有模板，创建空文档
                en_doc = Document()
                en_doc.add_heading("Daily Construction Statistics Table", level=2)
            
            # 如果有模板，更新日期和天气，并插入表格到指定位置
            if request.en_template_base64 or os.path.exists(DEFAULT_EN_TEMPLATE):
                update_document_date_weather(en_doc, is_english=True, 
                                             weather_data=weather_data, 
                                             yesterday=yesterday)
                if english_rows:
                    en_table = create_english_table(en_doc, english_rows)
                    if en_table:
                        if not insert_table_after_heading(en_doc, "Daily Construction Statistics Table", en_table):
                            print("DEBUG: Failed to insert EN table after heading 'Daily Construction Statistics Table'")
                else:
                    print("DEBUG: No english_rows provided, only updating date/weather in EN doc")
            elif english_rows:
                create_english_table(en_doc, english_rows)
            
            en_buffer = io.BytesIO()
            en_doc.save(en_buffer)
            en_buffer.seek(0)
            results['en_document_base64'] = base64.b64encode(en_buffer.read()).decode('utf-8')
            results['en_filename'] = "施工日报EN.docx"
        
        # 添加天气信息到返回结果
        results['weather_info'] = {
            "date": format_date_cn(yesterday),
            "weather": weather_data["weather_cn"],
            "temp": f"{weather_data['temp_min']}℃/{weather_data['temp_max']}℃",
            "api_success": weather_data["success"]
        }
        
        results['success'] = True
        return results
    
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"JSON 解析错误: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"生成失败: {str(e)}")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

