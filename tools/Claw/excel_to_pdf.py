# -*- coding: utf-8 -*-
"""
保定润民供电服务有限公司 2025年度财务报表 PDF生成器 - 专业版
支持完整的Excel数据读取和美观的PDF输出
"""

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.pdfgen import canvas
import os
import datetime

# 注册中文字体
pdfmetrics.registerFont(TTFont('SimHei', 'C:/Windows/Fonts/simhei.ttf'))
pdfmetrics.registerFont(TTFont('SimSun', 'C:/Windows/Fonts/simsunb.ttf'))
pdfmetrics.registerFont(TTFont('NotoSansSC', 'C:/Windows/Fonts/NotoSansSC-VF.ttf'))

# 页面尺寸 (A4横向更适合财务报表)
PAGE_WIDTH, PAGE_HEIGHT = A4

# 颜色定义
HEADER_BG = colors.Color(0.25, 0.45, 0.65)  # 深蓝色表头
SUBHEADER_BG = colors.Color(0.85, 0.9, 0.95)  # 浅蓝色次级表头
TOTAL_BG = colors.Color(0.92, 0.92, 0.92)  # 灰色汇总行

class NumberedCanvas(canvas.Canvas):
    """带页码的Canvas"""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []
    
    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()
    
    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)
    
    def draw_page_number(self, page_count):
        self.saveState()
        self.setFont('SimSun', 9)
        self.setFillColor(colors.gray)
        
        # 页码 - 底部居中
        page_num = self._pageNumber
        self.drawCentredString(PAGE_WIDTH / 2, 12*mm, f'{page_num} / {page_count}')
        
        # 公司名称 - 左下角
        self.drawString(15*mm, 12*mm, '保定润民供电服务有限公司')
        
        # 日期 - 右下角
        self.drawRightString(PAGE_WIDTH - 15*mm, 12*mm, '2025年12月31日')
        
        self.restoreState()

def read_excel_sheets(file_path):
    """读取Excel所有工作表"""
    xls = pd.ExcelFile(file_path)
    return {sheet: pd.read_excel(xls, sheet_name=sheet, header=None) for sheet in xls.sheet_names}

def format_number(num):
    """格式化数字"""
    if pd.isna(num):
        return ''
    if isinstance(num, (int, float)):
        if num == 0:
            return ''
        return f'{num:,.2f}'
    return str(num)

def read_full_sheet(df, sheet_name):
    """读取完整工作表数据"""
    data = []
    for idx in range(len(df)):
        row = df.iloc[idx]
        row_list = []
        for col in range(len(row)):
            val = row.iloc[col] if hasattr(row, 'iloc') else row[col]
            row_list.append(val)
        data.append(row_list)
    return data

def build_financial_table(data, col_widths, highlight_keywords=None):
    """构建财务报表表格"""
    if highlight_keywords is None:
        highlight_keywords = ['：', '合计', '总计', '小计']
    
    # 创建样式
    style_list = [
        # 表头样式
        ('BACKGROUND', (0, 0), (-1, 0), HEADER_BG),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'SimHei'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        
        # 数据样式
        ('FONTNAME', (0, 1), (-1, -1), 'SimSun'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 1), (1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 4),
        
        # 网格
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('LINEBELOW', (0, 0), (-1, 0), 1.5, colors.black),
    ]
    
    # 高亮汇总行
    for i, row in enumerate(data[1:], 1):
        if row and row[0]:
            text = str(row[0])
            if any(kw in text for kw in highlight_keywords):
                style_list.append(('BACKGROUND', (0, i), (-1, i), TOTAL_BG))
                style_list.append(('FONTNAME', (0, i), (-1, i), 'SimHei'))
                style_list.append(('LINEABOVE', (0, i), (-1, i), 1, colors.black))
    
    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(style_list))
    return table

def create_pdf(output_path, data_dict):
    """创建PDF文档"""
    
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=12*mm,
        leftMargin=12*mm,
        topMargin=15*mm,
        bottomMargin=25*mm
    )
    
    story = []
    
    # ========== 封面页 ==========
    story.append(Spacer(1, 50*mm))
    
    # 标题
    title_style = ParagraphStyle('MainTitle', fontName='SimHei', fontSize=22, 
                                  leading=32, alignment=TA_CENTER, spaceAfter=8)
    subtitle_style = ParagraphStyle('SubTitle', fontName='SimHei', fontSize=16, 
                                      leading=24, alignment=TA_CENTER, spaceAfter=20)
    info_style = ParagraphStyle('Info', fontName='SimSun', fontSize=12, 
                                  leading=20, alignment=TA_CENTER, spaceAfter=5)
    
    story.append(Paragraph('保定润民供电服务有限公司', title_style))
    story.append(Spacer(1, 10*mm))
    story.append(Paragraph('2025年度', subtitle_style))
    story.append(Paragraph('财务会计报表', subtitle_style))
    story.append(Spacer(1, 30*mm))
    
    # 报表信息
    story.append(Paragraph('编制单位：保定润民供电服务有限公司', info_style))
    story.append(Paragraph('报表期间：2025年1月1日至2025年12月31日', info_style))
    story.append(Paragraph('金额单位：人民币元', info_style))
    story.append(Spacer(1, 40*mm))
    
    # 报表目录
    toc_style = ParagraphStyle('TOC', fontName='SimSun', fontSize=11, 
                                 leading=18, alignment=TA_CENTER, textColor=colors.gray)
    story.append(Paragraph('────────────────────────────────────', toc_style))
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph('• 资产负债表', toc_style))
    story.append(Paragraph('• 利润表', toc_style))
    story.append(Paragraph('• 现金流量表', toc_style))
    story.append(Paragraph('• 所有者权益变动表', toc_style))
    
    story.append(PageBreak())
    
    # ========== 资产负债表 ==========
    sheet_title_style = ParagraphStyle('SheetTitle', fontName='SimHei', fontSize=14, 
                                        leading=20, alignment=TA_CENTER, spaceAfter=3)
    sheet_subtitle_style = ParagraphStyle('SheetSub', fontName='SimSun', fontSize=9, 
                                           leading=14, alignment=TA_CENTER, spaceAfter=8)
    
    story.append(Paragraph('资 产 负 债 表', sheet_title_style))
    story.append(Paragraph('2025年12月31日', sheet_subtitle_style))
    story.append(Paragraph('编制单位：保定润民供电服务有限公司    单位：元', sheet_subtitle_style))
    story.append(Spacer(1, 5*mm))
    
    if '1资产负债表' in data_dict:
        df = data_dict['1资产负债表']
        
        # 资产部分 - 前半部分
        table_data = [['项 目', '附注', '期末余额', '期初余额']]
        
        for idx in range(2, len(df)):
            row = df.iloc[idx]
            item = str(row[1]) if not pd.isna(row[1]) else ''
            ref = str(row[2]) if not pd.isna(row[2]) else ''
            end_bal = format_number(row[3]) if len(row) > 3 else ''
            beg_bal = format_number(row[4]) if len(row) > 4 else ''
            
            if item and item != 'nan' and item.strip():
                table_data.append([item, ref, end_bal, beg_bal])
        
        col_widths = [110*mm, 25*mm, 45*mm, 45*mm]
        table = build_financial_table(table_data, col_widths)
        story.append(table)
    
    story.append(PageBreak())
    
    # ========== 资产负债表（续） ==========
    story.append(Paragraph('资 产 负 债 表（续）', sheet_title_style))
    story.append(Paragraph('2025年12月31日', sheet_subtitle_style))
    story.append(Paragraph('编制单位：保定润民供电服务有限公司    单位：元', sheet_subtitle_style))
    story.append(Spacer(1, 5*mm))
    
    if '2资产负债表（续）' in data_dict:
        df = data_dict['2资产负债表（续）']
        
        table_data = [['项 目', '附注', '期末余额', '期初余额']]
        
        for idx in range(2, len(df)):
            row = df.iloc[idx]
            item = str(row[0]) if not pd.isna(row[0]) else ''
            ref = str(row[1]) if not pd.isna(row[1]) else ''
            end_bal = format_number(row[2]) if len(row) > 2 else ''
            beg_bal = format_number(row[3]) if len(row) > 3 else ''
            
            if item and item != 'nan' and item.strip():
                table_data.append([item, ref, end_bal, beg_bal])
        
        col_widths = [110*mm, 25*mm, 45*mm, 45*mm]
        table = build_financial_table(table_data, col_widths)
        story.append(table)
    
    story.append(PageBreak())
    
    # ========== 利润表 ==========
    story.append(Paragraph('利 润 表', sheet_title_style))
    story.append(Paragraph('2025年度', sheet_subtitle_style))
    story.append(Paragraph('编制单位：保定润民供电服务有限公司    单位：元', sheet_subtitle_style))
    story.append(Spacer(1, 5*mm))
    
    if '3利润表' in data_dict:
        df = data_dict['3利润表']
        
        table_data = [['项 目', '附注', '本期金额', '上期金额']]
        
        for idx in range(2, len(df)):
            row = df.iloc[idx]
            item = str(row[0]) if not pd.isna(row[0]) else ''
            ref = str(row[1]) if not pd.isna(row[1]) else ''
            current = format_number(row[2]) if len(row) > 2 else ''
            previous = format_number(row[3]) if len(row) > 3 else ''
            
            if item and item != 'nan' and item.strip():
                table_data.append([item, ref, current, previous])
        
        col_widths = [110*mm, 25*mm, 45*mm, 45*mm]
        table = build_financial_table(table_data, col_widths)
        story.append(table)
    
    story.append(PageBreak())
    
    # ========== 现金流量表 ==========
    story.append(Paragraph('现 金 流 量 表', sheet_title_style))
    story.append(Paragraph('2025年度', sheet_subtitle_style))
    story.append(Paragraph('编制单位：保定润民供电服务有限公司    单位：元', sheet_subtitle_style))
    story.append(Spacer(1, 5*mm))
    
    if '4现金流量表' in data_dict:
        df = data_dict['4现金流量表']
        
        table_data = [['项 目', '附注', '本期金额', '上期金额']]
        
        for idx in range(2, len(df)):
            row = df.iloc[idx]
            item = str(row[0]) if not pd.isna(row[0]) else ''
            ref = str(row[1]) if not pd.isna(row[1]) else ''
            current = format_number(row[2]) if len(row) > 2 else ''
            previous = format_number(row[3]) if len(row) > 3 else ''
            
            if item and item != 'nan' and item.strip():
                table_data.append([item, ref, current, previous])
        
        col_widths = [110*mm, 25*mm, 45*mm, 45*mm]
        table = build_financial_table(table_data, col_widths)
        story.append(table)
    
    story.append(PageBreak())
    
    # ========== 所有者权益变动表 ==========
    story.append(Paragraph('所有者权益变动表', sheet_title_style))
    story.append(Paragraph('2025年度', sheet_subtitle_style))
    story.append(Paragraph('编制单位：保定润民供电服务有限公司    单位：元', sheet_subtitle_style))
    story.append(Spacer(1, 5*mm))
    
    if '5权益变动表 ' in data_dict:
        df = data_dict['5权益变动表 ']
        
        # 本年数据
        table_data = [['项 目', '实收资本', '资本公积', '其他综合收益', '专项储备', '盈余公积', '未分配利润', '所有者权益合计']]
        
        for idx in range(5, min(20, len(df))):
            row = df.iloc[idx]
            item = str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else ''
            
            if item and item != 'nan' and item.strip():
                equity_row = [item]
                col_indices = [3, 5, 8, 9, 10, 11]
                for ci in col_indices:
                    val = ''
                    if ci < len(row):
                        v = row.iloc[ci] if hasattr(row, 'iloc') else row[ci]
                        val = format_number(v) if not pd.isna(v) else ''
                    equity_row.append(val)
                
                if len(equity_row) == 8:
                    table_data.append(equity_row)
        
        col_widths = [35*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 25*mm]
        table = build_financial_table(table_data, col_widths)
        story.append(table)
    
    story.append(Spacer(1, 8*mm))
    
    # 上年对比
    if '6权益变动表（续）' in data_dict:
        df = data_dict['6权益变动表（续）']
        
        table_data2 = [['项 目', '实收资本', '资本公积', '其他综合收益', '专项储备', '盈余公积', '未分配利润', '所有者权益合计']]
        
        for idx in range(5, min(20, len(df))):
            row = df.iloc[idx]
            item = str(row[1]) if len(row) > 1 and not pd.isna(row[1]) else ''
            
            if item and item != 'nan' and item.strip():
                equity_row = [item]
                col_indices = [3, 5, 8, 9, 10, 11]
                for ci in col_indices:
                    val = ''
                    if ci < len(row):
                        v = row.iloc[ci] if hasattr(row, 'iloc') else row[ci]
                        val = format_number(v) if not pd.isna(v) else ''
                    equity_row.append(val)
                
                if len(equity_row) == 8:
                    table_data2.append(equity_row)
        
        col_widths = [35*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 22*mm, 25*mm]
        table = build_financial_table(table_data2, col_widths)
        story.append(table)
    
    # ========== 生成PDF ==========
    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"✓ PDF生成完成: {output_path}")

if __name__ == '__main__':
    # 输入输出路径
    input_file = r'D:/Users/12844/Desktop/吉达建设、润民/23.保定润民供电服务有限公司/1.审计报告/3.2025年度财务报表-保定润民5.14.xls'
    output_file = r'D:/Users/12844/Desktop/吉达建设、润民/23.保定润民供电服务有限公司/1.审计报告/保定润民供电2025年度财务报表.pdf'
    
    # 读取数据
    print('读取Excel文件...')
    data = read_excel_sheets(input_file)
    print(f'共读取 {len(data)} 个工作表: {list(data.keys())}')
    
    # 生成PDF
    print('生成PDF文件...')
    create_pdf(output_file, data)
