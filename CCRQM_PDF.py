import pandas as pd
from opencc import OpenCC
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER
from PyPDF2 import PdfMerger
import tejapi
from ccrqm_text import ccrqm_text

tejapi.ApiConfig.api_key  = "Your Key"
tejapi.ApiConfig.api_base = "http://10.10.10.66"
tejapi.ApiConfig.ignoretz = True

start = input("輸入起始日期(ex:2024-01-01)：")
end = input("輸入結束日期(ex:2024-01-01)：")

print(ccrqm_text(start, end))

Date_range = {'gte':start, 'lte':end}
CFLTNEW1 = tejapi.get('CHN/CFLTNEW1',
                        mdate = Date_range,
                        paginate=True,
                        chinese_column_name = True
                        )

date_range_text = f"从 {Date_range['gte']} 到 {Date_range['lte']}"

# 撈取公司名稱
def get_name(coid):
    CCR1STD = tejapi.get('CHN/CCR1STD',
    coid=coid,
    paginate=True,
    chinese_column_name = True
    )
    return CCR1STD["公司簡稱(中)"].iloc[0]

# 定義字體
pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))

# 創建 OpenCC 用于繁體轉簡體
cc = OpenCC('t2s')

# 轉換繁體中文為簡體中文
CFLTNEW1['新聞內容'] = CFLTNEW1['新聞內容'].apply(lambda x: cc.convert(x))
CFLTNEW1['事件小分類名'] = CFLTNEW1['事件小分類名'].apply(lambda x: cc.convert(x))

# 合并相同公司码的资料，去重事件属性
CFLTNEW1_grouped = CFLTNEW1.groupby(['公司碼', '事件大分類'], as_index=False).agg({
    '新聞內容': lambda x: ' '.join(x),
    '事件小分類名': lambda x: ', '.join(pd.Series(x).drop_duplicates()),
    'CCRI-數量模型(資料年月)': lambda x: ', '.join(pd.Series(x).drop_duplicates())
})

# 定義頭尾 Call to Action 內容
cta_text = "订阅我们的服务，获取更详细的金融服务！"

# 創建 PDF 文件
pdf_file = '中国快讯.pdf'
pdf = SimpleDocTemplate(pdf_file, pagesize=A4, topMargin=0.3*inch)
elements = []
styles = getSampleStyleSheet()

# Call to Action 樣式
cta_style = ParagraphStyle(name='CTA', fontName='SimSun', fontSize=12, spaceAfter=12, alignment=TA_CENTER, textColor=HexColor('#003459'))
# 標題樣式
title_style = ParagraphStyle(name='Title', fontName='SimSun', fontSize=20, spaceAfter=12,leftIndent=-0.5*inch)
# 日期樣式
date_style = ParagraphStyle(name='Date', fontName='SimSun', fontSize=12, spaceAfter=12,leftIndent=-0.5*inch)
# 分類標題樣式
heading_style = ParagraphStyle(name='Heading2', fontName='SimSun', fontSize=16, spaceAfter=6, textColor=HexColor('#007EA7'),leftIndent=-0.6*inch)
# 普通文本樣式
normal_style = ParagraphStyle(name='Normal', fontName='SimSun', fontSize=10)

# 表格樣式
table_style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), HexColor('#003459')),  # 表頭背景色
    ('TEXTCOLOR', (0, 0), (-1, 0), HexColor('#FFFFFF')),   # 表頭文本顏色（白色）
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # 數據水平居中
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # 數據垂直居中
    ('FONTNAME', (0, 0), (-1, 0), 'SimSun'),  # 表頭字體
    ('FONTNAME', (0, 1), (-1, -1), 'SimSun'),  # 數據字體
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), HexColor('#FFFFFF')),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('ALIGN', (0, 0), (-1, 0), 'CENTER')  # 對表頭應用居中對齊
])

# 添加自定義頁面模板
def my_first_page(canvas, doc):
    canvas.saveState()
    # 添加 LOGO
    logo_path = "TEJ logo.png"  # LOGO 圖片路徑
    logo = Image(logo_path, width=1*inch, height=0.6*inch)
    logo.drawOn(canvas, A4[0] - 1.2*inch - 15, A4[1] - 0.7*inch - 15)  # LOGO 位置：右上角，邊距 20pt
    canvas.restoreState()

pdf.build(elements, onFirstPage=my_first_page)

# 添加 PDF 標題
elements.append(Paragraph(cta_text, cta_style))  # 在文檔頂部添加 Call to Action
elements.append(Paragraph("CCRQM 沪深上市企业风险事件早报", title_style))
elements.append(Paragraph(date_range_text, date_style))
elements.append(Spacer(1, 0.2 * inch))

# 瀏覽分類並整理數據
for category in CFLTNEW1_grouped['事件大分類'].unique():
    category_data = CFLTNEW1_grouped[CFLTNEW1_grouped['事件大分類'] == category]
    
    # 添加分類標題
    category_title = cc.convert(category)  # 转换分类标题为简体中文
    if category_title == "机会":
        category_title = "商业机会"
    elements.append(Paragraph(f"{category_title}", heading_style))
    elements.append(Spacer(1, 0.2 * inch))

    # 準備表格數據
    table_data = [['公司码', 'CCRQM 等级', '新闻摘要', '事件属性']]  # 表頭
    for _, row in category_data.iterrows():
        # 獲取公司名稱
        company_name = get_name(row['公司碼'])
        company_name_simplified = cc.convert(company_name)
        # 將公司名稱與公司碼组合
        company_info = f"{row['公司碼']} {company_name_simplified}"
        # 添加到表格數據中
        table_data.append([company_info, row['CCRI-數量模型(資料年月)'], row['新聞內容'], row['事件小分類名']])

    # 创建表格
    if len(table_data) > 1:  # 確保表格數據中包含行
        # 使用 Paragraph 處理自動换行
        wrapped_data = []
        
        # 處理表頭，确保應用正確的樣式
        header_row = [Paragraph(text, ParagraphStyle(name='Header', fontName='SimSun', fontSize=12, textColor=HexColor('#FFFFFF'), alignment=TA_CENTER)) for text in table_data[0]]
        wrapped_data.append(header_row)
        
        # 處理表格内容
        for row in table_data[1:]:
            wrapped_row = [Paragraph(text, normal_style) for text in row]
            wrapped_data.append(wrapped_row)
        
        table = Table(wrapped_data, colWidths=[1.3*inch, 1.1*inch, 3.7*inch, 1.4*inch])
        table.setStyle(table_style)
        
        elements.append(table)
    else:
        elements.append(Paragraph("没有数据", normal_style))

    elements.append(Spacer(1, 0.5 * inch))

# 添加尾部 Call to Action
elements.append(Spacer(1, 0.5 * inch))
elements.append(Paragraph(cta_text, cta_style))

# 生成PDF
pdf.build(elements)
print(f"PDF生成成功，已保存为 {pdf_file}")

# 合併PDF
pdfs_to_merge = ['中国快讯.pdf', 'CCRQM.pdf']  # 需合并的PDF文件列表
output_pdf_file = "CCRQM 沪深上市企业风险事件早报.pdf"

merger = PdfMerger()

for pdf in pdfs_to_merge:
    merger.append(pdf)

# 保存合併后的PDF文件
merger.write(output_pdf_file)
merger.close()

print(f"PDF文件已成功合併，並保存為 {output_pdf_file}")