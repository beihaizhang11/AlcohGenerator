import re
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# 1. 字体配置 Font Configuration
FONT_PATH = "/workspace/SimHei.ttf"
FONT_NAME = "SimHei"

def register_font():
    try:
        # 注册字体
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
        print(f"成功注册字体: {FONT_PATH}")
        return True
    except Exception as e:
        print(f"无法注册字体: {e}")
        return False

# 2. 解析 Markdown Parsing Markdown
def parse_menu(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    sections = re.split(r'^###\s+', content, flags=re.MULTILINE)
    
    drinks = []
    
    for section in sections:
        if not section.strip():
            continue
            
        lines = section.strip().split('\n')
        title = lines[0].strip()
        
        details = {}
        
        # 简单的关键字提取
        recipe_pattern = re.compile(r'\*\s*\*\*配方\*\*[：:]\s*(.+)')
        
        recipe = ""
        
        for line in lines[1:]:
            line = line.strip()
            r_match = recipe_pattern.match(line)
            
            if r_match:
                recipe = r_match.group(1)
                
        drinks.append({
            'title': title,
            'recipe': recipe
        })
        
    return drinks

# 3. 生成 PDF Generation
def generate_pdf(drinks, filename):
    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=1.5*cm,
        leftMargin=1.5*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )
    
    if not register_font():
        print("警告：未找到中文字体，中文将无法显示。")
        body_font = "Helvetica"
        title_font = "Helvetica-Bold"
    else:
        body_font = FONT_NAME
        title_font = FONT_NAME

    styles = getSampleStyleSheet()
    
    # 标题样式 - 调整字体大小和间距以适应单页 6 个
    title_style = ParagraphStyle(
        'DrinkTitle',
        parent=styles['Heading2'],
        fontName=title_font,
        fontSize=16, # 稍微加大
        leading=22,
        spaceAfter=6,
        textColor=colors.HexColor('#2c3e50'),
        alignment=TA_CENTER
    )
    
    # 正文样式 - 居中
    text_style = ParagraphStyle(
        'Content',
        parent=styles['BodyText'],
        fontName=body_font,
        fontSize=12, # 稍微加大
        leading=18,
        textColor=colors.HexColor('#555555'),
        alignment=TA_CENTER
    )

    story = []
    
    # 大标题
    story.append(Paragraph("Signature Cocktails", ParagraphStyle(
        'MainTitle', 
        parent=styles['Title'], 
        fontName=title_font, 
        fontSize=28, 
        spaceAfter=10,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#34495e')
    )))
    story.append(Paragraph("精选酒单", ParagraphStyle(
        'SubTitle', 
        parent=styles['Normal'], 
        fontName=title_font, 
        fontSize=14, 
        alignment=TA_CENTER,
        spaceAfter=20,
        textColor=colors.HexColor('#7f8c8d')
    )))

    # 构建表格数据 - 单列
    data = []
    
    for i, drink in enumerate(drinks):
        card_content = []
        
        # 1. Title
        card_content.append(Paragraph(drink['title'], title_style))
        
        # 2. Separator Line
        card_content.append(Spacer(1, 4))
        
        # 3. Details - 只显示配方
        if drink['recipe']:
            # 不再使用 colored label，直接居中显示内容，更简约
            p = Paragraph(f"{drink['recipe']}", text_style)
            card_content.append(p)
            
        # 增加单元格内的底部间距
        card_content.append(Spacer(1, 10))
        
        # 直接添加一行（包含一个单元格）
        data.append([card_content])

    # Calculation for card size
    page_width = A4[0] - 3*cm
    col_width = page_width 
    
    # Outer Table Style
    t = Table(data, colWidths=[col_width], spaceBefore=0, spaceAfter=0)
    
    # Table Styling
    style_cmds = [
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), # 垂直居中
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),  # 水平居中
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 15), # 增加行间距
        ('BOTTOMPADDING', (0,0), (-1,-1), 15),
    ]
    
    # 分割线：只在行与行之间画线
    # 我们可以使用 LINEBELOW
    style_cmds.append(('LINEBELOW', (0,0), (-1,-2), 0.5, colors.HexColor('#ecf0f1'))) # 除最后一行外的下划线

    t.setStyle(TableStyle(style_cmds))
    
    story.append(t)
    
    doc.build(story)
    print(f"PDF 已生成: {filename}")

if __name__ == "__main__":
    menu_file = "/workspace/menu.md"
    output_file = "/workspace/cocktail_menu.pdf"
    
    if os.path.exists(menu_file):
        drinks_data = parse_menu(menu_file)
        generate_pdf(drinks_data, output_file)
    else:
        print("Menu file not found.")
