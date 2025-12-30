import re
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Frame, PageTemplate
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
        method_pattern = re.compile(r'\*\s*\*\*做法\*\*[：:]\s*(.+)')
        garnish_pattern = re.compile(r'\*\s*\*\*点睛\*\*[：:]\s*(.+)')
        
        recipe = ""
        method = ""
        garnish = ""
        
        for line in lines[1:]:
            line = line.strip()
            r_match = recipe_pattern.match(line)
            m_match = method_pattern.match(line)
            g_match = garnish_pattern.match(line)
            
            if r_match:
                recipe = r_match.group(1)
            elif m_match:
                method = m_match.group(1)
            elif g_match:
                garnish = g_match.group(1)
                
        drinks.append({
            'title': title,
            'recipe': recipe,
            'method': method,
            'garnish': garnish
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
    
    # 标题样式
    title_style = ParagraphStyle(
        'DrinkTitle',
        parent=styles['Heading2'],
        fontName=title_font,
        fontSize=14,
        leading=20,
        spaceAfter=10,
        textColor=colors.HexColor('#2c3e50'),
        alignment=TA_CENTER
    )
    
    # 正文样式
    text_style = ParagraphStyle(
        'Content',
        parent=styles['BodyText'],
        fontName=body_font,
        fontSize=10,
        leading=16,
        textColor=colors.black
    )

    story = []
    
    # 大标题
    story.append(Paragraph("Signature Cocktails", ParagraphStyle(
        'MainTitle', 
        parent=styles['Title'], 
        fontName=title_font, 
        fontSize=28, 
        spaceAfter=10,
        textColor=colors.HexColor('#34495e')
    )))
    story.append(Spacer(1, 10))
    story.append(Paragraph("精选酒单", ParagraphStyle(
        'SubTitle', 
        parent=styles['Normal'], 
        fontName=title_font, 
        fontSize=14, 
        alignment=TA_CENTER,
        spaceAfter=30,
        textColor=colors.HexColor('#7f8c8d')
    )))

    # 构建表格数据
    data = []
    row = []
    
    for i, drink in enumerate(drinks):
        card_content = []
        
        # 1. Title
        card_content.append(Paragraph(drink['title'], title_style))
        
        # 2. Separator Line (using a Table within the cell or just spacing)
        # Using a Paragraph with a border is hard, let's just use space.
        card_content.append(Spacer(1, 6))
        
        # 3. Details
        # We construct a string with HTML tags for colors/bold
        # Note: ReportLab's HTML parser is simple.
        
        if drink['recipe']:
            # SimHei doesn't have a bold variant usually, but reportlab might synthesize or we just use color
            p = Paragraph(f"<font color='#8e44ad'><b>配方</b></font> | {drink['recipe']}", text_style)
            card_content.append(p)
            
        if drink['method']:
            p = Paragraph(f"<font color='#2980b9'><b>做法</b></font> | {drink['method']}", text_style)
            card_content.append(p)
            
        if drink['garnish']:
            p = Paragraph(f"<font color='#27ae60'><b>点睛</b></font> | {drink['garnish']}", text_style)
            card_content.append(p)
            
        # Add some padding at bottom of content inside the cell
        card_content.append(Spacer(1, 10))
        
        row.append(card_content)
        
        if len(row) == 2:
            data.append(row)
            row = []
            
    if row:
        row.append([]) # Empty cell placeholder
        data.append(row)

    # Calculation for card size
    page_width = A4[0] - 3*cm
    col_width = page_width / 2 - 0.5*cm # Leave some gap
    
    # Outer Table Style
    t = Table(data, colWidths=[col_width, col_width], spaceBefore=0, spaceAfter=0)
    
    # Table Styling for "Card" look
    # (col, row)
    style_cmds = [
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 12),
        ('RIGHTPADDING', (0,0), (-1,-1), 12),
        ('TOPPADDING', (0,0), (-1,-1), 12),
        ('BOTTOMPADDING', (0,0), (-1,-1), 12),
    ]
    
    # Apply borders to each cell individually to make them look like cards
    # Or just a grid.
    # Let's do a grid with some spacing.
    # Actually, to make them look like cards separated by space, we can't easily use a single Table with borders unless we use spacing.
    # ReportLab Tables can have 'cellStyles' but standard grids are continuous.
    # We will use a simple inner grid line for minimalism.
    
    style_cmds.append(('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#ecf0f1'))) # Light grey grid
    style_cmds.append(('BACKGROUND', (0,0), (-1,-1), colors.white))

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
