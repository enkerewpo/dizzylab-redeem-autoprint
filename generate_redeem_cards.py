import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas

import os
import yaml
import sys

def register_fonts():
    sf_mono_path = os.path.join(os.path.dirname(__file__), 'SFMono-Regular.ttf')
    if os.path.exists(sf_mono_path):
        pdfmetrics.registerFont(TTFont('SFMono', sf_mono_path))
    else:
        print(f"Warning: SF Mono font not found at {sf_mono_path}")
        sys.exit(1)
    
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    
    styles = getSampleStyleSheet()
    basestyle = ParagraphStyle(
        'BaseStyle',
        parent=styles['Normal'],
        fontName='STSong-Light',
        fontSize=8,
        leading=12,
        alignment=1
    )
    mono_style = ParagraphStyle(
        'MonoStyle',
        parent=styles['Normal'],
        fontName='SFMono',
        fontSize=14,
        leading=18,
        alignment=1,
        spaceBefore=12,
        spaceAfter=12,
        textColor=colors.HexColor('#000000')
    )
    return basestyle, mono_style

def load_config(config_file):
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        print(f"Error loading config file: {e}")
        sys.exit(1)

def create_redeem_cards(config):
    basestyle, mono_style = register_fonts()
    
    excel_file = config['input']['excel_file']
    output_pdf = config['output']['pdf_file']
    album_name = config['card']['album_name']
    additional_info = config['card']['additional_info']
    
    grid_cols = config['layout']['grid']['columns']
    grid_rows = config['layout']['grid']['rows']
    card_width = config['layout']['card']['width']
    card_height = config['layout']['card']['height']
    margin = config['layout']['card']['margin']
    font_size = config['layout']['card']['font_size']
    
    df = pd.read_excel(excel_file)
    
    if df.empty:
        print("Error: Excel file is empty")
        sys.exit(1)
    
    columns = df.columns.tolist()
    if len(columns) < 1:
        print("Error: Excel file must have at least one column")
        sys.exit(1)
    
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        rightMargin=margin*cm,
        leftMargin=margin*cm,
        topMargin=margin*cm,
        bottomMargin=margin*cm
    )
    
    # Filter out skipped rows first
    valid_rows = []
    skipped_count = 0
    for _, row in df.iterrows():
        if len(columns) > 3 and str(row[columns[3]]) == "skip":
            skipped_count += 1
            continue
        valid_rows.append(row)
    
    # Calculate total number of cards needed
    total_cards = len(valid_rows)
    total_pages = (total_cards + grid_cols * grid_rows - 1) // (grid_cols * grid_rows)
    
    print(f"Total rows in Excel: {len(df)}")
    print(f"Skipped rows: {skipped_count}")
    print(f"Valid redeem codes: {total_cards}")
    print(f"Generated PDF pages: {total_pages}")
    
    cards = []
    for row in valid_rows:
        redeem_code = str(row[columns[0]])
        product_number = str(row[columns[2]]) if len(columns) > 2 else ""
        
        card_content = f"""
        <font name="STSong-Light" color="#666666">电子版兑换码（dizzylab平台）</font><br/>
        <br/>
        <font name="SFMono" size="14" color="#000000">{redeem_code}</font><br/>
        <br/>
        <font name="STSong-Light" color="#666666">专辑名称: {album_name}</font><br/>
        <font name="STSong-Light" color="#666666">制品号: {product_number}</font><br/>
        """
        
        for info in additional_info:
            card_content += f'<font name="STSong-Light" color="#666666">{info}</font><br/>'
        
        cards.append(Paragraph(card_content, basestyle))
    
    all_tables = []
    for i in range(0, len(cards), grid_cols * grid_rows):
        page_cards = cards[i:i + grid_cols * grid_rows]
        grid_data = []
        for j in range(0, len(page_cards), grid_cols):
            row = page_cards[j:j + grid_cols]
            while len(row) < grid_cols:
                row.append(Paragraph("", basestyle))
            grid_data.append(row)
        
        while len(grid_data) < grid_rows:
            grid_data.append([Paragraph("", basestyle)] * grid_cols)
        
        table = Table(
            grid_data,
            colWidths=[card_width*cm]*grid_cols,
            rowHeights=[card_height*cm]*grid_rows
        )
        
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey, None, (2, 2)),
            ('LINEBELOW', (0, 0), (-1, -1), 0.5, colors.lightgrey, None, (2, 2)),
            ('LINEABOVE', (0, 0), (-1, -1), 0.5, colors.lightgrey, None, (2, 2)),
            ('LINEBEFORE', (0, 0), (-1, -1), 0.5, colors.lightgrey, None, (2, 2)),
            ('LINEAFTER', (0, 0), (-1, -1), 0.5, colors.lightgrey, None, (2, 2)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('PADDING', (0, 0), (-1, -1), 6),
        ]))
        
        all_tables.append(table)
    
    doc.build(all_tables)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python generate_redeem_cards.py <config_file>")
        sys.exit(1)
    
    config_file = sys.argv[1]
    if not os.path.exists(config_file):
        print(f"Error: Config file {config_file} does not exist.")
        sys.exit(1)
    
    config = load_config(config_file)
    create_redeem_cards(config)
    print(f"PDF generated successfully: {config['output']['pdf_file']}") 