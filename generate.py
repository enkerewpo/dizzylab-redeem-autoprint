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
    previously_redeemed_count = 0
    
    for index, row in df.iterrows():
            
        skip_flag = False
            
        # Skip if second column has any value (previously redeemed)
        if len(columns) > 1 and pd.notna(row[columns[1]]) and str(row[columns[1]]).strip() != "":
            try:
                user_id = int(float(str(row[columns[1]]).strip()))
                print(f"Skipping row {index + 2}: Previously redeemed (id: {user_id})")
            except ValueError:
                print(f"Skipping row {index + 2}: Previously redeemed (id: {row[columns[1]]})")
            previously_redeemed_count += 1
            skip_flag = True
        
        # Skip if marked as skip in fourth column
        if len(columns) > 3 and str(row[columns[3]]) == "skip":
            print(f"Skipping row {index + 2}: Marked as skip")
            skipped_count += 1
            skip_flag = True
            # continue
            
        if skip_flag:
            continue
            
        valid_rows.append(row)
    
    # Calculate total number of cards needed
    total_cards = len(valid_rows)
    cards_per_page = grid_cols * grid_rows
    total_pages = (total_cards + cards_per_page - 1) // cards_per_page
    
    print(f"Total rows in Excel: {len(df)}")
    print(f"Skipped rows (marked as skip): {skipped_count}")
    print(f"Skipped rows (previously redeemed): {previously_redeemed_count}")
    print(f"Valid redeem codes: {total_cards}")
    print(f"Cards per page: {cards_per_page} ({grid_cols}x{grid_rows})")
    print(f"Total pages: {total_pages}")
    
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
        
        # Calculate how many complete rows we need for this page
        cards_on_this_page = len(page_cards)
        complete_rows = (cards_on_this_page + grid_cols - 1) // grid_cols
        
        for j in range(0, cards_on_this_page, grid_cols):
            row = page_cards[j:j + grid_cols]
            # Only fill remaining cells in the last row if needed
            if j // grid_cols == complete_rows - 1:
                while len(row) < grid_cols:
                    row.append(Paragraph("", basestyle))
            grid_data.append(row)
        
        # Don't add empty rows if we're on the last page
        if i + grid_cols * grid_rows < len(cards):
            while len(grid_data) < grid_rows:
                grid_data.append([Paragraph("", basestyle)] * grid_cols)
        
        table = Table(
            grid_data,
            colWidths=[card_width*cm]*grid_cols,
            rowHeights=[card_height*cm]*len(grid_data)  # Only use as many rows as we have data
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
    # print actual pages
    # read pdf file and count pages
    from PyPDF2 import PdfReader
    reader = PdfReader(config['output']['pdf_file'])
    print(f"Actual PDF pages: {len(reader.pages)}")
