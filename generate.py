import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import os
import yaml
import sys
from pathlib import Path
from urllib.parse import urlparse
from colorama import init, Fore, Style
from PIL import Image, ImageOps, ImageFilter, ImageEnhance
import hashlib


init()


def log_info(message):
    print(f"{Fore.GREEN}[INFO]{Style.RESET_ALL} {message}")


def log_warn(message):
    print(f"{Fore.YELLOW}[WARN]{Style.RESET_ALL} {message}")


def log_error(message):
    print(f"{Fore.RED}[ERROR]{Style.RESET_ALL} {message}")


def load_config(config_file):
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        log_error(f"Error loading config file: {e}")
        sys.exit(1)


def get_cache_path(original_path, prefix):
    """生成缓存文件路径"""

    file_stat = os.stat(original_path)
    file_hash = hashlib.md5(
        f"{original_path}_{file_stat.st_mtime}".encode()
    ).hexdigest()
    cache_dir = Path("cache")
    cache_dir.mkdir(exist_ok=True)
    return cache_dir / f"{prefix}_{file_hash}.png"


def process_background_image(image_path):
    """处理背景图片为勾线效果"""
    try:

        cache_path = get_cache_path(image_path, "bg")
        if cache_path.exists():
            return str(cache_path)

        img = Image.open(image_path)

        img = img.convert("L")

        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)

        img = img.filter(ImageFilter.FIND_EDGES)

        img = ImageOps.invert(img)

        enhancer = ImageEnhance.Brightness(img)
        img = enhancer.enhance(2.2)

        width, height = img.size
        crop_margin = 5
        img = img.crop(
            (crop_margin, crop_margin, width - crop_margin, height - crop_margin)
        )

        img = img.convert("RGBA")

        data = img.getdata()

        new_data = []
        for item in data:

            new_data.append((item[0], item[1], item[2], 128))

        img.putdata(new_data)

        img.save(cache_path, "PNG")
        return str(cache_path)
    except Exception as e:
        log_error(f"Error processing background image: {e}")
        return None


def process_logo(image_path):
    """处理logo，去除背景"""
    try:

        cache_path = get_cache_path(image_path, "logo")
        if cache_path.exists():
            return str(cache_path)

        img = Image.open(image_path)

        img = img.convert("RGBA")

        data = img.getdata()

        new_data = []
        for item in data:

            if item[0] > 240 and item[1] > 240 and item[2] > 240:
                new_data.append((255, 255, 255, 0))
            else:
                new_data.append(item)

        img.putdata(new_data)

        img.save(cache_path, "PNG")
        return str(cache_path)
    except Exception as e:
        log_error(f"Error processing logo: {e}")
        return None


def process_logo_edges(image_path):
    """处理logo为勾线效果，线条更明显"""
    try:

        cache_path = get_cache_path(image_path, "logo_edges")
        if cache_path.exists():
            return str(cache_path)

        img = Image.open(image_path)

        img = img.convert("L")

        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(4.0)

        img = img.filter(ImageFilter.FIND_EDGES)

        img = ImageOps.invert(img)

        enhancer = ImageEnhance.Brightness(img)
        img = enhancer.enhance(1.2)

        width, height = img.size
        crop_margin = 5
        img = img.crop(
            (crop_margin, crop_margin, width - crop_margin, height - crop_margin)
        )

        img = img.convert("RGBA")

        data = img.getdata()

        new_data = []
        for item in data:

            new_data.append((item[0], item[1], item[2], 180))

        img.putdata(new_data)

        img.save(cache_path, "PNG")
        return str(cache_path)
    except Exception as e:
        log_error(f"Error processing logo edges: {e}")
        return None


def draw_rounded_rect(pdf, x, y, w, h, r, style=""):
    """绘制圆角矩形
    x, y: 左上角坐标
    w, h: 宽度和高度
    r: 圆角半径
    style: 填充样式，'F'为填充，'D'为描边，'FD'为填充+描边
    """

    pdf.ellipse(x + r, y + r, r * 2, r * 2, style)
    pdf.ellipse(x + w - r * 2, y + r, r * 2, r * 2, style)
    pdf.ellipse(x + r, y + h - r * 2, r * 2, r * 2, style)
    pdf.ellipse(x + w - r * 2, y + h - r * 2, r * 2, r * 2, style)

    pdf.rect(x + r, y, w - r * 2, h, style)
    pdf.rect(x, y + r, w, h - r * 2, style)


def create_redeem_cards(config):
    fonts_dir = Path("fonts")

    excel_file = config["input"]["excel_file"]
    output_pdf = config["output"]["pdf_file"]
    album_name = config["card"]["album_name"]
    additional_info = config["card"]["additional_info"]

    grid_cols = config["layout"]["grid"]["columns"]
    grid_rows = config["layout"]["grid"]["rows"]
    card_width = config["layout"]["card"]["width"]
    card_height = config["layout"]["card"]["height"]
    margin = config["layout"]["card"]["margin"]
    font_size = config["layout"]["card"]["font_size"]

    bg_path = None
    if "background" in config["card"] and config["card"]["background"]:
        original_bg_path = config["card"]["background"]
        if os.path.exists(original_bg_path):
            bg_path = process_background_image(original_bg_path)
            if not bg_path:
                log_warn("Using original background image due to processing error")
                bg_path = original_bg_path
        else:
            log_warn(f"Background image not found: {original_bg_path}")

    personal_logo_path = None
    if "personal_logo" in config["card"] and config["card"]["personal_logo"]:
        original_personal_logo_path = config["card"]["personal_logo"]
        if os.path.exists(original_personal_logo_path):
            personal_logo_path = process_logo_edges(original_personal_logo_path)
            if not personal_logo_path:
                log_warn("Using original personal logo due to processing error")
                personal_logo_path = original_personal_logo_path
        else:
            log_warn(f"Personal logo not found: {original_personal_logo_path}")

    logo_path = "logo/dl-n-88_2.jpg"
    processed_logo_path = None
    if os.path.exists(logo_path):
        processed_logo_path = process_logo(logo_path)
        if not processed_logo_path:
            log_warn("Using original logo due to processing error")
            processed_logo_path = logo_path
    else:
        log_warn(f"Logo not found: {logo_path}")

    df = pd.read_excel(excel_file)

    if df.empty:
        log_error("Excel file is empty")
        sys.exit(1)

    columns = df.columns.tolist()
    if len(columns) < 1:
        log_error("Excel file must have at least one column")
        sys.exit(1)

    valid_rows = []
    skipped_count = 0
    previously_redeemed_count = 0

    for index, row in df.iterrows():
        skip_flag = False

        if (
            len(columns) > 1
            and pd.notna(row[columns[1]])
            and str(row[columns[1]]).strip() != ""
        ):
            try:
                user_id = int(float(str(row[columns[1]]).strip()))
                log_warn(
                    f"Skipping row {index + 2}: Previously redeemed (id: {user_id})"
                )
            except ValueError:
                log_warn(
                    f"Skipping row {index + 2}: Previously redeemed (id: {row[columns[1]]})"
                )
            previously_redeemed_count += 1
            skip_flag = True

        if len(columns) > 3 and str(row[columns[3]]) == "skip":
            log_warn(f"Skipping row {index + 2}: Marked as skip")
            skipped_count += 1
            skip_flag = True

        if skip_flag:
            continue

        valid_rows.append((index, row))

    total_cards = len(valid_rows)
    cards_per_page = grid_cols * grid_rows
    total_pages = (total_cards + cards_per_page - 1) // cards_per_page

    log_info(f"Total rows in Excel: {len(df)}")
    log_info(f"Skipped rows (marked as skip): {skipped_count}")
    log_info(f"Skipped rows (previously redeemed): {previously_redeemed_count}")
    log_info(f"Valid redeem codes: {total_cards}")
    log_info(f"Cards per page: {cards_per_page} ({grid_cols}x{grid_rows})")
    log_info(f"Total pages: {total_pages}")

    pdf = FPDF()
    pdf.set_auto_page_break(False)

    page_width = 210
    page_height = 297
    card_width_mm = card_width * 10
    card_height_mm = card_height * 10
    margin_mm = margin * 10

    total_cards_width = grid_cols * card_width_mm
    total_cards_height = grid_rows * card_height_mm
    horizontal_gap = (
        (page_width - 2 * margin_mm - total_cards_width) / (grid_cols - 1)
        if grid_cols > 1
        else 0
    )
    vertical_gap = (
        (page_height - 2 * margin_mm - total_cards_height) / (grid_rows - 1)
        if grid_rows > 1
        else 0
    )

    start_x = (
        margin_mm
        + (
            page_width
            - 2 * margin_mm
            - total_cards_width
            - (grid_cols - 1) * horizontal_gap
        )
        / 2
    )
    start_y = (
        margin_mm
        + (
            page_height
            - 2 * margin_mm
            - total_cards_height
            - (grid_rows - 1) * vertical_gap
        )
        / 2
    )

    pdf.add_font("NotoSansSC", "", str(fonts_dir / "NotoSansSC-SemiBold.ttf"))
    pdf.add_font("SFMono", "", str(fonts_dir / "SFMono-Regular.ttf"))

    current_page = -1

    for i, (original_index, row) in enumerate(valid_rows):
        page_num = i // cards_per_page
        if page_num != current_page:
            pdf.add_page()
            current_page = page_num

        pos_on_page = i % cards_per_page
        col = pos_on_page % grid_cols
        row_num = pos_on_page // grid_cols

        x = start_x + col * (card_width_mm + horizontal_gap)
        y = start_y + row_num * (card_height_mm + vertical_gap)

        pdf.set_draw_color(230, 230, 230)
        pdf.set_line_width(0.2)
        pdf.set_dash_pattern(dash=2, gap=1)
        pdf.rect(x, y, card_width_mm, card_height_mm)
        pdf.set_dash_pattern()

        if bg_path:
            pdf.image(bg_path, x, y, w=card_width_mm, h=card_height_mm)

        pdf.set_font("NotoSansSC", "", 8)
        pdf.set_text_color(150, 150, 150)
        row_number = original_index + 2
        pdf.set_xy(x + card_width_mm - 20, y + 2)
        pdf.cell(10, 5, f"#{row_number}", align="R")

        if personal_logo_path:
            pdf.image(personal_logo_path, x + card_width_mm - 10, y + 2, w=10)

        if processed_logo_path:
            pdf.image(processed_logo_path, x + 2, y + 2, w=10)

        card_center_y = y + card_height_mm / 2

        redeem_code = str(row[columns[0]])
        product_number = str(row[columns[2]]) if len(columns) > 2 else ""

        pdf.set_font("NotoSansSC", "", 9)
        pdf.set_text_color(150, 150, 150)
        pdf.set_xy(x + 5, card_center_y - 20)
        pdf.cell(
            card_width_mm - 10,
            5,
            "电子版兑换码",
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
            align="C",
        )

        pdf.set_fill_color(0, 0, 0)

        pdf.set_font("SFMono", "", 14)
        code_width = pdf.get_string_width(redeem_code)

        bg_width = code_width + 20

        bg_x = x + (card_width_mm - bg_width) / 2
        pdf.rect(bg_x, card_center_y - 10, bg_width, 8, style="F")

        pdf.set_font("SFMono", "", 14)
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(x + 5, card_center_y - 10)
        pdf.cell(
            card_width_mm - 10,
            8,
            redeem_code,
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
            align="C",
        )

        pdf.set_font("NotoSansSC", "", 11)
        pdf.set_text_color(150, 150, 150)
        pdf.set_xy(x + 5, card_center_y + 2)
        pdf.cell(
            card_width_mm - 10,
            5,
            album_name,
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
            align="C",
        )

        pdf.set_font("SFMono", "", 9)
        pdf.set_xy(x + 5, card_center_y + 8)
        pdf.cell(
            card_width_mm - 10,
            5,
            product_number,
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT,
            align="C",
        )

        y_offset = 14
        for info in additional_info:
            pdf.set_xy(x + 5, card_center_y + y_offset)

            if any(
                domain in info.lower()
                for domain in [".com", ".cn", ".net", ".org", ".io", "dizzylab.com"]
            ):
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("NotoSansSC", "", 9)
            else:
                pdf.set_text_color(150, 150, 150)
                pdf.set_font("NotoSansSC", "", 9)
            pdf.cell(
                card_width_mm - 10,
                5,
                info,
                new_x=XPos.LMARGIN,
                new_y=YPos.NEXT,
                align="C",
            )
            y_offset += 6

    pdf.output(output_pdf)
    log_info(f"PDF generated successfully: {output_pdf}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        log_error("Usage: python generate_redeem_cards.py <config_file>")
        sys.exit(1)

    config_file = sys.argv[1]
    if not os.path.exists(config_file):
        log_error(f"Config file {config_file} does not exist.")
        sys.exit(1)

    config = load_config(config_file)
    create_redeem_cards(config)
