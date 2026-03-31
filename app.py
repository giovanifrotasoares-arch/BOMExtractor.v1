import streamlit as st
from streamlit_cropper import st_cropper
import fitz
import pandas as pd
import io
from PIL import Image
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# 1. MOTOR (Coração Anti-Estouro de RAM)
# ==========================================
def get_page_image(pdf_bytes, page_num):
    # Zoom Leve = 1.0 (Economiza aprox. 95% da memoria e evita que a Nuvem crashe!)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[page_num]
    mat = fitz.Matrix(1.0, 1.0)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    
    # Salva a matemática das larguras ANTES de executar a faxina do doc.close() !
    pt_width = page.rect.width
    pt_height = page.rect.height
    pix_w = pix.width
    pix_h = pix.height
    
    doc.close()
    
    return img, pt_width, pt_height, pix_w, pix_h

def get_highres_crop(pdf_bytes, page_num, bbox, img_width, img_height):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[page_num]
    
    ratio_x = float(page.rect.width) / float(img_width)
    ratio_y = float(page.rect.height) / float(img_height)
    
    pdf_rect = fitz.Rect(
        bbox['left'] * ratio_x,
        bbox['top'] * ratio_y,
        (bbox['left'] + bbox['width']) * ratio_x,
        (bbox['top'] + bbox['height']) * ratio_y
    )
    
    mat = fitz.Matrix(4.0, 4.0)
    pix = page.get_pixmap(matrix=mat, clip=pdf_rect)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img

def deduplicate_columns(cols):
    seen, new_cols = {}, []
    for c in cols:
        c_str = str(c).strip() if c else "Vazio"
        if not c_str: c_str = "Vazio"
        if c_str in seen:
            seen[c_str] += 1
            new_cols.append(f"{c_str}_{seen[c_str]}")
        else:
            seen[c_str] = 0
            new_cols.append(c_str)
    return new_cols

def extract_table_from_bbox(pdf_bytes, page_num, bbox, pt_width, pt_height, img_width, img_height):
    left, top, width, height = bbox['left'], bbox['top'], bbox['width'], bbox['height']
    right, bottom = left + width, top + height
    pdf_stream = io.BytesIO(pdf_bytes)
    
    with pdfplumber.open(pdf_stream) as pdf:
        page = pdf.pages[page_num]
        ratio_x = float(page.width) / float(img_width)
        ratio_y = float(page.height) / float(img_height)
        
        pdf_bbox = ((left * ratio_x) - 1, (top * ratio_y) - 1, (right * ratio_x) + 1, (bottom * ratio_y) + 1)
        cropped = page.within_bbox(pdf_bbox)
        
        table = cropped.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text", "intersection_y_tolerance": 5})
        
        if not table or len(table) == 0:
            text = cropped.extract_text()
            if text: table = [line.split() for line in text.split('\n') if line.strip()]
        
        if table and len(table) > 0:
            cleaned_table = []
            for row in table:
                c_row = [str(cell).strip() if cell else "" for cell in row]
                if any(cell not in ("", "-", "None") for cell in c_row):
                    c_row = ["" if cell == "-" else cell for cell in c_row]
                    cleaned_table.append(c_row)
                    
            table = cleaned_table
            if len(table) > 1: return pd.DataFrame(table[1:], columns=deduplicate_columns(table[0]))
            elif len(table) == 1: return pd.DataFrame(table)
        
        full_text = page.extract_text()
        if not full_text: raise ValueError("SCANNED_PDF")
        elif len(full_text.strip()) < 5: raise ValueError("SCANNED_PDF")
        else: raise ValueError("WRONG_BBOX")

def format_excel(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in worksheet[1]:
        cell.fill = header_fill; cell.font =
