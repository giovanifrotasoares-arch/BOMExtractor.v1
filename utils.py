import fitz
import pdfplumber
import pandas as pd
from PIL import Image
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# Aumentamos o Zoom para 3.0 (Altíssima Definição) para a Lupa não borrar
def get_page_image(pdf_bytes, page_num, zoom=3.0):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[page_num]
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img, page.rect.width, page.rect.height, pix.width, pix.height
def extract_table_from_bbox(pdf_bytes, page_num, bbox, pt_width, pt_height, img_width, img_height):
    ratio_x = pt_width / img_width
    ratio_y = pt_height / img_height
    
    left = bbox['left']
    top = bbox['top']
    width = bbox['width']
    height = bbox['height']
    right = left + width
    bottom = top + height
    
    # Margem de segurança sutil (expande o quadro 1pt matemático) para não cortar letras nas bordas
    pdf_bbox = (
        (left * ratio_x) - 1,
        (top * ratio_y) - 1,
        (right * ratio_x) + 1,
        (bottom * ratio_y) + 1
    )
    
    with pdfplumber.open(pdf_bytes) as pdf:
        page = pdf.pages[page_num]
        cropped = page.within_bbox(pdf_bbox)
        
        table_settings = {
            "vertical_strategy": "text", 
            "horizontal_strategy": "text",
            "intersection_y_tolerance": 5
        }
        table = cropped.extract_table(table_settings)
        
        if not table or len(table) == 0:
            text = cropped.extract_text()
            if text:
                table = [line.split() for line in text.split('\n') if line.strip()]
        
        if table and len(table) > 0:
            cleaned_table = []
            for row in table:
                c_row = [str(cell).strip() if cell else "" for cell in row]
                if any(cell not in ("", "-", "None") for cell in c_row):
                    c_row = ["" if cell == "-" else cell for cell in c_row]
                    cleaned_table.append(c_row)
                    
            table = cleaned_table
            if len(table) > 1:
                return pd.DataFrame(table[1:], columns=table[0])
            elif len(table) == 1:
                return pd.DataFrame(table)
            
    return None
def format_excel(writer, sheet_name):
    worksheet = writer.sheets[sheet_name]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    
    worksheet.freeze_panes = "A2"
    
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            cell.border = thin_border
            if cell.row > 1:
                cell.alignment = Alignment(vertical="center", horizontal="center")
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = min(adjusted_width, 45)
