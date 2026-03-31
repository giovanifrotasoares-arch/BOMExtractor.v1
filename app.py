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
        cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = thin_border
    worksheet.freeze_panes = "A2"
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            cell.border = thin_border
            if cell.row > 1: cell.alignment = Alignment(vertical="center", horizontal="center")
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        worksheet.column_dimensions[column].width = min((max_length + 2), 45)

# ==========================================
# 2. INTERFACE APP
# ==========================================
st.set_page_config(page_title="Extrator BOM Web", layout="wide", page_icon="🚜")
st.title("🚜 Extrator Automotivo BOM - Robô Visual")

st.markdown("""
<div style='background-color: #2F4F4F; padding: 15px; border-radius: 10px; color: white; font-size: 14px;'>
<strong>💡 Acesso Liberado e Sem Quedas:</strong> Desenhe o enquadramento no Mapa Geral, confira a leitura na Lupa abaixo (ela tem o Zoom nativo de 4K, enquanto o Mapa foi otimizado) e Extraia! 
</div>
""", unsafe_allow_html=True)

if 'pdf_bytes' not in st.session_state: st.session_state.pdf_bytes = None
if 'page_count' not in st.session_state: st.session_state.page_count = 0
if 'extracted_tables' not in st.session_state: st.session_state.extracted_tables = []

st.sidebar.markdown("### Painel de Nuvem")
uploaded_file = st.sidebar.file_uploader("1. Faça o Upload do PDF (Max 200MB)", type=['pdf'])

if uploaded_file is not None:
    st.session_state.pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=st.session_state.pdf_bytes, filetype="pdf")
    st.session_state.page_count = doc.page_count
    doc.close()
    
    page_num = st.sidebar.number_input("Página da Folha", min_value=1, max_value=st.session_state.page_count, value=1) - 1
    
    img, pt_w, pt_h, img_w, img_h = get_page_image(st.session_state.pdf_bytes, page_num)
    
    st.markdown("---")
    st.subheader("1. Mapa Geral (Selecione o Corte)")
    
    box = st_cropper(img, realtime_update=True, box_color='#FF0000', aspect_ratio=None, return_type='box')
    
    st.markdown("---")
    st.subheader("🔍 2. Lupa Fotográfica Exata (Confirme o enquadramento aqui)")
    
    lupa_img = get_highres_crop(st.session_state.pdf_bytes, page_num, box, img_w, img_h)
    st.image(lupa_img)
    
    st.markdown("---")
    st.subheader("⚙️ 3. Ação & Exportação Custeio Automático")
    
    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("▶️ Extrair Tabela Focalizada na Lupa", type="primary"):
            with st.spinner("Analisando micro-medidas de colunas..."):
                try:
                    df = extract_table_from_bbox(st.session_state.pdf_bytes, page_num, box, pt_w, pt_h, img_w, img_h)
                    if df is not None and not df.empty:
                        st.session_state.extracted_tables.append(df)
                        st.success(f"Nuvem validou! {len(df)} linhas limpas importadas.")
                except ValueError as ve:
                    if str(ve) == "SCANNED_PDF": st.error("🚨 PDF ESCANEADO (Bateu na Barreira!): O PDF não é vetor, logo o motor não consegue separar letras da tinta branca da foto. Exportar arquivos originais por favor!")
                    elif str(ve) == "WRONG_BBOX": st.error("🚨 CAIXA VAZIA (Z-Index Erro): A lupa ou está isolando uma parte gráfica vazia (que não tem letrinhas), ou um bug de margem do AutoCAD atacou. Se estiver pegando texto, avise o Dev.")
                    else: st.error(f"Erro Incomum: {str(ve)}")
                except Exception as e:
                    st.error(f"Erro Crítico Total: {str(e)}")

    with col_b:
        st.write(f"Estocadas na Memória Principal: **{len(st.session_state.extracted_tables)} aba(s)**")
        if st.button("🗑️ Esvaziar Extrator (Começar do Zero)"):
            st.session_state.extracted_tables = []
            st.rerun()

    if len(st.session_state.extracted_tables) > 0:
        st.markdown("### Prévia Oficial:")
        st.dataframe(st.session_state.extracted_tables[-1].head(10))
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for i, df in enumerate(st.session_state.extracted_tables):
                sheet_name = f"Tabela_{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel(writer, sheet_name)
        
        st.download_button(
            label="📁 Fazer Download da Planilha Excel OficialJD (Formatada)",
            data=excel_buffer.getvalue(),
            file_name="BOM_Dados_Gerais_Web_Sniper.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("👈 Operação aguardando PDF Vetorial de Engenharia no menu lateral.")

