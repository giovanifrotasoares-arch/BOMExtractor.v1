import streamlit as st
from streamlit_cropper import st_cropper
import fitz  # PyMuPDF
import pandas as pd
import io
import utils

# Configurations
st.set_page_config(page_title="Extrator Visual BOM", layout="wide", page_icon="🚜")
st.title("🚜 Extrator Automotivo BOM - Inteligência PDF")

# Session state initialization
if 'pdf_bytes' not in st.session_state:
    st.session_state.pdf_bytes = None
if 'page_count' not in st.session_state:
    st.session_state.page_count = 0
if 'extracted_tables' not in st.session_state:
    st.session_state.extracted_tables = []

st.sidebar.markdown("### Configuração")
uploaded_file = st.sidebar.file_uploader("1. Faça o Upload do PDF", type=['pdf'])

if uploaded_file is not None:
    st.session_state.pdf_bytes = uploaded_file.read()
    
    # Calculate pages
    doc = fitz.open(stream=st.session_state.pdf_bytes, filetype="pdf")
    st.session_state.page_count = doc.page_count
    
    page_num = st.sidebar.number_input("Página Atual", min_value=1, max_value=st.session_state.page_count, value=1) - 1
    
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Extração")
    st.sidebar.info("Ajuste o tamanho e a posição do Retângulo Vermelho em cima da tabela na tela central. Em seguida, clique em Extrair Área.")
    
    if st.sidebar.button("Extrair Área do Quadro Vermelho"):
        extrair = True
    else:
        extrair = False
    
    # Render page to PIL Image
    img, pt_w, pt_h, img_w, img_h = utils.get_page_image(st.session_state.pdf_bytes, page_num)
    
    col1, col2 = st.columns([2.5, 1])
    
    with col1:
        st.subheader("Documento e Seleção")
        # Visual Drag/Drop Bounding Box with Streamlit Cropper
        box = st_cropper(img, realtime_update=True, box_color='#FF0000', aspect_ratio=None, return_type='box')
    
    with col2:
        st.subheader("Processamento")
        
        # Se o botao da sidebar foi clicado
        if extrair:
            with st.spinner("Decodificando texto do PDF..."):
                df = utils.extract_table_from_bbox(
                    st.session_state.pdf_bytes, page_num, box, 
                    pt_w, pt_h, img_w, img_h
                )
                if df is not None and not df.empty:
                    st.session_state.extracted_tables.append(df)
                    st.success(f"Capturadas {len(df)} linhas limpas!")
                else:
                    st.error("Não foi possível ler texto nessa área.")
        
        st.write(f"Tabelas Prontas na Fila: **{len(st.session_state.extracted_tables)}**")
        
        if len(st.session_state.extracted_tables) > 0:
            st.markdown("---")
            st.write("Prévia da Última Aba:")
            st.dataframe(st.session_state.extracted_tables[-1].head(5), use_container_width=True)
            
            # Export to Excel
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                for i, df in enumerate(st.session_state.extracted_tables):
                    sheet_name = f"Tabela_{i+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    utils.format_excel(writer, sheet_name)
            
            st.download_button(
                label="📁 Download Planilha Excel (Limpa & Formatada)",
                data=excel_buffer.getvalue(),
                file_name="BOM_Dados_Gerais.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Limpar Dados Extraídos"):
                st.session_state.extracted_tables = []
                st.rerun()
                
else:
    st.info("👈 Use o painel lateral para carregar um documento de Engenharia e começar a análise visual de Custos.")
