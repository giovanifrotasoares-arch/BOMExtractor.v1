import streamlit as st
from streamlit_cropper import st_cropper
import fitz
import pandas as pd
import io
import utils

st.set_page_config(page_title="Extrator Visual BOM", layout="wide", page_icon="🚜")
st.title("🚜 Extrator Automotivo BOM - Inteligência PDF")

st.markdown("""
<div style='background-color: #2F4F4F; padding: 15px; border-radius: 10px; color: white; font-size: 14px;'>
<strong>💡 Lupa de Alta Definição Automática</strong><br>
Para extrair um PDF enorme com letras minúsculas sem errar o quadro, nós construímos uma Zoom Camera!<br>
1. Desenhe um quadrado por cima de onde está a Tabela.<br>
2. A Lupa logo abaixo projetará um Zoom Gigante e Nítido do texto! Ajuste o quadrado até que as palavras na lupa fiquem perfeitamente isoladas, e pressione Extrair.
</div>
""", unsafe_allow_html=True)

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
    
    doc = fitz.open(stream=st.session_state.pdf_bytes, filetype="pdf")
    st.session_state.page_count = doc.page_count
    
    page_num = st.sidebar.number_input("Página Atual", min_value=1, max_value=st.session_state.page_count, value=1) - 1
    
    img, pt_w, pt_h, img_w, img_h = utils.get_page_image(st.session_state.pdf_bytes, page_num)
    
    st.markdown("---")
    st.subheader("1. Mapa Geral (Selecione a Região)")
    
    box = st_cropper(img, realtime_update=True, box_color='#FF0000', aspect_ratio=None, return_type='box')
    
    st.markdown("---")
    st.subheader("🔍 2. Lupa de Precisão (Evita Extração Vazia)")
    st.info("Veja aqui se o quadrado capturou as bordas superiores, inferiores e laterais do texto inteiro.\n Ajuste o quadro acima e o Zoom focará automaticamente!")
    
    left, top, width, height = box['left'], box['top'], box['width'], box['height']
    cropped_preview = img.crop((left, top, left + width, top + height))
    
    st.image(cropped_preview, use_container_width=True)
    
    st.markdown("---")
    st.subheader("⚙️ 3. Processamento e Exportação")
    
    col_a, col_b = st.columns([1, 1])
    
    with col_a:
        if st.button("▶️ Extrair Tabela que aparece na Lupa", type="primary"):
            with st.spinner("Decodificando texto matemático do PDF..."):
                try:
                    df = utils.extract_table_from_bbox(
                        st.session_state.pdf_bytes, page_num, box, 
                        pt_w, pt_h, img_w, img_h
                    )
                    if df is not None and not df.empty:
                        st.session_state.extracted_tables.append(df)
                        st.success(f"Matriz extraída perfeitamente! Capturadas {len(df)} linhas.")
                except ValueError as ve:
                    if str(ve) == "SCANNED_PDF":
                        st.error("🚨 PDF ESCANEADO (IMAGEM) DETECTADO! Este documento é como uma foto (não tem texto vetorial embarcado). O motor só consegue extrair planilhas de desenhos PDF exportados nativamente do software de desenho.")
                    elif str(ve) == "WRONG_BBOX":
                        st.error("🚨 DESALINHAMENTO TÉCNICO: O PDF tem texto livre, mas o software de quem desenhou a planilha inverteu silenciosamente ou recuou o eixo de dimensão das páginas!")
                    else:
                        st.error(f"Erro Inesperado Python: {str(ve)}")
                except Exception as e:
                    st.error(f"Ocorreu um Erro Crítico: {str(e)}")

    with col_b:
        st.write(f"Tabelas Prontas na Fila: **{len(st.session_state.extracted_tables)}**")
        if st.button("🗑️ Limpar Fila (Reset)"):
            st.session_state.extracted_tables = []
            st.rerun()

    if len(st.session_state.extracted_tables) > 0:
        st.markdown("### Prévia da Última Tabela Estocada:")
        st.dataframe(st.session_state.extracted_tables[-1].head(10), use_container_width=True)
        
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for i, df in enumerate(st.session_state.extracted_tables):
                sheet_name = f"Tabela_{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                utils.format_excel(writer, sheet_name)
        
        st.download_button(
            label="📁 Fazer Download da Planilha Excel",
            data=excel_buffer.getvalue(),
            file_name="BOM_Dados_Gerais_Web.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("👈 Use o painel lateral para carregar um chicote/documento.")
