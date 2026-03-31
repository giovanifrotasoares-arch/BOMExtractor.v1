import streamlit as st
from streamlit_cropper import st_cropper
import fitz  # PyMuPDF
import pandas as pd
import io
import utils
# Configurations
st.set_page_config(page_title="Extrator Visual BOM", layout="wide", page_icon="🚜")
st.title("🚜 Extrator Automotivo BOM - Inteligência PDF")
st.markdown("""
<div style='background-color: #2F4F4F; padding: 15px; border-radius: 10px; color: white; font-size: 14px;'>
<strong>💡 Manual de Operação de Precisão Web:</strong><br><br>
Diferente dos programas de computador (onde o <code>Ctrl + Scroll</code> puxa e dá zoom no mapa), os <b>Navegadores de Internet bloqueiam o controle do mouse do seu Windows</b>. Se você der Ctrl+Scroll no navegador, as letras do site vão aumentar, mas a imagem do desenho ficará estática.<br><br>
<b>✔️ A SOLUÇÃO: Lupa de Alta Definição Automática</b><br>
Para extrair um PDF enorme com letras minúsculas sem errar o quadro, nós construímos uma <i>Zoom Camera</i>!<br>
1. Desenhe um quadrado "por cima" de onde você acha que está a Tabela na folha geral.<br>
2. Olhe para a Seção Ocular (Lupa) logo abaixo da imagem! Ela projetará um <b>Zoom Gigante e Nítido</b> do que o quadrado vermelho está tocando. Você ajusta o quadrado e a lupa ajusta instantaneamente pra você refinar a captura antes de clicar em Extrair!
</div>
""", unsafe_allow_html=True)
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
    
    # Render page to PIL Image in High 3.0x Resolution
    img, pt_w, pt_h, img_w, img_h = utils.get_page_image(st.session_state.pdf_bytes, page_num)
    
    st.markdown("---")
    st.subheader("1. Mapa Geral (Selecione a Região)")
    
    # Removeu-se as colunas para que a imagem do cropper fique com a largura máxima possível nativamente
    box = st_cropper(img, realtime_update=True, box_color='#FF0000', aspect_ratio=None, return_type='box')
    
    # --- NOVO RECURSO: LUPA CORTADA ---
    st.markdown("---")
    st.subheader("🔍 2. Lupa de Precisão (Evita Extração Vazia)")
    st.info("Veja aqui se o quadrado capturou as bordas superiores, inferiores e laterais do texto inteiro.\n Ajuste o quadro acima e o Zoom focará automaticamente!")
    
    left, top, width, height = box['left'], box['top'], box['width'], box['height']
    cropped_preview = img.crop((left, top, left + width, top + height))
    
    # Exibe a prévia da área selecionada usando 100% da largura da tela (o que gera o efeito de Lupa Gigante)
    st.image(cropped_preview, use_container_width=True)
    
    st.markdown("---")
    st.subheader("⚙️ 3. Processamento e Exportação")
    
    col_a, col_b = st.columns([1, 1])
    
    with col_a:
        if st.button("▶️ Extrair Tabela que aparece na Lupa", type="primary"):
            with st.spinner("Decodificando texto matemático do PDF..."):
                df = utils.extract_table_from_bbox(
                    st.session_state.pdf_bytes, page_num, box, 
                    pt_w, pt_h, img_w, img_h
                )
                if df is not None and not df.empty:
                    st.session_state.extracted_tables.append(df)
                    st.success(f"Matriz extraída perfeitamente! Capturadas {len(df)} linhas.")
                else:
                    st.error("Falha ao encontrar texto legível. A tabela está sendo enquadrada inteira na lupa? As margens do quadrado estão muito apertadas cortando letras?")
    with col_b:
        st.write(f"Tabelas Prontas na Fila: **{len(st.session_state.extracted_tables)}**")
        if st.button("🗑️ Limpar Fila (Reset)"):
            st.session_state.extracted_tables = []
            st.rerun()
    if len(st.session_state.extracted_tables) > 0:
        st.markdown("### Prévia da Última Tabela Estocada:")
        st.dataframe(st.session_state.extracted_tables[-1].head(10), use_container_width=True)
        
        # Gerenciamento de exportacao (Memory Buffer)
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            for i, df in enumerate(st.session_state.extracted_tables):
                sheet_name = f"Tabela_{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                utils.format_excel(writer, sheet_name)
        
        st.download_button(
            label="📁 Fazer Download da Planilha Excel (Limpa & Formatada)",
            data=excel_buffer.getvalue(),
            file_name="BOM_Dados_Gerais_Web.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
                
else:
    st.info("👈 Use o painel lateral para carregar um chicote/documento e começar a análise de engenharia.")
