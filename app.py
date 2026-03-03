# ============================================
# APP WEB PARA ANALISAR PROCESSOS SEI
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
import pandas as pd
import io
import base64

# Configuração da página (TÍTULO BONITO)
st.set_page_config(
    page_title="Analisador IPEm",
    page_icon="📋",
    layout="wide"
)

# TÍTULO PRINCIPAL
st.title("📋 Analisador de Processos SEI - IPEm/RJ")
st.markdown("---")

# SIDEBAR COM INFORMAÇÕES
with st.sidebar:
    st.header("ℹ️ Sobre")
    st.info("""
    **Analisador automático de processos SEI**
    
    - Extrai dados automaticamente
    - Verifica conformidade com Lei 14.133/2021
    - Gera relatório de auditoria
    
    **Como usar:**
    1. Faça upload do PDF
    2. Aguarde análise
    3. Baixe o relatório
    """)
    
    st.header("📊 Estatísticas")
    if 'processos_analisados' not in st.session_state:
        st.session_state.processos_analisados = 0
    st.metric("Processos analisados nesta sessão", st.session_state.processos_analisados)

# ÁREA PRINCIPAL
st.header("📂 Upload do Processo")

# Upload do arquivo
arquivo = st.file_uploader(
    "Selecione o arquivo PDF do processo SEI",
    type=['pdf'],
    help="Clique para escolher o arquivo no seu computador"
)

# FUNÇÃO PARA EXTRAIR DADOS
def extrair_dados_pdf(arquivo_pdf):
    """Extrai informações do PDF"""
    
    dados = {
        'processo': 'Não identificado',
        'data': 'Não identificado',
        'valor': 'Não identificado',
        'objeto': 'Não identificado',
        'orgao': 'Não identificado',
        'assinaturas': [],
        'conformidades': [],
        'inconformidades': []
    }
    
    with pdfplumber.open(io.BytesIO(arquivo_pdf.read())) as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto_completo += pagina.extract_text() + "\n"
    
    # Extrair número do processo
    match = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto_completo, re.IGNORECASE)
    if match:
        dados['processo'] = match.group(1).strip()
    
    # Extrair data
    match = re.search(r'(\d{1,2})\s*de\s*([A-ZÇa-zç]+)\s*de\s*(\d{4})', texto_completo)
    if match:
        dados['data'] = f"{match.group(1)}/{match.group(2)}/{match.group(3)}"
    
    # Extrair valor
    match = re.search(r'Valor\s*R\$\s*([\d.,]+)', texto_completo, re.IGNORECASE)
    if match:
        dados['valor'] = match.group(1)
    
    # Extrair objeto
    match = re.search(r'(objeto|objetivo)[:\s]*([^.]+)', texto_completo, re.IGNORECASE)
    if match:
        dados['objeto'] = match.group(2).strip()[:200]
    
    # Extrair assinaturas
    assinaturas = re.findall(r'assinado eletronicamente por (.*?),\s*em (\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
    for nome, data_assinatura in assinaturas:
        dados['assinaturas'].append(f"• {nome.strip()} - {data_assinatura}")
    
    # ANÁLISE DE CONFORMIDADE
    itens = {
        'Estudo Preliminar': r'estudo preliminar|etp',
        'Termo de Referência': r'termo de referência|projeto básico',
        'Pesquisa de Preços': r'pesquisa de preços|mapa de preços',
        'Parecer Jurídico': r'parecer jurídico|assessoria jurídica',
        'Justificativa': r'justificativa|motivação',
        'Publicação': r'publicação|diário oficial',
        'Lei 14.133': r'14\.133|lei 14.133'
    }
    
    for item, padrao in itens.items():
        if re.search(padrao, texto_completo, re.IGNORECASE):
            dados['conformidades'].append(f"✅ {item}")
        else:
            dados['inconformidades'].append(f"❌ {item}")
    
    return dados, texto_completo

# SE ARQUIVO FOI ENVIADO
if arquivo is not None:
    
    with st.spinner("🔍 Processando PDF... Aguarde..."):
        
        # Extrair dados
        dados, texto = extrair_dados_pdf(arquivo)
        
        # Incrementar contador
        st.session_state.processos_analisados += 1
        
        # MOSTRAR RESULTADOS
        st.success("✅ Análise concluída!")
        
        # Criar abas para organizar
        tab1, tab2, tab3 = st.tabs(["📋 Resumo", "✅ Checklist", "📊 Relatório"])
        
        with tab1:
            st.subheader("📋 Dados do Processo")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("Nº Processo", dados['processo'])
                st.metric("Data", dados['data'])
                st.metric("Valor", f"R$ {dados['valor']}")
            
            with col2:
                st.metric("Assinaturas", len(dados['assinaturas']))
                total_itens = len(dados['conformidades']) + len(dados['inconformidades'])
                perc = (len(dados['conformidades']) / total_itens * 100) if total_itens > 0 else 0
                st.metric("Conformidade", f"{perc:.1f}%")
            
            st.subheader("📝 Objeto")
            st.info(dados['objeto'])
            
            if dados['assinaturas']:
                st.subheader("✍️ Assinaturas")
                for assinatura in dados['assinaturas']:
                    st.write(assinatura)
        
        with tab2:
            st.subheader("✅ Checklist de Conformidade")
            
            for item in dados['conformidades']:
                st.success(item)
            
            for item in dados['inconformidades']:
                st.warning(item)
        
        with tab3:
            st.subheader("📊 Relatório Completo")
            
            # Gerar relatório
            relatorio = f"""
RELATÓRIO DE AUDITORIA
======================

PROCESSO: {dados['processo']}
DATA: {datetime.now().strftime('%d/%m/%Y %H:%M')}
VALOR: R$ {dados['valor']}

OBJETO:
-------
{dados['objeto']}

ANÁLISE DE CONFORMIDADE:
------------------------
"""
            for item in dados['conformidades'] + dados['inconformidades']:
                relatorio += f"{item}\n"
            
            relatorio += f"""
ASSINATURAS:
------------
"""
            for assinatura in dados['assinaturas']:
                relatorio += f"{assinatura}\n"
            
            # Botão de download
            st.download_button(
                label="📥 Baixar Relatório Completo",
                data=relatorio,
                file_name=f"relatorio_{dados['processo'].replace('/', '_')}.txt",
                mime="text/plain"
            )
            
            # Prévia do relatório
            st.text_area("Prévia do relatório", relatorio, height=300)

else:
    # Mostrar instruções iniciais
    st.info("👆 Clique no botão acima para selecionar um arquivo PDF")
    
    # Exemplo de como funciona
    with st.expander("📌 Como funciona?"):
        st.write("""
        1. **Upload**: Você seleciona o PDF do processo SEI
        2. **Análise**: O sistema extrai automaticamente:
           - Número do processo
           - Data e valor
           - Objeto da contratação
           - Assinaturas eletrônicas
        3. **Conformidade**: Verifica itens da Lei 14.133/2021
        4. **Relatório**: Gera relatório de auditoria completo
        """)

# Rodapé
st.markdown("---")
st.caption("© 2024 - Auditoria Interna IPEm/RJ - Versão 1.0")
