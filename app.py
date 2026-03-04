# ============================================
# ANALISADOR COM DESPACHO EM WORD (CORRIGIDO)
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
import pandas as pd
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Analisador IPEm - Despacho Automático",
    page_icon="📋",
    layout="wide"
)

# TÍTULO
st.title("📋 Analisador de Instrução Processual - IPEm/RJ")
st.markdown("**Geração automática de DESPACHO em WORD**")
st.markdown("---")

# SIDEBAR
with st.sidebar:
    st.header("ℹ️ Sobre")
    st.info("""
    **Finalidade:** Gerar despacho padronizado para juntada ao processo SEI
    
    **Formato de saída:** Microsoft Word (.docx)
    """)
    
    st.header("📊 Estatísticas")
    if 'processos_analisados' not in st.session_state:
        st.session_state.processos_analisados = 0
    st.metric("Processos analisados", st.session_state.processos_analisados)

# UPLOAD
st.header("📂 Processo SEI")
arquivo = st.file_uploader(
    "Selecione o PDF do processo",
    type=['pdf'],
    help="Upload do documento principal do processo"
)

# ============================================
# FUNÇÕES DE ANÁLISE
# ============================================

def extrair_texto_pdf(arquivo):
    """Extrai texto do PDF"""
    texto = ""
    with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
        for pagina in pdf.pages:
            if pagina.extract_text():
                texto += pagina.extract_text() + "\n"
    return texto

def verificar_documentos_obrigatorios(texto):
    """Verifica documentos obrigatórios"""
    
    documentos = {
        'Estudo Técnico Preliminar (Art. 18, I)': r'estudo preliminar|etp',
        'Termo de Referência (Art. 18, II)': r'termo de referência|projeto básico',
        'Pesquisa de Preços (Art. 23)': r'pesquisa de preços|mapa de preços',
        'Parecer Jurídico (Art. 53)': r'parecer jurídico|assessoria jurídica',
        'Justificativa (Art. 18, III)': r'justificativa|motivação',
        'Publicação (Art. 54)': r'publicação|diário oficial',
        'Designação de Fiscal (Art. 117)': r'fiscal|gestor do contrato'
    }
    
    presentes = []
    ausentes = []
    
    for doc, padrao in documentos.items():
        if re.search(padrao, texto, re.IGNORECASE):
            presentes.append(doc)
        else:
            ausentes.append(doc)
    
    return presentes, ausentes

def extrair_dados_basicos(texto):
    """Extrai informações básicas"""
    dados = {}
    
    # Número
    match = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    dados['processo'] = match.group(1).strip() if match else "Não identificado"
    
    # Valor
    match = re.search(r'Valor\s*R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    dados['valor'] = match.group(1) if match else "Não identificado"
    
    # Data
    match = re.search(r'(\d{1,2})\s*de\s*([A-ZÇa-zç]+)\s*de\s*(\d{4})', texto)
    if match:
        dados['data'] = f"{match.group(1)}/{match.group(2)}/{match.group(3)}"
    else:
        dados['data'] = "Não identificado"
    
    # Órgão
    match = re.search(r'([A-Z]{3,}(?:\/[A-Z]{3,})*)', texto)
    dados['orgao'] = match.group(1) if match else "Não identificado"
    
    return dados

# ============================================
# FUNÇÃO PARA GERAR DESPACHO EM WORD
# ============================================

def gerar_despacho_word(dados, presentes, ausentes, recomendacao):
    """
    Gera um despacho formatado em Word (.docx)
    """
    
    doc = Document()
    
    # CONFIGURAÇÃO DE ESTILOS
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    # ========================================
    # CABEÇALHO
    # ========================================
    
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header.add_run("INSTITUTO DE PESOS E MEDIDAS DO ESTADO DO RIO DE JANEIRO")
    run.bold = True
    run.font.size = Pt(14)
    
    header2 = doc.add_paragraph()
    header2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header2.add_run("AUDITORIA INTERNA")
    run.bold = True
    run.font.size = Pt(13)
    
    doc.add_paragraph()
    
    # ========================================
    # NÚMERO DO DESPACHO
    # ========================================
    
    num_despacho = f"DESPACHO AUDIT Nº {datetime.now().strftime('%Y')}/"
    p = doc.add_paragraph()
    run = p.add_run(num_despacho)
    run.bold = True
    
    doc.add_paragraph()
    
    # ========================================
    # DESTINATÁRIO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("Ao Sr. Presidente do IPEm/RJ")
    run.bold = True
    
    doc.add_paragraph()
    
    # ========================================
    # ASSUNTO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run(f"Assunto: Análise de instrução processual - Processo {dados['processo']}")
    run.bold = True
    
    doc.add_paragraph()
    
    # ========================================
    # CORPO DO DESPACHO
    # ========================================
    
    doc.add_paragraph(f"Reporto-me ao Processo {dados['processo']}, em trâmite por esta Auditoria Interna, para manifestação quanto à instrução processual para prosseguimento.")
    doc.add_paragraph()
    
    doc.add_paragraph("Em análise aos autos, verificou-se o seguinte:")
    doc.add_paragraph()
    
    # Documentos presentes
    if presentes:
        p = doc.add_paragraph()
        run = p.add_run("Documentos identificados:")
        run.bold = True
        
        for doc_item in presentes:
            doc.add_paragraph(f"✓ {doc_item}", style='List Bullet')
        
        doc.add_paragraph()
    
    # Documentos ausentes
    if ausentes:
        p = doc.add_paragraph()
        run = p.add_run("Documentos não localizados:")
        run.bold = True
        
        for doc_item in ausentes:
            doc.add_paragraph(f"✗ {doc_item}", style='List Bullet')
        
        doc.add_paragraph()
    
    # ========================================
    # CONCLUSÃO
    # ========================================
    
    doc.add_paragraph("Diante do exposto, esta Auditoria Interna manifesta-se:")
    doc.add_paragraph()
    
    if recomendacao == "PROSSEGUIMENTO":
        p = doc.add_paragraph()
        run = p.add_run("PELO PROSSEGUIMENTO do feito, eis que presentes os documentos essenciais à instrução processual.")
        run.bold = True
        
    elif recomendacao == "PROSSEGUIMENTO COM RESSALVAS":
        p = doc.add_paragraph()
        run = p.add_run("PELO PROSSEGUIMENTO COM RESSALVAS, recomendando-se a juntada dos documentos ausentes.")
        run.bold = True
        
    else:
        p = doc.add_paragraph()
        run = p.add_run("PELA DEVOLUÇÃO para complementação da instrução, conforme documentos faltantes listados.")
        run.bold = True
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    # ========================================
    # ASSINATURA
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run(f"Rio de Janeiro, {datetime.now().strftime('%d de %B de %Y')}")
    run.italic = True
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    run = p.add_run("___________________________________")
    run.bold = True
    
    p = doc.add_paragraph()
    run = p.add_run("Auditor Interno")
    run.italic = True
    
    p = doc.add_paragraph()
    run = p.add_run("IPEm/RJ")
    run.italic = True
    
    return doc

# ============================================
# PROCESSAMENTO PRINCIPAL
# ============================================

if arquivo is not None:
    
    with st.spinner("🔍 Analisando e gerando despacho..."):
        
        # Extrair texto
        texto = extrair_texto_pdf(arquivo)
        
        # Extrair dados
        dados = extrair_dados_basicos(texto)
        
        # Verificar documentos
        presentes, ausentes = verificar_documentos_obrigatorios(texto)
        
        # Incrementar contador
        st.session_state.processos_analisados += 1
        
        # Determinar recomendação
        if not ausentes:
            recomendacao = "PROSSEGUIMENTO"
        elif len(ausentes) <= 2 and not any(x in str(ausentes) for x in ['Parecer', 'Termo', 'Estudo']):
            recomendacao = "PROSSEGUIMENTO COM RESSALVAS"
        else:
            recomendacao = "AGUARDAR"
        
        # ========================================
        # EXIBIR RESULTADO NA TELA
        # ========================================
        
        st.success("✅ Análise concluída!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Processo", dados['processo'])
        with col2:
            st.metric("Documentos presentes", len(presentes))
        with col3:
            st.metric("Documentos ausentes", len(ausentes))
        
        if recomendacao == "PROSSEGUIMENTO":
            st.success(f"**RECOMENDAÇÃO: {recomendacao}**")
        elif recomendacao == "PROSSEGUIMENTO COM RESSALVAS":
            st.warning(f"**RECOMENDAÇÃO: {recomendacao}**")
        else:
            st.error(f"**RECOMENDAÇÃO: {recomendacao}**")
        
        col_left, col_right = st.columns(2)
        
        with col_left:
            st.subheader("✅ Documentos identificados")
            for doc in presentes:
                st.write(f"✓ {doc}")
        
        with col_right:
            if ausentes:
                st.subheader("❌ Documentos não localizados")
                for doc in ausentes:
                    st.write(f"✗ {doc}")
        
        st.markdown("---")
        
        # ========================================
        # BOTÃO DE DOWNLOAD SIMPLIFICADO
        # ========================================
        
        st.subheader("📥 Download do Despacho")
        
        # Gerar documento
        doc = gerar_despacho_word(dados, presentes, ausentes, recomendacao)
        
        # Salvar em arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name
        
        # Ler o arquivo
        with open(tmp_path, 'rb') as f:
            doc_bytes = f.read()
        
        # Remover arquivo temporário
        os.unlink(tmp_path)
        
        # Botão de download
        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DESPACHO (WORD)",
            data=doc_bytes,
            file_name=f"DESPACHO_{dados['processo'].replace('/', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        st.info("✅ Despacho gerado! Clique no botão acima para baixar.")

else:
    st.info("👆 Faça upload do PDF para gerar o despacho")
    
    with st.expander("📌 Como funciona?"):
        st.write("""
        1. Faça upload do PDF do processo
        2. O sistema analisa automaticamente
        3. Gera um despacho em Word
        4. Você baixa e anexa no SEI
        """)

# RODAPÉ
st.markdown("---")
st.caption(f"© 2026 - Auditoria Interna IPEm/RJ - Gerador de Despachos v2.0")
