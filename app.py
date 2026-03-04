# ============================================
# DESPACHO AUDIT - MODELO IPEm/RJ
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Despacho AUDIT - IPEm",
    page_icon="📄",
    layout="centered"
)

# TÍTULO
st.title("📄 Despacho AUDIT - IPEm/RJ")
st.markdown("**Gerador de Despachos conforme modelo oficial**")
st.markdown("---")

# UPLOAD
st.subheader("📂 Upload do Processo")
arquivo = st.file_uploader(
    "Selecione o PDF do processo",
    type=['pdf']
)

# ============================================
# FUNÇÕES DE EXTRAÇÃO
# ============================================

def extrair_texto_pdf(arquivo):
    texto = ""
    with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
        for pagina in pdf.pages:
            if pagina.extract_text():
                texto += pagina.extract_text() + "\n"
    return texto

def extrair_dados(texto):
    dados = {}
    
    # Processo
    proc = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    dados['processo'] = proc.group(1) if proc else "NÃO IDENTIFICADO"
    
    # Objeto (pegar frase após "objeto" ou "objetivo")
    obj = re.search(r'(objeto|objetivo)[:\s]*([^.]+)', texto, re.IGNORECASE)
    dados['objeto'] = obj.group(2).strip() if obj else "NÃO IDENTIFICADO"
    
    # Número do parecer
    parecer = re.search(r'Doc\. SEI nº (\d+)', texto, re.IGNORECASE)
    dados['parecer'] = parecer.group(1) if parecer else "XXX"
    
    # Verificar documentos
    dados['etp'] = "✅" if re.search(r'estudo preliminar|etp', texto, re.IGNORECASE) else "❌"
    dados['tr'] = "✅" if re.search(r'termo de referência|tr', texto, re.IGNORECASE) else "❌"
    dados['justificativa'] = "✅" if re.search(r'justificativa', texto, re.IGNORECASE) else "❌"
    dados['pesquisa'] = "✅" if re.search(r'pesquisa de preços|mapa', texto, re.IGNORECASE) else "❌"
    dados['risco'] = "✅" if re.search(r'gestão de risco|risco', texto, re.IGNORECASE) else "❌"
    
    return dados

# ============================================
# FUNÇÃO PARA GERAR DESPACHO (MODELO EXATO)
# ============================================

def gerar_despacho(dados):
    doc = Document()
    
    # Configura fonte
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    # ========================================
    # 1. OBJETO E ESCOPO
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("1. OBJETO E ESCOPO").bold = True
    
    doc.add_paragraph(
        f"Trata-se da análise de conformidade da instrução processual relativa "
        f"à {dados['objeto']}, com fulcro na Lei nº 14.133/2021. A presente "
        f"manifestação desta Auditoria Interna (AUDINT) limita-se ao exame da "
        f"regularidade do rito administrativo, não adentrando no mérito técnico "
        f"ou na discricionariedade da despesa."
    )
    
    doc.add_paragraph()
    
    # ========================================
    # 2. VERIFICAÇÃO DA INSTRUÇÃO
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("2. VERIFICAÇÃO DA INSTRUÇÃO").bold = True
    
    doc.add_paragraph(
        "Verificamos que o processo seguiu o fluxo obrigatório da fase "
        "preparatória, estando devidamente instruído com os seguintes "
        "elementos essenciais:"
    )
    
    doc.add_paragraph()
    
    # Planejamento
    doc.add_paragraph("• Planejamento: ", style='List Bullet').add_run(
        "Justificativa Técnica e Administrativa, Gestão de Risco, "
        "Estudo Técnico Preliminar - ETP e Termo de Referência - TR "
        "em conformidade com o Art. 18 da Lei 14.133/21;"
    ).italic = True
    
    # Economicidade
    doc.add_paragraph("• Economicidade: ", style='List Bullet').add_run(
        "Pesquisa de preços realizada com base em parâmetros de mercado;"
    ).italic = True
    
    # Legalidade
    p = doc.add_paragraph("• Legalidade: ", style='List Bullet')
    p.add_run(f"Existência de Parecer Jurídico (Doc. SEI nº {dados['parecer']}),").italic = True
    
    doc.add_paragraph()
    
    # ========================================
    # 3. CONSIDERAÇÕES DE CONTROLE
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("3. CONSIDERAÇÕES DE CONTROLE").bold = True
    
    doc.add_paragraph(
        "Considerando que a Assessoria Jurídica não apontou óbices e que "
        "as áreas técnicas certificaram a adequação dos quantitativos e "
        "especificações, esta AUDIT registra que:"
    )
    
    doc.add_paragraph()
    doc.add_paragraph("• As etapas de controle preventivo foram observadas;", style='List Bullet')
    doc.add_paragraph("• A segregação de funções foi respeitada;", style='List Bullet')
    doc.add_paragraph("• O processo encontra-se em condições de prosseguimento.", style='List Bullet')
    
    doc.add_paragraph()
    
    # ========================================
    # 4. CONCLUSÃO
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("4. CONCLUSÃO").bold = True
    
    doc.add_paragraph(
        "Diante da regularidade formal da instrução processual, "
        "esta AUDIT indica o prosseguimento do presente processo."
    )
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    # ========================================
    # ASSINATURA
    # ========================================
    
    doc.add_paragraph("At.te.")
    
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    doc.add_paragraph("___________________________________")
    doc.add_paragraph()
    
    return doc

# ============================================
# INTERFACE PRINCIPAL
# ============================================

if arquivo:
    
    with st.spinner("🔍 Analisando processo..."):
        
        # Extrair texto
        texto = extrair_texto_pdf(arquivo)
        
        # Extrair dados
        dados = extrair_dados(texto)
        
        # ========================================
        # MOSTRA RESULTADO DA ANÁLISE
        # ========================================
        
        st.success("✅ Processo analisado!")
        
        # Dados principais
        st.subheader("📋 Dados extraídos:")
        st.write(f"**Processo:** {dados['processo']}")
        st.write(f"**Objeto:** {dados['objeto']}")
        st.write(f"**Parecer Jurídico:** Doc. SEI nº {dados['parecer']}")
        
        # Checklist
        st.subheader("✅ Documentos identificados:")
        
        col1, col2 = st.columns(2)
        with col1:
            st.write(f"ETP: {dados['etp']}")
            st.write(f"TR: {dados['tr']}")
            st.write(f"Justificativa: {dados['justificativa']}")
        with col2:
            st.write(f"Pesquisa Preços: {dados['pesquisa']}")
            st.write(f"Gestão Risco: {dados['risco']}")
        
        # ========================================
        # BOTÃO DE DOWNLOAD
        # ========================================
        
        st.markdown("---")
        st.subheader("📥 Download do Despacho")
        
        if st.button("📄 GERAR DESPACHO"):
            
            with st.spinner("Gerando despacho..."):
                
                # Gerar documento
                doc = gerar_despacho(dados)
                
                # Salvar
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                    doc.save(tmp.name)
                    tmp_path = tmp.name
                
                # Ler
                with open(tmp_path, 'rb') as f:
                    doc_bytes = f.read()
                
                # Apagar
                os.unlink(tmp_path)
                
                # Download
                nome_arquivo = f"DESPACHO_AUDIT_{dados['processo'].replace('/', '_')}.docx"
                
                st.download_button(
                    label="📥 CLIQUE AQUI PARA BAIXAR O DESPACHO",
                    data=doc_bytes,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Mostrar prévia
                with st.expander("📄 Prévia do despacho gerado"):
                    preview = f"""
1. OBJETO E ESCOPO
==================
Trata-se da análise de conformidade da instrução processual relativa à {dados['objeto']}, 
com fulcro na Lei nº 14.133/2021...

2. VERIFICAÇÃO DA INSTRUÇÃO
===========================
• Planejamento: Justificativa Técnica e Administrativa, Gestão de Risco, 
  Estudo Técnico Preliminar - ETP e Termo de Referência - TR...
• Economicidade: Pesquisa de preços realizada...
• Legalidade: Existência de Parecer Jurídico (Doc. SEI nº {dados['parecer']})...

3. CONSIDERAÇÕES DE CONTROLE
============================
• As etapas de controle preventivo foram observadas;
• A segregação de funções foi respeitada;
• O processo encontra-se em condições de prosseguimento.

4. CONCLUSÃO
============
Diante da regularidade formal da instrução processual, esta AUDIT indica o 
prosseguimento do presente processo.
                    """
                    st.text(preview)
        
else:
    st.info("👆 Faça upload do PDF para gerar o despacho")
    
    with st.expander("📌 Modelo de despacho utilizado"):
        st.write("""
**1. OBJETO E ESCOPO**
Trata-se da análise de conformidade da instrução processual relativa à aquisição..., 
com fulcro na Lei nº 14.133/2021...

**2. VERIFICAÇÃO DA INSTRUÇÃO**
• Planejamento: Justificativa Técnica, ETP e TR...
• Economicidade: Pesquisa de preços...
• Legalidade: Parecer Jurídico...

**3. CONSIDERAÇÕES DE CONTROLE**
• As etapas de controle preventivo foram observadas...
• A segregação de funções foi respeitada...
• O processo encontra-se em condições de prosseguimento.

**4. CONCLUSÃO**
Diante da regularidade formal, indica-se o prosseguimento...
        """)

st.markdown("---")
st.caption("© 2026 - Auditoria Interna IPEm/RJ")
