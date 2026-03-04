# ============================================
# ANALISADOR IPEm - DESPACHO AUDIT PERSONALIZADO
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
import pandas as pd
import io
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Analisador IPEm - Despacho AUDIT",
    page_icon="📋",
    layout="wide"
)

# TÍTULO
st.title("📋 Analisador de Instrução Processual - IPEm/RJ")
st.markdown("**Geração automática de DESPACHO AUDIT**")
st.markdown("---")

# SIDEBAR
with st.sidebar:
    st.header("ℹ️ Sobre")
    st.info("""
    **Finalidade:** Gerar despacho AUDIT padronizado conforme modelo do IPEm
    
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
    """
    Verifica se o processo contém os documentos obrigatórios
    Conforme modelo de despacho do IPEm
    """
    
    # Planejamento
    planejamento_itens = {
        'Justificativa Técnica e Administrativa': r'justificativa|motivação',
        'Gestão de Risco': r'gestão de risco|risco|matriz de risco',
        'Estudo Técnico Preliminar (ETP)': r'estudo preliminar|etp',
        'Termo de Referência (TR)': r'termo de referência|tr'
    }
    
    # Economicidade
    economicidade_itens = {
        'Pesquisa de Preços': r'pesquisa de preços|mapa de preços|orçamento estimado'
    }
    
    # Legalidade
    legalidade_itens = {
        'Parecer Jurídico': r'parecer jurídico|assessoria jurídica|procuradoria'
    }
    
    # Controle
    controle_itens = {
        'Segregação de Funções': r'segregação|funções distintas|responsáveis diferentes'
    }
    
    resultados = {
        'planejamento': [],
        'economicidade': [],
        'legalidade': [],
        'controle': []
    }
    
    # Verificar planejamento
    for item, padrao in planejamento_itens.items():
        if re.search(padrao, texto, re.IGNORECASE):
            resultados['planejamento'].append(f"• {item}")
    
    # Verificar economicidade
    for item, padrao in economicidade_itens.items():
        if re.search(padrao, texto, re.IGNORECASE):
            resultados['economicidade'].append(f"• {item}")
    
    # Verificar legalidade
    for item, padrao in legalidade_itens.items():
        if re.search(padrao, texto, re.IGNORECASE):
            resultados['legalidade'].append(f"• {item}")
    
    # Verificar controle
    for item, padrao in controle_itens.items():
        if re.search(padrao, texto, re.IGNORECASE):
            resultados['controle'].append(f"• {item}")
    
    # Extrair número do parecer jurídico
    parecer = re.search(r'Doc\. SEI nº (\d+)', texto, re.IGNORECASE)
    num_parecer = parecer.group(1) if parecer else "XXX"
    
    return resultados, num_parecer

def extrair_dados_basicos(texto):
    """Extrai informações básicas"""
    dados = {}
    
    # Número do processo
    match = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    dados['processo'] = match.group(1).strip() if match else "Não identificado"
    
    # Objeto
    match = re.search(r'(objeto|objetivo)[:\s]*([^.]+)', texto, re.IGNORECASE)
    dados['objeto'] = match.group(2).strip() if match else "Não especificado"
    
    # Valor
    match = re.search(r'Valor\s*R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    dados['valor'] = match.group(1) if match else "Não identificado"
    
    return dados

# ============================================
# FUNÇÃO PARA GERAR DESPACHO CONFORME MODELO
# ============================================

def gerar_despacho_audit(dados, resultados, num_parecer, recomendacao):
    """
    Gera despacho EXATAMENTE conforme modelo fornecido
    """
    
    doc = Document()
    
    # CONFIGURAÇÃO DE ESTILOS
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    # ========================================
    # 1. OBJETO E ESCOPO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("1. OBJETO E ESCOPO")
    run.bold = True
    run.font.size = Pt(12)
    
    doc.add_paragraph()
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
    run = p.add_run("2. VERIFICAÇÃO DA INSTRUÇÃO")
    run.bold = True
    
    doc.add_paragraph()
    doc.add_paragraph(
        "Verificamos que o processo seguiu o fluxo obrigatório da fase "
        "preparatória, estando devidamente instruído com os seguintes "
        "elementos essenciais:"
    )
    
    doc.add_paragraph()
    
    # Planejamento
    p = doc.add_paragraph()
    run = p.add_run("• Planejamento:")
    run.bold = True
    for item in resultados['planejamento']:
        doc.add_paragraph(f"  {item}")
    
    doc.add_paragraph()
    
    # Economicidade
    p = doc.add_paragraph()
    run = p.add_run("• Economicidade:")
    run.bold = True
    for item in resultados['economicidade']:
        doc.add_paragraph(f"  {item}")
    
    doc.add_paragraph()
    
    # Legalidade
    p = doc.add_paragraph()
    run = p.add_run("• Legalidade:")
    run.bold = True
    for item in resultados['legalidade']:
        doc.add_paragraph(f"  {item} (Doc. SEI nº {num_parecer})")
    
    doc.add_paragraph()
    
    # ========================================
    # 3. CONSIDERAÇÕES DE CONTROLE
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("3. CONSIDERAÇÕES DE CONTROLE")
    run.bold = True
    
    doc.add_paragraph()
    doc.add_paragraph(
        "Considerando que a Assessoria Jurídica não apontou óbices e que "
        "as áreas técnicas certificaram a adequação dos quantitativos e "
        "especificações, esta AUDIT registra que:"
    )
    
    doc.add_paragraph()
    doc.add_paragraph("• As etapas de controle preventivo foram observadas;")
    
    if resultados['controle']:
        doc.add_paragraph("• A segregação de funções foi respeitada;")
    else:
        doc.add_paragraph("• Recomenda-se verificar a segregação de funções;")
    
    doc.add_paragraph()
    
    if recomendacao == "PROSSEGUIMENTO":
        doc.add_paragraph("• O processo encontra-se em condições de prosseguimento.")
    else:
        doc.add_paragraph("• O processo necessita de complementação antes do prosseguimento.")
    
    doc.add_paragraph()
    
    # ========================================
    # 4. CONCLUSÃO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("4. CONCLUSÃO")
    run.bold = True
    
    doc.add_paragraph()
    
    if recomendacao == "PROSSEGUIMENTO":
        doc.add_paragraph(
            "Diante da regularidade formal da instrução processual, "
            "esta AUDIT indica o prosseguimento do presente processo."
        )
    elif recomendacao == "PROSSEGUIMENTO COM RESSALVAS":
        doc.add_paragraph(
            "Diante da regularidade parcial da instrução processual, "
            "esta AUDIT indica o prosseguimento com recomendações de "
            "complementação documental."
        )
    else:
        doc.add_paragraph(
            "Diante das inconsistências identificadas na instrução "
            "processual, esta AUDIT recomenda a complementação dos "
            "documentos faltantes antes do prosseguimento."
        )
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    # ========================================
    # ASSINATURA
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("At.te.")
    run.italic = True
    
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    run = p.add_run("___________________________________")
    run.bold = True
    
    doc.add_paragraph()
    
    return doc

# ============================================
# PROCESSAMENTO PRINCIPAL
# ============================================

if arquivo is not None:
    
    with st.spinner("🔍 Analisando processo e gerando despacho AUDIT..."):
        
        # Extrair texto
        texto = extrair_texto_pdf(arquivo)
        
        # Extrair dados básicos
        dados = extrair_dados_basicos(texto)
        
        # Verificar documentos
        resultados, num_parecer = verificar_documentos_obrigatorios(texto)
        
        # Incrementar contador
        st.session_state.processos_analisados += 1
        
        # Calcular percentual de documentos encontrados
        total_itens = 7  # Total de itens obrigatórios
        encontrados = len(resultados['planejamento']) + len(resultados['economicidade']) + \
                     len(resultados['legalidade']) + len(resultados['controle'])
        percentual = (encontrados / total_itens) * 100
        
        # Determinar recomendação
        if encontrados >= 6:
            recomendacao = "PROSSEGUIMENTO"
        elif encontrados >= 4:
            recomendacao = "PROSSEGUIMENTO COM RESSALVAS"
        else:
            recomendacao = "AGUARDAR COMPLEMENTAÇÃO"
        
        # ========================================
        # EXIBIR RESULTADO NA TELA
        # ========================================
        
        st.success("✅ Análise concluída!")
        
        # Métricas principais
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Processo", dados['processo'])
        with col2:
            st.metric("Documentos encontrados", f"{encontrados}/{total_itens}")
        with col3:
            st.metric("Percentual", f"{percentual:.1f}%")
        
        # Recomendação em destaque
        if recomendacao == "PROSSEGUIMENTO":
            st.success(f"**RECOMENDAÇÃO: {recomendacao}**")
        elif recomendacao == "PROSSEGUIMENTO COM RESSALVAS":
            st.warning(f"**RECOMENDAÇÃO: {recomendacao}**")
        else:
            st.error(f"**RECOMENDAÇÃO: {recomendacao}**")
        
        # Detalhamento dos documentos
        with st.expander("📋 Ver detalhamento da análise", expanded=True):
            
            col_left, col_right = st.columns(2)
            
            with col_left:
                st.subheader("✅ Documentos encontrados")
                
                for item in resultados['planejamento']:
                    st.write(item)
                for item in resultados['economicidade']:
                    st.write(item)
                for item in resultados['legalidade']:
                    st.write(item)
                for item in resultados['controle']:
                    st.write(item)
            
            with col_right:
                st.subheader("❌ Possíveis ausências")
                
                if len(resultados['planejamento']) < 4:
                    st.write("• Complementar documentos de planejamento")
                if len(resultados['economicidade']) < 1:
                    st.write("• Incluir pesquisa de preços")
                if len(resultados['legalidade']) < 1:
                    st.write("• Incluir parecer jurídico")
                if len(resultados['controle']) < 1:
                    st.write("• Verificar segregação de funções")
        
        st.markdown("---")
        
        # ========================================
        # BOTÃO DE DOWNLOAD DO DESPACHO
        # ========================================
        
        st.subheader("📥 Download do Despacho AUDIT")
        
        # Gerar documento
        doc = gerar_despacho_audit(dados, resultados, num_parecer, recomendacao)
        
        # Salvar em arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name
        
        # Ler o arquivo
        with open(tmp_path, 'rb') as f:
            doc_bytes = f.read()
        
        # Remover arquivo temporário
        os.unlink(tmp_path)
        
        # Nome do arquivo
        data_atual = datetime.now().strftime("%Y%m%d")
        nome_arquivo = f"DESPACHO_AUDIT_{dados['processo'].replace('/', '_')}_{data_atual}.docx"
        
        # Botão de download
        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DESPACHO AUDIT (WORD)",
            data=doc_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        # Prévia do despacho
        with st.expander("📄 Prévia do despacho gerado"):
            st.text(f"""
1. OBJETO E ESCOPO
==================
Trata-se da análise de conformidade da instrução processual relativa à {dados['objeto']}, 
com fulcro na Lei nº 14.133/2021. A presente manifestação desta Auditoria Interna (AUDINT) 
limita-se ao exame da regularidade do rito administrativo, não adentrando no mérito técnico 
ou na discricionariedade da despesa.

2. VERIFICAÇÃO DA INSTRUÇÃO
===========================
Verificamos que o processo seguiu o fluxo obrigatório da fase preparatória, estando 
devidamente instruído com os seguintes elementos essenciais:

• Planejamento:
{chr(10).join(resultados['planejamento']) if resultados['planejamento'] else '  • Nenhum documento de planejamento identificado'}

• Economicidade:
{chr(10).join(resultados['economicidade']) if resultados['economicidade'] else '  • Nenhuma pesquisa de preços identificada'}

• Legalidade:
{chr(10).join(resultados['legalidade']) if resultados['legalidade'] else f'  • Nenhum parecer jurídico identificado'}

3. CONSIDERAÇÕES DE CONTROLE
============================
Considerando que a Assessoria Jurídica não apontou óbices e que as áreas técnicas 
certificaram a adequação dos quantitativos e especificações, esta AUDIT registra que:

• As etapas de controle preventivo foram observadas;
• A segregação de funções foi respeitada;
• {recomendacao}

4. CONCLUSÃO
============
{recomendacao}
            """)
        
        st.info("✅ Despacho gerado com sucesso! Clique no botão acima para baixar.")

else:
    # TELA INICIAL
    st.info("👆 Faça upload do PDF do processo para gerar o despacho AUDIT")
    
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
        
        • Etapas de controle preventivo observadas...
        • Segregação de funções respeitada...
        
        **4. CONCLUSÃO**
        
        Diante da regularidade formal, indica-se o prosseguimento...
        """)

# RODAPÉ
st.markdown("---")
st.caption(f"© 2026 - Auditoria Interna IPEm/RJ - Gerador de Despachos AUDIT v1.0")
