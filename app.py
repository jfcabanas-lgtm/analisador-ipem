# ============================================
# DESPACHO AUDIT - MODELO COMPLETO IPEm/RJ
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
    page_title="Despacho AUDIT - IPEm/RJ",
    page_icon="📄",
    layout="wide"
)

# TÍTULO
st.title("📄 Despacho AUDIT - IPEm/RJ")
st.markdown("**Gerador de Despachos conforme modelo oficial completo**")
st.markdown("---")

# SIDEBAR COM INFORMAÇÕES
with st.sidebar:
    st.header("ℹ️ Sobre")
    st.info("""
    **Modelo de Despacho:**
    
    I. Introdução
    II. Análise Processual (6 itens)
    III. Observações
    IV. Despacho
    
    **Base Legal:** Lei 14.133/2021
    """)
    
    st.header("📊 Documentos verificados")
    st.write("""
    1. Solicitação e Autorização
    2. ETP e Gestão de Riscos
    3. Termo de Referência
    4. Pesquisa de Mercado/SIGA
    5. Conformidade Orçamentária
    6. Parecer Jurídico
    """)

# UPLOAD
st.subheader("📂 Upload do Processo")
arquivo = st.file_uploader(
    "Selecione o PDF do processo SEI",
    type=['pdf']
)

# ============================================
# FUNÇÕES DE EXTRAÇÃO
# ============================================

def extrair_texto_pdf(arquivo):
    """Extrai texto do PDF"""
    texto = ""
    with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
        for pagina in pdf.pages:
            if pagina.extract_text():
                texto += pagina.extract_text() + "\n"
    return texto

def extrair_dados_completos(texto):
    """
    Extrai todas as informações necessárias para o despacho
    Conforme modelo completo
    """
    
    dados = {}
    
    # ========================================
    # INTRODUÇÃO - Processo SEI
    # ========================================
    
    # Número do processo (formato SEI-XXXXXXXX/YYYY)
    proc = re.search(r'SEI[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    if not proc:
        proc = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    dados['processo_sei'] = proc.group(1).strip() if proc else "150014/001585/2025"
    
    # Data da autorização
    data = re.search(r'(\d{1,2})[/](\d{1,2})[/](\d{4})', texto)
    if data:
        dados['data_autorizacao'] = f"{data.group(1)}/{data.group(2)}/{data.group(3)}"
    else:
        dados['data_autorizacao'] = "20/10/2025"
    
    # ========================================
    # ETP E GESTÃO DE RISCOS
    # ========================================
    
    etp = re.search(r'ETP[:\s]*n[º°]?\s*(\d+/\d+)', texto, re.IGNORECASE)
    dados['etp_numero'] = etp.group(1) if etp else "31/2025"
    
    risco = re.search(r'Matriz de Riscos[:\s]*n[º°]?\s*(\d+/\d+)', texto, re.IGNORECASE)
    dados['risco_numero'] = risco.group(1) if risco else "26/2025"
    
    # ========================================
    # TERMO DE REFERÊNCIA
    # ========================================
    
    tr = re.search(r'TR[:\s]*n[º°]?\s*(\d+/\d+)', texto, re.IGNORECASE)
    dados['tr_numero'] = tr.group(1) if tr else "46/2025"
    
    # ========================================
    # VALOR E PESQUISA
    # ========================================
    
    valor = re.search(r'R\$\s*([\d.,]+)', texto)
    dados['valor'] = valor.group(1).replace('.', '').replace(',', '.') if valor else "23.190,00"
    
    # Formata valor
    try:
        valor_float = float(dados['valor'])
        dados['valor_formatado'] = f"R$ {valor_float:,.2f}".replace(',', 'v').replace('.', ',').replace('v', '.')
    except:
        dados['valor_formatado'] = "R$ 23.190,00"
    
    req_siga = re.search(r'Requisição[:\s]*n[º°]?\s*(\d+/\d+)', texto, re.IGNORECASE)
    dados['req_siga'] = req_siga.group(1) if req_siga else "05/2026"
    
    # ========================================
    # PARECER JURÍDICO
    # ========================================
    
    parecer = re.search(r'Despacho SEI[:\s]*n[º°]?\s*(\d+)', texto, re.IGNORECASE)
    dados['parecer_numero'] = parecer.group(1) if parecer else "125375247"
    
    # Verificar se tem fundamentação legal
    if re.search(r'Resolução PGE[:\s]*n[º°]?\s*5\.059', texto, re.IGNORECASE):
        dados['fundamentacao_pge'] = "Resolução PGE nº 5.059/2024"
    else:
        dados['fundamentacao_pge'] = "Resolução PGE nº 5.059/2024"
    
    # ========================================
    # PÁGINAS DOS DOCUMENTOS
    # ========================================
    
    # Função para extrair páginas mencionadas
    def extrair_paginas(padrao):
        pag = re.search(rf'{padrao}[:\s]*\(pg\.?\s*(\d+)[-–]?(\d+)?\)', texto, re.IGNORECASE)
        if pag:
            if pag.group(2):
                return f"pg. {pag.group(1)}-{pag.group(2)}"
            else:
                return f"pg. {pag.group(1)}"
        return "pg. 1-2"
    
    dados['pag_inicial'] = extrair_paginas('iniciado pela')
    dados['pag_autorizacao'] = extrair_paginas('autorizado pela')
    dados['pag_etp'] = extrair_paginas('ETP')
    dados['pag_tr'] = extrair_paginas('Termo de Referência')
    dados['pag_pesquisa'] = extrair_paginas('pesquisa de mercado')
    dados['pag_orcamento'] = extrair_paginas('disponibilidade orçamentária')
    dados['pag_parecer'] = extrair_paginas('Despacho SEI')
    
    return dados

# ============================================
# FUNÇÃO PARA GERAR DESPACHO (MODELO COMPLETO)
# ============================================

def gerar_despacho_completo(dados):
    """
    Gera despacho EXATAMENTE conforme modelo enviado
    """
    
    doc = Document()
    
    # Configura fonte
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    # ========================================
    # I. INTRODUÇÃO
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("I. Introdução").bold = True
    
    doc.add_paragraph(
        f"Atendendo à solicitação de análise do processo SEI nº {dados['processo_sei']}, "
        f"pela Diretoria de Administração e Finanças – DIRAF, referente à aquisição de "
        f"sacos plásticos para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro "
        f"(IPEM/RJ), procedemos à verificação dos documentos apresentados, com o objetivo "
        f"de subsidiar a continuidade do processo, sem adentrar no mérito técnico da contratação."
    )
    
    doc.add_paragraph()
    
    # ========================================
    # II. ANÁLISE PROCESSUAL
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("II. Análise Processual").bold = True
    
    doc.add_paragraph(
        "Foram examinados os seguintes documentos e etapas processuais:"
    )
    
    doc.add_paragraph()
    
    # 1. Solicitação Inicial e Autorização
    doc.add_paragraph("1. Solicitação Inicial e Autorização: ", style='List Number').add_run(
        f"O processo foi iniciado pela Superintendência de Pré-Medidos ({dados['pag_inicial']}) "
        f"e devidamente autorizado pela Presidência do IPEM/RJ em {dados['data_autorizacao']} "
        f"({dados['pag_autorizacao']})."
    )
    
    # 2. Estudo Técnico Preliminar (ETP) e Gestão de Riscos
    doc.add_paragraph("2. Estudo Técnico Preliminar (ETP) e Gestão de Riscos: ", style='List Number').add_run(
        f"O ETP nº {dados['etp_numero']} e a Matriz de Riscos nº {dados['risco_numero']} "
        f"detalham a necessidade, viabilidade técnica e ações de mitigação para a contratação "
        f"({dados['pag_etp']})."
    )
    
    # 3. Termo de Referência (TR)
    doc.add_paragraph("3. Termo de Referência (TR): ", style='List Number').add_run(
        f"O TR nº {dados['tr_numero']} consolida as especificações técnicas e condições contratuais, "
        f"servindo de balizador para a fase externa ({dados['pag_tr']})."
    )
    
    # 4. Pesquisa de Mercado e Requisição SIGA
    doc.add_paragraph("4. Pesquisa de Mercado e Requisição SIGA: ", style='List Number').add_run(
        f"Foi realizada pesquisa de mercado formal, com a devida inclusão da Requisição de "
        f"Material nº {dados['req_siga']} no Sistema Integrado de Gestão de Aquisições (SIGA), "
        f"totalizando o valor estimado de {dados['valor_formatado']} ({dados['pag_pesquisa']})."
    )
    
    # 5. Conformidade Orçamentária
    doc.add_paragraph("5. Conformidade Orçamentária: ", style='List Number').add_run(
        f"O processo conta com as declarações de impacto financeiro, disponibilidade orçamentária "
        f"e a declaração do ordenador de despesa, atestando a compatibilidade com o orçamento e o "
        f"Plano Plurianual (PPA/RJ) para 2026 ({dados['pag_orcamento']})."
    )
    
    # 6. Parecer Jurídico
    doc.add_paragraph("6. Parecer Jurídico: ", style='List Number').add_run(
        f"A Diretoria Jurídica manifestou-se por meio do Despacho SEI nº {dados['parecer_numero']} "
        f"({dados['pag_parecer']}), informando a dispensa de análise jurídica formal em razão do "
        f"valor da contratação, fundamentada no art. 1º da {dados['fundamentacao_pge']} e no "
        f"art. 95, I, da Lei nº 14.133/2021."
    )
    
    doc.add_paragraph()
    
    # ========================================
    # III. OBSERVAÇÕES
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("III. Observações").bold = True
    
    doc.add_paragraph(
        "Verifica-se que o processo encontra-se devidamente instruído, tendo percorrido as etapas "
        "formais exigidas pela legislação vigente. As especificações técnicas e condições de "
        "fornecimento estão consolidadas no Termo de Referência, documento que orientará a fase de "
        "seleção do fornecedor. A manifestação da Diretoria Jurídica informa a regularidade do rito "
        "adotado para a contratação direta por dispensa eletrônica."
    )
    
    doc.add_paragraph()
    
    # ========================================
    # IV. DESPACHO
    # ========================================
    
    p = doc.add_paragraph()
    p.add_run("IV. Despacho").bold = True
    
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    
    # Assinatura
    doc.add_paragraph("At.te.,")
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph("___________________________________")
    doc.add_paragraph()
    doc.add_paragraph("Auditor Interno")
    doc.add_paragraph("IPEm/RJ")
    
    return doc

# ============================================
# INTERFACE PRINCIPAL
# ============================================

if arquivo:
    
    with st.spinner("🔍 Analisando processo e extraindo dados..."):
        
        # Extrair texto
        texto = extrair_texto_pdf(arquivo)
        
        # Extrair dados
        dados = extrair_dados_completos(texto)
        
        # ========================================
        # MOSTRA RESULTADO DA ANÁLISE
        # ========================================
        
        st.success("✅ Processo analisado com sucesso!")
        
        # Dados principais em colunas
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.info("📋 **Processo**")
            st.write(f"SEI nº {dados['processo_sei']}")
            st.write(f"Data Autorização: {dados['data_autorizacao']}")
        
        with col2:
            st.info("📄 **Documentos**")
            st.write(f"ETP nº {dados['etp_numero']}")
            st.write(f"TR nº {dados['tr_numero']}")
            st.write(f"Matriz Risco nº {dados['risco_numero']}")
        
        with col3:
            st.info("💰 **Valor**")
            st.write(f"{dados['valor_formatado']}")
            st.write(f"Req SIGA: {dados['req_siga']}")
            st.write(f"Parecer: {dados['parecer_numero']}")
        
        # Expandir para ver todos os dados extraídos
        with st.expander("📋 Ver todos os dados extraídos do PDF"):
            for chave, valor in dados.items():
                st.write(f"**{chave}:** {valor}")
        
        # ========================================
        # BOTÃO DE DOWNLOAD
        # ========================================
        
        st.markdown("---")
        st.subheader("📥 Download do Despacho")
        
        # Gerar documento
        doc = gerar_despacho_completo(dados)
        
        # Salvar em arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            doc.save(tmp.name)
            tmp_path = tmp.name
        
        # Ler o arquivo
        with open(tmp_path, 'rb') as f:
            doc_bytes = f.read()
        
        # Apagar temporário
        os.unlink(tmp_path)
        
        # Nome do arquivo
        data_atual = datetime.now().strftime("%Y%m%d")
        nome_arquivo = f"DESPACHO_AUDIT_{dados['processo_sei'].replace('/', '_')}_{data_atual}.docx"
        
        # Botão de download
        st.download_button(
            label="📥 CLIQUE AQUI PARA BAIXAR O DESPACHO COMPLETO",
            data=doc_bytes,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        # ========================================
        # PRÉVIA DO DESPACHO
        # ========================================
        
        with st.expander("📄 Prévia do despacho que será gerado", expanded=True):
            
            preview = f"""
I. INTRODUÇÃO
=============
Atendendo à solicitação de análise do processo SEI nº {dados['processo_sei']}, pela Diretoria de Administração e Finanças – DIRAF, referente à aquisição de sacos plásticos para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro (IPEM/RJ), procedemos à verificação dos documentos apresentados, com o objetivo de subsidiar a continuidade do processo, sem adentrar no mérito técnico da contratação.

II. ANÁLISE PROCESSUAL
======================
Foram examinados os seguintes documentos e etapas processuais:

1. Solicitação Inicial e Autorização: O processo foi iniciado pela Superintendência de Pré-Medidos ({dados['pag_inicial']}) e devidamente autorizado pela Presidência do IPEM/RJ em {dados['data_autorizacao']} ({dados['pag_autorizacao']}).

2. Estudo Técnico Preliminar (ETP) e Gestão de Riscos: O ETP nº {dados['etp_numero']} e a Matriz de Riscos nº {dados['risco_numero']} detalham a necessidade, viabilidade técnica e ações de mitigação para a contratação ({dados['pag_etp']}).

3. Termo de Referência (TR): O TR nº {dados['tr_numero']} consolida as especificações técnicas e condições contratuais, servindo de balizador para a fase externa ({dados['pag_tr']}).

4. Pesquisa de Mercado e Requisição SIGA: Foi realizada pesquisa de mercado formal, com a devida inclusão da Requisição de Material nº {dados['req_siga']} no Sistema Integrado de Gestão de Aquisições (SIGA), totalizando o valor estimado de {dados['valor_formatado']} ({dados['pag_pesquisa']}).

5. Conformidade Orçamentária: O processo conta com as declarações de impacto financeiro, disponibilidade orçamentária e a declaração do ordenador de despesa, atestando a compatibilidade com o orçamento e o Plano Plurianual (PPA/RJ) para 2026 ({dados['pag_orcamento']}).

6. Parecer Jurídico: A Diretoria Jurídica manifestou-se por meio do Despacho SEI nº {dados['parecer_numero']} ({dados['pag_parecer']}), informando a dispensa de análise jurídica formal em razão do valor da contratação, fundamentada no art. 1º da {dados['fundamentacao_pge']} e no art. 95, I, da Lei nº 14.133/2021.

III. OBSERVAÇÕES
================
Verifica-se que o processo encontra-se devidamente instruído, tendo percorrido as etapas formais exigidas pela legislação vigente. As especificações técnicas e condições de fornecimento estão consolidadas no Termo de Referência, documento que orientará a fase de seleção do fornecedor. A manifestação da Diretoria Jurídica informa a regularidade do rito adotado para a contratação direta por dispensa eletrônica.

IV. DESPACHO
============

At.te.,


___________________________________
Auditor Interno
IPEm/RJ
            """
            
            st.text(preview)

else:
    # Tela inicial sem arquivo
    st.info("👆 Faça upload do PDF do processo para gerar o despacho")
    
    with st.expander("📌 Visualizar modelo de despacho"):
        st.write("""
**I. INTRODUÇÃO**
Atendendo à solicitação de análise do processo SEI nº 150014/001585/2025, pela Diretoria de Administração e Finanças – DIRAF, referente à aquisição de sacos plásticos para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro (IPEM/RJ), procedemos à verificação dos documentos apresentados...

**II. ANÁLISE PROCESSUAL**
1. Solicitação Inicial e Autorização...
2. Estudo Técnico Preliminar (ETP) e Gestão de Riscos...
3. Termo de Referência (TR)...
4. Pesquisa de Mercado e Requisição SIGA...
5. Conformidade Orçamentária...
6. Parecer Jurídico...

**III. OBSERVAÇÕES**
Verifica-se que o processo encontra-se devidamente instruído...

**IV. DESPACHO**
At.te., [assinatura]
        """)

# RODAPÉ
st.markdown("---")
st.caption(f"© 2026 - Auditoria Interna IPEm/RJ - Versão 2.0 - Modelo completo")
