# ============================================
# DESPACHO AUDIT - VERSÃO TRANSPARENTE
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
import io
from docx import Document
from docx.shared import Pt
import tempfile
import os

st.set_page_config(page_title="Despacho AUDIT - IPEm/RJ", page_icon="📄", layout="wide")
st.title("📄 Despacho AUDIT - IPEm/RJ")
st.markdown("---")

# UPLOAD
arquivo = st.file_uploader("Selecione o PDF do processo", type=['pdf'])

def extrair_texto_pdf(arquivo):
    texto = ""
    with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
        for pagina in pdf.pages:
            if pagina.extract_text():
                texto += pagina.extract_text() + "\n"
    return texto

def extrair_dados_do_texto(texto):
    """
    Extrai dados REALMENTE do texto, com valores padrão apenas se não encontrar
    """
    dados = {}
    
    # ========================================
    # MOSTRAR O QUE ESTÁ SENDO ENCONTRADO
    # ========================================
    
    st.subheader("🔍 O que o app encontrou no seu PDF:")
    
    # 1. Processo SEI
    proc = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    if proc:
        dados['processo_sei'] = proc.group(1)
        st.success(f"✅ Processo encontrado: {dados['processo_sei']}")
    else:
        dados['processo_sei'] = "NÃO ENCONTRADO"
        st.error("❌ Processo não encontrado no PDF")
    
    # 2. Data
    data = re.search(r'(\d{1,2})[/](\d{1,2})[/](\d{4})', texto)
    if data:
        dados['data_autorizacao'] = f"{data.group(1)}/{data.group(2)}/{data.group(3)}"
        st.success(f"✅ Data encontrada: {dados['data_autorizacao']}")
    else:
        dados['data_autorizacao'] = "NÃO ENCONTRADA"
        st.error("❌ Data não encontrada")
    
    # 3. ETP
    etp = re.search(r'ETP[:\s]*n[º°]?\s*(\d+/\d+)', texto, re.IGNORECASE)
    if etp:
        dados['etp_numero'] = etp.group(1)
        st.success(f"✅ ETP encontrado: {dados['etp_numero']}")
    else:
        dados['etp_numero'] = "NÃO ENCONTRADO"
        st.warning("⚠️ ETP não encontrado")
    
    # 4. TR
    tr = re.search(r'TR[:\s]*n[º°]?\s*(\d+/\d+)', texto, re.IGNORECASE)
    if tr:
        dados['tr_numero'] = tr.group(1)
        st.success(f"✅ TR encontrado: {dados['tr_numero']}")
    else:
        dados['tr_numero'] = "NÃO ENCONTRADO"
        st.warning("⚠️ TR não encontrado")
    
    # 5. Valor
    valor = re.search(r'R\$\s*([\d.,]+)', texto)
    if valor:
        dados['valor'] = valor.group(1)
        st.success(f"✅ Valor encontrado: R$ {dados['valor']}")
    else:
        dados['valor'] = "NÃO ENCONTRADO"
        st.error("❌ Valor não encontrado")
    
    # 6. Parecer
    parecer = re.search(r'Despacho SEI[:\s]*n[º°]?\s*(\d+)', texto, re.IGNORECASE)
    if parecer:
        dados['parecer_numero'] = parecer.group(1)
        st.success(f"✅ Parecer encontrado: {dados['parecer_numero']}")
    else:
        dados['parecer_numero'] = "NÃO ENCONTRADO"
        st.warning("⚠️ Parecer não encontrado")
    
    return dados

if arquivo:
    
    with st.spinner("🔍 Analisando processo..."):
        
        # Extrair texto
        texto = extrair_texto_pdf(arquivo)
        
        # Extrair dados (já mostra na tela o que encontrou)
        dados = extrair_dados_do_texto(texto)
        
        st.markdown("---")
        
        # ========================================
        # DECISÃO: PODE GERAR O DESPACHO?
        # ========================================
        
        # Verificar se encontrou dados suficientes
        dados_essenciais = ['processo_sei', 'valor']
        dados_encontrados = sum(1 for d in dados_essenciais if dados.get(d) != "NÃO ENCONTRADO")
        
        if dados_encontrados >= 1:
            
            st.success("✅ Dados suficientes encontrados! Pode gerar o despacho.")
            
            # Botão para gerar
            if st.button("📄 GERAR DESPACHO"):
                
                # Criar documento Word
                doc = Document()
                style = doc.styles['Normal']
                style.font.name = 'Arial'
                style.font.size = Pt(12)
                
                # I. INTRODUÇÃO
                doc.add_paragraph().add_run("I. Introdução").bold = True
                doc.add_paragraph(
                    f"Atendendo à solicitação de análise do processo SEI nº {dados['processo_sei']}, "
                    f"pela Diretoria de Administração e Finanças – DIRAF, referente à aquisição de "
                    f"sacos plásticos para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro "
                    f"(IPEM/RJ), procedemos à verificação dos documentos apresentados..."
                )
                
                doc.add_paragraph()
                
                # II. ANÁLISE PROCESSUAL
                doc.add_paragraph().add_run("II. Análise Processual").bold = True
                doc.add_paragraph("Foram examinados os seguintes documentos e etapas processuais:")
                doc.add_paragraph()
                
                # Item 1
                doc.add_paragraph("1. Solicitação Inicial e Autorização: ", style='List Number').add_run(
                    f"O processo foi iniciado pela Superintendência de Pré-Medidos "
                    f"e devidamente autorizado pela Presidência do IPEM/RJ em {dados['data_autorizacao']}."
                )
                
                # Item 2
                doc.add_paragraph("2. Estudo Técnico Preliminar (ETP) e Gestão de Riscos: ", style='List Number').add_run(
                    f"O ETP nº {dados['etp_numero'] if dados['etp_numero'] != 'NÃO ENCONTRADO' else 'não identificado'} "
                    f"e a Matriz de Riscos detalham a necessidade da contratação."
                )
                
                # Item 3
                doc.add_paragraph("3. Termo de Referência (TR): ", style='List Number').add_run(
                    f"O TR nº {dados['tr_numero'] if dados['tr_numero'] != 'NÃO ENCONTRADO' else 'não identificado'} "
                    f"consolida as especificações técnicas."
                )
                
                # Item 4
                doc.add_paragraph("4. Pesquisa de Mercado e Requisição SIGA: ", style='List Number').add_run(
                    f"Foi realizada pesquisa de mercado formal, totalizando o valor estimado de R$ {dados['valor']}."
                )
                
                # Item 5
                doc.add_paragraph("5. Conformidade Orçamentária: ", style='List Number').add_run(
                    f"O processo conta com as declarações de impacto financeiro e disponibilidade orçamentária."
                )
                
                # Item 6
                doc.add_paragraph("6. Parecer Jurídico: ", style='List Number').add_run(
                    f"A Diretoria Jurídica manifestou-se por meio do Despacho SEI nº {dados['parecer_numero']}."
                )
                
                doc.add_paragraph()
                
                # III. OBSERVAÇÕES
                doc.add_paragraph().add_run("III. Observações").bold = True
                doc.add_paragraph(
                    "Verifica-se que o processo encontra-se devidamente instruído, tendo percorrido as etapas "
                    "formais exigidas pela legislação vigente."
                )
                
                doc.add_paragraph()
                
                # IV. DESPACHO
                doc.add_paragraph().add_run("IV. Despacho").bold = True
                doc.add_paragraph()
                doc.add_paragraph()
                doc.add_paragraph("At.te.,")
                doc.add_paragraph()
                doc.add_paragraph("___________________________________")
                doc.add_paragraph("Auditor Interno")
                doc.add_paragraph("IPEm/RJ")
                
                # Salvar e disponibilizar
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                    doc.save(tmp.name)
                    tmp_path = tmp.name
                
                with open(tmp_path, 'rb') as f:
                    doc_bytes = f.read()
                
                os.unlink(tmp_path)
                
                st.download_button(
                    label="📥 BAIXAR DESPACHO",
                    data=doc_bytes,
                    file_name=f"DESPACHO_AUDIT.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        
        else:
            st.error("❌ Não foi possível encontrar dados suficientes no PDF para gerar o despacho.")
            st.info("💡 Dica: Certifique-se de que o PDF contém informações como número do processo e valor.")
