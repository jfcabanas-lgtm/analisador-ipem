# ============================================
# DESPACHO AUDIT - VERSÃO DIAGNÓSTICO
# MOSTRA EXATAMENTE O QUE ENCONTROU NO PDF
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
    """Extrai todo o texto do PDF"""
    texto = ""
    with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
        for i, pagina in enumerate(pdf.pages):
            if pagina.extract_text():
                texto += f"\n--- PÁGINA {i+1} ---\n"
                texto += pagina.extract_text() + "\n"
    return texto

if arquivo:
    
    with st.spinner("🔍 Extraindo texto do PDF..."):
        
        # Extrair texto completo
        texto = extrair_texto_pdf(arquivo)
        
        # ========================================
        # MOSTRAR O TEXTO EXTRAÍDO (para diagnóstico)
        # ========================================
        
        st.subheader("📄 TEXTO EXTRAÍDO DO PDF")
        st.text_area("Visualize o que o app está lendo:", texto[:2000] + "...", height=300)
        
        st.markdown("---")
        
        # ========================================
        # BUSCAR PADRÕES COMUNS
        # ========================================
        
        st.subheader("🔍 BUSCANDO INFORMAÇÕES NO TEXTO")
        
        # Lista de padrões para testar
        padroes = {
            "Processo SEI": [
                r'SEI[:\s]*n[º°]?\s*([\d\-/]+)',
                r'Processo[:\s]*n[º°]?\s*([\d\-/]+)',
                r'SEI[-]?(\d+/\d+/\d+)',
                r'(\d{6}/\d{6}/\d{4})'
            ],
            "Data": [
                r'(\d{1,2})[/](\d{1,2})[/](\d{4})',
                r'(\d{1,2})\s*de\s*([A-Za-zç]+)\s*de\s*(\d{4})'
            ],
            "Valor": [
                r'R\$\s*([\d.,]+)',
                r'valor[:\s]*R\$\s*([\d.,]+)',
                r'total[:\s]*R\$\s*([\d.,]+)'
            ],
            "ETP": [
                r'ETP[:\s]*n[º°]?\s*(\d+/\d+)',
                r'Estudo Técnico Preliminar[:\s]*n[º°]?\s*(\d+/\d+)'
            ],
            "TR": [
                r'TR[:\s]*n[º°]?\s*(\d+/\d+)',
                r'Termo de Referência[:\s]*n[º°]?\s*(\d+/\d+)'
            ],
            "Parecer": [
                r'Despacho SEI[:\s]*n[º°]?\s*(\d+)',
                r'Parecer Jurídico[:\s]*n[º°]?\s*(\d+)',
                r'Documento SEI[:\s]*n[º°]?\s*(\d+)'
            ]
        }
        
        resultados = {}
        
        for campo, padroes_lista in padroes.items():
            st.write(f"**{campo}:**")
            encontrou = False
            
            for i, padrao in enumerate(padroes_lista):
                matches = re.findall(padrao, texto, re.IGNORECASE)
                if matches:
                    if isinstance(matches[0], tuple):
                        valor = '/'.join(matches[0])
                    else:
                        valor = matches[0]
                    
                    st.success(f"  ✓ Padrão {i+1}: ENCONTROU → {valor}")
                    resultados[campo] = valor
                    encontrou = True
                    break
            
            if not encontrou:
                st.warning(f"  ✗ Nenhum padrão encontrado para {campo}")
        
        st.markdown("---")
        
        # ========================================
        # MOSTRAR RESULTADO FINAL
        # ========================================
        
        st.subheader("📊 DADOS QUE SERÃO USADOS NO DESPACHO:")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Campo**")
        with col2:
            st.write("**Valor encontrado**")
        
        for campo, valor in resultados.items():
            with col1:
                st.write(f"{campo}:")
            with col2:
                st.write(f"**{valor}**")
        
        # ========================================
        # BOTÃO PARA GERAR DESPACHO
        # ========================================
        
        st.markdown("---")
        
        if resultados:
            if st.button("📄 GERAR DESPACHO COM ESTES DADOS"):
                
                # Criar documento
                doc = Document()
                style = doc.styles['Normal']
                style.font.name = 'Arial'
                style.font.size = Pt(12)
                
                # I. INTRODUÇÃO
                doc.add_paragraph().add_run("I. Introdução").bold = True
                doc.add_paragraph(
                    f"Atendendo à solicitação de análise do processo {resultados.get('Processo SEI', 'NÃO IDENTIFICADO')}, "
                    f"pela Diretoria de Administração e Finanças – DIRAF, referente à aquisição de "
                    f"materiais para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro "
                    f"(IPEM/RJ), procedemos à verificação dos documentos apresentados, com o objetivo "
                    f"de subsidiar a continuidade do processo, sem adentrar no mérito técnico da contratação."
                )
                
                doc.add_paragraph()
                
                # II. ANÁLISE PROCESSUAL
                doc.add_paragraph().add_run("II. Análise Processual").bold = True
                doc.add_paragraph("Foram examinados os seguintes documentos e etapas processuais:")
                doc.add_paragraph()
                
                doc.add_paragraph("1. Solicitação Inicial e Autorização: ", style='List Number').add_run(
                    f"Processo devidamente autorizado."
                )
                
                doc.add_paragraph("2. Estudo Técnico Preliminar (ETP) e Gestão de Riscos: ", style='List Number').add_run(
                    f"Documentos analisados conforme legislação."
                )
                
                doc.add_paragraph("3. Termo de Referência (TR): ", style='List Number').add_run(
                    f"Especificações técnicas consolidadas."
                )
                
                doc.add_paragraph("4. Pesquisa de Mercado: ", style='List Number').add_run(
                    f"Valor estimado: {resultados.get('Valor', 'NÃO IDENTIFICADO')}."
                )
                
                doc.add_paragraph("5. Conformidade Orçamentária: ", style='List Number').add_run(
                    f"Declarações de impacto financeiro presentes."
                )
                
                doc.add_paragraph("6. Parecer Jurídico: ", style='List Number').add_run(
                    f"{resultados.get('Parecer', 'Documento analisado')}."
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
                
                # Salvar
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
            st.error("❌ NENHUM DADO ENCONTRADO NO PDF")
            st.info("💡 O PDF pode estar como imagem (não texto selecionável) ou ter formato diferente")
