# ============================================
# DESPACHO AUDIT - VERSÃO HÍBRIDA
# ANALISA O PDF E PREENCHE O FORMULÁRIO!
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt
import tempfile
import os
import io

# CONFIGURAÇÃO
st.set_page_config(page_title="IPEm - Despacho Inteligente", page_icon="⚖️", layout="wide")

# CSS
st.markdown("""
<style>
    .header {background: linear-gradient(90deg, #003366 0%, #0047ab 100%); padding: 2rem; border-radius: 20px; color: white; text-align: center; margin-bottom: 2rem;}
    .section-title {color: #003366; font-size: 1.3rem; font-weight: 600; margin: 1.5rem 0 1rem 0; border-bottom: 2px solid #eef2f6; padding-bottom: 0.5rem;}
    .success-box {background: #d4edda; color: #155724; padding: 1rem; border-radius: 10px; border-left: 4px solid #28a745;}
    .warning-box {background: #fff3cd; color: #856404; padding: 1rem; border-radius: 10px; border-left: 4px solid #ffc107;}
</style>
""", unsafe_allow_html=True)

# CABEÇALHO
st.markdown('<div class="header"><h1>⚖️ IPEm - Despacho Inteligente</h1><p>Upload do PDF → Extrai dados → Você confirma → Gera despacho</p></div>', unsafe_allow_html=True)

# ============================================
# PASSO 1: UPLOAD DO PDF
# ============================================

st.markdown('<div class="section-title">📂 PASSO 1: Upload do Processo</div>', unsafe_allow_html=True)

arquivo = st.file_uploader("Selecione o PDF do processo", type=['pdf'])

if arquivo:
    
    with st.spinner("🔍 Analisando PDF e extraindo dados..."):
        
        # Extrair texto do PDF
        texto = ""
        with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
            for pagina in pdf.pages:
                if pagina.extract_text():
                    texto += pagina.extract_text() + "\n"
        
        # ========================================
        # EXTRAIR DADOS DO TEXTO
        # ========================================
        
        def extrair_campo(padroes, texto, default=""):
            for padrao in padroes:
                match = re.search(padrao, texto, re.IGNORECASE)
                if match:
                    return match.group(1).strip()
            return default
        
        dados_extraidos = {
            'processo_sei': extrair_campo([
                r'Processo[:\s]*n[º°]?\s*([\d\-/]+)',
                r'SEI[:\s]*n[º°]?\s*([\d\-/]+)',
                r'(\d{6,}/\d{6,}/\d{4})'
            ], texto, "Não identificado"),
            
            'objeto': extrair_campo([
                r'objeto[:\s]*([^.]+)',
                r'aquisição[:\s]*([^.]+)',
                r'contratação[:\s]*([^.]+)'
            ], texto, "Não identificado"),
            
            'valor': extrair_campo([
                r'R\$\s*([\d.,]+)',
                r'valor[:\s]*R\$\s*([\d.,]+)',
                r'total[:\s]*R\$\s*([\d.,]+)'
            ], texto, ""),
            
            'etp_numero': extrair_campo([
                r'ETP[:\s]*n[º°]?\s*(\d+/\d+)',
                r'Estudo Técnico Preliminar[:\s]*n[º°]?\s*(\d+/\d+)'
            ], texto, ""),
            
            'tr_numero': extrair_campo([
                r'TR[:\s]*n[º°]?\s*(\d+/\d+)',
                r'Termo de Referência[:\s]*n[º°]?\s*(\d+/\d+)'
            ], texto, ""),
            
            'risco_numero': extrair_campo([
                r'Matriz de Riscos[:\s]*n[º°]?\s*(\d+/\d+)',
                r'Gestão de Risco[:\s]*n[º°]?\s*(\d+/\d+)'
            ], texto, ""),
            
            'req_siga': extrair_campo([
                r'Requisição[:\s]*n[º°]?\s*(\d+/\d+)',
                r'SIGA[:\s]*n[º°]?\s*(\d+/\d+)'
            ], texto, ""),
            
            'parecer_numero': extrair_campo([
                r'Despacho SEI[:\s]*n[º°]?\s*(\d+)',
                r'Parecer[:\s]*n[º°]?\s*(\d+)'
            ], texto, ""),
            
            'data_autorizacao': extrair_campo([
                r'autorizado[:\s]*em[:\s]*(\d{1,2}[/]\d{1,2}[/]\d{4})',
                r'(\d{1,2}[/]\d{1,2}[/]\d{4})'
            ], texto, "")
        }
        
        # Extrair SEIs (todos os números de documento)
        seis_encontrados = re.findall(r'SEI[:\s]*n[º°]?\s*(\d+)', texto, re.IGNORECASE)
        
        # ========================================
        # PASSO 2: MOSTRAR RESULTADO DA ANÁLISE
        # ========================================
        
        st.markdown('<div class="section-title">🔍 PASSO 2: Dados Encontrados no PDF</div>', unsafe_allow_html=True)
        
        col_res1, col_res2, col_res3 = st.columns(3)
        
        with col_res1:
            st.info("📋 **Processo**")
            st.write(f"**Nº:** {dados_extraidos['processo_sei']}")
            st.write(f"**Data:** {dados_extraidos['data_autorizacao']}")
        
        with col_res2:
            st.info("💰 **Valor**")
            if dados_extraidos['valor']:
                try:
                    valor_float = float(dados_extraidos['valor'].replace('.', '').replace(',', '.'))
                    st.write(f"**R$ {valor_float:,.2f}**")
                except:
                    st.write(f"**R$ {dados_extraidos['valor']}**")
            else:
                st.write("**Não identificado**")
        
        with col_res3:
            st.info("📑 **Documentos**")
            st.write(f"ETP: {dados_extraidos['etp_numero'] or '❌'}")
            st.write(f"TR: {dados_extraidos['tr_numero'] or '❌'}")
            st.write(f"Parecer: {dados_extraidos['parecer_numero'] or '❌'}")
        
        # Mostrar SEIs encontrados
        if seis_encontrados:
            with st.expander(f"📎 {len(seis_encontrados)} números SEI encontrados"):
                for sei in seis_encontrados[:10]:
                    st.write(f"• SEI {sei}")
        
        # ========================================
        # PASSO 3: FORMULÁRIO PARA CONFIRMAÇÃO
        # ========================================
        
        st.markdown('<div class="section-title">✏️ PASSO 3: Confirme e Ajuste os Dados</div>', unsafe_allow_html=True)
        
        with st.form("form_confirmacao"):
            
            st.markdown("##### 📋 Dados do Processo")
            col1, col2 = st.columns(2)
            
            with col1:
                processo_sei = st.text_input(
                    "Nº do Processo SEI *", 
                    value=dados_extraidos['processo_sei'],
                    help="Confira se o número está correto"
                )
                
                objeto = st.text_input(
                    "Objeto da Contratação *",
                    value=dados_extraidos['objeto'],
                    help="Descreva o objeto"
                )
                
                data_autorizacao = st.date_input(
                    "Data da Autorização",
                    value=datetime.strptime(dados_extraidos['data_autorizacao'], '%d/%m/%Y') if dados_extraidos['data_autorizacao'] != "Não identificado" and '/' in dados_extraidos['data_autorizacao'] else None,
                    format="DD/MM/YYYY"
                )
            
            with col2:
                valor_input = st.text_input(
                    "Valor (R$)",
                    value=dados_extraidos['valor'],
                    help="Ex: 23.190,00"
                )
                
                sei_inicio = st.text_input(
                    "SEI da Solicitação Inicial",
                    value=seis_encontrados[0] if len(seis_encontrados) > 0 else ""
                )
                
                sei_autorizacao = st.text_input(
                    "SEI da Autorização",
                    value=seis_encontrados[1] if len(seis_encontrados) > 1 else ""
                )
            
            st.markdown("##### 📑 Documentos Técnicos")
            col3, col4 = st.columns(2)
            
            with col3:
                etp_numero = st.text_input("Nº do ETP", value=dados_extraidos['etp_numero'])
                sei_etp = st.text_input("SEI do ETP", value=seis_encontrados[2] if len(seis_encontrados) > 2 else "")
                tr_numero = st.text_input("Nº do TR", value=dados_extraidos['tr_numero'])
                sei_tr = st.text_input("SEI do TR", value=seis_encontrados[3] if len(seis_encontrados) > 3 else "")
            
            with col4:
                risco_numero = st.text_input("Nº da Matriz de Riscos", value=dados_extraidos['risco_numero'])
                sei_risco = st.text_input("SEI da Matriz de Riscos", value=seis_encontrados[4] if len(seis_encontrados) > 4 else "")
                req_siga = st.text_input("Nº da Requisição SIGA", value=dados_extraidos['req_siga'])
                parecer_numero = st.text_input("Nº do Parecer Jurídico", value=dados_extraidos['parecer_numero'])
            
            st.markdown("##### ⚖️ Documentos Jurídicos e Orçamentários")
            col5, col6 = st.columns(2)
            
            with col5:
                sei_impacto = st.text_input("SEI da Declaração de Impacto", value=seis_encontrados[5] if len(seis_encontrados) > 5 else "")
                sei_disponibilidade = st.text_input("SEI da Disponibilidade Orçamentária", value=seis_encontrados[6] if len(seis_encontrados) > 6 else "")
                sei_ordenador = st.text_input("SEI da Declaração do Ordenador", value=seis_encontrados[7] if len(seis_encontrados) > 7 else "")
            
            with col6:
                fundamentacao = st.text_area(
                    "Fundamentação Legal",
                    value="art. 1º da Resolução PGE nº 5.059/2024 e no art. 95, I, da Lei nº 14.133/2021",
                    height=100
                )
            
            observacoes = st.text_area(
                "Observações",
                placeholder="Inclua observações adicionais se necessário...",
                height=100
            )
            
            st.markdown("---")
            
            # BOTÃO PARA GERAR DESPACHO
            gerar = st.form_submit_button(
                "✅ GERAR DESPACHO EM WORD",
                use_container_width=True,
                type="primary"
            )
            
            if gerar:
                
                if not processo_sei or not objeto:
                    st.error("❌ Processo e Objeto são obrigatórios!")
                else:
                    
                    with st.spinner("Gerando despacho..."):
                        
                        # Criar documento Word
                        doc = Document()
                        style = doc.styles['Normal']
                        style.font.name = 'Arial'
                        style.font.size = Pt(12)
                        
                        # I. INTRODUÇÃO
                        doc.add_paragraph().add_run("I. Introdução").bold = True
                        doc.add_paragraph(
                            f"Atendendo à solicitação de análise do processo SEI nº {processo_sei}, "
                            f"pela Diretoria de Administração e Finanças – DIRAF, referente à {objeto} "
                            f"para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro (IPEM/RJ), "
                            f"procedemos à verificação dos documentos apresentados, com o objetivo de "
                            f"subsidiar a continuidade do processo, sem adentrar no mérito técnico da contratação."
                        )
                        doc.add_paragraph()
                        
                        # II. ANÁLISE PROCESSUAL
                        doc.add_paragraph().add_run("II. Análise Processual").bold = True
                        doc.add_paragraph("Foram examinados os seguintes documentos e etapas processuais:")
                        doc.add_paragraph()
                        
                        # Item 1
                        p = doc.add_paragraph()
                        p.add_run("1. Solicitação Inicial e Autorização: ").bold = True
                        p.add_run(
                            f"O processo foi iniciado pela Superintendência de Pré-Medidos "
                            f"{'SEI ' + sei_inicio if sei_inicio else ''} e devidamente autorizado "
                            f"pela Presidência do IPEM/RJ em {data_autorizacao.strftime('%d/%m/%Y') if data_autorizacao else 'data não informada'} "
                            f"{'SEI ' + sei_autorizacao if sei_autorizacao else ''}."
                        )
                        
                        # Item 2
                        p = doc.add_paragraph()
                        p.add_run("2. Estudo Técnico Preliminar-ETP e Gestão de Riscos: ").bold = True
                        p.add_run(
                            f"O ETP nº {etp_numero if etp_numero else 'não informado'} "
                            f"{'SEI ' + sei_etp if sei_etp else ''} e a Matriz de Riscos nº {risco_numero if risco_numero else 'não informada'} "
                            f"{'SEI ' + sei_risco if sei_risco else ''} detalham a necessidade, viabilidade técnica e ações de mitigação para a contratação."
                        )
                        
                        # Item 3
                        p = doc.add_paragraph()
                        p.add_run("3. Termo de Referência-TR: ").bold = True
                        p.add_run(
                            f"O TR nº {tr_numero if tr_numero else 'não informado'}, "
                            f"{'SEI ' + sei_tr if sei_tr else ''}, consolida as especificações técnicas e condições contratuais, "
                            f"servindo de balizador para a fase externa."
                        )
                        
                        # Item 4
                        p = doc.add_paragraph()
                        p.add_run("4. Pesquisa de Mercado e Requisição SIGA: ").bold = True
                        texto4 = (
                            f"Foi realizada pesquisa de mercado formal, com a devida inclusão da Requisição de Material nº {req_siga if req_siga else 'não informada'} "
                            f"no Sistema Integrado de Gestão de Aquisições (SIGA)"
                        )
                        if valor_input:
                            texto4 += f", totalizando o valor estimado de R$ {valor_input}"
                        texto4 += "."
                        p.add_run(texto4)
                        
                        # Item 5
                        p = doc.add_paragraph()
                        p.add_run("5. Conformidade Orçamentária: ").bold = True
                        p.add_run(
                            f"O processo conta com as declarações de impacto financeiro {'SEI ' + sei_impacto if sei_impacto else ''}, "
                            f"disponibilidade orçamentária {'SEI ' + sei_disponibilidade if sei_disponibilidade else ''} e a declaração do ordenador de despesa "
                            f"{'SEI ' + sei_ordenador if sei_ordenador else ''}, atestando a compatibilidade com o orçamento e o Plano Plurianual (PPA/RJ) para 2026."
                        )
                        
                        # Item 6
                        p = doc.add_paragraph()
                        p.add_run("6. Parecer Jurídico: ").bold = True
                        p.add_run(
                            f"A Diretoria Jurídica manifestou-se por meio do Despacho SEI {parecer_numero if parecer_numero else 'não informado'}, "
                            f"informando a dispensa de análise jurídica formal em razão do valor da contratação, fundamentada na {fundamentacao}."
                        )
                        
                        doc.add_paragraph()
                        
                        # III. OBSERVAÇÕES
                        doc.add_paragraph().add_run("III. Observações").bold = True
                        if observacoes:
                            doc.add_paragraph(observacoes)
                        else:
                            doc.add_paragraph(
                                "Verifica-se que o processo encontra-se devidamente instruído, tendo percorrido as etapas "
                                "formais exigidas pela legislação vigente. As especificações técnicas e condições de "
                                "fornecimento estão consolidadas no Termo de Referência, documento que orientará a fase de "
                                "seleção do fornecedor."
                            )
                        
                        doc.add_paragraph()
                        
                        # IV. DESPACHO
                        doc.add_paragraph().add_run("IV. Despacho").bold = True
                        doc.add_paragraph(
                            "Dessa forma, e considerando que os atos administrativos até o presente momento se mostram "
                            "formalmente adequados e em conformidade com a Lei nº 14.133/2021 e demais normas aplicáveis, "
                            "indicamos à continuidade do processo."
                        )
                        
                        doc.add_paragraph()
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
                        
                        # Sucesso
                        st.success("✅ DESPACHO GERADO COM SUCESSO!")
                        st.balloons()
                        
                        st.download_button(
                            label="📥 BAIXAR DESPACHO",
                            data=doc_bytes,
                            file_name=f"DESPACHO_{processo_sei.replace('/', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
