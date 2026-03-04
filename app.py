# ============================================
# DESPACHO AUDIT - VERSÃO FORMULÁRIO
# PREENCHA OS DADOS E O DESPACHO SAI PERFEITO!
# ============================================

import streamlit as st
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os
import io

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="IPEm - Gerador de Despachos",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS PERSONALIZADO
st.markdown("""
<style>
    .header {
        background: linear-gradient(90deg, #003366 0%, #0047ab 100%);
        padding: 2rem;
        border-radius: 20px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    .header h1 {
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
    }
    .header p {
        font-size: 1.1rem;
        opacity: 0.9;
    }
    .card {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        border: 1px solid #eef2f6;
        margin-bottom: 1rem;
    }
    .section-title {
        color: #003366;
        font-size: 1.3rem;
        font-weight: 600;
        margin-bottom: 1.5rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #eef2f6;
    }
    .info-box {
        background: #f0f4ff;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #003366;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# CABEÇALHO
st.markdown("""
<div class="header">
    <h1>⚖️ IPEm - Gerador de Despachos</h1>
    <p>Preencha os dados do processo e gere o despacho automaticamente</p>
    <p style="font-size: 0.9rem; margin-top: 1rem;">Modelo conforme Lei 14.133/2021</p>
</div>
""", unsafe_allow_html=True)

# ============================================
# FORMULÁRIO DE ENTRADA
# ============================================

with st.form("form_despacho"):
    
    st.markdown('<div class="section-title">📋 DADOS DO PROCESSO</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        processo_sei = st.text_input(
            "Nº do Processo SEI *",
            placeholder="Ex: 150014/001585/2025",
            help="Número completo do processo"
        )
        
        objeto = st.text_input(
            "Objeto da Contratação *",
            placeholder="Ex: aquisição de sacos plásticos",
            help="O que está sendo contratado"
        )
        
        data_autorizacao = st.date_input(
            "Data da Autorização",
            format="DD/MM/YYYY",
            help="Data em que o processo foi autorizado"
        )
    
    with col2:
        sei_inicio = st.text_input(
            "SEI da Solicitação Inicial",
            placeholder="Ex: 12345678",
            help="Número do documento SEI da solicitação"
        )
        
        sei_autorizacao = st.text_input(
            "SEI da Autorização",
            placeholder="Ex: 12345679",
            help="Número do documento SEI da autorização"
        )
        
        data_inicio = st.date_input(
            "Data de Início",
            format="DD/MM/YYYY",
            help="Data em que o processo foi iniciado"
        )
    
    st.markdown('<div class="section-title">📑 DOCUMENTOS TÉCNICOS</div>', unsafe_allow_html=True)
    
    col3, col4 = st.columns(2)
    
    with col3:
        etp_numero = st.text_input(
            "Nº do ETP",
            placeholder="Ex: 31/2025",
            help="Estudo Técnico Preliminar"
        )
        
        sei_etp = st.text_input(
            "SEI do ETP",
            placeholder="Ex: 12345680"
        )
        
        tr_numero = st.text_input(
            "Nº do TR",
            placeholder="Ex: 46/2025",
            help="Termo de Referência"
        )
        
        sei_tr = st.text_input(
            "SEI do TR",
            placeholder="Ex: 12345681"
        )
    
    with col4:
        risco_numero = st.text_input(
            "Nº da Matriz de Riscos",
            placeholder="Ex: 26/2025"
        )
        
        sei_risco = st.text_input(
            "SEI da Matriz de Riscos",
            placeholder="Ex: 12345682"
        )
        
        req_siga = st.text_input(
            "Nº da Requisição SIGA",
            placeholder="Ex: 05/2026",
            help="Requisição de Material no SIGA"
        )
        
        valor = st.number_input(
            "Valor Total (R$)",
            min_value=0.0,
            step=100.0,
            format="%.2f",
            help="Valor estimado da contratação"
        )
    
    st.markdown('<div class="section-title">⚖️ DOCUMENTOS JURÍDICOS E ORÇAMENTÁRIOS</div>', unsafe_allow_html=True)
    
    col5, col6 = st.columns(2)
    
    with col5:
        sei_impacto = st.text_input(
            "SEI da Declaração de Impacto",
            placeholder="Ex: 12345683"
        )
        
        sei_disponibilidade = st.text_input(
            "SEI da Disponibilidade Orçamentária",
            placeholder="Ex: 12345684"
        )
        
        sei_ordenador = st.text_input(
            "SEI da Declaração do Ordenador",
            placeholder="Ex: 12345685"
        )
    
    with col6:
        parecer_numero = st.text_input(
            "Nº do Despacho Jurídico",
            placeholder="Ex: 125375247"
        )
        
        fundamentacao = st.text_area(
            "Fundamentação Legal",
            value="art. 1º da Resolução PGE nº 5.059/2024 e no art. 95, I, da Lei nº 14.133/2021",
            height=80
        )
    
    st.markdown('<div class="section-title">📝 OBSERVAÇÕES</div>', unsafe_allow_html=True)
    
    observacoes = st.text_area(
        "Observações (opcional)",
        placeholder="Inclua aqui qualquer observação adicional relevante...",
        height=100
    )
    
    st.markdown("---")
    
    # BOTÃO DE ENVIO
    submitted = st.form_submit_button(
        "✅ GERAR DESPACHO EM WORD",
        use_container_width=True,
        type="primary"
    )

# ============================================
# PROCESSAR FORMULÁRIO
# ============================================

if submitted:
    
    # Validar campos obrigatórios
    if not processo_sei or not objeto:
        st.error("❌ Preencha pelo menos o número do processo e o objeto da contratação!")
    else:
        
        with st.spinner("⏳ Gerando despacho..."):
            
            # Criar documento Word
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Arial'
            style.font.size = Pt(12)
            
            # ========================================
            # I. INTRODUÇÃO
            # ========================================
            p = doc.add_paragraph()
            p.add_run("I. Introdução").bold = True
            
            texto_intro = (
                f"Atendendo à solicitação de análise do processo SEI nº {processo_sei}, "
                f"pela Diretoria de Administração e Finanças – DIRAF, referente à {objeto} "
                f"para o Instituto de Pesos e Medidas do Estado do Rio de Janeiro (IPEM/RJ), "
                f"procedemos à verificação dos documentos apresentados, com o objetivo de "
                f"subsidiar a continuidade do processo, sem adentrar no mérito técnico da contratação."
            )
            doc.add_paragraph(texto_intro)
            doc.add_paragraph()
            
            # ========================================
            # II. ANÁLISE PROCESSUAL
            # ========================================
            p = doc.add_paragraph()
            p.add_run("II. Análise Processual").bold = True
            doc.add_paragraph("Foram examinados os seguintes documentos e etapas processuais:")
            doc.add_paragraph()
            
            # Item 1
            p = doc.add_paragraph()
            p.add_run("1. Solicitação Inicial e Autorização: ").bold = True
            texto1 = (
                f"O processo foi iniciado pela Superintendência de Pré-Medidos "
                f"{'SEI ' + sei_inicio if sei_inicio else ''} e devidamente autorizado "
                f"pela Presidência do IPEM/RJ em {data_autorizacao.strftime('%d/%m/%Y') if data_autorizacao else 'data não informada'} "
                f"{'SEI ' + sei_autorizacao if sei_autorizacao else ''}."
            )
            p.add_run(texto1)
            
            # Item 2
            p = doc.add_paragraph()
            p.add_run("2. Estudo Técnico Preliminar-ETP e Gestão de Riscos: ").bold = True
            texto2 = (
                f"O ETP nº {etp_numero if etp_numero else 'não informado'} "
                f"{'SEI ' + sei_etp if sei_etp else ''} e a Matriz de Riscos nº {risco_numero if risco_numero else 'não informada'} "
                f"{'SEI ' + sei_risco if sei_risco else ''} detalham a necessidade, viabilidade técnica e ações de mitigação para a contratação."
            )
            p.add_run(texto2)
            
            # Item 3
            p = doc.add_paragraph()
            p.add_run("3. Termo de Referência-TR: ").bold = True
            texto3 = (
                f"O TR nº {tr_numero if tr_numero else 'não informado'}, "
                f"{'SEI ' + sei_tr if sei_tr else ''}, consolida as especificações técnicas e condições contratuais, "
                f"servindo de balizador para a fase externa."
            )
            p.add_run(texto3)
            
            # Item 4
            p = doc.add_paragraph()
            p.add_run("4. Pesquisa de Mercado e Requisição SIGA: ").bold = True
            texto4 = (
                f"Foi realizada pesquisa de mercado formal, com a devida inclusão da Requisição de Material nº {req_siga if req_siga else 'não informada'} "
                f"no Sistema Integrado de Gestão de Aquisições (SIGA), totalizando o valor estimado de R$ {valor:,.2f}." if valor else
                f"Foi realizada pesquisa de mercado formal, com a devida inclusão da Requisição de Material nº {req_siga if req_siga else 'não informada'} "
                f"no Sistema Integrado de Gestão de Aquisições (SIGA)."
            )
            p.add_run(texto4)
            
            # Item 5
            p = doc.add_paragraph()
            p.add_run("5. Conformidade Orçamentária: ").bold = True
            texto5 = (
                f"O processo conta com as declarações de impacto financeiro {'SEI ' + sei_impacto if sei_impacto else ''}, "
                f"disponibilidade orçamentária {'SEI ' + sei_disponibilidade if sei_disponibilidade else ''} e a declaração do ordenador de despesa "
                f"{'SEI ' + sei_ordenador if sei_ordenador else ''}, atestando a compatibilidade com o orçamento e o Plano Plurianual (PPA/RJ) para 2026."
            )
            p.add_run(texto5)
            
            # Item 6
            p = doc.add_paragraph()
            p.add_run("6. Parecer Jurídico: ").bold = True
            texto6 = (
                f"A Diretoria Jurídica manifestou-se por meio do Despacho SEI {parecer_numero if parecer_numero else 'não informado'}, "
                f"informando a dispensa de análise jurídica formal em razão do valor da contratação, fundamentada na {fundamentacao}."
            )
            p.add_run(texto6)
            
            doc.add_paragraph()
            
            # ========================================
            # III. OBSERVAÇÕES
            # ========================================
            p = doc.add_paragraph()
            p.add_run("III. Observações").bold = True
            
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
            
            # ========================================
            # IV. DESPACHO
            # ========================================
            p = doc.add_paragraph()
            p.add_run("IV. Despacho").bold = True
            
            doc.add_paragraph(
                "Dessa forma, e considerando que os atos administrativos até o presente momento se mostram "
                "formalmente adequados e em conformidade com a Lei nº 14.133/2021 e demais normas aplicáveis, "
                "indicamos à continuidade do processo."
            )
            
            doc.add_paragraph()
            doc.add_paragraph()
            doc.add_paragraph()
            
            # Assinatura
            doc.add_paragraph("At.te.,")
            doc.add_paragraph()
            doc.add_paragraph("___________________________________")
            doc.add_paragraph("Auditor Interno")
            doc.add_paragraph("IPEm/RJ")
            
            # Salvar documento
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
                doc.save(tmp.name)
                tmp_path = tmp.name
            
            with open(tmp_path, 'rb') as f:
                doc_bytes = f.read()
            
            os.unlink(tmp_path)
            
            # Nome do arquivo
            data_atual = datetime.now().strftime("%Y%m%d_%H%M")
            nome_arquivo = f"DESPACHO_{processo_sei.replace('/', '_')}_{data_atual}.docx"
            
            # Mostrar sucesso
            st.success("✅ DESPACHO GERADO COM SUCESSO!")
            st.balloons()
            
            # Botão de download
            st.download_button(
                label="📥 CLIQUE AQUI PARA BAIXAR O DESPACHO",
                data=doc_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            # Mostrar prévia
            with st.expander("📄 Visualizar prévia do despacho"):
                st.write(f"**Processo:** {processo_sei}")
                st.write(f"**Objeto:** {objeto}")
                st.write(f"**Valor:** R$ {valor:,.2f}" if valor else "**Valor:** não informado")

# ============================================
# RODAPÉ
# ============================================

st.markdown("---")
st.markdown(f"""
<div style="text-align: center; color: #64748b; font-size: 0.9rem;">
    © 2026 - Auditoria Interna IPEm/RJ • Versão 4.0<br>
    Última atualização: {datetime.now().strftime('%d/%m/%Y %H:%M')}
</div>
""", unsafe_allow_html=True)
