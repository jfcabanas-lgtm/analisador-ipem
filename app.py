# ============================================
# DESPACHO AUDIT - IPEM/RJ
# VERSÃO PREMIUM - FRASE CORRIGIDA
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
from PIL import Image
import base64

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="IPEM - Sistema de Despacho Inteligente",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# CSS PERSONALIZADO - DESIGN PREMIUM
# ============================================

st.markdown("""
<style>
    /* FONTE PRINCIPAL */
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif;
    }
    
    /* HEADER PRINCIPAL - AZUL INSTITUCIONAL */
    .main-header {
        background: linear-gradient(135deg, #001529 0%, #003366 50%, #0047ab 100%);
        padding: 2.5rem 2rem;
        border-radius: 30px;
        margin-bottom: 2rem;
        box-shadow: 0 20px 40px rgba(0,20,50,0.3);
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.1);
    }
    
    .main-header::before {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none"><polygon points="0,0 100,0 50,100 0,100" fill="rgba(255,255,255,0.03)"/></svg>');
        background-size: 50px 50px;
        opacity: 0.1;
        pointer-events: none;
    }
    
    .header-content {
        display: flex;
        align-items: center;
        justify-content: space-between;
        position: relative;
        z-index: 2;
    }
    
    .header-title {
        flex: 1;
    }
    
    .header-title h1 {
        color: white;
        font-size: 3.5rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .header-title h3 {
        color: rgba(255,255,255,0.9);
        font-size: 1.5rem;
        font-weight: 300;
        margin: 0.5rem 0 0 0;
        letter-spacing: 1px;
    }
    
    .header-seal {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1.2rem 2rem;
        border-radius: 20px;
        text-align: center;
        border: 2px solid #ffd700;
        box-shadow: 0 10px 20px rgba(0,0,0,0.2);
        min-width: 200px;
    }
    
    .header-seal h2 {
        color: #003366;
        font-size: 1.8rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.2;
    }
    
    .header-seal p {
        color: #666;
        font-size: 0.9rem;
        margin: 0.3rem 0 0 0;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    /* SELO OFICIAL */
    .official-seal {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1rem;
        border-radius: 15px;
        text-align: center;
        border: 1px solid #dee2e6;
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.05);
    }
    
    .official-seal span {
        font-size: 2rem;
        display: block;
        margin-bottom: 0.5rem;
    }
    
    .official-seal p {
        color: #495057;
        font-size: 0.8rem;
        margin: 0;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    /* CARTÕES DE ESTATÍSTICAS */
    .stats-container {
        display: flex;
        gap: 1.5rem;
        margin: 2rem 0;
    }
    
    .stat-card {
        background: white;
        padding: 1.5rem;
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,51,102,0.1);
        flex: 1;
        text-align: center;
        border: 1px solid #e9ecef;
        transition: transform 0.3s ease;
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 40px rgba(0,51,102,0.15);
    }
    
    .stat-icon {
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
    }
    
    .stat-label {
        color: #6c757d;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    .stat-value {
        color: #003366;
        font-size: 2rem;
        font-weight: 700;
        margin: 0.5rem 0 0 0;
    }
    
    /* SEÇÕES PRINCIPAIS */
    .section-premium {
        background: white;
        padding: 2rem;
        border-radius: 20px;
        margin: 2rem 0;
        box-shadow: 0 10px 30px rgba(0,0,0,0.05);
        border-left: 5px solid #003366;
        position: relative;
    }
    
    .section-premium::before {
        content: "⚖️";
        position: absolute;
        top: -15px;
        left: 30px;
        background: #003366;
        color: white;
        font-size: 1.5rem;
        padding: 0.5rem 1rem;
        border-radius: 50px;
        box-shadow: 0 5px 15px rgba(0,51,102,0.3);
    }
    
    .section-title-premium {
        color: #003366;
        font-size: 1.8rem;
        font-weight: 600;
        margin-bottom: 1.5rem;
        padding-left: 3rem;
    }
    
    /* INFO CARDS */
    .info-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
        gap: 1.5rem;
        margin: 2rem 0;
    }
    
    .info-card {
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        padding: 1.5rem;
        border-radius: 15px;
        border: 1px solid #dee2e6;
    }
    
    .info-card h4 {
        color: #003366;
        font-size: 1.2rem;
        font-weight: 600;
        margin: 0 0 1rem 0;
        border-bottom: 2px solid #003366;
        padding-bottom: 0.5rem;
    }
    
    .info-card ul {
        list-style: none;
        padding: 0;
        margin: 0;
    }
    
    .info-card li {
        padding: 0.5rem 0;
        border-bottom: 1px solid #dee2e6;
        color: #495057;
    }
    
    .info-card li:last-child {
        border-bottom: none;
    }
    
    .info-card li::before {
        content: "•";
        color: #003366;
        font-weight: bold;
        margin-right: 0.5rem;
    }
    
    /* UPLOAD ÁREA */
    .upload-premium {
        border: 3px dashed #003366;
        border-radius: 30px;
        padding: 3rem;
        text-align: center;
        background: linear-gradient(135deg, rgba(0,51,102,0.02) 0%, rgba(0,71,171,0.02) 100%);
        transition: all 0.3s ease;
        margin: 2rem 0;
    }
    
    .upload-premium:hover {
        border-color: #0047ab;
        background: linear-gradient(135deg, rgba(0,51,102,0.05) 0%, rgba(0,71,171,0.05) 100%);
        transform: scale(1.02);
    }
    
    .upload-premium span {
        font-size: 4rem;
        display: block;
        margin-bottom: 1rem;
    }
    
    .upload-premium h3 {
        color: #003366;
        font-size: 1.8rem;
        font-weight: 600;
        margin: 0.5rem 0;
    }
    
    .upload-premium p {
        color: #6c757d;
        font-size: 1.1rem;
        max-width: 600px;
        margin: 1rem auto;
    }
    
    /* BOTÕES */
    .stButton > button {
        background: linear-gradient(135deg, #003366 0%, #0047ab 100%);
        color: white;
        font-weight: 600;
        padding: 1rem 2rem;
        border-radius: 15px;
        border: none;
        box-shadow: 0 10px 20px rgba(0,51,102,0.3);
        transition: all 0.3s ease;
        font-size: 1.2rem;
        letter-spacing: 0.5px;
        width: 100%;
        margin: 0.5rem 0;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #0047ab 0%, #003366 100%);
        box-shadow: 0 15px 30px rgba(0,51,102,0.4);
        transform: translateY(-2px);
    }
    
    /* RODAPÉ */
    .footer-premium {
        background: linear-gradient(135deg, #001529 0%, #003366 100%);
        padding: 2rem;
        border-radius: 30px 30px 0 0;
        margin-top: 3rem;
        color: white;
        text-align: center;
        position: relative;
        overflow: hidden;
    }
    
    .footer-premium::before {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 3px;
        background: linear-gradient(90deg, #ffd700, #ffffff, #ffd700);
    }
    
    .footer-premium p {
        margin: 0.5rem 0;
        font-size: 1rem;
        opacity: 0.9;
    }
    
    .footer-premium strong {
        color: #ffd700;
        font-weight: 600;
    }
    
    /* TIMELINE */
    .timeline {
        display: flex;
        justify-content: space-between;
        margin: 2rem 0;
        position: relative;
    }
    
    .timeline::before {
        content: "";
        position: absolute;
        top: 30px;
        left: 50px;
        right: 50px;
        height: 3px;
        background: linear-gradient(90deg, #003366, #0047ab, #003366);
        z-index: 1;
    }
    
    .timeline-step {
        text-align: center;
        position: relative;
        z-index: 2;
        flex: 1;
    }
    
    .timeline-icon {
        background: white;
        width: 60px;
        height: 60px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 1rem;
        border: 3px solid #003366;
        box-shadow: 0 5px 15px rgba(0,51,102,0.2);
        font-size: 1.8rem;
    }
    
    .timeline-label {
        background: #f8f9fa;
        padding: 0.5rem 1rem;
        border-radius: 30px;
        font-weight: 600;
        color: #003366;
        font-size: 0.9rem;
        display: inline-block;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    
    /* MENSAGENS DE SUCESSO */
    .success-premium {
        background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
        color: #155724;
        padding: 1.5rem;
        border-radius: 15px;
        border-left: 5px solid #28a745;
        margin: 1rem 0;
        font-size: 1.1rem;
        font-weight: 500;
        box-shadow: 0 5px 15px rgba(40,167,69,0.2);
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# HEADER PRINCIPAL - FRASE CORRIGIDA
# ============================================

st.markdown("""
<div class="main-header">
    <div class="header-content">
        <div class="header-title">
            <h1>IPEM/RJ</h1>
            <h3>INSTITUTO DE PESOS E MEDIDAS DO ESTADO DO RIO DE JANEIRO</h3>
        </div>
        <div class="header-seal">
            <h2>AUDITORIA<br>INTERNA</h2>
            <p>ANÁLISE DE PROCESSO DE LICITAÇÃO E DISPENSA</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================
# MENSAGEM DE BOAS-VINDAS
# ============================================

st.markdown("""
<div class="section-premium">
    <div class="section-title-premium">Bem-vindo ao Sistema de Despacho Inteligente</div>
    <p style="font-size: 1.2rem; color: #495057; line-height: 1.6;">
        Este sistema foi desenvolvido para automatizar a geração de despachos de auditoria,
        garantindo padronização, agilidade e conformidade com a Lei nº 14.133/2021.
        Através da análise inteligente de documentos, você poderá gerar despachos
        completos em poucos minutos.
    </p>
</div>
""", unsafe_allow_html=True)

# ============================================
# TIMELINE DO PROCESSO
# ============================================

st.markdown("""
<div class="timeline">
    <div class="timeline-step">
        <div class="timeline-icon">📂</div>
        <span class="timeline-label">Upload do PDF</span>
    </div>
    <div class="timeline-step">
        <div class="timeline-icon">🔍</div>
        <span class="timeline-label">Análise Automática</span>
    </div>
    <div class="timeline-step">
        <div class="timeline-icon">✏️</div>
        <span class="timeline-label">Confirmação de Dados</span>
    </div>
    <div class="timeline-step">
        <div class="timeline-icon">📥</div>
        <span class="timeline-label">Download do Despacho</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================
# INICIALIZAÇÃO DO SESSION STATE
# ============================================

if 'dados_extraidos' not in st.session_state:
    st.session_state.dados_extraidos = None
if 'texto_extraido' not in st.session_state:
    st.session_state.texto_extraido = None
if 'seis_encontrados' not in st.session_state:
    st.session_state.seis_encontrados = []
if 'doc_bytes' not in st.session_state:
    st.session_state.doc_bytes = None
if 'nome_arquivo' not in st.session_state:
    st.session_state.nome_arquivo = None
if 'processos_analisados' not in st.session_state:
    st.session_state.processos_analisados = 0

# ============================================
# ESTATÍSTICAS (somente se não houver análise em andamento)
# ============================================

if not st.session_state.dados_extraidos and not st.session_state.doc_bytes:
    
    st.markdown(f"""
    <div class="stats-container">
        <div class="stat-card">
            <div class="stat-icon">⚡</div>
            <div class="stat-label">Processos Analisados</div>
            <div class="stat-value">{st.session_state.processos_analisados}</div>
        </div>
        <div class="stat-card">
            <div class="stat-icon">📊</div>
            <div class="stat-label">Despachos Gerados</div>
            <div class="stat-value">{st.session_state.processos_analisados}</div>
        </div>
        <div class="stat-card">
            <div class="stat-icon">⏱️</div>
            <div class="stat-label">Tempo Médio</div>
            <div class="stat-value">2 min</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

# ============================================
# PASSO 1: UPLOAD DO PDF
# ============================================

st.markdown("""
<div class="upload-premium">
    <span>📎</span>
    <h3>Upload do Processo</h3>
    <p>Arraste o arquivo PDF do processo SEI ou clique para selecionar</p>
</div>
""", unsafe_allow_html=True)

arquivo = st.file_uploader("", type=['pdf'], label_visibility="collapsed")

# ============================================
# FUNÇÃO PARA EXTRAIR DADOS (MANTIDA IGUAL)
# ============================================

if arquivo and st.session_state.dados_extraidos is None:
    
    with st.spinner("🔍 Analisando PDF e extraindo dados..."):
        
        texto = ""
        with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
            for pagina in pdf.pages:
                if pagina.extract_text():
                    texto += pagina.extract_text() + "\n"
        
        st.session_state.texto_extraido = texto
        
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
            ], texto, ""),
            
            'objeto': extrair_campo([
                r'objeto[:\s]*([^.]+)',
                r'aquisição[:\s]*([^.]+)',
                r'contratação[:\s]*([^.]+)'
            ], texto, ""),
            
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
        
        seis_encontrados = re.findall(r'SEI[:\s]*n[º°]?\s*(\d+)', texto, re.IGNORECASE)
        
        st.session_state.dados_extraidos = dados_extraidos
        st.session_state.seis_encontrados = seis_encontrados
        st.session_state.processos_analisados += 1
        
        st.rerun()

# ============================================
# SE HOUVER DADOS EXTRAÍDOS, MOSTRA O FORMULÁRIO
# ============================================

if st.session_state.dados_extraidos:
    
    dados = st.session_state.dados_extraidos
    seis = st.session_state.seis_encontrados
    
    st.markdown("""
    <div class="section-premium">
        <div class="section-title-premium">📊 Dados Encontrados no PDF</div>
    </div>
    """, unsafe_allow_html=True)
    
    col_res1, col_res2, col_res3 = st.columns(3)
    
    with col_res1:
        st.info("📋 **Processo**")
        st.write(f"**Nº:** {dados['processo_sei'] or 'Não encontrado'}")
        st.write(f"**Data:** {dados['data_autorizacao'] or 'Não encontrada'}")
    
    with col_res2:
        st.info("💰 **Valor**")
        if dados['valor']:
            try:
                valor_float = float(dados['valor'].replace('.', '').replace(',', '.'))
                st.write(f"**R$ {valor_float:,.2f}**")
            except:
                st.write(f"**R$ {dados['valor']}**")
        else:
            st.write("**Não identificado**")
    
    with col_res3:
        st.info("📑 **Documentos**")
        st.write(f"ETP: {dados['etp_numero'] or '❌'}")
        st.write(f"TR: {dados['tr_numero'] or '❌'}")
        st.write(f"Parecer: {dados['parecer_numero'] or '❌'}")
    
    if seis:
        with st.expander(f"📎 {len(seis)} números SEI encontrados"):
            for i, sei in enumerate(seis[:10]):
                st.write(f"• SEI {sei}")
    
    # FORMULÁRIO
    st.markdown("""
    <div class="section-premium">
        <div class="section-title-premium">✏️ Confirmação de Dados</div>
    </div>
    """, unsafe_allow_html=True)
    
    with st.form("form_confirmacao"):
        
        col1, col2 = st.columns(2)
        
        with col1:
            processo_sei = st.text_input("Nº do Processo SEI *", value=dados['processo_sei'])
            objeto = st.text_input("Objeto da Contratação *", value=dados['objeto'])
            
            try:
                data_valor = datetime.strptime(dados['data_autorizacao'], '%d/%m/%Y') if dados['data_autorizacao'] and '/' in dados['data_autorizacao'] else None
            except:
                data_valor = None
            
            data_autorizacao = st.date_input("Data da Autorização", value=data_valor, format="DD/MM/YYYY")
        
        with col2:
            valor_input = st.text_input("Valor (R$)", value=dados['valor'])
            sei_inicio = st.text_input("SEI da Solicitação Inicial", value=seis[0] if len(seis) > 0 else "")
            sei_autorizacao = st.text_input("SEI da Autorização", value=seis[1] if len(seis) > 1 else "")
        
        col3, col4 = st.columns(2)
        
        with col3:
            etp_numero = st.text_input("Nº do ETP", value=dados['etp_numero'])
            sei_etp = st.text_input("SEI do ETP", value=seis[2] if len(seis) > 2 else "")
            tr_numero = st.text_input("Nº do TR", value=dados['tr_numero'])
            sei_tr = st.text_input("SEI do TR", value=seis[3] if len(seis) > 3 else "")
        
        with col4:
            risco_numero = st.text_input("Nº da Matriz de Riscos", value=dados['risco_numero'])
            sei_risco = st.text_input("SEI da Matriz de Riscos", value=seis[4] if len(seis) > 4 else "")
            req_siga = st.text_input("Nº da Requisição SIGA", value=dados['req_siga'])
            parecer_numero = st.text_input("Nº do Parecer Jurídico", value=dados['parecer_numero'])
        
        col5, col6 = st.columns(2)
        
        with col5:
            sei_impacto = st.text_input("SEI da Declaração de Impacto", value=seis[5] if len(seis) > 5 else "")
            sei_disponibilidade = st.text_input("SEI da Disponibilidade Orçamentária", value=seis[6] if len(seis) > 6 else "")
            sei_ordenador = st.text_input("SEI da Declaração do Ordenador", value=seis[7] if len(seis) > 7 else "")
        
        with col6:
            fundamentacao = st.text_area("Fundamentação Legal", 
                value="art. 1º da Resolução PGE nº 5.059/2024 e no art. 95, I, da Lei nº 14.133/2021",
                height=100)
        
        observacoes = st.text_area("Observações", height=100)
        
        submitted = st.form_submit_button("✅ CONFIRMAR DADOS E GERAR DESPACHO", use_container_width=True)
    
    # GERAÇÃO DO DESPACHO
    if submitted:
        
        if not processo_sei or not objeto:
            st.error("❌ Processo e Objeto são obrigatórios!")
        else:
            
            with st.spinner("Gerando despacho..."):
                
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
                    f"procedemos à verificação dos documentos apresentados."
                )
                doc.add_paragraph()
                
                # II. ANÁLISE PROCESSUAL
                doc.add_paragraph().add_run("II. Análise Processual").bold = True
                doc.add_paragraph("Foram examinados os seguintes documentos e etapas processuais:")
                doc.add_paragraph()
                
                # Item 1
                p = doc.add_paragraph()
                p.add_run("1. Solicitação Inicial e Autorização: ").bold = True
                data_aut_str = data_autorizacao.strftime('%d/%m/%Y') if data_autorizacao else 'data não informada'
                p.add_run(
                    f"O processo foi iniciado pela Superintendência de Pré-Medidos "
                    f"{'SEI ' + sei_inicio if sei_inicio else ''} e autorizado pela Presidência "
                    f"em {data_aut_str} {'SEI ' + sei_autorizacao if sei_autorizacao else ''}."
                )
                
                # Item 2
                p = doc.add_paragraph()
                p.add_run("2. Estudo Técnico Preliminar-ETP e Gestão de Riscos: ").bold = True
                p.add_run(
                    f"O ETP nº {etp_numero if etp_numero else 'não informado'} "
                    f"{'SEI ' + sei_etp if sei_etp else ''} e a Matriz de Riscos nº {risco_numero if risco_numero else 'não informada'} "
                    f"{'SEI ' + sei_risco if sei_risco else ''} detalham a necessidade e viabilidade."
                )
                
                # Item 3
                p = doc.add_paragraph()
                p.add_run("3. Termo de Referência-TR: ").bold = True
                p.add_run(
                    f"O TR nº {tr_numero if tr_numero else 'não informado'} "
                    f"{'SEI ' + sei_tr if sei_tr else ''} consolida as especificações técnicas."
                )
                
                # Item 4
                p = doc.add_paragraph()
                p.add_run("4. Pesquisa de Mercado e Requisição SIGA: ").bold = True
                texto4 = f"Requisição nº {req_siga if req_siga else 'não informada'} no SIGA"
                if valor_input:
                    texto4 += f", valor estimado de R$ {valor_input}"
                p.add_run(texto4 + ".")
                
                # Item 5
                p = doc.add_paragraph()
                p.add_run("5. Conformidade Orçamentária: ").bold = True
                p.add_run(
                    f"Declarações SEI {sei_impacto if sei_impacto else ''}, "
                    f"{sei_disponibilidade if sei_disponibilidade else ''} e "
                    f"{sei_ordenador if sei_ordenador else ''} atestam compatibilidade orçamentária."
                )
                
                # Item 6
                p = doc.add_paragraph()
                p.add_run("6. Parecer Jurídico: ").bold = True
                p.add_run(
                    f"Despacho SEI {parecer_numero if parecer_numero else 'não informado'}, "
                    f"fundamentado em {fundamentacao}."
                )
                
                doc.add_paragraph()
                
                # III. OBSERVAÇÕES
                doc.add_paragraph().add_run("III. Observações").bold = True
                if observacoes:
                    doc.add_paragraph(observacoes)
                else:
                    doc.add_paragraph(
                        "Verifica-se que o processo encontra-se devidamente instruído."
                    )
                
                doc.add_paragraph()
                
                # IV. DESPACHO
                doc.add_paragraph().add_run("IV. Despacho").bold = True
                doc.add_paragraph(
                    "Dessa forma, considerando a regularidade formal, indicamos a continuidade do processo."
                )
                
                doc.add_paragraph()
                doc.add_paragraph()
                doc.add_paragraph()
                doc.add_paragraph("At.te.,")
                doc.add_paragraph()
                doc.add_paragraph("___________________________________")
                doc.add_paragraph("Auditor Interno")
                doc.add_paragraph("IPEM/RJ")
                
                doc_bytes = io.BytesIO()
                doc.save(doc_bytes)
                doc_bytes.seek(0)
                
                st.session_state.doc_bytes = doc_bytes.getvalue()
                st.session_state.nome_arquivo = f"DESPACHO_{processo_sei.replace('/', '_')}.docx"
                
                st.rerun()

# ============================================
# DOWNLOAD DO DESPACHO
# ============================================

if st.session_state.doc_bytes:
    
    st.markdown("""
    <div class="section-premium">
        <div class="section-title-premium">✅ Despacho Gerado com Sucesso!</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.balloons()
    
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    
    with col_btn2:
        st.download_button(
            label="📥 BAIXAR DESPACHO",
            data=st.session_state.doc_bytes,
            file_name=st.session_state.nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        if st.button("🔄 NOVO PROCESSO", use_container_width=True):
            for key in ['dados_extraidos', 'texto_extraido', 'seis_encontrados', 'doc_bytes', 'nome_arquivo']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

# ============================================
# RODAPÉ INSTITUCIONAL
# ============================================

st.markdown("""
<div class="footer-premium">
    <p><strong>IPEM/RJ - INSTITUTO DE PESOS E MEDIDAS DO ESTADO DO RIO DE JANEIRO</strong></p>
    <p>AUDITORIA INTERNA - ANÁLISE DE PROCESSO DE LICITAÇÃO E DISPENSA</p>
    <p style="font-size: 0.9rem; opacity: 0.8;">Lei nº 14.133/2021 • Versão 7.2 Premium</p>
    <p style="font-size: 0.8rem; opacity: 0.6;">© 2026 - Todos os direitos reservados</p>
</div>
""", unsafe_allow_html=True)
