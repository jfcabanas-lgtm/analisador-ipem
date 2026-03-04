# ============================================
# DESPACHO AUDIT - IPEM/RJ
# VERSÃO PREMIUM - COM DESPACHO DETALHADO
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os
import io
from PIL import Image

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="IPEM - Despacho Inteligente",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# CSS PERSONALIZADO
# ============================================

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif;
    }
    
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
        min-width: 250px;
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
        font-size: 1rem;
        margin: 0.5rem 0 0 0;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .header-seal .sei-link {
        display: inline-block;
        margin-top: 10px;
        padding: 5px 15px;
        background: #ffd700;
        color: #003366;
        text-decoration: none;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    
    .header-seal .sei-link:hover {
        background: #ffffff;
        transform: scale(1.05);
    }
    
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
    
    .footer-premium a {
        color: #ffd700;
        text-decoration: none;
        font-weight: 600;
        padding: 5px 20px;
        border: 1px solid #ffd700;
        border-radius: 20px;
        transition: all 0.3s ease;
        display: inline-block;
        margin: 1rem 0;
    }
    
    .footer-premium a:hover {
        background: #ffd700;
        color: #003366;
    }
    
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
# HEADER PRINCIPAL COM LINK DO SEI
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
            <a href="https://sei.rj.gov.br/sei/" target="_blank" class="sei-link">
                🔐 ACESSAR SEI
            </a>
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
# ESTATÍSTICAS
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
# FUNÇÃO PARA EXTRAIR DADOS
# ============================================

def extrair_campo(padroes, texto, default=""):
    for padrao in padroes:
        match = re.search(padrao, texto, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return default

if arquivo and st.session_state.dados_extraidos is None:
    
    with st.spinner("🔍 Analisando PDF e extraindo dados..."):
        
        texto = ""
        with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
            for pagina in pdf.pages:
                if pagina.extract_text():
                    texto += pagina.extract_text() + "\n"
        
        st.session_state.texto_extraido = texto
        
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
# FUNÇÃO PARA GERAR DESPACHO DETALHADO
# ============================================

def gerar_despacho_detalhado(processo_sei, objeto, data_autorizacao, 
                            valor_input, seis, 
                            etp_numero, sei_etp,
                            tr_numero, sei_tr,
                            risco_numero, sei_risco,
                            req_siga, parecer_numero,
                            sei_impacto, sei_disponibilidade, sei_ordenador,
                            fundamentacao, observacoes):
    
    doc = Document()
    
    # CONFIGURAÇÃO DE ESTILOS
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)
    
    # ========================================
    # CABEÇALHO INSTITUCIONAL
    # ========================================
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("ESTADO DO RIO DE JANEIRO")
    run.bold = True
    run.font.size = Pt(14)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("IPEM/RJ - INSTITUTO DE PESOS E MEDIDAS DO ESTADO DO RIO DE JANEIRO")
    run.bold = True
    run.font.size = Pt(13)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("AUDITORIA INTERNA")
    run.bold = True
    run.font.size = Pt(13)
    run.underline = True
    
    doc.add_paragraph()
    
    # ========================================
    # NÚMERO DO DESPACHO E DATA
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run(f"DESPACHO AUDIT Nº {datetime.now().strftime('%Y')}/")
    run.bold = True
    
    p = doc.add_paragraph()
    run = p.add_run(f"Data: {datetime.now().strftime('%d/%m/%Y')}")
    run.italic = True
    
    doc.add_paragraph()
    
    # ========================================
    # DESTINATÁRIO E ASSUNTO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("À Diretoria de Administração e Finanças - DIRAF")
    run.bold = True
    
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    run = p.add_run(f"Assunto: Análise de Instrução Processual - SEI nº {processo_sei}")
    run.bold = True
    
    doc.add_paragraph()
    doc.add_paragraph()
    
    # ========================================
    # 1. INTRODUÇÃO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("1. INTRODUÇÃO")
    run.bold = True
    run.font.size = Pt(13)
    
    doc.add_paragraph(
        f"Em atenção à solicitação dessa Diretoria, referente ao Processo SEI nº {processo_sei}, "
        f"que trata da {objeto}, esta Auditoria Interna procedeu à análise da instrução processual, "
        f"com fundamento na Lei Federal nº 14.133, de 1º de abril de 2021, e demais normativos aplicáveis."
    )
    
    doc.add_paragraph(
        "A presente manifestação limita-se à verificação da regularidade formal do processo, "
        "não adentrando no mérito técnico da contratação ou na discricionariedade da despesa."
    )
    
    doc.add_paragraph()
    
    # ========================================
    # 2. ANÁLISE DA INSTRUÇÃO PROCESSUAL
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("2. ANÁLISE DA INSTRUÇÃO PROCESSUAL")
    run.bold = True
    run.font.size = Pt(13)
    
    doc.add_paragraph(
        "Foram examinados os seguintes documentos e etapas processuais:"
    )
    
    doc.add_paragraph()
    
    # 2.1. PLANEJAMENTO
    p = doc.add_paragraph()
    run = p.add_run("2.1. Planejamento da Contratação")
    run.bold = True
    run.italic = True
    
    # ETP
    p = doc.add_paragraph()
    p.add_run("• Estudo Técnico Preliminar (ETP): ").bold = True
    if etp_numero:
        p.add_run(f"Presente - ETP nº {etp_numero} {f'(SEI {sei_etp})' if sei_etp else ''}. ")
        p.add_run("O documento aborda a necessidade da contratação, alternativas de mercado e viabilidade técnica.")
    else:
        p.add_run("NÃO LOCALIZADO. Documento obrigatório conforme Art. 18, I da Lei 14.133/2021.")
    
    # Gestão de Riscos
    p = doc.add_paragraph()
    p.add_run("• Gestão de Riscos: ").bold = True
    if risco_numero:
        p.add_run(f"Presente - Matriz de Riscos nº {risco_numero} {f'(SEI {sei_risco})' if sei_risco else ''}. ")
        p.add_run("Identificados os riscos e definidas medidas de mitigação.")
    else:
        p.add_run("NÃO LOCALIZADA. Recomenda-se a elaboração de matriz de riscos.")
    
    # TR
    p = doc.add_paragraph()
    p.add_run("• Termo de Referência (TR): ").bold = True
    if tr_numero:
        p.add_run(f"Presente - TR nº {tr_numero} {f'(SEI {sei_tr})' if sei_tr else ''}. ")
        p.add_run("Contém especificações técnicas detalhadas e condições contratuais.")
    else:
        p.add_run("NÃO LOCALIZADO. Documento essencial para a fase externa.")
    
    doc.add_paragraph()
    
    # 2.2. ECONOMICIDADE
    p = doc.add_paragraph()
    run = p.add_run("2.2. Economicidade")
    run.bold = True
    run.italic = True
    
    p = doc.add_paragraph()
    p.add_run("• Pesquisa de Preços: ").bold = True
    if valor_input:
        p.add_run(f"Realizada - Valor estimado: R$ {valor_input}. ")
        if req_siga:
            p.add_run(f"Requisição SIGA nº {req_siga}.")
        p.add_run("\n  Foram consultados os seguintes parâmetros:")
        p.add_run("\n  - Painel de Preços do Governo Federal")
        p.add_run("\n  - Contratações similares de outros órgãos")
        p.add_run("\n  - Cotação direta com fornecedores")
    else:
        p.add_run("NÃO LOCALIZADA. Necessário realizar pesquisa de preços para embasar o valor estimado.")
    
    doc.add_paragraph()
    
    # 2.3. LEGALIDADE
    p = doc.add_paragraph()
    run = p.add_run("2.3. Legalidade")
    run.bold = True
    run.italic = True
    
    p = doc.add_paragraph()
    p.add_run("• Parecer Jurídico: ").bold = True
    if parecer_numero:
        p.add_run(f"Presente - Despacho SEI nº {parecer_numero}. ")
        p.add_run(f"Fundamentação: {fundamentacao}.")
    else:
        p.add_run("NÃO LOCALIZADO. Necessário manifestação da Assessoria Jurídica, conforme Art. 53 da Lei 14.133/2021.")
    
    p = doc.add_paragraph()
    p.add_run("• Fundamentação Legal: ").bold = True
    p.add_run(f"Lei 14.133/2021, especialmente arts. 18, 23, 53 e 72.")
    
    doc.add_paragraph()
    
    # 2.4. CONTROLE E CONFORMIDADE
    p = doc.add_paragraph()
    run = p.add_run("2.4. Controle e Conformidade")
    run.bold = True
    run.italic = True
    
    p = doc.add_paragraph()
    p.add_run("• Disponibilidade Orçamentária: ").bold = True
    if sei_disponibilidade:
        p.add_run(f"Presente - SEI {sei_disponibilidade}.")
    else:
        p.add_run("NÃO LOCALIZADA.")
    
    p = doc.add_paragraph()
    p.add_run("• Declaração do Ordenador de Despesas: ").bold = True
    if sei_ordenador:
        p.add_run(f"Presente - SEI {sei_ordenador}.")
    else:
        p.add_run("NÃO LOCALIZADA.")
    
    p = doc.add_paragraph()
    p.add_run("• Impacto Financeiro: ").bold = True
    if sei_impacto:
        p.add_run(f"Presente - SEI {sei_impacto}.")
    else:
        p.add_run("NÃO LOCALIZADA.")
    
    doc.add_paragraph()
    
    # ========================================
    # 3. QUADRO RESUMO
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("3. QUADRO RESUMO DA INSTRUÇÃO")
    run.bold = True
    run.font.size = Pt(13)
    
    # Criar tabela
    tabela = doc.add_table(rows=1, cols=3)
    tabela.style = 'Table Grid'
    
    cabecalho = tabela.rows[0].cells
    cabecalho[0].text = "Documento/Etapa"
    cabecalho[1].text = "Status"
    cabecalho[2].text = "Observação"
    
    dados_tabela = [
        ["Solicitação Inicial", "✅ Presente" if seis else "❌ Ausente", f"SEI {seis[0] if len(seis) > 0 else 'não localizado'}"],
        ["Autorização", "✅ Presente" if data_autorizacao else "❌ Ausente", f"Data: {data_autorizacao.strftime('%d/%m/%Y') if data_autorizacao else 'não localizada'}"],
        ["ETP", "✅ Presente" if etp_numero else "❌ Ausente", f"Nº {etp_numero if etp_numero else '-'}"],
        ["TR", "✅ Presente" if tr_numero else "❌ Ausente", f"Nº {tr_numero if tr_numero else '-'}"],
        ["Matriz de Riscos", "✅ Presente" if risco_numero else "❌ Ausente", f"Nº {risco_numero if risco_numero else '-'}"],
        ["Pesquisa de Preços", "✅ Presente" if valor_input else "❌ Ausente", f"R$ {valor_input if valor_input else '-'}"],
        ["Parecer Jurídico", "✅ Presente" if parecer_numero else "❌ Ausente", f"SEI {parecer_numero if parecer_numero else '-'}"],
        ["Disponibilidade Orçamentária", "✅ Presente" if sei_disponibilidade else "❌ Ausente", f"SEI {sei_disponibilidade if sei_disponibilidade else '-'}"],
    ]
    
    for linha in dados_tabela:
        cells = tabela.add_row().cells
        cells[0].text = linha[0]
        cells[1].text = linha[1]
        cells[2].text = linha[2]
    
    doc.add_paragraph()
    
    # ========================================
    # 4. ANÁLISE DE RISCOS
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("4. ANÁLISE DE RISCOS")
    run.bold = True
    run.font.size = Pt(13)
    
    total_docs = len(dados_tabela)
    docs_presentes = sum(1 for linha in dados_tabela if "✅" in linha[1])
    percentual = (docs_presentes / total_docs) * 100
    
    if percentual >= 80:
        risco = "BAIXO"
        cor_risco = "Verde"
        desc_risco = "A maioria dos documentos obrigatórios está presente, indicando boa instrução processual."
    elif percentual >= 60:
        risco = "MÉDIO"
        cor_risco = "Amarelo"
        desc_risco = "Parte dos documentos está ausente, necessitando complementação."
    else:
        risco = "ALTO"
        cor_risco = "Vermelho"
        desc_risco = "Documentos essenciais ausentes, comprometendo a instrução processual."
    
    doc.add_paragraph(f"• Percentual de documentos presentes: {percentual:.1f}%")
    doc.add_paragraph(f"• Nível de risco identificado: {risco} ({cor_risco})")
    doc.add_paragraph(f"• {desc_risco}")
    
    doc.add_paragraph()
    
    # ========================================
    # 5. OBSERVAÇÕES E RECOMENDAÇÕES
    # ========================================
    
    p = doc.add_paragraph()
    run = p.add_run("5. OBSERVAÇÕES E RECOMENDAÇÕES")
    run.bold = True
    run.font.size = Pt(13)
    
    if observacoes:
        doc.add_paragraph(observacoes)
    else:
        doc.add_paragraph("Não foram identificadas irregularidades que impeçam o prosseguimento do feito.")
    
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    run = p.add_run
