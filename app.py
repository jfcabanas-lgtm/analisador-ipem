# ============================================
# ANALISADOR DE INSTRUÇÃO PROCESSUAL - LEI 14.133/2021
# PARA AUDITORIA INTERNA DO IPEm/RJ
# ============================================

import streamlit as st
import pdfplumber
import re
from datetime import datetime
import pandas as pd
import io
import plotly.graph_objects as go

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(
    page_title="Analisador IPEm - Instrução Processual",
    page_icon="📋",
    layout="wide"
)

# TÍTULO
st.title("📋 Analisador de Instrução Processual - IPEm/RJ")
st.markdown("**Lei 14.133/2021 - Verificação para prosseguimento**")
st.markdown("---")

# SIDEBAR
with st.sidebar:
    st.header("ℹ️ Sobre")
    st.info("""
    **Finalidade:** Verificar se o processo está devidamente instruído 
    para prosseguimento, conforme arts. 11, 18, 53 e 72 da Lei 14.133/2021.
    
    **Documentos obrigatórios:**
    - Estudo Técnico Preliminar
    - Termo de Referência/Projeto Básico
    - Pesquisa de Preços
    - Parecer Jurídico
    - Publicação/Transparência
    - Justificativa da Contratação
    - Designação de Fiscal
    """)
    
    st.header("📊 Estatísticas")
    if 'processos_analisados' not in st.session_state:
        st.session_state.processos_analisados = 0
    st.metric("Processos analisados", st.session_state.processos_analisados)

# UPLOAD
st.header("📂 Processo SEI")
arquivo = st.file_uploader(
    "Selecione o PDF do processo (portaria, despacho, parecer)",
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
    Baseado nos arts. 11, 18, 53, 72 da Lei 14.133/2021
    """
    
    # Documentos a verificar
    documentos = {
        # Planejamento (Art. 18)
        '📄 Estudo Técnico Preliminar (Art. 18, I)': {
            'padrao': r'estudo preliminar|etp|estudo técnico preliminar',
            'fundamento': 'Art. 18, I - obrigatório para toda contratação',
            'peso': 3  # Essencial
        },
        '📄 Termo de Referência/Projeto Básico (Art. 18, II)': {
            'padrao': r'termo de referência|projeto básico|especificações',
            'fundamento': 'Art. 18, II - define o objeto e condições',
            'peso': 3  # Essencial
        },
        '📄 Pesquisa de Preços (Art. 23)': {
            'padrao': r'pesquisa de preços|mapa de preços|orçamento estimado',
            'fundamento': 'Art. 23 - estimativa de preços',
            'peso': 3  # Essencial
        },
        '📄 Justificativa da Contratação (Art. 18, III)': {
            'padrao': r'justificativa|motivação da contratação|razão da escolha',
            'fundamento': 'Art. 18, III - demonstra necessidade',
            'peso': 2  # Importante
        },
        
        # Jurídico (Art. 53)
        '⚖️ Parecer Jurídico (Art. 53)': {
            'padrao': r'parecer jurídico|assessoria jurídica|manifestação jurídica|procuradoria',
            'fundamento': 'Art. 53 - obrigatório antes da ratificação',
            'peso': 3  # Essencial
        },
        
        # Publicidade (Art. 54)
        '📢 Publicação Oficial (Art. 54)': {
            'padrao': r'publicação|diário oficial|d\.o\.|DOU|DOERJ|publicado no',
            'fundamento': 'Art. 54 - transparência',
            'peso': 2  # Importante
        },
        
        # Controle (Art. 7º, 117)
        '👤 Designação de Fiscal (Art. 117)': {
            'padrao': r'fiscal do contrato|gestor do contrato|comissão de fiscalização|designa|responsável pelo acompanhamento',
            'fundamento': 'Art. 117 - obrigatório nomear fiscal',
            'peso': 2  # Importante
        },
        
        # Integridade (Art. 25)
        '🛡️ Programa de Integridade (quando exigível)': {
            'padrao': r'programa de integridade|compliance|código de ética',
            'fundamento': 'Art. 25, §4º - para obras de grande vulto',
            'peso': 1  # Complementar
        }
    }
    
    resultados = []
    pontuacao = 0
    pontuacao_maxima = 0
    
    for doc, info in documentos.items():
        if re.search(info['padrao'], texto, re.IGNORECASE):
            status = "✅ PRESENTE"
            pontuacao += info['peso']
            cor = "green"
        else:
            status = "❌ AUSENTE"
            cor = "red"
        
        pontuacao_maxima += info['peso']
        
        resultados.append({
            'documento': doc,
            'status': status,
            'fundamento': info['fundamento'],
            'cor': cor,
            'peso': info['peso']
        })
    
    percentual = (pontuacao / pontuacao_maxima * 100) if pontuacao_maxima > 0 else 0
    
    return resultados, percentual, pontuacao, pontuacao_maxima

def verificar_modalidade(texto):
    """Verifica qual modalidade está sendo usada"""
    modalidades = {
        'Dispensa (Art. 75)': r'dispensa|inexigibilidade|art\. 75',
        'Pregão (Art. 6º, XLI)': r'pregão|pregao',
        'Concorrência (Art. 6º, XXXVIII)': r'concorrência|concorrencia',
        'Tomada de Preços': r'tomada de preços',
        'Leilão (Art. 6º, XL)': r'leilão|leilao',
        'Diálogo Competitivo (Art. 6º, XLII)': r'diálogo competitivo|dialogo competitivo'
    }
    
    for modalidade, padrao in modalidades.items():
        if re.search(padrao, texto, re.IGNORECASE):
            return modalidade
    
    return "Não identificada"

def verificar_fundamentacao_legal(texto):
    """Verifica se o processo menciona a lei correta"""
    
    if re.search(r'14\.133|lei 14\.133|nova lei de licitações', texto, re.IGNORECASE):
        return "✅ Correta (Lei 14.133/2021)"
    elif re.search(r'8\.666|lei 8\.666|lei antiga', texto, re.IGNORECASE):
        return "⚠️ ATENÇÃO: Lei 8.666/93 (revogada)"
    else:
        return "❌ Não identificada"

def extrair_dados_basicos(texto):
    """Extrai informações básicas do processo"""
    dados = {}
    
    # Número do processo
    match = re.search(r'Processo[:\s]*n[º°]?\s*([\d\-/]+)', texto, re.IGNORECASE)
    dados['processo'] = match.group(1).strip() if match else "Não identificado"
    
    # Valor
    match = re.search(r'Valor\s*R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    dados['valor'] = match.group(1) if match else "Não identificado"
    
    # Data
    match = re.search(r'(\d{1,2})\s*de\s*([A-ZÇa-zç]+)\s*de\s*(\d{4})', texto)
    if match:
        meses = {
            'janeiro':'01', 'fevereiro':'02', 'março':'03', 'abril':'04',
            'maio':'05', 'junho':'06', 'julho':'07', 'agosto':'08',
            'setembro':'09', 'outubro':'10', 'novembro':'11', 'dezembro':'12'
        }
        mes = meses.get(match.group(2).lower(), '00')
        dados['data'] = f"{match.group(1)}/{mes}/{match.group(3)}"
    else:
        dados['data'] = "Não identificado"
    
    return dados

def recomendar_prosseguimento(pontuacao, pontuacao_maxima, documentos_essenciais):
    """
    Recomenda se o processo pode prosseguir
    Baseado na presença de documentos essenciais
    """
    
    # Verificar se todos os documentos essenciais (peso 3) estão presentes
    essenciais_presentes = all([
        doc['status'] == "✅ PRESENTE" 
        for doc in documentos_essenciais 
        if doc['peso'] == 3
    ])
    
    percentual = (pontuacao / pontuacao_maxima * 100) if pontuacao_maxima > 0 else 0
    
    if essenciais_presentes and percentual >= 80:
        return {
            'decisao': "✅ RECOMENDA-SE PROSSEGUIMENTO",
            'cor': 'green',
            'justificativa': "Documentos essenciais presentes e alto percentual de conformidade"
        }
    elif essenciais_presentes:
        return {
            'decisao': "⚠️ PROSSEGUIMENTO COM RESSALVAS",
            'cor': 'orange',
            'justificativa': "Documentos essenciais presentes, mas recomenda-se complementar documentação"
        }
    else:
        return {
            'decisao': "❌ AGUARDAR COMPLEMENTAÇÃO",
            'cor': 'red',
            'justificativa': "Documentos essenciais ausentes - não recomendado prosseguir"
        }

# ============================================
# PROCESSAMENTO PRINCIPAL
# ============================================

if arquivo is not None:
    
    with st.spinner("🔍 Analisando instrução processual..."):
        
        # Extrair texto
        texto = extrair_texto_pdf(arquivo)
        
        # Extrair dados básicos
        dados = extrair_dados_basicos(texto)
        
        # Verificar documentos obrigatórios
        resultados, percentual, pontuacao, pontuacao_maxima = verificar_documentos_obrigatorios(texto)
        
        # Verificar modalidade
        modalidade = verificar_modalidade(texto)
        
        # Verificar fundamentação legal
        fundamentacao = verificar_fundamentacao_legal(texto)
        
        # Recomendar prosseguimento
        recomendacao = recomendar_prosseguimento(pontuacao, pontuacao_maxima, resultados)
        
        # Incrementar contador
        st.session_state.processos_analisados += 1
        
        # ============================================
        # EXIBIR RESULTADOS
        # ============================================
        
        # CABEÇALHO COM DECISÃO
        st.markdown(f"""
        <div style='background-color: {recomendacao['cor']}; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;'>
            <h2 style='color: white; margin: 0;'>{recomendacao['decisao']}</h2>
            <p style='color: white; margin: 5px 0 0 0;'>{recomendacao['justificativa']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        # MÉTRICAS PRINCIPAIS
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Nº Processo", dados['processo'], help="Número do processo SEI")
        
        with col2:
            st.metric("Data", dados['data'], help="Data do documento")
        
        with col3:
            st.metric("Modalidade", modalidade, help="Modalidade de licitação identificada")
        
        with col4:
            st.metric("Fundamentação", "✅ OK" if "Correta" in fundamentacao else "❌", 
                     help=fundamentacao)
        
        # BARRAS DE PROGRESSO
        st.subheader("📊 Grau de Instrução Processual")
        
        col_prog1, col_prog2 = st.columns(2)
        
        with col_prog1:
            st.metric("Pontuação", f"{pontuacao}/{pontuacao_maxima}")
            st.progress(percentual/100)
        
        with col_prog2:
            if percentual >= 80:
                st.success(f"**{percentual:.1f}%** - Processo bem instruído")
            elif percentual >= 60:
                st.warning(f"**{percentual:.1f}%** - Instrução parcial")
            else:
                st.error(f"**{percentual:.1f}%** - Instrução insuficiente")
        
        # TABELA DE DOCUMENTOS
        st.subheader("📋 Documentos Obrigatórios")
        
        df = pd.DataFrame(resultados)
        
        # Colorir linhas
        def colorir_linhas(row):
            if row['status'] == "✅ PRESENTE":
                return ['background-color: #d4edda'] * len(row)
            else:
                return ['background-color: #f8d7da'] * len(row)
        
        styled_df = df.style.apply(colorir_linhas, axis=1)
        st.dataframe(styled_df, use_container_width=True)
        
        # DOCUMENTOS ESSENCIAIS FALTANTES
        faltantes_essenciais = [
            f"• {doc['documento']}" 
            for doc in resultados 
            if doc['status'] == "❌ AUSENTE" and doc['peso'] == 3
        ]
        
        if faltantes_essenciais:
            st.error("🚨 **DOCUMENTOS ESSENCIAIS FALTANTES (Art. 18 e 53):**")
            for falta in faltantes_essenciais:
                st.write(falta)
        
        # RECOMENDAÇÕES
        st.subheader("📌 Recomendações para Instrução")
        
        if percentual == 100:
            st.success("✅ Processo totalmente instruído. Pode prosseguir.")
        else:
            st.warning("**Documentos a incluir antes do prosseguimento:**")
            for doc in resultados:
                if doc['status'] == "❌ AUSENTE":
                    st.write(f"• {doc['documento']} - *{doc['fundamento']}*")
        
        # BOTÃO PARA RELATÓRIO
        if st.button("📥 Gerar Relatório de Instrução"):
            
            relatorio = f"""
=============================================================================
        INSTITUTO DE PESOS E MEDIDAS DO ESTADO DO RIO DE JANEIRO
               AUDITORIA INTERNA - PARECER DE INSTRUÇÃO PROCESSUAL
=============================================================================

PROCESSO: {dados['processo']}
DATA DA ANÁLISE: {datetime.now().strftime('%d/%m/%Y %H:%M')}
MODALIDADE: {modalidade}
FUNDAMENTAÇÃO: {fundamentacao}

=============================================================================
DECISÃO RECOMENDADA
=============================================================================
{recomendacao['decisao']} - {recomendacao['justificativa']}

=============================================================================
GRAU DE INSTRUÇÃO
=============================================================================
Pontuação obtida: {pontuacao}/{pontuacao_maxima}
Percentual de instrução: {percentual:.1f}%

=============================================================================
DOCUMENTOS VERIFICADOS
=============================================================================
"""
            for doc in resultados:
                relatorio += f"{doc['status']} - {doc['documento']}\n"
            
            if faltantes_essenciais:
                relatorio += """
=============================================================================
DOCUMENTOS ESSENCIAIS FALTANTES
=============================================================================
"""
                for falta in faltantes_essenciais:
                    relatorio += f"{falta}\n"
            
            relatorio += """
=============================================================================
FUNDAMENTAÇÃO LEGAL
=============================================================================
Arts. 11, 18, 53 e 72 da Lei Federal nº 14.133, de 1º de abril de 2021
- Art. 11: Objetivos da licitação
- Art. 18: Fase preparatória e documentos obrigatórios
- Art. 53: Obrigatoriedade de parecer jurídico
- Art. 72: Instrução processual

=============================================================================
                        FIM DO PARECER
=============================================================================
"""
            
            st.download_button(
                label="📥 Baixar Parecer (TXT)",
                data=relatorio,
                file_name=f"parecer_{dados['processo'].replace('/', '_')}.txt",
                mime="text/plain"
            )

else:
    # TELA INICIAL
    st.info("👆 Faça upload do PDF do processo para iniciar a análise")
    
    with st.expander("📌 O que este analisador verifica?"):
        st.write("""
        **Documentos obrigatórios (Lei 14.133/2021):**
        
        1. **Estudo Técnico Preliminar (Art. 18, I)** - Viabilidade da contratação
        2. **Termo de Referência/Projeto Básico (Art. 18, II)** - Objeto e condições
        3. **Pesquisa de Preços (Art. 23)** - Estimativa de custos
        4. **Parecer Jurídico (Art. 53)** - Análise de legalidade
        5. **Justificativa da Contratação (Art. 18, III)** - Motivação
        6. **Publicação (Art. 54)** - Transparência
        7. **Designação de Fiscal (Art. 117)** - Controle da execução
        
        **Resultado:** Recomendação de prosseguimento ou não
        """)

# RODAPÉ
st.markdown("---")
st.caption(f"© 2026 - Auditoria Interna IPEm/RJ - Versão 2.0 - Análise atualizada em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
