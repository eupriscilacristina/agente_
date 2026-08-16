import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from docx import Document

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="AG-MTE | Agente Master de Tecnologia e Processos",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO VISUAL CORPORATIVA AVANÇADA ---
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
    /* Reset e Base */
    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }
    
    /* Título Principal */
    .main-header {
        font-size: 2.5rem;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 0;
        letter-spacing: -0.02em;
        line-height: 1.2;
    }
    
    .sub-header {
        font-size: 1.15rem;
        color: #64748B;
        margin-bottom: 25px;
        font-weight: 400;
        letter-spacing: 0.01em;
    }
    
    /* Cards de Métricas */
    .metric-card {
        background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFC 100%);
        padding: 24px;
        border-radius: 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        transition: all 0.3s ease;
    }
    
    .metric-card:hover {
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08), 0 8px 10px -6px rgba(0, 0, 0, 0.04);
        transform: translateY(-2px);
    }
    
    .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0;
        line-height: 1;
    }
    
    .metric-label {
        font-size: 0.85rem;
        color: #64748B;
        margin-top: 8px;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Card de Resultado */
    .result-card {
        background: linear-gradient(135deg, #F0F9FF 0%, #E0F2FE 100%);
        border-left: 5px solid #0EA5E9;
        padding: 28px;
        border-radius: 12px;
        margin-top: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .result-card h3 {
        color: #0C4A6E;
        font-weight: 700;
        margin-bottom: 16px;
        font-size: 1.3rem;
    }
    
    .result-item {
        display: flex;
        align-items: flex-start;
        margin-bottom: 12px;
        padding: 12px;
        background: rgba(255, 255, 255, 0.7);
        border-radius: 8px;
    }
    
    .result-icon {
        font-size: 1.4rem;
        margin-right: 12px;
        flex-shrink: 0;
    }
    
    .result-content {
        flex: 1;
    }
    
    .result-label {
        font-weight: 600;
        color: #0F172A;
        font-size: 0.9rem;
    }
    
    .result-value {
        color: #475569;
        font-size: 0.95rem;
        margin-top: 4px;
    }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0F172A 0%, #1E293B 100%);
    }
    
    section[data-testid="stSidebar"] .stRadio label {
        color: #F1F5F9 !important;
        font-weight: 500;
    }
    
    section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        color: #94A3B8;
    }
    
    /* Botões */
    .stButton > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 28px;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.3);
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
        box-shadow: 0 6px 12px -1px rgba(37, 99, 235, 0.4);
        transform: translateY(-1px);
    }
    
    /* Tabelas */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    /* Formulários */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div {
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        padding: 12px 16px;
        font-size: 1rem;
    }
    
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #2563EB;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
    }
    
    /* Alertas */
    .stSuccess {
        background: linear-gradient(135deg, #F0FDF4 0%, #DCFCE7 100%);
        border: 1px solid #86EFAC;
        border-radius: 12px;
    }
    
    .stWarning {
        background: linear-gradient(135deg, #FFFBEB 0%, #FEF3C7 100%);
        border: 1px solid #FCD34D;
        border-radius: 12px;
    }
    
    .stError {
        background: linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%);
        border: 1px solid #FCA5A5;
        border-radius: 12px;
    }
    
    .stInfo {
        background: linear-gradient(135deg, #F0F9FF 0%, #E0F2FE 100%);
        border: 1px solid #7DD3FC;
        border-radius: 12px;
    }
    
    /* Dividers */
    hr {
        border: none;
        border-top: 1px solid #E2E8F0;
        margin: 24px 0;
    }
    
    /* Headers de Seção */
    .section-header {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 8px;
    }
    
    .section-subheader {
        font-size: 1rem;
        color: #64748B;
        margin-bottom: 20px;
    }
    
    /* Badges */
    .badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    .badge-blue {
        background: #DBEAFE;
        color: #1E40AF;
    }
    
    .badge-green {
        background: #DCFCE7;
        color: #166534;
    }
    
    .badge-yellow {
        background: #FEF3C7;
        color: #92400E;
    }
    
    .badge-red {
        background: #FEE2E2;
        color: #991B1B;
    }
</style>
""", unsafe_allow_html=True)

# --- BANCO DE DADOS LOCAL (JSON) ---
DB_PATH = "data/demands_db.json"

def carregar_dados():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    if not os.path.exists(DB_PATH):
        with open(DB_PATH, "w", encoding="utf-8") as f:
            json.dump([], f)
    with open(DB_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar_dados(dados):
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with open(DB_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)

# --- MOTOR DE GERAÇÃO DE RELATÓRIOS ---
def gerar_word(demanda):
    doc = Document()
    doc.add_heading('Relatório Executivo de Status - AG-MTE', 0)
    doc.add_paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Departamento de Tecnologia e Processos")
    
    doc.add_heading('1. Visão Geral da Demanda', level=1)
    doc.add_paragraph(f"Título: {demanda['titulo']}")
    doc.add_paragraph(f"Descrição: {demanda['descricao']}")
    doc.add_paragraph(f"Fase Atual: {demanda['fase']}")
    
    doc.add_heading('2. Cronograma e SLA Preditivo', level=1)
    doc.add_paragraph(f"Prazo Estimado: {demanda['prazo_sugerido_dias']} dias úteis")
    
    doc.add_heading('3. Gestão de Riscos e Plano de Contingência', level=1)
    doc.add_paragraph(f"Risco Mapeado: {demanda['risco_principal']}")
    doc.add_paragraph(f"Plano de Mitigação: {demanda['plano_contingencia']}")
    
    filename = f"relatorio_{demanda['id']}.docx"
    doc.save(filename)
    return filename

def gerar_excel(demandas):
    df = pd.DataFrame(demandas)
    filename = "relatorio_demandas_completo.xlsx"
    df.to_excel(filename, index=False, engine='openpyxl')
    return filename

# --- INTERFACE PRINCIPAL ---
st.markdown('<p class="main-header">⚡ AG-MTE</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Agente Master de Tecnologia e Processos — Plataforma Inteligente de Gestão de Ciclo de Vida, SLA e Governança de Entregas</p>', unsafe_allow_html=True)
st.markdown("---")

# Menu Lateral de Navegação
with st.sidebar:
    st.markdown("## ⚡ Navegação")
    st.markdown("---")
    menu = st.radio(
        "Selecione o Módulo:",
        ["🚀 Central de Execução & Nova Demanda", "📊 Painel de Controle & Métricas", "📑 Central de Relatórios (Diretoria)"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.markdown("""
    <div style='padding: 16px; background: rgba(255,255,255,0.05); border-radius: 12px; margin-top: 20px;'>
        <p style='color: #94A3B8; font-size: 0.85rem; margin: 0;'>
            <strong style='color: #F1F5F9;'>AG-MTE v2.0</strong><br>
            Departamento de Tecnologia e Processos<br>
            NIT — Núcleo de Inovação e Tecnologia
        </p>
    </div>
    """, unsafe_allow_html=True)

demandas = carregar_dados()

# --- MÓDULO 1: NOVA DEMANDA E ANÁLISE ---
if menu == "🚀 Central de Execução & Nova Demanda":
    st.markdown('<p class="section-header">Nova Demanda / Solicitação da Diretoria</p>', unsafe_allow_html=True)
    st.markdown('<p class="section-subheader">Insira os dados da atividade para que o agente calcule prazos, mapeie riscos e estruture o fluxo de ponta a ponta.</p>', unsafe_allow_html=True)
    
    with st.form("form_demanda"):
        col1, col2 = st.columns([3, 1])
        with col1:
            titulo = st.text_input("📋 Título do Projeto / Demanda:", placeholder="Ex: Automação de Relatórios de TI")
        with col2:
            complexidade = st.selectbox("🎯 Complexidade:", [1, 2, 3, 4, 5], index=2, format_func=lambda x: f"{x} - {'Muito Baixa' if x==1 else 'Baixa' if x==2 else 'Média' if x==3 else 'Alta' if x==4 else 'Crítica'}")
            
        descricao = st.text_area("📝 Escopo Detalhado / Requisitos:", placeholder="Descreva o que foi pedido, os objetivos e restrições...", height=120)
        
        st.markdown("---")
        btn_executar = st.form_submit_button("🤖 Processar Demanda com IA", use_container_width=True)
        
        if btn_executar:
            if not titulo.strip():
                st.error("❌ O título da demanda é obrigatório.")
            else:
                # Lógica preditiva de SLA e Riscos
                dias_estimados = complexidade * 3
                
                if complexidade <= 2:
                    risco = "Baixo risco de desvio. Tarefa operacional padrão."
                    contingencia = "Execução direta com validação pontual."
                elif complexidade == 3:
                    risco = "Risco moderado. Dependências externas podem impactar prazo."
                    contingencia = "Validação incremental semanal e ajuste ágil de escopo."
                elif complexidade == 4:
                    risco = "Risco alto. Possível gargalo técnico ou dependência cruzada de infraestrutura."
                    contingencia = "Reuniões de alinhamento a cada 3 dias. Plano B para integrações."
                else:
                    risco = "Risco crítico. Alta complexidade com múltiplas dependências."
                    contingencia = "Comitê de crise. Revisão diária. Escopo mínimo viável (MVP)."
                
                # Fase baseada na complexidade
                fases = {
                    1: "Execução Rápida",
                    2: "Planejamento / Início",
                    3: "Desenvolvimento / Execução",
                    4: "Desenvolvimento / Validação",
                    5: "Planejamento Estratégico / Execução Completa"
                }
                
                nova = {
                    "id": len(demandas) + 1,
                    "titulo": titulo,
                    "descricao": descricao,
                    "complexidade": complexidade,
                    "prazo_sugerido_dias": dias_estimados,
                    "fase": fases.get(complexidade, "Planejamento / Início"),
                    "risco_principal": risco,
                    "plano_contingencia": contingencia,
                    "data_criacao": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                
                demandas.append(nova)
                salvar_dados(demandas)
                
                st.success("✅ Demanda processada, planejada e salva com sucesso!")
                
                st.markdown(f"""
                <div class="result-card">
                    <h3>📋 Resultado da Análise de Engenharia</h3>
                    <div class="result-item">
                        <span class="result-icon">⏱️</span>
                        <div class="result-content">
                            <div class="result-label">Prazo Adequado (SLA)</div>
                            <div class="result-value">{dias_estimados} dias úteis</div>
                        </div>
                    </div>
                    <div class="result-item">
                        <span class="result-icon">📊</span>
                        <div class="result-content">
                            <div class="result-label">Fase do Projeto</div>
                            <div class="result-value">{fases.get(complexidade, "Planejamento / Início")}</div>
                        </div>
                    </div>
                    <div class="result-item">
                        <span class="result-icon">⚠️</span>
                        <div class="result-content">
                            <div class="result-label">Risco Identificado</div>
                            <div class="result-value">{risco}</div>
                        </div>
                    </div>
                    <div class="result-item">
                        <span class="result-icon">🛡️</span>
                        <div class="result-content">
                            <div class="result-label">Plano de Contingência</div>
                            <div class="result-value">{contingencia}</div>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

# --- MÓDULO 2: PAINEL DE CONTROLE ---
elif menu == "📊 Painel de Controle & Métricas":
    st.markdown('<p class="section-header">Painel de Controle de Entregas</p>', unsafe_allow_html=True)
    st.markdown('<p class="section-subheader">Acompanhe o andamento de tudo o que foi iniciado.</p>', unsafe_allow_html=True)
    
    if not demandas:
        st.info("📭 Nenhuma demanda cadastrada ainda. Utilize a aba de Execução para adicionar.")
    else:
        df = pd.DataFrame(demandas)
        
        # Métricas
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("""
            <div class="metric-card">
                <div class="metric-value">""" + str(len(demandas)) + """</div>
                <div class="metric-label">Total de Demandas</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            media_prazo = df["prazo_sugerido_dias"].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{media_prazo:.1f}</div>
                <div class="metric-label">Média de Prazo (Dias)</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            media_complexidade = df["complexidade"].mean()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-value">{media_complexidade:.1f}</div>
                <div class="metric-label">Complexidade Média</div>
            </div>
            """, unsafe_allow_html=True)
        with col4:
            st.markdown("""
            <div class="metric-card">
                <div class="metric-value">100%</div>
                <div class="metric-label">Status Operacional</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown('<p class="section-header">📋 Histórico Geral de Atividades</p>', unsafe_allow_html=True)
        st.dataframe(df, use_container_width=True, height=400)

# --- MÓDULO 3: RELATÓRIOS EXECUTIVOS ---
elif menu == "📑 Central de Relatórios (Diretoria)":
    st.markdown('<p class="section-header">Emissão de Relatórios para Apresentação Executiva</p>', unsafe_allow_html=True)
    st.markdown('<p class="section-subheader">Selecione a demanda desejada para compilar o documento formal para a diretoria.</p>', unsafe_allow_html=True)
    
    if not demandas:
        st.warning("⚠️ Cadastre demandas para poder gerar os relatórios.")
    else:
        titulos_demandas = [d['titulo'] for d in demandas]
        escolha = st.selectbox("🔍 Selecione a Demanda:", titulos_demandas)
        
        demanda_selecionada = next(d for d in demandas if d['titulo'] == escolha)
        
        # Card com resumo da demanda selecionada
        st.markdown(f"""
        <div class="result-card" style="margin-top: 16px;">
            <h3>📄 Demanda Selecionada</h3>
            <div class="result-item">
                <span class="result-icon">📋</span>
                <div class="result-content">
                    <div class="result-label">Título</div>
                    <div class="result-value">{demanda_selecionada['titulo']}</div>
                </div>
            </div>
            <div class="result-item">
                <span class="result-icon">📊</span>
                <div class="result-content">
                    <div class="result-label">Complexidade</div>
                    <div class="result-value">{demanda_selecionada['complexidade']}/5</div>
                </div>
            </div>
            <div class="result-item">
                <span class="result-icon">⏱️</span>
                <div class="result-content">
                    <div class="result-label">Prazo Estimado</div>
                    <div class="result-value">{demanda_selecionada['prazo_sugerido_dias']} dias úteis</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        col_w, col_e = st.columns(2)
        
        with col_w:
            st.markdown("### 📄 Documento Executivo (Word)")
            st.markdown("Gere um relatório formatado em `.docx` pronto para apresentação.")
            if st.button("📥 Gerar e Baixar Relatório Word", use_container_width=True):
                arquivo_word = gerar_word(demanda_selecionada)
                with open(arquivo_word, "rb") as f:
                    st.download_button(
                        label="📥 Clique para Baixar",
                        data=f,
                        file_name=arquivo_word,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                st.success("✅ Relatório compilado com sucesso!")
                    
        with col_e:
            st.markdown("### 📊 Planilha de Controle (Excel)")
            st.markdown("Exporte todos os dados em formato `.xlsx` para análise.")
            if st.button("📥 Gerar e Baixar Planilha Excel", use_container_width=True):
                arquivo_excel = gerar_excel(demandas)
                with open(arquivo_excel, "rb") as f:
                    st.download_button(
                        label="📥 Clique para Baixar",
                        data=f,
                        file_name=arquivo_excel,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                st.success("✅ Planilha gerada com sucesso!")
