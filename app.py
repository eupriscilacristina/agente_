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

# --- ESTILIZAÇÃO VISUAL CORPORATIVA (CSS) ---
st.markdown("""
<style>
    .main-title { font-size: 2.2rem; font-weight: 800; color: #0F172A; margin-bottom: 0px; }
    .sub-title { font-size: 1.1rem; color: #475569; margin-bottom: 20px; }
    .card-metric { background: #FFFFFF; padding: 20px; border-radius: 10px; border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .card-result { background: #F8FAFC; border-left: 5px solid #2563EB; padding: 20px; border-radius: 6px; margin-top: 20px; border-top: 1px solid #E2E8F0; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0; }
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
st.markdown('<p class="main-title">⚡ AG-MTE: Agente Master de Tecnologia e Processos</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Plataforma Inteligente de Gestão de Ciclo de Vida, SLA e Governança de Entregas</p>', unsafe_allow_html=True)
st.markdown("---")

# Menu Lateral de Navegação
menu = st.sidebar.selectbox("Navegação da Ferramenta", [
    "🚀 Central de Execução & Nova Demanda",
    "📊 Painel de Controle & Métricas",
    "📑 Central de Relatórios (Diretoria)"
])

demandas = carregar_dados()

# --- MÓDULO 1: NOVA DEMANDA E ANÁLISE ---
if menu == "🚀 Central de Execução & Nova Demanda":
    st.subheader("Nova Demanda / Solicitação da Diretoria")
    st.write("Insira os dados da atividade para que o agente calcule prazos, mapeie riscos e estruture o fluxo de ponta a ponta.")
    
    with st.form("form_demanda"):
        col1, col2 = st.columns([3, 1])
        with col1:
            titulo = st.text_input("Título do Projeto / Demanda:")
        with col2:
            complexidade = st.selectbox("Complexidade (1 a 5):", [1, 2, 3, 4, 5], index=2)
            
        descricao = st.text_area("Escopo Detalhado / Requisitos:")
        
        btn_executar = st.form_submit_button("🤖 Processar Demanda com IA")
        
        if btn_executar:
            if not titulo.strip():
                st.error("O título da demanda é obrigatório.")
            else:
                # Lógica preditiva de SLA e Riscos
                dias_estimados = complexidade * 3
                risco = "Possível gargalo técnico ou dependência cruzada de infraestrutura." if complexidade > 2 else "Baixo risco de desvio."
                contingencia = "Validação incremental semanal com a diretoria e ajuste rápido de escopo."
                
                nova = {
                    "id": len(demandas) + 1,
                    "titulo": titulo,
                    "descricao": descricao,
                    "complexidade": complexidade,
                    "prazo_sugerido_dias": dias_estimados,
                    "fase": "Planejamento / Início",
                    "risco_principal": risco,
                    "plano_contingencia": contingencia,
                    "data_criacao": datetime.now().strftime('%Y-%m-%d')
                }
                
                demandas.append(nova)
                salvar_dados(demandas)
                
                st.success("Demanda processada, planejada e salva com sucesso!")
                
                st.markdown(f"""
                <div class="card-result">
                    <h3>📋 Resultado da Análise de Engenharia</h3>
                    <p><b>⏱ Prazo Adequado (SLA):</b> {dias_estimados} dias úteis</p>
                    <p><b>⚠️ Risco Identificado:</b> {risco}</p>
                    <p><b>🛡 Plano de Contingência:</b> {contingencia}</p>
                </div>
                """, unsafe_allow_html=True)

# --- MÓDULO 2: PAINEL DE CONTROLE ---
elif menu == "📊 Painel de Controle & Métricas":
    st.subheader("Painel de Controle de Entregas")
    
    if not demandas:
        st.info("Nenhuma demanda cadastrada ainda. Utilize a aba de Execução para adicionar.")
    else:
        df = pd.DataFrame(demandas)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Total de Demandas", len(demandas))
        with c2:
            media_prazo = df["prazo_sugerido_dias"].mean()
            st.metric("Média de Prazo", f"{media_prazo:.1f} dias")
        with c3:
            st.metric("Status Operacional", "100% Online")
            
        st.markdown("### 📋 Histórico Geral de Atividades")
        st.dataframe(df, use_container_width=True)

# --- MÓDULO 3: RELATÓRIOS EXECUTIVOS ---
elif menu == "📑 Central de Relatórios (Diretoria)":
    st.subheader("Emissão de Relatórios para Apresentação Executiva")
    
    if not demandas:
        st.warning("Cadastre demandas para poder gerar os relatórios.")
    else:
        st.write("Selecione a demanda desejada para compilar o documento formal para a diretoria:")
        
        titulos_demandas = [d['titulo'] for d in demandas]
        escolha = st.selectbox("Selecione a Demanda:", titulos_demandas)
        
        demanda_selecionada = next(d for d in demandas if d['titulo'] == escolha)
        
        col_w, col_e = st.columns(2)
        
        with col_w:
            st.markdown("#### Documento Executivo (Word)")
            if st.button("Gerar Relatório .docx"):
                arquivo_word = gerar_word(demanda_selecionada)
                st.success("Relatório compilado com sucesso!")
                with open(arquivo_word, "rb") as f:
                    st.download_button("📥 Baixar Arquivo Word", f, file_name=arquivo_word)
                    
        with col_e:
            st.markdown("#### Planilha de Controle (Excel)")
            if st.button("Exportar Base .xlsx"):
                arquivo_excel = gerar_excel(demandas)
                st.success("Planilha gerada com sucesso!")
                with open(arquivo_excel, "rb") as f:
                    st.download_button("📥 Baixar Planilha Excel", f, file_name=arquivo_excel)
