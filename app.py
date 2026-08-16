import streamlit as st
import pandas as pd
from core.agent_logic import ProcessAgent
from core.excel_word_gen import ReportGenerator

# Configuracao da pagina
st.set_page_config(
    page_title="AG-MTE - Agente Master de Tecnologia e Processos",
    page_icon="⚡",
    layout="wide"
)

# Inicializar agente
agent = ProcessAgent()

# Cabecalho
st.title("AG-MTE: Agente Master de Tecnologia e Processos")
st.markdown("---")

# Sidebar com menu
menu = st.sidebar.selectbox(
    "Menu",
    ["🏠 Inicio", "➕ Nova Demanda", "📋 Listar Demandas", "📊 Gerar Relatorios"]
)

# Pagina Inicio
if menu == "🏠 Inicio":
    st.header("Bem-vindo ao AG-MTE")
    st.markdown("""
    ### Funcionalidades
    - **Gestao Inteligente de Demandas:** Rastreamento de ponta a ponta
    - **SLA Preditivo:** Estimativa de prazos por complexidade
    - **Automacao de Relatorios:** Geracao de .docx e .xlsx
    - **Gestao de Riscos:** Mapeamento de gargalos e contingencias
    """)
    
    # Metricas
    demandas = agent.listar_demandas()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total de Demandas", len(demandas))
    with col2:
        if demandas:
            prazo_medio = sum(d["prazo_sugerido_dias"] for d in demandas) / len(demandas)
            st.metric("Prazo Medio", f"{prazo_medio:.1f} dias")
        else:
            st.metric("Prazo Medio", "0 dias")
    with col3:
        if demandas:
            complexidade_media = sum(d["complexidade"] for d in demandas) / len(demandas)
            st.metric("Complexidade Media", f"{complexidade_media:.1f}/5")
        else:
            st.metric("Complexidade Media", "0/5")

# Pagina Nova Demanda
elif menu == "➕ Nova Demanda":
    st.header("Cadastrar Nova Demanda")
    
    with st.form("nova_demanda"):
        titulo = st.text_input("Titulo da Demanda")
        descricao = st.text_area("Descricao Detalhada")
        complexidade = st.slider("Complexidade", 1, 5, 3)
        
        submitted = st.form_submit_button("Processar Demanda")
        
        if submitted:
            if titulo and descricao:
                resultado = agent.analisar_demanda(titulo, descricao, complexidade)
                st.success("Demanda processada com sucesso!")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"**Prazo Sugerido:** {resultado['prazo_sugerido_dias']} dias")
                    st.info(f"**Fase:** {resultado['fase']}")
                with col2:
                    st.warning(f"**Risco:** {resultado['risco_principal']}")
                    st.warning(f"**Contingencia:** {resultado['plano_contingencia']}")
            else:
                st.error("Preencha todos os campos!")

# Pagina Listar Demandas
elif menu == "📋 Listar Demandas":
    st.header("Demandas Registradas")
    
    demandas = agent.listar_demandas()
    
    if demandas:
        df = pd.DataFrame(demandas)
        st.dataframe(
            df[["id", "titulo", "complexidade", "prazo_sugerido_dias", "fase", "status"]],
            use_container_width=True
        )
    else:
        st.info("Nenhuma demanda cadastrada.")

# Pagina Gerar Relatorios
elif menu == "📊 Gerar Relatorios":
    st.header("Gerar Relatorios Executivos")
    
    demandas = agent.listar_demandas()
    
    if demandas:
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Gerar Relatorio Word"):
                ultimo = demandas[-1]
                resumo = f"Demanda: {ultimo['titulo']}. Fase: {ultimo['fase']}. Prazo: {ultimo['prazo_sugerido_dias']} dias."
                detalhes = f"Risco: {ultimo['risco_principal']}. Contingencia: {ultimo['plano_contingencia']}"
                msg = ReportGenerator.generate_word(resumo, detalhes)
                st.success(msg)
        
        with col2:
            if st.button("Gerar Planilha Excel"):
                msg = ReportGenerator.generate_excel(demandas)
                st.success(msg)
    else:
        st.info("Cadastre demandas antes de gerar relatorios.")

# Rodape
st.markdown("---")
st.markdown("*AG-MTE - Departamento de Tecnologia e Processos - NIT*")
