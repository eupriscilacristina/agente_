import streamlit as st
import pandas as pd
from core.agent_logic import ProcessAgent
from core.excel_word_gen import ReportGenerator

# Configuração da página (Layout Executivo)
st.set_page_config(
    page_title="AG-MTE | Gestão de Tecnologia e Processos",
    page_icon="⚡",
    layout="wide"
)

agent = ProcessAgent()

# Cabeçalho Visual
st.title("⚡ AG-MTE: Agente Master de Tecnologia e Processos")
st.markdown("Plataforma de inteligência operacional para planejamento, prazos e relatórios gerenciais.")
st.markdown("---")

# Menu Lateral (Sidebar) para Navegação
st.sidebar.title("Navegação")
opcao = st.sidebar.radio(
    "Escolha uma opção:",
    ["📊 Painel de Demandas", "➕ Nova Demanda & Análise", "📑 Gerar Relatórios Executivos"]
)

# --- ABA 1: PAINEL DE DEMANDAS ---
if opcao == "📊 Painel de Demandas":
    st.subheader("Visão Geral do Ciclo de Vida de Demandas")
    demandas = agent._carregar_dados()
    
    if not demandas:
        st.info("Nenhuma demanda cadastrada no momento. Utilize a aba lateral para cadastrar.")
    else:
        df = pd.DataFrame(demandas)
        # Exibição em tabela interativa
        st.dataframe(df, use_container_width=True)
        
        # Métricas rápidas
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total de Demandas", len(demandas))
        with col2:
            media_dias = df["prazo_sugerido_dias"].mean() if not df.empty else 0
            st.metric("Média de Prazo (Dias)", f"{media_dias:.1f}")
        with col3:
            st.metric("Status do Sistema", "Online / Estável", delta="100%")

# --- ABA 2: NOVA DEMANDA ---
elif opcao == "➕ Nova Demanda & Análise":
    st.subheader("Cadastro e Análise Preditiva de Demanda")
    
    with st.form("form_demanda"):
        titulo = st.text_input("Título da Demanda / Projeto")
        descricao = st.text_area("Descrição Detalhada / Escopo")
        complexidade = st.slider("Nível de Complexidade Técnica (1 a 5):", min_value=1, max_value=5, value=3)
        
        submitted = st.form_submit_button("Processar Demanda com IA")
        
        if submitted:
            if titulo.strip() == "":
                st.error("O título da demanda é obrigatório.")
            else:
                resultado = agent.analisar_demanda(titulo, descricao, complexidade)
                st.success("Demanda processada e salva com sucesso!")
                
                # Exibindo o resultado estruturado
                st.markdown("### 📋 Análise Gerada pelo Agente")
                st.info(f"**Prazo Sugerido (SLA):** {resultado['prazo_sugerido_dias']} dias úteis")
                st.warning(f"**Risco Principal Mapeado:** {resultado['risco_principal']}")
                st.success(f"**Plano de Contingência:** {resultado['plano_contingencia']}")

# --- ABA 3: RELATÓRIOS EXECUTIVOS ---
elif opcao == "📑 Gerar Relatórios Executivos":
    st.subheader("Exportação de Status Reports para a Diretoria")
    
    demandas = agent._carregar_dados()
    if not demandas:
        st.warning("Cadastre ao menos uma demanda para gerar relatórios.")
    else:
        ultimo = demandas[-1]
        
        st.markdown(f"**Demanda Selecionada para Relatório:** `{ultimo['titulo']}`")
        
        col_w, col_e = st.columns(2)
        
        with col_w:
            if st.button("Gerar Relatório em Word (.docx)"):
                resumo_exec = f"A demanda '{ultimo['titulo']}' encontra-se na fase de {ultimo['fase']} com complexidade {ultimo['complexidade']}. Prazo estimado: {ultimo['prazo_sugerido_dias']} dias úteis."
                detalhes = f"Descrição: {ultimo['descricao']}\nRisco Principal: {ultimo['risco_principal']}\nPlano de Contingência: {ultimo['plano_contingencia']}"
                
                arquivo = "status_report_diretoria.docx"
                ReportGenerator.generate_word(resumo_exec, detalhes, arquivo)
                st.success(f"Relatório Word gerado: `{arquivo}`")
                
                with open(arquivo, "rb") as f:
                    st.download_button("📥 Baixar Relatório Word", f, file_name=arquivo)

        with col_e:
            if st.button("Exportar Planilha Excel (.xlsx)"):
                arquivo = "relatorio_demandas.xlsx"
                ReportGenerator.generate_excel(demandas, arquivo)
                st.success(f"Planilha gerada: `{arquivo}`")
                
                with open(arquivo, "rb") as f:
                    st.download_button("📥 Baixar Planilha Excel", f, file_name=arquivo)
