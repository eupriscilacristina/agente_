# AG-MTE - Main Entry Point
# Agente Master de Gestao de Processos, Tecnologia e Entregas

from core.agent_logic import ProcessAgent
from core.excel_word_gen import ReportGenerator


def main():
    """Ponto de entrada principal do sistema AG-MTE."""
    print("=" * 60)
    print("  AG-MTE - Agente Master de Tecnologia e Processos")
    print("  Departamento de Tecnologia e Processos - NIT")
    print("=" * 60)

    agent = ProcessAgent()

    # Exemplo de entrada de uma nova demanda vinda da diretoria
    titulo = "Automacao de Relatorios de Processos de TI"
    descricao = (
        "Criar fluxo automatizado para coletar metricas de entregas "
        "e gerar visibilidade executiva."
    )
    complexidade = 4  # Escala de 1 a 5

    print(f"\nProcessando demanda: '{titulo}'...")
    print(f"Complexidade: {complexidade}/5")

    try:
        resultado = agent.analisar_demanda(titulo, descricao, complexidade)

        print("\n" + "-" * 60)
        print("[+] Analise Concluida com Sucesso:")
        print("-" * 60)
        for k, v in resultado.items():
            print(f"  {k.upper():.<30} {v}")

        # Gerando os arquivos para entrega a diretoria
        print("\n" + "-" * 60)
        print("[+] Gerando relatorios executivos...")
        print("-" * 60)

        # Gerar Word
        resumo_exec = (
            f"A demanda '{titulo}' foi avaliada com complexidade {complexidade}. "
            f"Prazo estimado de entrega: {resultado['prazo_sugerido_dias']} dias uteis."
        )
        detalhes = (
            f"Risco mapeado: {resultado['risco_principal']}\n"
            f"Mitigacao: {resultado['plano_contingencia']}"
        )
        print(ReportGenerator.generate_word(resumo_exec, detalhes))

        # Gerar Excel com o historico atualizado
        todas_demandas = agent.listar_demandas()
        print(ReportGenerator.generate_excel(todas_demandas))

        print("\n" + "=" * 60)
        print("  Arquivos gerados na pasta do projeto:")
        print("  - status_report_diretoria.docx")
        print("  - relatorio_demandas.xlsx")
        print("=" * 60)

    except Exception as e:
        print(f"\n[ERRO] Falha ao processar demanda: {e}")


if __name__ == "__main__":
    main()
