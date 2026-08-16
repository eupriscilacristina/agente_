# AG-MTE - Main Entry Point
# Agente Master de Gestao de Processos, Tecnologia e Entregas

import sys
from core.agent_logic import ProcessAgent
from core.excel_word_gen import ReportGenerator


def main():
    """Ponto de entrada principal do sistema AG-MTE."""
    # Verificar se deve usar CLI interativa
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        from cli import main as cli_main
        cli_main()
    else:
        # Modo simples (execucao direta)
        print("=" * 60)
        print("  AG-MTE - Agente Master de Tecnologia e Processos")
        print("  Departamento de Tecnologia e Processos - NIT")
        print("=" * 60)
        print("\n  Para usar a interface interativa, execute:")
        print("  python main.py --cli")
        print("\n  Ou execute diretamente:")
        print("  python cli.py")
        print("=" * 60)

        # Executar exemplo padrao
        agent = ProcessAgent()

        titulo = "Automacao de Relatorios de Processos de TI"
        descricao = (
            "Criar fluxo automatizado para coletar metricas de entregas "
            "e gerar visibilidade executiva."
        )
        complexidade = 4

        print(f"\nProcessando demanda: '{titulo}'...")

        try:
            resultado = agent.analisar_demanda(titulo, descricao, complexidade)

            print("\n[+] Analise Concluida com Sucesso:")
            for k, v in resultado.items():
                print(f"  {k.upper():.<30} {v}")

            print("\n[+] Gerando relatorios executivos...")

            resumo_exec = (
                f"A demanda '{titulo}' foi avaliada com complexidade {complexidade}. "
                f"Prazo estimado: {resultado['prazo_sugerido_dias']} dias uteis."
            )
            detalhes = (
                f"Risco mapeado: {resultado['risco_principal']}\n"
                f"Mitigacao: {resultado['plano_contingencia']}"
            )
            print(ReportGenerator.generate_word(resumo_exec, detalhes))

            todas_demandas = agent.listar_demandas()
            print(ReportGenerator.generate_excel(todas_demandas))

            print("\n" + "=" * 60)
            print("  Arquivos gerados:")
            print("  - status_report_diretoria.docx")
            print("  - relatorio_demandas.xlsx")
            print("=" * 60)

        except Exception as e:
            print(f"\n[ERRO] Falha ao processar demanda: {e}")


if __name__ == "__main__":
    main()
