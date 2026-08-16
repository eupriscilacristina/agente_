import os
import questionary
from rich import print as rprint
from rich.panel import Panel
from rich.table import Table
from core.agent_logic import ProcessAgent
from core.excel_word_gen import ReportGenerator

def exibir_cabecalho():
    rprint(Panel.fit(
        "[bold cyan]AG-MTE: AGENTE MASTER DE TECNOLOGIA E PROCESSOS[/bold cyan]\n"
        "[italic]Gerenciamento de Demandas, Prazos e Relatórios Executivos[/italic]",
        border_style="cyan"
    ))

def main():
    agent = ProcessAgent()
    
    while True:
        exibir_cabecalho()
        
        opcao = questionary.select(
            "O que você deseja fazer?",
            choices=[
                "1. Adicionar e Analisar Nova Demanda",
                "2. Listar Demandas Atuais",
                "3. Gerar Relatório Executivo (Word)",
                "4. Exportar Planilha (Excel)",
                "5. Sair"
            ]
        ).ask()
        
        if opcao.startswith("1"):
            rprint("\n[bold yellow]--- Cadastro de Nova Demanda ---[/bold yellow]")
            titulo = questionary.text("Título da demanda/projeto:").ask()
            descricao = questionary.text("Descrição detalhada:").ask()
            
            complexidade = questionary.select(
                "Nível de Complexidade (1 a 5):",
                choices=["1", "2", "3", "4", "5"]
            ).ask()
            
            resultado = agent.analisar_demanda(titulo, descricao, int(complexidade))
            
            rprint("\n[bold green]✔ Demanda processada e salva com sucesso![/bold green]")
            rprint(f"⏱ [cyan]Prazo Sugerido:[/cyan] {resultado['prazo_sugerido_dias']} dias")
            rprint(f"⚠️ [yellow]Risco Mapeado:[/yellow] {resultado['risco_principal']}\n")
            
            questionary.press_any_key_to_continue().ask()

        elif opcao.startswith("2"):
            demandas = agent._carregar_dados()
            if not demandas:
                rprint("\n[yellow]Nenhuma demanda cadastrada no momento.[/yellow]\n")
            else:
                table = Table(title="Painel de Demandas Atuais")
                table.add_column("ID", style="cyan", justify="center")
                table.add_column("Título", style="magenta")
                table.add_column("Complexidade", justify="center")
                table.add_column("Prazo (Dias)", justify="center")
                table.add_column("Fase", style="green")

                for d in demandas:
                    table.add_row(
                        str(d["id"]),
                        d["titulo"],
                        str(d["complexidade"]),
                        str(d["prazo_sugerido_dias"]),
                        d["fase"]
                    )
                rprint(table)
            
            questionary.press_any_key_to_continue().ask()

        elif opcao.startswith("3"):
            demandas = agent._carregar_dados()
            if not demandas:
                rprint("\n[yellow]Cadastre demandas antes de gerar relatórios.[/yellow]\n")
            else:
                ultimo = demandas[-1]
                resumo_exec = f"A demanda '{ultimo['titulo']}' encontra-se na fase de {ultimo['fase']} com complexidade {ultimo['complexidade']}. Prazo estimado: {ultimo['prazo_sugerido_dias']} dias úteis."
                detalhes = f"Descrição: {ultimo['descricao']}\nRisco Principal: {ultimo['risco_principal']}\nPlano de Contingência: {ultimo['plano_contingencia']}"
                
                msg = ReportGenerator.generate_word(resumo_exec, detalhes)
                rprint(f"\n[bold green]✔ {msg}[/bold green]\n")
            
            questionary.press_any_key_to_continue().ask()

        elif opcao.startswith("4"):
            demandas = agent._carregar_dados()
            if not demandas:
                rprint("\n[yellow]Nenhuma demanda para exportar.[/yellow]\n")
            else:
                msg = ReportGenerator.generate_excel(demandas)
                rprint(f"\n[bold green]✔ {msg}[/bold green]\n")
            
            questionary.press_any_key_to_continue().ask()

        elif opcao.startswith("5"):
            rprint("\n[bold cyan]Até logo! Fechando o Agente Master.[/bold cyan]")
            break

if __name__ == "__main__":
    main()
