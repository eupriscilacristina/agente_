# AG-MTE - CLI Interface
# Interface de linha de comando interativa e formatada

import questionary
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.syntax import Syntax
from rich import box
from datetime import datetime

from core.agent_logic import ProcessAgent
from core.excel_word_gen import ReportGenerator


console = Console()


def show_banner():
    """Exibe o banner principal do AG-MTE."""
    banner = """
[bold cyan]+============================================================+
|                                                              |
|   AG-MTE                                                     |
|   Agente Master de Gestao de Processos, Tecnologia e Entregas|
|   Departamento de Tecnologia e Processos - NIT               |
|                                                              |
+============================================================+[/bold cyan]"""
    console.print(banner)


def show_menu():
    """Exibe o menu principal e retorna a opcao escolhida."""
    console.print("\n[bold yellow]═══ MENU PRINCIPAL ═══[/bold yellow]\n")

    opcoes = [
        "1. Nova Demanda",
        "2. Listar Demandas",
        "3. Buscar Demanda por ID",
        "4. Gerar Relatorios",
        "5. Sair"
    ]

    for opcao in opcoes:
        console.print(f"  [cyan]{opcao}[/cyan]")

    escolha = questionary.select(
        "\nSelecione uma opcao:",
        choices=["1", "2", "3", "4", "5"]
    ).ask()

    return escolha


def nova_demanda(agent: ProcessAgent):
    """Coleta dados de uma nova demanda e a processa."""
    console.print("\n[bold green]═══ NOVA DEMANDA ═══[/bold green]\n")

    titulo = questionary.text("Titulo da demanda:").ask()
    if not titulo:
        console.print("[red]Titulo nao pode ser vazio![/red]")
        return

    descricao = questionary.text("Descricao detalhada:").ask()
    if not descricao:
        console.print("[red]Descricao nao pode ser vazia![/red]")
        return

    complexidade = questionary.select(
        "Nivel de complexidade (1-5):",
        choices=[
            "1 - Simples (operacional)",
            "2 - Baixa (ajustes menores)",
            "3 - Media (dependencias externas)",
            "4 - Alta (integracao tecnica)",
            "5 - Critica (alta complexidade)"
        ]
    ).ask()

    # Extrair numero da complexidade
    num_complexidade = int(complexidade[0])

    # Processar demanda
    with console.status("[bold green]Processando demanda...[/bold green]"):
        resultado = agent.analisar_demanda(titulo, descricao, num_complexidade)

    # Exibir resultado formatado
    exibir_resultado_analise(resultado)


def exibir_resultado_analise(resultado: dict):
    """Exibe o resultado da analise de forma formatada."""
    console.print("\n[bold cyan]═══ ANALISE CONCLUIDA ═══[/bold cyan]\n")

    # Tabela de resultados
    table = Table(title="Resultado da Analise", box=box.ROUNDED)
    table.add_column("Campo", style="cyan", width=25)
    table.add_column("Valor", style="white")

    table.add_row("ID", str(resultado["id"]))
    table.add_row("Titulo", resultado["titulo"])
    table.add_row("Descricao", resultado["descricao"][:50] + "...")
    table.add_row("Complexidade", f"{resultado['complexidade']}/5")
    table.add_row("Prazo Estimado", f"{resultado['prazo_sugerido_dias']} dias")
    table.add_row("Fase", resultado["fase"])
    table.add_row("Risco Principal", resultado["risco_principal"])
    table.add_row("Plano Contingencia", resultado["plano_contingencia"])
    table.add_row("Data Criacao", resultado["data_criacao"])
    table.add_row("Status", resultado["status"])

    console.print(table)

    # Perguntar se quer gerar relatorios
    gerar = questionary.confirm("Deseja gerar os relatorios agora?").ask()
    if gerar:
        gerar_relatorios(resultado)


def listar_demandas(agent: ProcessAgent):
    """Lista todas as demandas registradas."""
    console.print("\n[bold green]═══ DEMANDAS REGISTRADAS ═══[/bold green]\n")

    demandas = agent.listar_demandas()

    if not demandas:
        console.print("[yellow]Nenhuma demanda registrada.[/yellow]")
        return

    # Tabela de demandas
    table = Table(title=f"Total: {len(demandas)} demandas", box=box.ROUNDED)
    table.add_column("ID", style="cyan", width=5)
    table.add_column("Titulo", style="white", width=30)
    table.add_column("Complexidade", style="yellow", width=12)
    table.add_column("Prazo", style="green", width=8)
    table.add_column("Status", style="magenta", width=12)

    for d in demandas:
        table.add_row(
            str(d["id"]),
            d["titulo"][:28] + "..." if len(d["titulo"]) > 28 else d["titulo"],
            f"{d['complexidade']}/5",
            f"{d['prazo_sugerido_dias']}d",
            d["status"]
        )

    console.print(table)


def buscar_demanda(agent: ProcessAgent):
    """Busca uma demanda especifica por ID."""
    console.print("\n[bold green]═══ BUSCAR DEMANDA ═══[/bold green]\n")

    demanda_id = questionary.text("ID da demanda:").ask()
    if not demanda_id or not demanda_id.isdigit():
        console.print("[red]ID invalido![/red]")
        return

    demanda = agent.buscar_demanda_por_id(int(demanda_id))

    if not demanda:
        console.print(f"[red]Demanda {demanda_id} nao encontrada![/red]")
        return

    exibir_resultado_analise(demanda)


def gerar_relatorios(resultado: dict = None):
    """Gera os relatorios executivos."""
    console.print("\n[bold green]═══ GERAR RELATORIOS ═══[/bold green]\n")

    agent = ProcessAgent()
    todas_demandas = agent.listar_demandas()

    with console.status("[bold yellow]Gerando relatorios...[/bold yellow]"):
        # Gerar Word
        if resultado:
            resumo_exec = (
                f"A demanda '{resultado['titulo']}' foi avaliada com complexidade "
                f"{resultado['complexidade']}. Prazo estimado: "
                f"{resultado['prazo_sugerido_dias']} dias uteis."
            )
            detalhes = (
                f"Risco mapeado: {resultado['risco_principal']}\n"
                f"Mitigacao: {resultado['plano_contingencia']}"
            )
            word_result = ReportGenerator.generate_word(resumo_exec, detalhes)
        else:
            word_result = ReportGenerator.generate_word(
                "Relatorio consolidado de todas as demandas registradas.",
                "Veja planilha para detalhes."
            )

        # Gerar Excel
        excel_result = ReportGenerator.generate_excel(todas_demandas)

    # Exibir resultados
    console.print(Panel(
        f"[green]{word_result}[/green]\n[green]{excel_result}[/green]",
        title="[bold]Arquivos Gerados[/bold]",
        border_style="green"
    ))


def main():
    """Funcao principal da CLI."""
    show_banner()

    agent = ProcessAgent()

    while True:
        escolha = show_menu()

        if escolha == "1":
            nova_demanda(agent)
        elif escolha == "2":
            listar_demandas(agent)
        elif escolha == "3":
            buscar_demanda(agent)
        elif escolha == "4":
            gerar_relatorios()
        elif escolha == "5":
            console.print("\n[bold cyan]Ate logo! AG-MTE encerrado.[/bold cyan]\n")
            break
        else:
            console.print("[red]Opcao invalida![/red]")


if __name__ == "__main__":
    main()
