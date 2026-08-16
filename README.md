# AG-MTE: Agente Master de Gestão de Tecnologia e Processos

O **AG-MTE** é uma solução de inteligência operacional desenhada para otimizar o fluxo de trabalho, mensurar a produtividade técnica e automatizar a geração de relatórios gerenciais executivos para o Departamento de Tecnologia e Processos.

## Funcionalidades Principais

- **Gestão Inteligente de Demandas:** Rastreamento de ponta a ponta (início, meio e fim).
- **SLA Preditivo:** Algoritmo de estimativa de prazos baseado em complexidade técnica.
- **Automação de Relatórios:** Geração instantânea de documentos formatados (.docx) e planilhas estratégicas (.xlsx).
- **Gestão de Riscos:** Mapeamento proativo de gargalos e planos de contingência.

## Estrutura do Projeto

```
agente_/
├── core/
│   ├── __init__.py           # Modulo principal
│   ├── agent_logic.py        # Logica central do agente
│   └── excel_word_gen.py     # Geradores Excel/Word
├── data/
│   └── demands_db.json       # Banco de dados local
├── main.py                   # Ponto de entrada
├── cli.py                    # Interface interativa
├── requirements.txt          # Dependencias
├── .gitignore                # Arquivos ignorados
└── README.md                 # Documentacao
```

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/eupriscilacristina/agente_.git
   ```

2. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

3. Execute a interface:
   ```bash
   python cli.py
   ```

## Stack Tecnológico

- Python 3.x
- Rich (UI/UX Terminal)
- Pandas / Openpyxl (Data Processing)
- python-docx (Document Automation)

## Uso

A interface CLI interativa oferece as seguintes opções:

| Opção | Descrição |
|-------|-----------|
| 1. Adicionar Nova Demanda | Cadastra e analisa uma nova demanda |
| 2. Listar Demandas | Exibe todas as demandas registradas |
| 3. Gerar Relatório Word | Cria relatório executivo em .docx |
| 4. Exportar Planilha Excel | Exporta dados em .xlsx |
| 5. Sair | Encerra o programa |

## Responsavel

- **Nome:** Priscila Cristina
- **Departamento:** NIT - Departamento de Tecnologia e Processos
- **GitHub:** [@eupriscilacristina](https://github.com/eupriscilacristina)

---

*Documento gerado pelo AG-MTE em 16/08/2026*
