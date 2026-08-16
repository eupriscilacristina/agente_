# AG-MTE - Agente Master de Gestao de Processos, Tecnologia e Entregas

## Visao Geral

O **AG-MTE** e um assistente de IA senior estruturado para atuar como o braco direito tecnico e estrategico do **Departamento de Tecnologia e Processos (NIT)**, otimizando o fluxo de trabalho de ponta a ponta.

## Missoes

1. **Gestao de Ciclo de Vida de Demandas** - Acompanhamento rigoroso de tudo o que e iniciado, desenvolvido, inventado ou reinventado
2. **Estimativa Preditiva de Prazos** - Analise da complexidade de tarefas para sugerir prazos viaveis e competitivos
3. **Analise de Impacto Sistemico** - Avaliacao preventiva de como cada demanda afeta as demais areas
4. **Gestao de Riscos e Resiliencia** - Identificacao proativa de gargalos e desenho de planos de contingencia
5. **Geracao de Relatorios Executivos** - Estruturacao de dados para exportacao em Word/Excel

## Estrutura do Projeto

```
agente_/
├── core/
│   ├── __init__.py           # Modulo principal do AG-MTE
│   ├── agent_logic.py        # Logica central do agente e tratamento de dados
│   └── excel_word_gen.py     # Geradores de arquivos .xlsx e .docx
│
├── data/
│   └── demands_db.json       # Banco de dados local leve para historico
│
├── main.py                   # Ponto de entrada do sistema
├── cli.py                    # Interface de linha de comando interativa
├── requirements.txt          # Dependencias do projeto
├── .gitignore                # Arquivos ignorados pelo Git
└── README.md                 # Documentacao
```

## Stack Tecnologica

- **Python 3.10+** - Linguagem principal
- **Pandas** - Manipulacao de dados e planilhas
- **openpyxl** - Geracao de arquivos Excel
- **python-docx** - Geracao de documentos Word
- **rich** - Formatacao de terminal e tabelas
- **questionary** - Interface CLI interativa
- **JSON** - Banco de dados local leve

## Instalacao

1. Clone o repositorio:
   ```bash
   git clone https://github.com/eupriscilacristina/agente_.git
   ```

2. Navegue ate a pasta do projeto:
   ```bash
   cd agente_
   ```

3. Instale as dependencias:
   ```bash
   pip install -r requirements.txt
   ```

4. Execute o agente:
   ```bash
   python main.py
   ```

## Como Funciona

### Modos de Execucao

**1. Modo CLI Interativo (Recomendado):**
```bash
python cli.py
# ou
python main.py --cli
```

**2. Modo Simples (Exemplo padrao):**
```bash
python main.py
```

### Funcionalidades da CLI

- **Nova Demanda:** Coleta dados interativamente e processa
- **Listar Demandas:** Exibe todas as demandas registradas
- **Buscar por ID:** Detalha uma demanda especifica
- **Gerar Relatorios:** Cria arquivos Word e Excel

### Entrada
O usuario fornece:
- **Titulo** da demanda
- **Descricao** detalhada
- **Complexidade** (1 a 5)

### Processamento
O AG-MTE:
- Calcula prazo estimado (complexidade x 3 dias)
- Mapeia fase do projeto
- Identifica riscos principais
- Gera plano de contingencia

### Saida
Sao gerados dois arquivos:
- `status_report_diretoria.docx` - Relatorio executivo em Word
- `relatorio_demandas.xlsx` - Planilha de acompanhamento em Excel

## Framework de Saida

Cada demanda e processada seguindo o framework:

| Etapa | Descricao |
|-------|-----------|
| 1. Resumo Executivo | Entendimento do escopo e classificacao |
| 2. Cronograma SLA | Estimativa realista com marcos e prazos |
| 3. Impacto Cruzado | Analise multi-area (TI, Processos, Operacao) |
| 4. Gestao de Riscos | Identificacao de riscos e planos de contingencia |
| 5. Relatorio Diretoria | Formato Word (.docx) e Excel (.xlsx) |

## Status do Projeto

| Fase | Status |
|------|--------|
| Configuracao GitHub | Concluido |
| Estrutura do Projeto | Concluido |
| Logica do Agente | Concluido |
| Gerador de Relatorios | Concluido |
| Interface CLI | Concluido |
| Testes Unitarios | Pendente |
| Producao | Pendente |

## Responsavel

- **Nome:** Priscila Cristina
- **Departamento:** NIT - Departamento de Tecnologia e Processos
- **GitHub:** [@eupriscilacristina](https://github.com/eupriscilacristina)

## Licenca

Projeto interno - Uso restrito a equipe NIT.

---

*Documento gerado pelo AG-MTE em 16/08/2026*
