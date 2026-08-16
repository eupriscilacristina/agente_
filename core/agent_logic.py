# AG-MTE - Agent Logic Module
# Aqui reside a inteligencia que processa a demanda bruta,
# calcula o SLA inteligente, avalia os riscos e formata os outputs.

import json
import os
from datetime import datetime


class ProcessAgent:
    """
    Classe principal do AG-MTE.
    Processa demandas, calcula SLA, avalia riscos e gerencia historico.
    """

    def __init__(self, db_path="data/demands_db.json"):
        """
        Inicializa o agente com o caminho do banco de dados local.

        Args:
            db_path: Caminho para o arquivo JSON do banco de dados.
        """
        self.db_path = db_path
        self._ensure_db()

    def _ensure_db(self):
        """Garante que o diretorio e o arquivo do banco de dados existam."""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        if not os.path.exists(self.db_path):
            with open(self.db_path, "w", encoding="utf-8") as f:
                json.dump([], f)

    def analisar_demanda(self, titulo: str, descricao: str, complexidade: int) -> dict:
        """
        Analisa a demanda, estima prazos baseados na complexidade (1 a 5)
        e mapeia riscos e impactos.

        Args:
            titulo: Titulo curto da demanda.
            descricao: Descricao detalhada do que precisa ser entregue.
            complexidade: Nivel de complexidade de 1 (simples) a 5 (critico).

        Returns:
            dict: Dicionario com a analise completa da demanda.
        """
        if not 1 <= complexidade <= 5:
            raise ValueError("Complexidade deve estar entre 1 e 5.")

        # Calculo preditivo de prazo baseado na complexidade
        dias_estimados = complexidade * 3

        # Mapeamento de fases
        fases = {
            1: "Execucao Rapida",
            2: "Planejamento / Inicio",
            3: "Desenvolvimento / Execucao",
            4: "Desenvolvimento / Validação",
            5: "Planejamento Estrategico / Execucao Completa"
        }

        # Mapeamento de riscos
        riscos = {
            1: "Risco minimo - tarefa operacional padrao.",
            2: "Baixo risco - possivel necessidade de ajustes menores.",
            3: "Risco moderado - dependencias externas podem impactar prazo.",
            4: "Risco alto - possivel gargalo de integracao tecnica ou dependencia de terceiros.",
            5: "Risco critico - alta complexidade, requer validacao constante com a diretoria."
        }

        analise = {
            "id": len(self._carregar_dados()) + 1,
            "titulo": titulo,
            "descricao": descricao,
            "complexidade": complexidade,
            "prazo_sugerido_dias": dias_estimados,
            "fase": fases.get(complexidade, "Planejamento / Inicio"),
            "risco_principal": riscos.get(complexidade, "Risco nao mapeado."),
            "plano_contingencia": self._gerar_plano_contingencia(complexidade),
            "data_criacao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "status": "Em Analise"
        }

        self._salvar_demanda(analise)
        return analise

    def _gerar_plano_contingencia(self, complexidade: int) -> str:
        """
        Gera um plano de contingencia baseado na complexidade.

        Args:
            complexidade: Nivel de complexidade (1-5).

        Returns:
            str: Plano de contingencia adequado.
        """
        planos = {
            1: "Manter execucao padrao. Sem necessidade de plano especial.",
            2: "Validacao incremental com revisoes rapidas.",
            3: "Ajuste de escopo rapido e validacao semanal com a equipe.",
            4: "Ajuste de escopo rapido e validacao incremental com a diretoria. Definir marcos claros.",
            5: "Revisao diaria de progresso. Escopo minimo viavel (MVP). Comite de crise se necessario."
        }
        return planos.get(complexidade, "Plano padrao de contingencia.")

    def _carregar_dados(self) -> list:
        """
        Carrega os dados do banco de dados local.

        Returns:
            list: Lista de demandas registradas.
        """
        if not os.path.exists(self.db_path):
            return []
        with open(self.db_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _salvar_demanda(self, nova_demanda: dict):
        """
        Salva uma nova demanda no banco de dados local.

        Args:
            nova_demanda: Dicionario com os dados da demanda.
        """
        dados = self._carregar_dados()
        dados.append(nova_demanda)
        with open(self.db_path, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=4)

    def listar_demandas(self) -> list:
        """
        Lista todas as demandas registradas.

        Returns:
            list: Lista de todas as demandas.
        """
        return self._carregar_dados()

    def buscar_demanda_por_id(self, demanda_id: int) -> dict | None:
        """
        Busca uma demanda especifica pelo ID.

        Args:
            demanda_id: ID da demanda a ser buscada.

        Returns:
            dict or None: Dados da demanda ou None se nao encontrada.
        """
        dados = self._carregar_dados()
        for demanda in dados:
            if demanda.get("id") == demanda_id:
                return demanda
        return None
