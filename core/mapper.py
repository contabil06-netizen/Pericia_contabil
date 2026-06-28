"""
Mapeador de contas.
Carrega o arquivo YAML do cliente e agrupa os saldos extraidos
nas linhas correspondentes do modelo Excel.

v2: propaga dre_estrutura (ordem e grupos da DRE) para o RelatorioFinal.
"""
from __future__ import annotations
import re
from decimal import Decimal
from pathlib import Path

import yaml

from core.models import (
    Balancete, ContaContabil, GrupoMapeado,
    NaturezaSaldo, RelatorioFinal, TipoConta,
)


def carregar_mapa_cliente(cliente_id: str, base_dir: str = "clientes") -> dict:
    """Carrega o arquivo YAML de configuracao do cliente."""
    path = Path(base_dir) / f"{cliente_id}.yaml"
    if not path.exists():
        raise FileNotFoundError(
            f"YAML do cliente nao encontrado: {path}\n"
            f"Crie o arquivo '{cliente_id}.yaml' na pasta 'clientes/' "
            f"antes de processar este cliente."
        )
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def mapear(balancete: Balancete, config_cliente: dict) -> RelatorioFinal:
    """
    Aplica o mapa de contas do cliente ao balancete extraido.
    Retorna RelatorioFinal com grupos_balanco, grupos_dre e dre_estrutura.
    """
    relatorio = RelatorioFinal(balancete=balancete)
    relatorio.tipo_pessoa  = config_cliente.get("tipo_pessoa", "")
    relatorio.nome_cliente = config_cliente.get("cliente", "")
    erros = []

    # Balanco Patrimonial
    if "balanco" in config_cliente:
        for secao_nome, secao in config_cliente["balanco"].items():
            for grupo_nome, grupo_cfg in secao.items():
                if not isinstance(grupo_cfg, dict):
                    continue
                resultado = _somar_grupo(balancete, grupo_cfg, grupo_nome)
                resultado.label = grupo_cfg.get("label", grupo_nome)
                resultado.secao = secao_nome
                resultado.sinal = int(grupo_cfg.get("sinal", 1))
                chave = f"{secao_nome}.{grupo_nome}"
                relatorio.grupos_balanco[chave] = resultado
                if resultado.valor == Decimal("0") and grupo_cfg.get("obrigatorio", False):
                    erros.append(
                        f"Grupo '{chave}' obrigatorio retornou zero — verifique o mapeamento."
                    )

    # DRE legada (grupos com celula_excel)
    if "dre" in config_cliente:
        for grupo_nome, grupo_cfg in config_cliente["dre"].items():
            if not isinstance(grupo_cfg, dict):
                continue
            resultado = _somar_grupo(balancete, grupo_cfg, grupo_nome)
            relatorio.grupos_dre[grupo_nome] = resultado

    # Estrutura da DRE (ordem de exibicao)
    # Le a secao dre_estrutura do YAML e monta lista ordenada de grupos
    # para o exporter renderizar na ordem correta com contas analiticas.
    if "dre_estrutura" in config_cliente:
        estrutura = []
        for chave, grupo_cfg in config_cliente["dre_estrutura"].items():
            if not isinstance(grupo_cfg, dict):
                continue
            estrutura.append({
                "chave":     chave,
                "label":     grupo_cfg.get("label", chave),
                "tipo":      grupo_cfg.get("tipo", "despesa"),
                "prefixos":  grupo_cfg.get("prefixos", []),
                "formula":   grupo_cfg.get("formula", ""),
                "subgrupos": grupo_cfg.get("subgrupos", {}),
            })
        relatorio.dre_estrutura = estrutura

    # Totais derivados
    relatorio.total_ativo      = _calcular_total(relatorio.grupos_balanco, "ativo_")
    relatorio.total_passivo    = _calcular_total(relatorio.grupos_balanco, "passivo_")
    relatorio.resultado_periodo = _calcular_resultado(relatorio.grupos_dre)
    relatorio.erros_mapeamento = erros

    return relatorio


def _somar_grupo(balancete: Balancete, grupo_cfg: dict, nome_grupo: str) -> GrupoMapeado:
    total = Decimal("0")
    contas_incluidas = []
    celula = grupo_cfg.get("celula_excel", "")

    termos = (
        grupo_cfg.get("descricoes")
        or grupo_cfg.get("grupos")
        or grupo_cfg.get("codigos")
        or []
    )
    estrategia = (
        "codigo"    if "codigos"    in grupo_cfg else
        "descricao" if "descricoes" in grupo_cfg else
        "auto"
    )
    sinal = Decimal(str(grupo_cfg.get("sinal", 1)))

    for termo in termos:
        for conta in _buscar_contas(balancete, termo, estrategia):
            valor = conta.saldo_atual
            if conta.natureza == NaturezaSaldo.CREDOR:
                valor = -valor
            total += valor * sinal
            contas_incluidas.append(f"{conta.codigo} | {conta.descricao}")

    return GrupoMapeado(
        nome=nome_grupo,
        celula_excel=celula,
        valor=total,
        contas_incluidas=list(set(contas_incluidas)),
    )


def _buscar_contas(balancete: Balancete, termo: str, estrategia: str) -> list[ContaContabil]:
    eh_codigo_exato = (
        (estrategia == "codigo" or (estrategia == "auto" and re.match(r'^[\d\.]+$', termo)))
        and "." not in termo
    )
    # ContaCerta: IDs são sequenciais — sem sobreposição de prefixo entre pai e filhos.
    # Match exato em balancete.contas captura o saldo já agregado da SINTETICA,
    # corrigindo zeros no BP para contas como [392],[483],[833],[1491],[1554].
    if eh_codigo_exato and balancete.layout_detectado == "contacerta":
        return [c for c in balancete.contas if c.codigo == termo]

    contas = balancete.contas_analiticas
    if estrategia == "codigo" or (estrategia == "auto" and re.match(r'^[\d\.]+$', termo)):
        # Códigos inteiros puros (sem ponto) usam correspondência EXATA.
        # Evita que "3" capture 3695, 3719, 3749 etc. no layout fs_sequencial.
        # Códigos hierárquicos com ponto (ex: "1.1.1") continuam usando startswith.
        if "." not in termo:
            return [c for c in contas if c.codigo == termo]
        return [c for c in contas if c.codigo.startswith(termo)]
    termo_lower = termo.lower()
    return [c for c in contas if termo_lower in c.descricao.lower()]


def _calcular_total(grupos: dict[str, GrupoMapeado], prefixo: str) -> Decimal:
    return sum(g.valor for k, g in grupos.items() if k.startswith(prefixo))


def _calcular_resultado(grupos_dre: dict[str, GrupoMapeado]) -> Decimal:
    receitas  = sum(g.valor for k, g in grupos_dre.items() if "receita"  in k.lower())
    custos    = sum(g.valor for k, g in grupos_dre.items() if "custo"    in k.lower())
    despesas  = sum(g.valor for k, g in grupos_dre.items() if "despesa"  in k.lower())
    return receitas - custos - despesas
