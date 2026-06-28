"""
Módulo de validação contábil.
Executa todas as verificações matemáticas e de integridade
antes de exportar para o Excel.
"""
from __future__ import annotations
from decimal import Decimal

from core.models import (
    Balancete, ContaContabil, ErroValidacao,
    NaturezaSaldo, SeveridadeErro, TipoConta,
)

TOLERANCIA_CENTAVOS = Decimal("0.05")


def validar_balancete(balancete: Balancete) -> list[ErroValidacao]:
    """
    Executa todas as validações. Retorna lista de erros encontrados.
    """
    erros: list[ErroValidacao] = []

    erros.extend(_validar_partidas_dobradas(balancete))
    erros.extend(_validar_equacao_balancete(balancete))
    erros.extend(_validar_contas_sem_movimento(balancete))

    return erros


def _validar_partidas_dobradas(balancete: Balancete) -> list[ErroValidacao]:
    """
    Verifica se Total Débito == Total Crédito (partidas dobradas).
    """
    erros = []
    analiticas = balancete.contas_analiticas

    total_debito = sum(c.debito for c in analiticas)
    total_credito = sum(c.credito for c in analiticas)
    diferenca = abs(total_debito - total_credito)

    if diferenca > TOLERANCIA_CENTAVOS:
        erros.append(ErroValidacao(
            severidade=SeveridadeErro.CRITICO,
            conta_codigo="GERAL",
            conta_descricao="Verificação de Partidas Dobradas",
            mensagem=f"Total Débito ≠ Total Crédito. Diferença: R$ {diferenca:,.2f}",
            valor_esperado=total_debito,
            valor_encontrado=total_credito,
            diferenca=diferenca,
        ))

    return erros


def _validar_equacao_balancete(balancete: Balancete) -> list[ErroValidacao]:
    """
    Verifica a equação patrimonial para balancete SEM ENCERRAMENTO:
    Ativo = Passivo + PL + (Receitas - Despesas do período)

    Em balancetes sem encerramento, as contas de resultado ainda não foram
    transferidas para o PL, então a equação padrão Ativo = Passivo + PL
    não se aplica diretamente.
    """
    erros = []
    contas = balancete.contas_analiticas

    def soma_grupo(prefixos: list[str]) -> Decimal:
        total = Decimal("0")
        for c in contas:
            for p in prefixos:
                if c.codigo.startswith(p):
                    if c.natureza == NaturezaSaldo.DEVEDOR:
                        total += c.saldo_atual
                    else:
                        total -= c.saldo_atual
                    break
        return total

    # Heurística baseada no dígito inicial do código
    ativo = soma_grupo(["1"])
    passivo = soma_grupo(["2"])

    # Equilíbrio geral (considerando que contas de resultado compensam)
    desequilibrio = abs(abs(ativo) - abs(passivo))

    if desequilibrio > Decimal("1.00"):
        erros.append(ErroValidacao(
            severidade=SeveridadeErro.AVISO,
            conta_codigo="EQUAÇÃO",
            conta_descricao="Equilíbrio Patrimonial",
            mensagem=(
                f"Ativo (R$ {ativo:,.2f}) ≠ Passivo+PL (R$ {passivo:,.2f}). "
                f"Diferença: R$ {desequilibrio:,.2f}. "
                "Em balancete sem encerramento, diferença pode refletir resultado do período."
            ),
            valor_esperado=ativo,
            valor_encontrado=passivo,
            diferenca=desequilibrio,
        ))

    return erros


def _validar_contas_sem_movimento(balancete: Balancete) -> list[ErroValidacao]:
    """
    Identifica contas onde Saldo Atual ≠ equação contábil correta.

    A equação respeita a natureza da conta:
      Conta Devedora  (Ativo, Despesa):  Saldo Atual = Anterior + Débito  − Crédito
      Conta Credora   (Passivo, Receita): Saldo Atual = Anterior + Crédito − Débito

    Diferenças dentro da tolerância de centavos são ignoradas (arredondamentos).
    """
    erros = []

    for conta in balancete.contas_analiticas:
        if conta.natureza == NaturezaSaldo.CREDOR:
            # Passivo, PL e Receitas: aumentam pelo crédito
            esperado = conta.saldo_anterior + conta.credito - conta.debito
        else:
            # Ativo e Despesas: aumentam pelo débito
            esperado = conta.saldo_anterior + conta.debito - conta.credito

        diferenca = abs(esperado - conta.saldo_atual)

        if diferenca > TOLERANCIA_CENTAVOS:
            erros.append(ErroValidacao(
                severidade=SeveridadeErro.AVISO,
                conta_codigo=conta.codigo,
                conta_descricao=conta.descricao,
                mensagem=(
                    f"Saldo Atual calculado (R$ {esperado:,.2f}) ≠ "
                    f"Saldo Atual informado (R$ {conta.saldo_atual:,.2f}). "
                    f"Diferença: R$ {diferenca:,.2f}. "
                    "Natureza: "
                    + ("Credora (Ant + Cré − Déb)" if conta.natureza == NaturezaSaldo.CREDOR
                       else "Devedora (Ant + Déb − Cré)")
                    + ". Verifique se o parser leu corretamente o D/C desta conta."
                ),
                valor_esperado=esperado,
                valor_encontrado=conta.saldo_atual,
                diferenca=diferenca,
            ))

    return erros


def resumo_validacao(erros: list[ErroValidacao]) -> dict:
    """Retorna contagem por severidade para exibição no relatório."""
    return {
        "criticos": sum(1 for e in erros if e.severidade == SeveridadeErro.CRITICO),
        "avisos": sum(1 for e in erros if e.severidade == SeveridadeErro.AVISO),
        "info": sum(1 for e in erros if e.severidade == SeveridadeErro.INFO),
        "total": len(erros),
        "aprovado": not any(e.severidade == SeveridadeErro.CRITICO for e in erros),
    }
