"""
Modelos de dados centrais do sistema de perícia contábil.
Todas as entidades são validadas via Pydantic com tipagem estrita.

CORREÇÃO APLICADA (v4.1):
  O field_validator normalizar_decimal foi reescrito para não destruir
  o separador decimal quando recebe um Decimal já convertido pelos parsers.
  Bug original: str(Decimal('16435.07')) -> replace('.','') -> '1643507'
"""
from __future__ import annotations
from decimal import Decimal
from enum import Enum
from typing import Optional
from pydantic import BaseModel, field_validator


class NaturezaSaldo(str, Enum):
    DEVEDOR = "D"
    CREDOR  = "C"
    ZERO    = "Z"


class TipoConta(str, Enum):
    SINTETICA = "sintetica"
    ANALITICA = "analitica"


class SeveridadeErro(str, Enum):
    CRITICO = "CRITICO"
    AVISO   = "AVISO"
    INFO    = "INFO"


class ContaContabil(BaseModel):
    codigo:          str
    descricao:       str
    tipo:            TipoConta
    natureza:        NaturezaSaldo
    saldo_anterior:  Decimal
    debito:          Decimal
    credito:         Decimal
    saldo_atual:     Decimal
    nivel:           int = 0

    @field_validator("saldo_anterior", "debito", "credito", "saldo_atual", mode="before")
    @classmethod
    def normalizar_decimal(cls, v):
        if v is None:
            return Decimal("0")
        if isinstance(v, Decimal):
            return v
        if isinstance(v, (int, float)):
            return Decimal(str(v))
        s = str(v).strip()
        s = s.rstrip("DCdc").strip()
        if not s or s == "-":
            return Decimal("0")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return Decimal(s)
        except Exception:
            return Decimal("0")

    @property
    def movimento_liquido(self) -> Decimal:
        return self.debito - self.credito


class ErroValidacao(BaseModel):
    severidade:       SeveridadeErro
    conta_codigo:     str
    conta_descricao:  str
    mensagem:         str
    valor_esperado:   Optional[Decimal] = None
    valor_encontrado: Optional[Decimal] = None
    diferenca:        Optional[Decimal] = None


class Balancete(BaseModel):
    empresa:          str
    cnpj:             str
    periodo_inicio:   str
    periodo_fim:      str
    sistema_contabil: str
    layout_detectado: str
    contas:           list[ContaContabil]
    erros_validacao:  list[ErroValidacao] = []

    @property
    def contas_analiticas(self) -> list[ContaContabil]:
        return [c for c in self.contas if c.tipo == TipoConta.ANALITICA]

    @property
    def contas_sinteticas(self) -> list[ContaContabil]:
        return [c for c in self.contas if c.tipo == TipoConta.SINTETICA]

    def buscar_conta(self, codigo: str) -> Optional[ContaContabil]:
        for c in self.contas:
            if c.codigo == codigo:
                return c
        return None

    def buscar_por_descricao(self, termo: str) -> list[ContaContabil]:
        termo_lower = termo.lower()
        return [c for c in self.contas if termo_lower in c.descricao.lower()]

    def soma_grupo(self, codigos_prefixo: list[str], apenas_analiticas: bool = True) -> Decimal:
        total = Decimal("0")
        contas = self.contas_analiticas if apenas_analiticas else self.contas
        for conta in contas:
            for prefixo in codigos_prefixo:
                if conta.codigo.startswith(prefixo):
                    if conta.natureza == NaturezaSaldo.DEVEDOR:
                        total += conta.saldo_atual
                    else:
                        total -= conta.saldo_atual
                    break
        return total


class GrupoMapeado(BaseModel):
    nome:             str
    celula_excel:     str = ""
    valor:            Decimal = Decimal("0")
    contas_incluidas: list[str] = []
    label:            str = ""   # label legível para exibição no BP (vem do YAML)
    secao:            str = ""   # seção do BP (ativo_circulante, passivo_circulante, etc.)
    sinal:            int = 1    # sinal aplicado no mapeamento (-1 para contas redutoras do PL)


class RelatorioFinal(BaseModel):
    balancete:         Balancete
    grupos_balanco:    dict[str, GrupoMapeado] = {}
    grupos_dre:        dict[str, GrupoMapeado] = {}
    erros_mapeamento:  list[str] = []
    total_ativo:       Decimal = Decimal("0")
    total_passivo:     Decimal = Decimal("0")
    resultado_periodo: Decimal = Decimal("0")
    # Estrutura da DRE carregada do YAML — lista ordenada de grupos
    # Cada item: {"chave": str, "label": str, "tipo": str, "prefixos": list[str]}
    dre_estrutura:     list[dict] = []
