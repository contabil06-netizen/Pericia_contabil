---
name: pericia-validacao-contabil
description: >
  Use este skill sempre que precisar criar, corrigir ou entender as validações
  contábeis do projeto de perícia. Triggers incluem: partidas dobradas,
  equação patrimonial, balancete sem encerramento, validação de saldo,
  erro de consistência, aviso de desequilíbrio, natureza D/C incorreta,
  validar extração, verificar se valores batem, diferença de centavos,
  tolerância de validação, Decimal vs float, precisão monetária.
  SEMPRE use este skill antes de editar validator.py ou models.py.
---

# Skill: Validação Contábil

## Contexto do domínio

Os clientes são empresas em **Recuperação Judicial** que entregam
**balancetes sem encerramento contábil** mensalmente.
Sem encerramento = as contas de receita e despesa ainda estão abertas,
não foram transferidas ao Patrimônio Líquido.

## As 3 validações implementadas

### 1. Partidas Dobradas
**Regra:** Total de Débitos do período = Total de Créditos do período

```python
total_debito  = sum(c.debito  for c in contas_analiticas)
total_credito = sum(c.credito for c in contas_analiticas)
diferenca = abs(total_debito - total_credito)

if diferenca > TOLERANCIA:   # TOLERANCIA = Decimal("0.05")
    # CRÍTICO — indica erro de extração ou PDF com problema
```

**Quando falha:** geralmente indica que o parser perdeu contas,
ou que a natureza D/C foi invertida em algum grupo.

### 2. Equação Patrimonial (sem encerramento)
**Regra correta:**
```
Ativo = Passivo + PL + (Receitas − Despesas do período)
```

**NÃO usar** `Ativo = Passivo + PL` — essa equação só vale após encerramento.

```python
ativo   = soma contas que começam com "1"
passivo = soma contas que começam com "2"
# diferença pode ser o resultado do período (normal em balancete aberto)
desequilibrio = abs(abs(ativo) - abs(passivo))

if desequilibrio > Decimal("1.00"):
    # AVISO (não CRÍTICO) — pode ser apenas o resultado do período
```

### 3. Consistência de cada conta
**Regra:** Saldo Atual = Saldo Anterior + Débito − Crédito

```python
for conta in contas_analiticas:
    esperado  = conta.saldo_anterior + conta.debito - conta.credito
    diferenca = abs(esperado - conta.saldo_atual)
    if diferenca > TOLERANCIA:
        # AVISO — possível erro de leitura ou natureza D/C incorreta
```

**Quando falha em massa:** indica problema no parser (ex: coluna errada).
**Quando falha em poucas contas:** pode ser ajuste de período ou conta diferencial.

## Precisão decimal — regra obrigatória

```python
from decimal import Decimal

# CERTO — Decimal preserva precisão monetária exata
saldo = Decimal("16435.07")
total = Decimal("0")
total += saldo   # sem erros de arredondamento

# ERRADO — float causa erros em centavos (IEEE 754)
saldo = 16435.07
total = 0.0
total += saldo   # pode virar 16435.070000000001
```

### Bug crítico — normalizar_decimal em models.py

O `field_validator` recebe valores de duas fontes:
1. Strings do PDF: `'1.234,56D'` → normalizar separadores BR
2. Decimals dos parsers: `Decimal('16435.07')` → retornar direto

**BUG a NÃO reproduzir:**
```python
# ERRADO — destrói o separador decimal quando recebe Decimal
s = str(Decimal('16435.07'))   # → '16435.07'
s = s.replace('.', '')          # → '1643507'  ← multiplicou por 100!
```

**CERTO — verificar tipo antes de normalizar:**
```python
def normalizar_decimal(cls, v):
    if isinstance(v, Decimal):
        return v          # já correto — não toca
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    # apenas para strings do PDF:
    s = str(v).strip().rstrip("DCdc").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")  # BR: 1.234,56 → 1234.56
    elif "," in s:
        s = s.replace(",", ".")
    return Decimal(s) if s else Decimal("0")
```

## Severidades dos erros

| Severidade | Quando usar | Impacto |
|---|---|---|
| `CRITICO` | Partidas dobradas quebradas | Bloqueia confiança no relatório |
| `AVISO` | Conta individual inconsistente, desequilíbrio patrimonial | Registra para revisão |
| `INFO` | Contas sem movimento, informações gerais | Apenas informativo |

## Como interpretar os avisos no Excel

Na aba **VALIDAÇÃO** do Excel gerado:

- **Vermelho (CRÍTICO):** revisar o mapeamento ou o PDF — algo está errado
- **Amarelo (AVISO):** pode ser normal (ex: ajuste de período, conta diferencial)
  verificar se a diferença é grande (> R$ 1,00) ou apenas centavos

## Tolerâncias configuráveis

```python
# em validator.py
TOLERANCIA_CENTAVOS = Decimal("0.05")  # por conta individual
TOLERANCIA_EQUACAO  = Decimal("1.00")  # para equação patrimonial
```

Aumentar tolerância só se o sistema contábil do cliente usar arredondamentos
diferentes. Não aumentar acima de R$ 0,10 para contas individuais.

## Checklist de validação ao desenvolver novo parser

```python
from core.parsers import get_parser
from core.validator import validar_balancete
from core.models import Balancete

# 1. Extrair
parser = get_parser("layout_id")
contas = parser.parsear("balancete.pdf")
print(f"Contas: {len(contas)} ({sum(1 for c in contas if c.tipo.value=='analitica')} analíticas)")

# 2. Montar balancete
b = Balancete(empresa="Teste", cnpj="", periodo_inicio="",
              periodo_fim="", sistema_contabil="x", layout_detectado="x",
              contas=contas)

# 3. Validar
erros = validar_balancete(b)
criticos = [e for e in erros if e.severidade.value == "CRITICO"]
avisos   = [e for e in erros if e.severidade.value == "AVISO"]
print(f"Críticos: {len(criticos)} | Avisos: {len(avisos)}")

# 4. Meta: 0 críticos, avisos < 10% das contas
# Se avisos > 10%: provavelmente natureza D/C está sendo inferida errado
```
