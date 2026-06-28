---
name: pericia-yaml-cliente
description: >
  Use este skill sempre que precisar criar, editar, corrigir ou explicar
  arquivos YAML de configuração de clientes do projeto de perícia contábil.
  Triggers incluem: cadastrar novo cliente, mapear contas, célula errada,
  prefixo de conta incorreto, grupo retornou zero, cliente novo no sistema,
  como mapear DRE, como mapear Balanço, sinal invertido, deduções, prejuízo.
  SEMPRE use este skill antes de criar ou editar qualquer arquivo em clientes/.
---

# Skill: Configuração YAML de Clientes

## O que é o arquivo YAML de cliente

É o único arquivo que precisa ser criado para cada novo cliente.
Ele instrui o sistema: "some as contas com esses prefixos e coloque
o resultado na célula X da planilha modelo."

Nenhuma linha de código Python precisa ser escrita para um novo cliente.

## Estrutura completa comentada

```yaml
# ─────────────────────────────────────────────────────────────
# IDENTIFICAÇÃO DO CLIENTE
# ─────────────────────────────────────────────────────────────
cliente: "Nome Completo da Empresa Ltda"
cnpj: "00.000.000/0001-00"

# layout: qual dos parsers usar para ler os PDFs desta empresa
# Valores possíveis: socium | hierarquico_5niveis | contacerta | somar | ctc007
layout: "socium"

# ─────────────────────────────────────────────────────────────
# MAPEAMENTO DO BALANÇO PATRIMONIAL
# Estrutura: balanco > seção > nome_do_grupo > configuração
# ─────────────────────────────────────────────────────────────
balanco:

  ativo_circulante:

    disponibilidades:
      celula_excel: "C8"         # célula da aba BALANÇO no modelo
      codigos: ["1.1.1"]         # soma todas as contas que começam com 1.1.1
      # obrigatorio: true        # descomente para avisar se retornar zero

    clientes:
      celula_excel: "C9"
      codigos: ["1.1.2"]

    contas_a_receber:
      celula_excel: "C10"
      codigos: ["1.1.3", "1.1.5"]   # pode listar vários prefixos

    estoques:
      celula_excel: "C11"
      codigos: ["1.1.4"]

    impostos_a_recuperar:
      celula_excel: "C12"
      codigos: ["1.1.6", "1.1.7"]

  ativo_nao_circulante:

    depositos_judiciais:
      celula_excel: "C17"
      codigos: ["1.2.1"]

    imobilizado:
      celula_excel: "C18"
      codigos: ["1.2.3"]

  passivo_circulante:

    fornecedores:
      celula_excel: "X8"
      codigos: ["2.1.3"]

    emprestimos_financiamentos_cp:
      celula_excel: "X9"
      codigos: ["2.1.1"]

    obrigacoes_trabalhistas:
      celula_excel: "X10"
      codigos: ["2.1.5"]

    obrigacoes_tributarias:
      celula_excel: "X11"
      codigos: ["2.1.4"]

  passivo_nao_circulante:

    emprestimos_financiamentos_lp:
      celula_excel: "X16"
      codigos: ["2.2.1", "2.2.2"]

    credores_recuperacao_judicial:
      celula_excel: "X17"
      codigos: ["2.1.9", "2.2.3"]     # passivos da RJ ficam aqui

  patrimonio_liquido:

    capital_social:
      celula_excel: "X21"
      codigos: ["2.3.1", "2.4.1"]

    prejuizos_acumulados:
      celula_excel: "X22"
      codigos: ["2.3.5", "2.4.3"]
      sinal: -1     # IMPORTANTE: prejuízo é devedor — entra negativo no PL

# ─────────────────────────────────────────────────────────────
# MAPEAMENTO DA DRE
# Estrutura: dre > nome_do_grupo > configuração
# ─────────────────────────────────────────────────────────────
dre:

  receita_bruta:
    celula_excel: "C5"
    codigos: ["4.1.1", "3.1.1.1"]

  deducoes_receita:
    celula_excel: "C6"
    codigos: ["4.1.2", "3.1.1.3"]
    sinal: -1       # deduções reduzem a receita

  custo_servicos_produtos:
    celula_excel: "C8"
    codigos: ["3.1.2.1", "4.8"]
    sinal: -1

  despesas_pessoal:
    celula_excel: "C11"
    codigos: ["3.2.2.01", "3.1.2.3.01"]
    sinal: -1

  despesas_administrativas:
    celula_excel: "C12"
    codigos: ["3.2.2.04", "3.1.2.3"]
    sinal: -1

  despesas_financeiras:
    celula_excel: "C13"
    codigos: ["3.2.2.05", "3.1.3", "3.4.02"]
    sinal: -1

  receitas_financeiras:
    celula_excel: "C14"
    codigos: ["3.4.01", "4.2"]

  outras_receitas:
    celula_excel: "C15"
    codigos: ["3.5", "4.2.1"]
```

## Campos disponíveis por grupo

| Campo | Tipo | Obrigatório | Descrição |
|---|---|---|---|
| `celula_excel` | string | Sim | Referência de célula: `"C8"` ou `"BALANÇO!C8"` |
| `codigos` | lista | Sim* | Prefixos de código de conta a somar |
| `descricoes` | lista | Sim* | Alternativa: busca por substring na descrição |
| `sinal` | -1 ou 1 | Não | `-1` para inverter (deduções, custos, prejuízos). Padrão: 1 |
| `obrigatorio` | bool | Não | Gera aviso se o grupo retornar zero. Padrão: false |

*`codigos` OU `descricoes` — pelo menos um obrigatório por grupo.

## Quando usar `descricoes` em vez de `codigos`

Use `descricoes` quando o plano de contas do cliente não tem hierarquia
numérica clara, ou quando quer agrupar por nome independente do código:

```yaml
disponibilidades:
  celula_excel: "C8"
  descricoes: ["CAIXA", "BANCO", "APLICAÇÃO FINANCEIRA"]
  # busca case-insensitive e parcial na descrição de cada conta
```

## Fluxo para cadastrar um novo cliente

### Passo 1 — Extração bruta (sem cliente)
Processar o PDF sem selecionar cliente no dropdown.
O Excel gerado terá uma aba **EXTRAÇÃO** com todas as contas:
```
Código             | Descrição              | Tipo      | Natureza | Sal.Atual
1.1.1.01.001      | CAIXA GERAL            | analitica | D        | 16.435,07
1.1.1.02.001      | SICOOB - C/C 150.022-8 | analitica | D        |    43,87
2.1.3.01.104      | ACTUAL PNEUS LTDA      | analitica | C        |  2.776,67
```

### Passo 2 — Identificar os prefixos de cada grupo
Com a aba EXTRAÇÃO aberta, olhar os códigos de cada seção:
- Disponibilidades: todos os `1.1.1.*` → `codigos: ["1.1.1"]`
- Clientes: todos os `1.1.2.*` → `codigos: ["1.1.2"]`
- Fornecedores: todos os `2.1.3.*` → `codigos: ["2.1.3"]`

### Passo 3 — Identificar as células do modelo Excel
Abrir MODELO_DE_PERICIA.xlsx e anotar em qual célula fica cada linha:
- Disponibilidades → célula C8 da aba BALANÇO
- Clientes → célula C9
- etc.

### Passo 4 — Criar o YAML
Copiar o template acima, preencher os campos e salvar em:
`pericia_contabil/clientes/nome_empresa.yaml`

O nome do arquivo (sem .yaml) é o `--cliente` passado no processamento.

### Passo 5 — Testar
Processar com o cliente e verificar a aba EXTRAÇÃO do Excel gerado.
Se algum grupo ficou zerado, o prefixo provavelmente está errado —
confirmar na aba EXTRAÇÃO quais são os códigos reais.

## Casos especiais

### Empresa com Recuperação Judicial (Classe III, I, IV)
O Bom Pastor tem contas classificadas por classe de credores:
```yaml
passivo_nao_circulante:
  credores_quirografarios_III:
    celula_excel: "X19"
    codigos: ["14626248"]    # código direto da conta sintética
  credores_trabalhistas_I:
    celula_excel: "X20"
    codigos: ["90920"]
```

### Empresa com Passivo em sinal negativo (Statera)
A Statera registra passivo com saldo negativo no PDF:
```yaml
passivo_circulante:
  fornecedores:
    celula_excel: "X8"
    codigos: ["2.1.1.1"]
    # O parser já normaliza o sinal negativo para credor
    # Não precisa de sinal: -1 aqui
```

### Prejuízo acumulado (sempre sinal: -1)
```yaml
patrimonio_liquido:
  prejuizos_acumulados:
    celula_excel: "X22"
    codigos: ["2.3.5", "2.4.3"]
    sinal: -1    # SEMPRE -1 para prejuízos
```

### DRE com deduções e abatimentos
```yaml
dre:
  receita_bruta:
    celula_excel: "C5"
    codigos: ["4.1.1"]
    # sem sinal — receita entra positiva

  deducoes:
    celula_excel: "C6"
    codigos: ["4.1.2"]
    sinal: -1    # SEMPRE -1 para deduções
    # O modelo provavelmente já subtrai esta linha — ajustar conforme o modelo
```

## Validação do YAML

Após criar o YAML, verificar:
1. Todos os grupos obrigatórios têm `celula_excel` preenchida
2. Os prefixos em `codigos` existem na aba EXTRAÇÃO do Excel bruto
3. Grupos com `sinal: -1` fazem sentido (deduções, custos, prejuízos)
4. Nenhuma conta analítica está mapeada em dois grupos diferentes
   (causaria dupla contagem no balanço)
