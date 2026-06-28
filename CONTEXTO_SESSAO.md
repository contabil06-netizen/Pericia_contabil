# CONTEXTO — Projeto PERITUS (Automação Contábil Pericial)

---

## Identidade do projeto

Sistema Python de automação para perícias contábeis em empresas sob Recuperação Judicial.
Extrai dados de balancetes em PDF e gera Excel com DRE estruturada, Balanço Patrimonial e indicadores.

**Pasta raiz do projeto:** `D:\Claude Perícia Contábil\pericia_contabil\`
**Escritório:** Inocêncio de Paula — Perícia Contábil

---

## Estrutura de pastas

```
D:\Claude Perícia Contábil\pericia_contabil\
│
├── main.py                        ← GUI (tkinter) com identidade visual do escritório
├── MODELO_DE_PERICIA.xlsx
│
├── core\
│   ├── __init__.py
│   ├── detector.py                ← detecta layout pelo texto do PDF
│   ├── engine.py                  ← orquestra o pipeline (processar / processar_lote)
│   ├── exporter.py                ← gera Excel (DRE, BP, validação, extração)
│   ├── mapper.py                  ← carrega YAML do cliente e mapeia contas
│   ├── models.py                  ← entidades Pydantic v2 (ContaContabil, Balancete, RelatorioFinal)
│   ├── parsers.py                 ← parsers por layout de PDF
│   └── validator.py               ← validações matemáticas (Ativo = Passivo + PL)
│
├── clientes\
│   ├── sagrados.yaml              ← COMPLETO e em produção
│   ├── statera.yaml               ← esqueleto, DRE a completar
│   └── rodomelo.yaml              ← esqueleto
│
├── output\
│   └── [empresa]\[empresa]_[competência]_(...).xlsx
│
└── Testes Balancetes Reais\
    ├── Bom Pastor\   (4 PDFs — SOMAR Contabilidade)
    ├── Carolina\     (4 PDFs — CONTACERTA)
    ├── SAGRADOS\     (4 PDFs — CTC007) ← já funcionando
    └── Statera Transportes\ (2 PDFs — layout hierárquico simples)
```

> **Atenção:** há uma pasta `Automação de fechamento contábil + criação de planilha Excel\`
> dentro de `pericia_contabil\` com cópias antigas de `core\` e `clientes\` e os Excels
> gerados pela GUI. Os arquivos ativos que devem ser editados são os da **raiz**
> (`pericia_contabil\core\` e `pericia_contabil\clientes\`).

---

## Parsers implementados (`core/parsers.py` — 605 linhas)

| Classe | Layout | Identificador YAML |
|--------|--------|--------------------|
| `ParserSocium` | Socium — hierárquico com D/C inline | `"socium"` |
| `ParserGenerico` | Fallback genérico | `"generico"` |
| `ParserContaCerta` | CONTACERTA — `[ID]Descrição` colchete no início | `"contacerta"` |
| `ParserSomar` | SOMAR Contabilidade — `Descrição - [ID]` colchete no fim | `"somar"` |
| `ParserCTC007` | CTC007 — 10 colunas, hierárquico 5 segmentos | `"ctc007"` |
| `ParserVerificacao` | Hierárquico simples 6 segmentos (Statera) | `"verificacao"` |

**Função de roteamento:** `get_parser(layout_id: str) -> BaseParser`

### Detalhes críticos do ParserCTC007 (Sagrados)
- `NIVEL_ANALITICO = 5` — contas com 5 segmentos separados por ponto são analíticas
- Regex `RE_COM_CC` e `RE_SEM_CC` usam `[-\d\.,]+` para aceitar saldos negativos
- `_mov(conta)`: DEVEDOR → `debito - credito`; CREDOR → `credito - debito`
- Conta `3.1.01.04.001` (CMV) é CREDOR com movimento devedor → `_mov` retorna negativo

---

## Exporter (`core/exporter.py` — 918 linhas) — funções-chave

| Função | Papel |
|--------|-------|
| `exportar_sem_modelo(relatorio, output_dir, relatorio_anterior)` | Entry point principal |
| `_construir_dre(wb, rel, rel_ant, ...)` | Gera aba DRE a partir de `dre_estrutura` do YAML |
| `_processar_bloco_dre(ws, row, grupo, rel, rel_ant, ...)` | Renderiza cada bloco da DRE recursivamente |
| `_construir_acc(rel)` | Monta acumulador `{tipo: Decimal}` para fórmulas de subtotal |
| `_acc_estrutura(estrutura, rel, acc)` | Percorre `dre_estrutura` e acumula movimentos |
| `_somar_mov_grupo(contas, prefixos, excluir)` | Soma `_mov` das analíticas por prefixo, com exclusões |
| `_contas_do_grupo(todas, prefixos, excluir)` | Filtra contas por prefixo com exclusões |
| `_calcular_rec_liq(rel)` | Receita Líquida = Receita - Deduções (base para AV%) |
| `_avaliar_formula(formula, acc)` | Avalia fórmula string como `"receita - deducao - custo"` |
| `_av(valor, base)` | Análise vertical — retorna `None` se base=0 (exibe `-`) |

### Convenção de sinais na DRE

```python
_TIPO_SINAL = {
    "receita": 1, "deducao": -1, "custo": -1, "despesa": -1,
    "outra_receita": 1, "outra_despesa": -1, "depreciacao": 1,
    "receita_financeira": 1, "despesa_financeira": -1, "provisao_ir": -1,
}
```

**`sinal_acc`** no YAML: campo opcional nos subgrupos. Quando `-1`, inverte o `_mov`
antes de acumular. Necessário para contas CREDOR em que `_mov` é negativo mas a
fórmula espera valor positivo.

Exemplo: CMV (`3.1.01.04`) é CREDOR → `_mov = -234.354` → `sinal_acc: -1` →
`acc["custo"] = +234.354` → `receita - custo = 208K - 234K = -26K` ✓

**`excluir_prefixos`** no YAML: lista de prefixos a excluir de um subgrupo.
Implementado em `_somar_mov_grupo` e `_contas_do_grupo`.

---

## Formato do YAML de cliente

```yaml
cliente: "Nome da Empresa Ltda"
cnpj: "XX.XXX.XXX/0001-XX"
layout: "ctc007"   # determina qual parser usar

balanco:
  ativo_circulante:
    disponibilidades:
      codigos: ["1.1.01"]
  # ... demais grupos do BP

dre_estrutura:
  receita_bruta:
    label: "01. Receita Bruta de Vendas"
    tipo: receita          # tipos: receita, deducao, custo, despesa, outra_receita,
    prefixos:              # outra_despesa, depreciacao, receita_financeira,
      - "3.1.01.01"        # despesa_financeira, provisao_ir, agrupador, subtotal

  rec_op_liq:
    label: "03. Receita Operacional Liquida"
    tipo: subtotal
    formula: "receita - deducao"

  custos_operacionais:
    label: "04. (-) Custos Operacionais"
    tipo: agrupador
    subgrupos:
      cmv:
        label: "    CMV"
        tipo: custo
        sinal_acc: -1              # opcional — inverte _mov para acumulação
        prefixos: ["3.1.01.04"]
        excluir_prefixos: []       # opcional — exclui contas específicas
```

---

## YAML do Sagrados — estado atual (completo e funcionando)

- `clientes/sagrados.yaml` — 258 linhas, DRE de 14 linhas estruturada
- Validado com 4 balancetes (Jan–Abr 2026)
- Resultados verificados: Lucro Bruto JAN = -26.354 ✓, CMV = 234.354 ✓

---

## Análise dos 4 PDFs disponíveis para novos YAMLs

### Sagrados (CTC007) — JÁ PRONTO
- Parser: `ParserCTC007` ✅
- YAML: `clientes/sagrados.yaml` ✅

### Statera Transportes — PENDENTE
- **Sistema:** desconhecido (provável Alterdata/Domínio)
- **Parser:** `ParserVerificacao` (hierárquico 6 segmentos, sem D/C suffix)
- **Layout colunas:** `Conta  ID  Descrição  SaldoAnterior  Débito  Crédito  SaldoAtual`
- **Código analítico:** `x.x.x.x.xx.00001` (6 segmentos, último com 5 dígitos zerados)
- **DRE extraída do PDF:**
  - Classe 3 = Resultado (não 4.x como Sagrados)
  - `3.1.1.1.02` = Receita Bruta (Vendas de Serviços)
  - `3.1.1.3.01` = Deduções (COFINS, PIS, ICMS, CPP)
  - `3.1.2.1.02` = Custos dos Serviços Prestados
  - `3.1.2.3.01` = Despesas c/ Pessoal (Adm)
  - `3.1.2.3.02` = Serviços de Terceiros
  - `3.1.2.3.03` = C/ Ocupação
  - `3.1.2.3.04` = Serviços e Utilidades
  - `3.1.2.3.08` = Despesas Gerais
  - `3.1.3.1.01` = Despesas Financeiras
  - Depreciação: `3.1.2.1.02.00018` e `3.1.2.3.08.00007`
- **Setor:** Transportes

### Bom Pastor (SOMAR Contabilidade) — PENDENTE
- **Sistema:** SOMAR CONTABILIDADE S/S LTDA
- **Parser:** `ParserSomar` (ID sequencial entre colchetes no FIM: `Descrição - [268]`)
- **Layout colunas:** `Descrição - [ID]  SaldoAnt D/C  Débito  Crédito  SaldoAtual D/C`
- **Código de conta:** ID sequencial — NÃO hierárquico. O mapeamento no YAML usa IDs, não prefixos
- **DRE extraída do PDF:**
  - `[329]` = Vendas a Prazo de Produtos (Receita)
  - `[335–342]` = Deduções (ICMS, PIS, COFINS, Cancelamentos)
  - `[359]` = Custo dos Produtos Vendidos
  - `[273–280]` = Despesas Adm (Salários, Encargos, FGTS)
  - `[282–317]` = Despesas Adm (Escrituração, Material, Seguros, Impostos, etc.)
  - `[309–312]` = Despesas Financeiras
  - `[382–669]` = Despesas Comerciais
- **Setor:** Distribuição de papel (Bom Pastor Papeis Ltda)
- **CNPJ:** 16.772.642/0001-49

### Carolina (CONTACERTA) — PENDENTE
- **Sistema:** CONTACERTA ORGANIZACAO CONTABIL LTDA
- **Parser:** `ParserContaCerta` (ID sequencial entre colchetes no INÍCIO: `[1925]Descrição`)
- **Layout colunas:** `[ID]Descrição  SaldoAnt D/C  Débito  Crédito  SaldoAtual D/C`
- **Código de conta:** ID sequencial — NÃO hierárquico
- **DRE extraída do PDF:**
  - `[2037]` = Receitas Diversas (única receita — PF agropecuária)
  - `[2128–2184]` = Custo Fabril / Gastos c/ Pessoal de Fabricação
  - `[2443–2450]` = Despesas Operacionais
  - `[2625–2632]` = Administrativas e Gerais (energia, combustível, fertilizantes, etc.)
  - `[2884–5694]` = Despesas Financeiras (IOF)
  - `[2940–9461]` = Despesas Tributárias
- **Setor:** Pessoa Física / Agropecuária (fazenda)
- **CNPJ:** 09.936.527/607 (CPF)

---

## Decisão arquitetural: YAML por empresa, não por layout

Cada empresa usa um ERP diferente **E** tem plano de contas único.
Bom Pastor e Carolina usam IDs sequenciais não portáveis entre empresas.
Statera e Sagrados usam hierarquia com pontos, mas prefixos completamente distintos.

**Conclusão:** 4 YAMLs individuais, sem compartilhamento de layout YAML.
A arquitetura de "layout compartilhado" seria útil no futuro se o escritório
atender outras empresas CTC007 com o mesmo plano de contas padronizado.

---

## Regras críticas de desenvolvimento

1. **Nunca usar a ferramenta `Edit` para edições em `exporter.py` ou `parsers.py`** —
   ela trunca arquivos grandes. Usar Python `str.replace()` via `mcp__workspace__bash`:
   ```python
   with open(path, 'r') as f: src = f.read()
   src = src.replace(old, new)
   with open(path, 'w') as f: f.write(src)
   ```

2. **`_processar_bloco_dre` e `_construir_acc` devem estar SEMPRE sincronizados** —
   ambos acumulam `acc[tipo]` de forma independente. Se mudar a lógica de acumulação
   em um, mudar no outro também.

3. **Nunca filtrar analíticas por saldo zero** — todas devem aparecer na DRE,
   inclusive as com movimento zero.

4. **`sinal_acc` deve ser propagado em ambos os loops** de `_processar_bloco_dre`:
   o loop de acumulação (`acc[sub_tipo] += sinal_acc * mov`) e o loop de cálculo
   de `total_a_d` para o cabeçalho do agrupador.

5. **Outputs em** `output\[empresa]\` (pipeline) ou na pasta
   `Automação de fechamento contábil + criação de planilha Excel\` (GUI).

---

## Próximas tarefas

- [ ] Criar `clientes/statera.yaml` completo com DRE de 14 linhas
- [ ] Criar `clientes/bom_pastor.yaml` completo (parser usa IDs `[xxx]`)
- [ ] Criar `clientes/carolina.yaml` completo (parser usa IDs `[xxx]`)
- [ ] Verificar se `ParserSomar` e `ParserContaCerta` estão funcionais com os PDFs reais
- [ ] Verificar se `ParserVerificacao` parseia corretamente o balancete da Statera

---

## Caminho dos PDFs de teste

```
D:\Claude Perícia Contábil\pericia_contabil\Testes Balancetes Reais\
├── SAGRADOS\
│   ├── BALANCETE ANALITICO- SAGRADOS JANEIRO 2026.pdf
│   ├── BALANCETE ANALITICO - SAGRADOS FEVEREIRO 2026.pdf
│   ├── BALANCETE ANALITICO - SAGRADOS MARÇO 2026.pdf    ← nome com encoding: MARÃ_O
│   └── BALANCETE ANALITICO - SAGRADOS ABRIL 2026.pdf
├── Statera Transportes\
│   ├── Balancete 01.2026 - Sem Encerramento.pdf
│   └── Balancete 12.2025.pdf
├── Bom Pastor\
│   ├── Balancete Analitico sem encerramento Janeiro 2026.pdf
│   ├── Balancete Analitico sem encerramento Dezembro 2025.pdf
│   ├── Balancente Analitico sem encerramento Novembro 2025.pdf
│   └── Balancete Analitico sem Encerramento Fevereiro 2026.pdf
└── Carolina\
    ├── BALANCETE PF 01-26.pdf
    ├── BALANCETE PF 02-26.pdf
    ├── BALANCETE PF 03-26.pdf
    └── BALANCETE  PF 12-25.pdf
```
