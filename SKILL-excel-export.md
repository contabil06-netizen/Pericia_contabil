---
name: pericia-excel-export
description: >
  Use este skill sempre que precisar criar, corrigir ou melhorar a geração
  de arquivos Excel no projeto de perícia contábil. Triggers incluem:
  célula não preenchida, aba errada, Excel em branco, modelo não encontrado,
  formatação perdida, criar nova aba, adicionar coluna, gráfico, indicador,
  comparativo dois períodos, série histórica, fórmula Excel, openpyxl,
  exportar planilha, preencher modelo, MODELO_DE_PERICIA.xlsx.
  SEMPRE use este skill antes de editar exporter.py.
---

# Skill: Exportação Excel de Perícia Contábil

## Princípio fundamental

**Nunca sobrescrever o MODELO_DE_PERICIA.xlsx.**
Sempre copiar o modelo para `output/` com nome único antes de editar.

```python
import shutil
from datetime import datetime

ts = datetime.now().strftime("%Y%m%d_%H%M%S")
output_path = f"output/{empresa_slug}_{periodo}_{ts}.xlsx"
shutil.copy("MODELO_DE_PERICIA.xlsx", output_path)   # copia primeiro
wb = openpyxl.load_workbook(output_path)              # abre a cópia
```

## Estrutura do Excel gerado

O arquivo Excel de saída tem 4 abas:

| Aba | Conteúdo | Criada como |
|---|---|---|
| `BALANÇO` | Balanço Patrimonial preenchido | Cópia do modelo |
| `DRE` | Demonstração de Resultado preenchida | Cópia do modelo |
| `VALIDAÇÃO` | Erros e avisos da validação matemática | Criada do zero |
| `EXTRAÇÃO` | Todas as contas brutas do PDF | Criada do zero |

## Escrever em célula por referência

```python
def _escrever_celula(ws, referencia: str, valor):
    """
    Escreve valor em célula.
    Suporta: 'C8' ou 'BALANÇO!C8' (descarta o prefixo da aba)
    """
    if "!" in referencia:
        referencia = referencia.split("!")[-1]
    ws[referencia] = valor
```

**Importante:** escrever como `float`, não `Decimal` — openpyxl não aceita Decimal:
```python
_escrever_celula(ws, "C8", float(grupo.valor))   # correto
_escrever_celula(ws, "C8", grupo.valor)           # pode falhar
```

## Selecionar aba correta

```python
# Aba do Balanço (pode ter acento ou não dependendo do modelo)
def _obter_aba(wb, nomes_possiveis: list[str]):
    for nome in nomes_possiveis:
        if nome in wb.sheetnames:
            return wb[nome]
    return wb.active   # fallback para primeira aba

ws_balanco = _obter_aba(wb, ["BALANÇO", "BALANCO", "Balanço"])
ws_dre     = _obter_aba(wb, ["DRE", "dre", "Resultado"])
```

## Criar aba de VALIDAÇÃO

```python
from openpyxl.styles import Font, PatternFill, Alignment

COR_CRITICO  = "FFD7D7"   # vermelho claro
COR_AVISO    = "FFF3CD"   # amarelo claro
COR_CABEC    = "1a1a2e"   # azul escuro

def criar_aba_validacao(wb, erros):
    nome = "VALIDAÇÃO"
    if nome in wb.sheetnames:
        del wb[nome]
    ws = wb.create_sheet(nome)

    # Cabeçalho
    colunas = ["Severidade", "Conta", "Descrição", "Mensagem",
               "Valor Esperado", "Valor Encontrado", "Diferença"]
    for col, titulo in enumerate(colunas, 1):
        cel = ws.cell(row=1, column=col, value=titulo)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor=COR_CABEC)

    if not erros:
        ws.cell(row=2, column=1, value="✓ Nenhum erro encontrado")
        return

    for row, erro in enumerate(erros, 2):
        cor = COR_CRITICO if erro.severidade.value == "CRITICO" else COR_AVISO
        valores = [
            erro.severidade.value,
            erro.conta_codigo,
            erro.conta_descricao,
            erro.mensagem,
            float(erro.valor_esperado)   if erro.valor_esperado   else "",
            float(erro.valor_encontrado) if erro.valor_encontrado else "",
            float(erro.diferenca)        if erro.diferenca        else "",
        ]
        for col, val in enumerate(valores, 1):
            cel = ws.cell(row=row, column=col, value=val)
            cel.fill = PatternFill("solid", fgColor=cor)

    # Larguras úteis
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 70
```

## Criar aba de EXTRAÇÃO (dados brutos)

```python
def criar_aba_extracao(wb, balancete):
    nome = "EXTRAÇÃO"
    if nome in wb.sheetnames:
        del wb[nome]
    ws = wb.create_sheet(nome)

    # Metadados
    ws["A1"] = "Empresa";    ws["B1"] = balancete.empresa
    ws["A2"] = "CNPJ";       ws["B2"] = balancete.cnpj
    ws["A3"] = "Período";    ws["B3"] = f"{balancete.periodo_inicio} a {balancete.periodo_fim}"
    ws["A4"] = "Layout";     ws["B4"] = balancete.layout_detectado
    ws["A5"] = "Total contas";ws["B5"] = len(balancete.contas)

    # Tabela de contas (a partir da linha 8)
    headers = ["Código", "Descrição", "Tipo", "Natureza",
               "Sal. Anterior", "Débito", "Crédito", "Sal. Atual"]
    for col, h in enumerate(headers, 1):
        cel = ws.cell(row=7, column=col, value=h)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="1a1a2e")

    for row, conta in enumerate(balancete.contas, 8):
        ws.cell(row=row, column=1, value=conta.codigo)
        ws.cell(row=row, column=2, value=conta.descricao)
        ws.cell(row=row, column=3, value=conta.tipo.value)
        ws.cell(row=row, column=4, value=conta.natureza.value)
        ws.cell(row=row, column=5, value=float(conta.saldo_anterior))
        ws.cell(row=row, column=6, value=float(conta.debito))
        ws.cell(row=row, column=7, value=float(conta.credito))
        ws.cell(row=row, column=8, value=float(conta.saldo_atual))
        # formato monetário
        for col in range(5, 9):
            ws.cell(row=row, column=col).number_format = '#,##0.00'
        # zebra
        if row % 2 == 0:
            for col in range(1, 9):
                ws.cell(row=row, column=col).fill = PatternFill("solid", fgColor="F5F5F5")

    # Ajuste de colunas
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 50
```

## Comparativo dois períodos no mesmo Excel

Quando dois PDFs são processados (mês atual + mês anterior), o exporter
deve criar colunas adjacentes:

```python
def preencher_dois_periodos(wb, relatorio_atual, relatorio_anterior):
    ws = wb["BALANÇO"]

    for chave, grupo in relatorio_atual.grupos_balanco.items():
        if grupo.celula_excel:
            _escrever_celula(ws, grupo.celula_excel, float(grupo.valor))

    if relatorio_anterior:
        for chave, grupo in relatorio_anterior.grupos_balanco.items():
            if grupo.celula_excel:
                # Mês anterior vai na coluna D (ou conforme o modelo define)
                cel_anterior = grupo.celula_excel.replace("C", "D")
                _escrever_celula(ws, cel_anterior, float(grupo.valor))
```

**Nota:** as referências de colunas do mês anterior dependem do modelo Excel.
Verificar o MODELO_DE_PERICIA.xlsx para confirmar onde o mês anterior fica.

## Formato monetário e número

```python
from openpyxl.styles import numbers

# Monetário BR
ws["C8"].number_format = '#,##0.00'

# Percentual
ws["E8"].number_format = '0.00%'

# Inteiro
ws["F8"].number_format = '#,##0'
```

## Geração do nome do arquivo de saída

```python
import re
from datetime import datetime

def gerar_nome_arquivo(empresa: str, periodo_fim: str) -> str:
    """
    Gera nome limpo para o arquivo de saída.
    Ex: 'rodomelo_transportes_31-01-2026_20260115_143022.xlsx'
    """
    slug = re.sub(r'[^a-z0-9\s]', '', empresa.lower())
    slug = re.sub(r'\s+', '_', slug.strip())[:30]
    periodo = periodo_fim.replace("/", "-") if periodo_fim else "sem_data"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"output/{slug}_{periodo}_{ts}.xlsx"
```

## Erros comuns ao exportar

| Erro | Causa | Solução |
|---|---|---|
| `KeyError: 'BALANÇO'` | Aba tem nome diferente no modelo | Usar `_obter_aba()` com lista de nomes alternativos |
| `ValueError: Decimal` | Passou Decimal para openpyxl | Converter com `float()` antes |
| Células em branco | `celula_excel` não está no YAML | Verificar YAML do cliente |
| Modelo não encontrado | `MODELO_DE_PERICIA.xlsx` ausente | Usar `_criar_workbook_vazio()` como fallback |
| Formatação perdida | Abriu o modelo com `data_only=True` | Não passar esse parâmetro |
