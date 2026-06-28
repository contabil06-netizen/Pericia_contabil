---
name: pericia-pdf-parser
description: >
  Use este skill sempre que precisar criar, corrigir, depurar ou melhorar
  qualquer parser de PDF do projeto de perícia contábil. Triggers incluem:
  novo sistema contábil não reconhecido, valores extraídos errados, contas
  faltando na extração, regex não batendo, layout desconhecido, PDF de novo
  cliente, taxa de extração baixa, ou qualquer menção a parsers, layouts,
  pdfplumber, detector, extração de contas, balancete PDF.
  SEMPRE use este skill antes de criar ou editar parsers.py ou detector.py.
---

# Skill: Parser de PDF de Balancetes Contábeis

## Contexto do projeto
Sistema de extração de dados de balancetes em PDF para perícia contábil.
Os PDFs são gerados por sistemas contábeis terceiros — cada sistema tem
uma estrutura diferente. Leia o CLAUDE.md para contexto completo.

## Como o pdfplumber funciona

```python
import pdfplumber

with pdfplumber.open("balancete.pdf") as pdf:
    for page in pdf.pages:
        texto = page.extract_text()   # string com todo o texto da página
        # O texto preserva colunas via espaçamento proporcional à posição X
        # Cada linha do PDF vira uma linha de texto separada por \n
```

**Limitação crítica:** só funciona em PDFs com camada de texto (gerados
por software). PDFs escaneados (fotos/imagens) retornam string vazia ou
com muito pouco texto. Detectar: `if len(texto.strip()) < 50: # provavelmente scan`

## Diagnóstico antes de criar/corrigir qualquer parser

Execute sempre este diagnóstico primeiro:

```python
import pdfplumber, re

with pdfplumber.open("balancete.pdf") as pdf:
    npages = len(pdf.pages)
    texto_p1 = pdf.pages[0].extract_text() or ""

# 1. Verificar se tem texto
print(f"Páginas: {npages} | Texto extraível: {len(texto_p1.strip()) > 50}")

# 2. Ver as primeiras 15 linhas (cabeçalho e primeiros dados)
linhas = [l.strip() for l in texto_p1.split('\n') if l.strip()]
for i, l in enumerate(linhas[:15]):
    print(f"L{i+1:02d}: {repr(l[:100])}")

# 3. Ver linhas típicas de dados (com números)
linhas_dados = [l for l in linhas if re.search(r'\d[\d\.,]{3,}', l)]
print("\nLinhas com dados numéricos (primeiras 8):")
for l in linhas_dados[:8]:
    print(f"  {repr(l[:100])}")
```

## Os 5 layouts existentes e seus padrões

### 1. `socium` — Socium / Gestão Empresarial
```
Identificador: "CONTABILIDADE E GESTAO EMPRESARIAL SOCIUM"
Colunas: ID_interno · Código · Descrição · Sal.Ant · Débito · Crédito · Sal.Atual
Exemplo: '5 1.1.1.01.001 CAIXA GERAL 15.084,05D 3.467,17 2.116,15 16.435,07D'
Sintética: nível < 5 segmentos (ex: 1.1 = sintética; 1.1.1.01.001 = analítica)
Natureza D/C: sufixo na coluna Sal.Atual ('D' ou 'C')
```

### 2. `hierarquico_5niveis` — Balancete de Verificação (Statera)
```
Identificador: "BALANCETE DE VERIFICACAO"
Colunas: Conta · [num_interno] · Descrição · Sal.Ant · Débito · Crédito · Sal.Atual
Exemplo analítica: '1.1.1.1.01.00001 1 Caixa Geral 0,00 30.830,00 30.830,00 0,00'
Exemplo sintética: '1.1.1 Disponivel 628,93 89.337,02 89.858,67 107,28'
Analítica: código termina em .00XXX (5 dígitos)
Passivo: saldo negativo → inverter sinal e marcar como Credor
```

### 3. `contacerta` — ContaCerta Organização Contábil
```
Identificador: "CONTACERTA ORGANIZACAO CONTABIL"
Colunas: [id]Descrição · Sal.Ant · Débito · Crédito · Sal.Atual
Exemplo: '[42]Caixa 57.401,29D 0,00 13.809,31 43.591,98D'
Sintética: palavras-chave como ATIVO, PASSIVO, CIRCULANTE, BANCOS, etc.
Analítica: tudo que tem os 4 valores numéricos na linha
```

### 4. `somar` — Somar Contabilidade
```
Identificador: "SOMAR CONTABILIDADE"
Colunas: Descrição · [id] · Sal.Ant · Débito · Crédito · Sal.Atual
Analítica: 'Caixa - [5] 8.757,41D 55.341,32 27.317,11 36.781,62D'
Sintética (=): '=CAIXA 8.757,41D 55.341,32 27.317,11 36.781,62D'
Grupo sem valores: 'CAIXA - [4]'  ← ignorar
Total (=T o t a l): ignorar — redundante com as sintéticas
```

### 5. `ctc007` — Sistema CTC007
```
Identificador: "CTC007"
Colunas (9): Conta · C.Res · [C.C D/C] · D/C · Nome · Sal.Ant · Déb · Créd · Mov · Sal.Atual
Com CC: '1.1.01.01.002 0007 D-BAN D CAIXA GERAL 259.439,58 0,00 0,00 0,00 259.439,58'
Sem CC: '1.1.01.02 0010 D BANCOS CONTA MOVIMENTO 2.334,88 398.204,80 370.268,31 27.936,49 30.271,37'
ATENÇÃO: usar a 9ª coluna como Sal.Atual — NÃO calcular via Sal.Ant + Mov
Analítica: nível >= 4 segmentos no código
```

## Regex de referência para cada layout

```python
# SOCIUM (7 grupos)
RE_SOCIUM = re.compile(
    r'^(\d+)\s+([\d\.]+)\s+(.+?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s*$'
)

# VERIFICAÇÃO — analítica (código termina em .00XXX)
RE_VER_ANALITICA = re.compile(
    r'^([\d\.]+\.\d{5})\s+(\d+)\s+(.+?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s*$'
)

# CONTACERTA (4 valores após [id])
RE_CC = re.compile(
    r'^\s*\[(\d+)\](.+?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+[DC]?)\s*$'
)

# SOMAR — analítica
RE_SOMAR_A = re.compile(
    r'^(.+?)\s+-\s+\[(\d+)\]\s+([\d\.,]+[DC]?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+[DC]?)\s*$'
)

# CTC007 com modificador CC (D-BAN, D-REC, D-PAG) — 9 colunas
RE_CTC_COM_CC = re.compile(
    r'^([\d\.]+)\s+(\d+)\s+([\w\-]+)\s+([DC])\s+(.+?)'
    r'\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,\-]+)\s+([\d\.,]+)\s*$'
)

# CTC007 sem modificador CC — 9 colunas
RE_CTC_SEM_CC = re.compile(
    r'^([\d\.]+)\s+(\d+)\s+([DC])\s+(.+?)'
    r'\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,\-]+)\s+([\d\.,]+)\s*$'
)
```

## Função _limpar_valor — conversão de valores monetários

```python
from decimal import Decimal
from core.models import NaturezaSaldo

def _limpar_valor(texto: str) -> tuple[Decimal, NaturezaSaldo]:
    """
    Converte string do PDF para Decimal + NaturezaSaldo.
    Suporta: '1.234,56D', '1.234,56C', '-1.234,56', '1234.56', '0,00'
    """
    if not texto or not texto.strip():
        return Decimal("0"), NaturezaSaldo.ZERO

    s = texto.strip()
    natureza = NaturezaSaldo.ZERO

    # Detectar D/C pelo sufixo
    if s.endswith("D"):
        natureza = NaturezaSaldo.DEVEDOR
        s = s[:-1].strip()
    elif s.endswith("C"):
        natureza = NaturezaSaldo.CREDOR
        s = s[:-1].strip()
    elif s.startswith("-"):
        natureza = NaturezaSaldo.CREDOR
        s = s[1:].strip()

    # Normalizar separadores BR: '1.234,56' → '1234.56'
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        val = Decimal(s)
        if natureza == NaturezaSaldo.ZERO and val != 0:
            natureza = NaturezaSaldo.DEVEDOR
        return val, natureza
    except Exception:
        return Decimal("0"), NaturezaSaldo.ZERO
```

## Como criar um parser para novo sistema

1. Executar o diagnóstico acima no PDF do novo cliente
2. Identificar padrão das linhas (quantas colunas, separador, presença de D/C)
3. Escrever regex testando com linhas reais antes de colocar no código
4. Criar classe herdando de `BaseParser`
5. Registrar na factory `_PARSERS` em `parsers.py`
6. Adicionar identificador em `detector.py`

```python
class ParserNovoSistema(BaseParser):
    RE_LINHA = re.compile(r'...')  # regex calibrado para o layout

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas = []
        for texto in self._extrair_texto_paginas(pdf_path):
            for linha in texto.split("\n"):
                conta = self._parsear_linha(linha.strip())
                if conta:
                    contas.append(conta)
        return contas

    def _parsear_linha(self, linha: str) -> ContaContabil | None:
        m = self.RE_LINHA.match(linha)
        if not m:
            return None
        # extrair grupos, chamar _limpar_valor, retornar ContaContabil
        ...
```

## Como medir taxa de extração

```python
import pdfplumber, re
from core.parsers import get_parser

pdf = "balancete.pdf"
layout = "socium"

# Contar linhas com dados no PDF
with pdfplumber.open(pdf) as p:
    linhas_com_dados = sum(
        1 for page in p.pages
        for l in (page.extract_text() or "").split("\n")
        if re.search(r'\d[\d\.,]{4,}', l)
    )

# Contar contas extraídas
parser = get_parser(layout)
contas = parser.parsear(pdf)

print(f"Linhas com dados no PDF: {linhas_com_dados}")
print(f"Contas extraídas: {len(contas)}")
print(f"Taxa: {len(contas)/linhas_com_dados*100:.1f}%")
# Esperado: > 85% para parsers bem calibrados
```

## Validação de valores extraídos

```python
from decimal import Decimal

# Verificar conta específica contra valor conhecido do PDF
def verificar_conta(contas, codigo, saldo_esperado, natureza_esperada):
    c = next((c for c in contas if c.codigo == codigo), None)
    if not c:
        print(f"AUSENTE: {codigo}")
        return False
    ok_val = abs(c.saldo_atual - Decimal(str(saldo_esperado))) < Decimal("0.05")
    ok_nat = c.natureza.value == natureza_esperada
    status = "OK" if (ok_val and ok_nat) else "ERRO"
    print(f"[{status}] {codigo} | val={float(c.saldo_atual):.2f} (esp={saldo_esperado}) nat={c.natureza.value}")
    return ok_val and ok_nat
```

## Erros comuns e soluções

| Problema | Causa provável | Solução |
|---|---|---|
| 0 contas extraídas | PDF escaneado (imagem) | Verificar com diagnóstico; informar ao usuário |
| Valores 100x maiores | Bug no normalizar_decimal (ver CLAUDE.md) | Verificar field_validator em models.py |
| Natureza sempre D | Regex não captura sufixo D/C | Verificar grupo do saldo_atual no regex |
| Taxa < 50% | Regex muito restrito | Inspecionar linhas que não batem com `re.match` |
| Saldo Atual errado no CTC007 | Usando Sal.Ant + Mov em vez da 9ª coluna | Usar a última coluna diretamente |
| Descrição truncada | Regex `.+?` guloso demais ou de menos | Ajustar âncoras dos grupos numéricos |
