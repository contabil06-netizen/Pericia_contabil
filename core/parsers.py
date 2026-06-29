"""
Parsers especializados por layout de sistema contábil.
Cada parser implementa a interface BaseParsear e retorna
uma lista de ContaContabil normalizada.
"""
from __future__ import annotations
import re
from abc import ABC, abstractmethod
from decimal import Decimal

import pdfplumber

from core.models import ContaContabil, NaturezaSaldo, TipoConta


# ---------------------------------------------------------------------------
# Utilitários compartilhados
# ---------------------------------------------------------------------------

def _limpar_valor(texto: str) -> tuple[Decimal, NaturezaSaldo]:
    """
    Converte string de valor contábil para Decimal e NaturezaSaldo.
    Suporta: '1.234,56D', '1.234,56C', '-1.234,56', '1234.56'
    """
    if not texto or not texto.strip():
        return Decimal("0"), NaturezaSaldo.ZERO

    s = texto.strip()

    # Detectar natureza pelo sufixo D/C
    natureza = NaturezaSaldo.ZERO
    if s.endswith("D"):
        natureza = NaturezaSaldo.DEVEDOR
        s = s[:-1].strip()
    elif s.endswith("C"):
        natureza = NaturezaSaldo.CREDOR
        s = s[:-1].strip()
    elif s.startswith("-"):
        natureza = NaturezaSaldo.CREDOR
        s = s[1:].strip()

    # Normalizar separadores: "1.234,56" → "1234.56"
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        valor = Decimal(s)
        if natureza == NaturezaSaldo.ZERO and valor != 0:
            natureza = NaturezaSaldo.DEVEDOR
        return valor, natureza
    except Exception:
        return Decimal("0"), NaturezaSaldo.ZERO


def _nivel_codigo(codigo: str) -> int:
    """Conta a profundidade hierárquica de um código de conta."""
    return len(codigo.split("."))


# ---------------------------------------------------------------------------
# Interface base
# ---------------------------------------------------------------------------

class BaseParser(ABC):
    @abstractmethod
    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        pass

    def _extrair_texto_paginas(self, pdf_path: str) -> list[str]:
        from core.ocr import is_scanned_pdf, extrair_texto_ocr
        if is_scanned_pdf(pdf_path):
            return extrair_texto_ocr(pdf_path)  # RuntimeError propaga naturalmente
        paginas = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                texto = page.extract_text() or ""
                paginas.append(texto)
        return paginas


# ---------------------------------------------------------------------------
# Parser: Socium / Hierárquico 5 níveis (Rodomelo)
# ---------------------------------------------------------------------------

class ParserSocium(BaseParser):
    """
    Layout: Código · Classificação · Descrição · Sal.Ant · Débito · Crédito · Sal.Atual
    Sintética: código com menos de 5 segmentos (ex: 1.1.1)
    Analítica: código com 5+ segmentos (ex: 1.1.1.01.001)
    """

    NIVEL_ANALITICO = 5
    RE_LINHA = re.compile(
        r'^(\d+)\s+([\d\.]+)\s+(.+?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s*$'
    )

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas: list[ContaContabil] = []
        # Rastreia códigos já vistos para resolver duplicatas do Socium.
        # O Socium pode emitir o mesmo código para:
        #   Caso 1 — mesma conta com descrição diferente (uma zerada, outra com saldo):
        #            ex: 1.2.1.02.002 "BANCOS CONTA VINCULADA" (zero) e
        #                1.2.1.02.002 "VALORES PAGOS INDEVIDAMENTE" (1.662,00)
        #            → manter a que tem saldo, descartar a zerada.
        #   Caso 2 — contas DISTINTAS com mesmo código por erro do plano de contas:
        #            ex: 2.1.1.03.005 "FINANCIAMENTO RANDON" (180.848,68) e
        #                2.1.1.03.005 "FINANCIAMENTO VOLKS VW/29.530" (1.243.513,15)
        #            → manter as DUAS como linhas separadas (não somar).
        #            O segundo recebe sufixo "_dup{n}" para diferenciar internamente.
        contador_dups: dict[str, int] = {}  # codigo → quantas vezes já apareceu
        idx_zero: dict[str, int] = {}       # codigo → posição da ocorrência zerada

        paginas = self._extrair_texto_paginas(pdf_path)

        for texto in paginas:
            for linha in texto.split("\n"):
                conta = self._parsear_linha(linha.strip())
                if not conta:
                    continue

                cod = conta.codigo
                tem_saldo = conta.saldo_atual > 0 or conta.saldo_anterior > 0

                if cod not in contador_dups:
                    # Primeira ocorrência — registrar normalmente
                    contador_dups[cod] = 1
                    if not tem_saldo:
                        idx_zero[cod] = len(contas)
                    contas.append(conta)
                else:
                    # Código já visto — distinguir os dois casos
                    if not tem_saldo:
                        # Nova ocorrência também zerada → ignorar completamente
                        continue

                    # Nova ocorrência TEM saldo
                    if cod in idx_zero:
                        # Caso 1: a ocorrência anterior estava zerada → substituir no lugar
                        pos = idx_zero.pop(cod)
                        contas[pos] = conta
                    else:
                        # Caso 2: ambas têm saldo → contas distintas com mesmo código
                        # Adicionar como linha separada com sufixo interno
                        contador_dups[cod] += 1
                        n = contador_dups[cod]
                        conta_dup = ContaContabil(
                            codigo=f"{cod}_dup{n}",
                            descricao=conta.descricao,
                            tipo=conta.tipo,
                            natureza=conta.natureza,
                            saldo_anterior=conta.saldo_anterior,
                            debito=conta.debito,
                            credito=conta.credito,
                            saldo_atual=conta.saldo_atual,
                            nivel=conta.nivel,
                        )
                        contas.append(conta_dup)

        return contas

    def _parsear_linha(self, linha: str) -> ContaContabil | None:
        m = self.RE_LINHA.match(linha)
        if not m:
            return None

        _id, classificacao, descricao, sal_ant, debito, credito, sal_atual = m.groups()

        nivel = _nivel_codigo(classificacao)
        tipo = TipoConta.ANALITICA if nivel >= self.NIVEL_ANALITICO else TipoConta.SINTETICA

        val_ant, _ = _limpar_valor(sal_ant)
        val_deb, _ = _limpar_valor(debito)
        val_cre, _ = _limpar_valor(credito)
        val_atu, nat = _limpar_valor(sal_atual)

        # Inferir natureza pelo grupo (1=Ativo=D, 2=Passivo=C)
        if nat == NaturezaSaldo.ZERO:
            nat = NaturezaSaldo.DEVEDOR if classificacao.startswith("1") or classificacao.startswith("3") else NaturezaSaldo.CREDOR

        return ContaContabil(
            codigo=classificacao,
            descricao=descricao.strip(),
            tipo=tipo,
            natureza=nat,
            saldo_anterior=val_ant,
            debito=val_deb,
            credito=val_cre,
            saldo_atual=val_atu,
            nivel=nivel,
        )


# ---------------------------------------------------------------------------
# Parser: Genérico / Statera (código hierárquico sem coluna ID)
# ---------------------------------------------------------------------------

class ParserGenerico(BaseParser):
    """
    Layout: Conta · Descrição · Sal.Ant · Débito · Crédito · Sal.Atual
    Identifica sintéticas pelo número de segmentos no código.
    """

    NIVEL_ANALITICO = 5
    RE_LINHA = re.compile(
        r'^([\d\.]+)\s+(.+?)\s+([\-\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+[DC]?)\s+([\-\d\.,]+[DC]?)\s*$'
    )

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas = []
        paginas = self._extrair_texto_paginas(pdf_path)

        for texto in paginas:
            for linha in texto.split("\n"):
                linha = linha.strip()
                if not linha or linha.startswith("-"):
                    continue
                conta = self._parsear_linha(linha)
                if conta:
                    contas.append(conta)

        return contas

    def _parsear_linha(self, linha: str) -> ContaContabil | None:
        m = self.RE_LINHA.match(linha)
        if not m:
            return None

        codigo, descricao, sal_ant, debito, credito, sal_atual = m.groups()

        # Filtrar linhas que não são contas (totais, cabeçalhos)
        if not re.match(r'^\d', codigo):
            return None

        nivel = _nivel_codigo(codigo)
        tipo = TipoConta.ANALITICA if nivel >= self.NIVEL_ANALITICO else TipoConta.SINTETICA

        val_ant, _ = _limpar_valor(sal_ant)
        val_deb, _ = _limpar_valor(debito)
        val_cre, _ = _limpar_valor(credito)
        val_atu, nat = _limpar_valor(sal_atual)

        if nat == NaturezaSaldo.ZERO:
            nat = NaturezaSaldo.DEVEDOR if codigo.startswith("1") or codigo.startswith("3") else NaturezaSaldo.CREDOR

        return ContaContabil(
            codigo=codigo,
            descricao=descricao.strip(),
            tipo=tipo,
            natureza=nat,
            saldo_anterior=val_ant,
            debito=val_deb,
            credito=val_cre,
            saldo_atual=val_atu,
            nivel=nivel,
        )


# ---------------------------------------------------------------------------
# Parser: ContaCerta (Carolina Ribeiro) — usa [id] como prefixo
# ---------------------------------------------------------------------------

class ParserContaCerta(BaseParser):
    """
    Layout com [id] numérico antes de cada conta.
    Sintéticas: linhas sem valor numérico no saldo ou nível de código <= 2
    """

    RE_LINHA = re.compile(
        r'^\s*\[(\d+)\](.+?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+[DC]?)\s*$'
    )
    RE_SINTETICA = re.compile(
        r'^\s*\[(\d+)\]([A-ZÁÉÍÓÚÀÃÕÇÜ\s\/\(\)\-\.]+)\s+([\d\.,]+[DC])\s*$'
    )

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas = []
        paginas = self._extrair_texto_paginas(pdf_path)

        # Neste layout precisamos inferir hierarquia pela indentação
        # e pelo padrão do código interno [id]
        for texto in paginas:
            linhas = texto.split("\n")
            for linha in linhas:
                conta = self._parsear_linha(linha)
                if conta:
                    contas.append(conta)

        return contas

    def _parsear_linha(self, linha: str) -> ContaContabil | None:
        m = self.RE_LINHA.match(linha)
        if not m:
            return None

        _id, descricao, sal_ant, debito, credito, sal_atual = m.groups()

        val_ant, _ = _limpar_valor(sal_ant)
        val_deb, _ = _limpar_valor(debito)
        val_cre, _ = _limpar_valor(credito)
        val_atu, nat = _limpar_valor(sal_atual)

        # Heurística de tipo: se débito e crédito são 0.00, provavelmente sintética
        tipo = TipoConta.ANALITICA
        descricao_clean = descricao.strip().upper()

        # Palavras-chave que indicam conta sintética/totalizadora
        SINTETICAS_KEYWORDS = [
            "ATIVO", "PASSIVO", "CIRCULANTE", "PERMANENTE", "PATRIMÔNIO",
            "RECEITAS", "DESPESAS", "CUSTOS", "RESULTADO", "NUMERÁRIOS",
            "BANCOS", "CRÉDITOS", "ADIANTAMENTO", "EMPRÉSTIMOS",
        ]
        if any(kw in descricao_clean for kw in SINTETICAS_KEYWORDS):
            tipo = TipoConta.SINTETICA

        if nat == NaturezaSaldo.ZERO:
            nat = NaturezaSaldo.DEVEDOR

        return ContaContabil(
            codigo=_id,
            descricao=descricao.strip(),
            tipo=tipo,
            natureza=nat,
            saldo_anterior=val_ant,
            debito=val_deb,
            credito=val_cre,
            saldo_atual=val_atu,
            nivel=1,
        )


# ---------------------------------------------------------------------------
# Parser: Somar (Bom Pastor) — totalizadores com "="
# ---------------------------------------------------------------------------

class ParserSomar(BaseParser):
    """
    Layout: Descrição · [id] · Sal.Ant · Débito · Crédito · Sal.Atual
    Sintéticas: linhas que começam com "=" ou "=T o t a l"
    """

    RE_ANALITICA = re.compile(
        r'^(.+?)\s+-\s+\[(\d+)\]\s+([\d\.,]+[DC]?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+[DC]?)\s*$'
    )
    RE_SINTETICA = re.compile(r'^=(.+?)\s+([\d\.,]+[DC]?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+[DC]?)\s*$')
    RE_TOTAL = re.compile(r'^=T\s*o\s*t\s*a\s*l')

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas = []
        paginas = self._extrair_texto_paginas(pdf_path)
        grupo_atual = "RAIZ"

        for texto in paginas:
            for linha in texto.split("\n"):
                linha = linha.strip()
                if not linha:
                    continue

                # Linha de total (ignorar — são redundantes com sintéticas)
                if self.RE_TOTAL.match(linha):
                    continue

                # Conta sintética (grupo)
                ms = self.RE_SINTETICA.match(linha)
                if ms:
                    descricao, sal_ant, debito, credito, sal_atual = ms.groups()
                    grupo_atual = descricao.strip()
                    val_ant, _ = _limpar_valor(sal_ant)
                    val_deb, _ = _limpar_valor(debito)
                    val_cre, _ = _limpar_valor(credito)
                    val_atu, nat = _limpar_valor(sal_atual)
                    contas.append(ContaContabil(
                        codigo=grupo_atual[:30],
                        descricao=grupo_atual,
                        tipo=TipoConta.SINTETICA,
                        natureza=nat if nat != NaturezaSaldo.ZERO else NaturezaSaldo.DEVEDOR,
                        saldo_anterior=val_ant,
                        debito=val_deb,
                        credito=val_cre,
                        saldo_atual=val_atu,
                        nivel=1,
                    ))
                    continue

                # Conta analítica
                ma = self.RE_ANALITICA.match(linha)
                if ma:
                    descricao, _id, sal_ant, debito, credito, sal_atual = ma.groups()
                    val_ant, _ = _limpar_valor(sal_ant)
                    val_deb, _ = _limpar_valor(debito)
                    val_cre, _ = _limpar_valor(credito)
                    val_atu, nat = _limpar_valor(sal_atual)
                    contas.append(ContaContabil(
                        codigo=_id,
                        descricao=descricao.strip(),
                        tipo=TipoConta.ANALITICA,
                        natureza=nat if nat != NaturezaSaldo.ZERO else NaturezaSaldo.DEVEDOR,
                        saldo_anterior=val_ant,
                        debito=val_deb,
                        credito=val_cre,
                        saldo_atual=val_atu,
                        nivel=2,
                    ))

        return contas


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Parser: CTC007 (Sagrados Corações) — 9 colunas com MOV. MENSAL separado
# ---------------------------------------------------------------------------

class ParserCTC007(BaseParser):
    """
    Layout CTC007 com 9 colunas por linha de conta:
      Conta · C.Res · [C.C D/C] · D/C · Nome · Sal.Ant · Débito · Crédito · Mov.Mensal · Sal.Atual

    Particularidades identificadas nos PDFs reais (Sagrados Corações):
      1. Coluna C.C pode ter modificadores 'D-BAN', 'D-REC', 'D-PAG' seguidos de D/C
      2. O SALDO ATUAL é a 9ª coluna — não calcular via Sal.Ant + Mov.Mensal
      3. Mov.Mensal pode ser negativo (ex: -100.132,01)
    """

    # Regex para linhas COM modificador de C.C (D-BAN, D-REC, D-PAG, etc.)
    # Ex: '1.1.01.01.002 0007 D-BAN D CAIXA GERAL 259.439,58 0,00 0,00 0,00 259.439,58'
    RE_COM_CC = re.compile(
        r'^([\d\.]+)'          # codigo da conta
        r'\s+(\d+)'            # C.Res
        r'\s+([\w\-]+)'        # modificador C.C (D-BAN, D-REC, etc.)
        r'\s+([DC])'            # D/C real da conta
        r'\s+(.+?)'             # descricao
        r'\s+([-\d\.,]+)'      # saldo anterior (pode ser negativo)
        r'\s+([\d\.,]+)'       # debito
        r'\s+([\d\.,]+)'       # credito
        r'\s+([-\d\.,]+)'      # mov. mensal (pode ser negativo)
        r'\s+([-\d\.,]+)\s*$'  # saldo atual (pode ser negativo — ex: conta CREDOR com saldo devedor)
    )

    # Regex para linhas SEM modificador de C.C
    # Ex: '1.1.01.02 0010 D BANCOS CONTA MOVIMENTO 2.334,88 398.204,80 370.268,31 27.936,49 30.271,37'
    RE_SEM_CC = re.compile(
        r'^([\d\.]+)'          # codigo
        r'\s+(\d+)'            # C.Res
        r'\s+([DC])'            # D/C
        r'\s+(.+?)'             # descricao
        r'\s+([-\d\.,]+)'      # saldo anterior (pode ser negativo)
        r'\s+([\d\.,]+)'       # debito
        r'\s+([\d\.,]+)'       # credito
        r'\s+([-\d\.,]+)'      # mov. mensal (pode ser negativo)
        r'\s+([-\d\.,]+)\s*$'  # saldo atual (pode ser negativo)
    )

    NIVEL_ANALITICO = 5

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas = []
        paginas = self._extrair_texto_paginas(pdf_path)

        for texto in paginas:
            for linha in texto.split("\n"):
                linha = linha.strip()
                if not linha or linha.startswith("-"):
                    continue
                conta = self._parsear_linha(linha)
                if conta:
                    contas.append(conta)

        return contas

    def _parsear_linha(self, linha: str) -> ContaContabil | None:
        # Tenta primeiro com modificador de C.C (D-BAN, D-REC, etc.)
        m = self.RE_COM_CC.match(linha)
        if m:
            codigo, _cres, _cc, dc, descricao, sa, deb, cre, _mov, satu = m.groups()
            return self._montar(codigo, dc, descricao, sa, deb, cre, satu)

        # Tenta sem modificador
        m = self.RE_SEM_CC.match(linha)
        if m:
            codigo, _cres, dc, descricao, sa, deb, cre, _mov, satu = m.groups()
            return self._montar(codigo, dc, descricao, sa, deb, cre, satu)

        return None

    def _montar(self, codigo, dc, descricao, sa, deb, cre, satu) -> ContaContabil:
        """Monta ContaContabil a partir dos campos já separados pela regex."""
        nivel = _nivel_codigo(codigo)
        tipo  = TipoConta.ANALITICA if nivel >= self.NIVEL_ANALITICO else TipoConta.SINTETICA
        nat   = NaturezaSaldo.DEVEDOR if dc.upper() == "D" else NaturezaSaldo.CREDOR

        val_ant, _ = _limpar_valor(sa)
        val_deb, _ = _limpar_valor(deb)
        val_cre, _ = _limpar_valor(cre)
        val_atu, _ = _limpar_valor(satu)  # usa a 9a coluna — saldo atual real do PDF

        return ContaContabil(
            codigo=codigo,
            descricao=descricao.strip(),
            tipo=tipo,
            natureza=nat,
            saldo_anterior=val_ant,
            debito=val_deb,
            credito=val_cre,
            saldo_atual=val_atu,
            nivel=nivel,
        )


# ---------------------------------------------------------------------------
# Parser: Statera / Verificação Padrão — código + num_interno + descrição
# ---------------------------------------------------------------------------

class ParserVerificacao(BaseParser):
    """
    Layout: Conta · [num_interno] · Descrição · Sal.Ant · Débito · Crédito · Sal.Atual
    Analítica: código com segmento "00XXX" (ex: 1.1.1.1.01.00001)
    Sintética: código sem segmento "00XXX"
    """

    NIVEL_ANALITICO = 6  # ex: 1.1.1.1.01.00001

    RE_ANALITICA = re.compile(
        r'^([\d\.]+\.\d{5})\s+(\d+)\s+(.+?)\s+(-?[\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(-?[\d\.,]+)\s*$'
    )
    RE_SINTETICA = re.compile(
        r'^([\d\.]+)\s+(.+?)\s+(-?[\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+(-?[\d\.,]+)\s*$'
    )

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas = []
        paginas = self._extrair_texto_paginas(pdf_path)

        for texto in paginas:
            for linha in texto.split("\n"):
                linha = linha.strip()
                if not linha or linha.startswith("-") or linha.startswith("="):
                    continue

                conta = self._parsear_analitica(linha) or self._parsear_sintetica(linha)
                if conta:
                    contas.append(conta)

        return contas

    def _parsear_analitica(self, linha: str) -> ContaContabil | None:
        m = self.RE_ANALITICA.match(linha)
        if not m:
            return None
        codigo, _num, descricao, sal_ant, debito, credito, sal_atual = m.groups()
        return self._montar_conta(codigo, descricao, sal_ant, debito, credito, sal_atual, TipoConta.ANALITICA)

    def _parsear_sintetica(self, linha: str) -> ContaContabil | None:
        m = self.RE_SINTETICA.match(linha)
        if not m:
            return None
        codigo, descricao, sal_ant, debito, credito, sal_atual = m.groups()
        # Filtrar linhas de cabeçalho
        if any(kw in descricao for kw in ["Conta", "Descricao", "Saldo", "Folha"]):
            return None
        nivel = _nivel_codigo(codigo)
        tipo = TipoConta.ANALITICA if nivel >= self.NIVEL_ANALITICO else TipoConta.SINTETICA
        return self._montar_conta(codigo, descricao, sal_ant, debito, credito, sal_atual, tipo)

    def _montar_conta(self, codigo, descricao, sal_ant, debito, credito, sal_atual, tipo) -> ContaContabil:
        val_ant, _ = _limpar_valor(sal_ant)
        val_deb, _ = _limpar_valor(debito)
        val_cre, _ = _limpar_valor(credito)
        val_atu, nat = _limpar_valor(sal_atual)

        # Para o Statera, passivo tem saldo negativo — normalizar
        if val_atu < 0:
            val_atu = -val_atu
            nat = NaturezaSaldo.CREDOR
        elif nat == NaturezaSaldo.ZERO:
            nat = NaturezaSaldo.DEVEDOR if codigo.startswith("1") or codigo.startswith("3") else NaturezaSaldo.CREDOR

        return ContaContabil(
            codigo=codigo,
            descricao=descricao.strip(),
            tipo=tipo,
            natureza=nat,
            saldo_anterior=abs(val_ant),
            debito=abs(val_deb),
            credito=abs(val_cre),
            saldo_atual=abs(val_atu),
            nivel=_nivel_codigo(codigo),
        )


# ---------------------------------------------------------------------------
# Parser: FS Sequencial (F S Soluções Contábeis) — códigos inteiros sequenciais
# ---------------------------------------------------------------------------

class ParserFSSequencial(BaseParser):
    """
    Layout FS Sequencial: códigos inteiros sem separador hierárquico.
    Ex: 1=ATIVO, 3=DISPONIVEL, 3695=TARGET, 12=CLIENTES, 504=CLIENTES DIVERSOS

    Características:
    - Saldo sempre com sufixo D/C: '5.197.229,63D', '4.069.944,64C'
    - 6 colunas: Código  Descrição  Sal.Ant  Débito  Crédito  Sal.Atual
    - PDFs escaneados — texto extraído via OCR (artefatos de ruído esperados)
    - TODAS as contas marcadas como ANALITICA para o mapper conseguir acessá-las.
      O mapeamento usa correspondência EXATA de código (não prefixo).
    """

    RE_LINHA = re.compile(
        r'^(\d+)'                       # código inteiro (1, 12, 3695, ...)
        r'\s+(.+?)'                     # descrição (lazy — absorve ruído OCR)
        r'\s+([\d\.,]+[DC]?)'           # saldo anterior
        r'\s+([\d\.,]+[DC]?)'           # débito
        r'\s+([\d\.,]+[DC]?)'           # crédito
        r'\s+([\d\.,]+[DC]?)'           # saldo atual
        r'\s*$'
    )

    # Palavras-chave de linhas de cabeçalho/rodapé que devem ser ignoradas
    _IGNORAR = re.compile(
        r'C[oó]digo|Descri|Saldo|Folha|Empresa|C\.N\.P|Per[ií]odo|Sistema|BALANCETE',
        re.IGNORECASE
    )

    def parsear(self, pdf_path: str) -> list[ContaContabil]:
        contas: list[ContaContabil] = []
        paginas = self._extrair_texto_paginas(pdf_path)

        for texto in paginas:
            for linha in texto.split("\n"):
                linha = linha.strip()
                if not linha:
                    continue
                # Rejeitar linhas de cabeçalho/rodapé antes de tentar o regex
                if self._IGNORAR.search(linha):
                    continue
                conta = self._parsear_linha(linha)
                if conta:
                    contas.append(conta)

        return contas

    def _parsear_linha(self, linha: str) -> ContaContabil | None:
        m = self.RE_LINHA.match(linha)
        if not m:
            return None

        codigo, descricao, sal_ant, debito, credito, sal_atual = m.groups()

        # Limpar artefatos OCR do início da descrição (_ATIVO, —ATIVO, etc.)
        descricao = re.sub(r'^[_\-—\s]+', '', descricao).strip()
        # Remover tokens de ruído isolados no meio da descrição (ex: ". :", "' :")
        descricao = re.sub(r'\s*[\.\'\":;/]\s*:\s*', ' ', descricao).strip()

        val_ant, _  = _limpar_valor(sal_ant)
        val_deb, _  = _limpar_valor(debito)
        val_cre, _  = _limpar_valor(credito)
        val_atu, nat = _limpar_valor(sal_atual)

        # Inferir natureza pelo grupo contábil se OCR não trouxe D/C
        if nat == NaturezaSaldo.ZERO and val_atu > 0:
            nat = NaturezaSaldo.DEVEDOR if codigo.startswith(("1", "3", "4")) \
                  else NaturezaSaldo.CREDOR

        return ContaContabil(
            codigo=codigo,
            descricao=descricao,
            tipo=TipoConta.ANALITICA,   # todas analíticas — mapper usa exact match
            natureza=nat,
            saldo_anterior=val_ant,
            debito=val_deb,
            credito=val_cre,
            saldo_atual=val_atu,
            nivel=1,
        )


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

_PARSERS: dict[str, type[BaseParser]] = {
    "socium": ParserSocium,
    "hierarquico_5niveis": ParserVerificacao,
    "generico": ParserVerificacao,
    "contacerta": ParserContaCerta,
    "somar": ParserSomar,
    "ctc007": ParserCTC007,
    "fs_sequencial": ParserFSSequencial,
}


def get_parser(layout_id: str) -> BaseParser:
    """Retorna o parser correto para o layout detectado."""
    cls = _PARSERS.get(layout_id, ParserVerificacao)
    return cls()
