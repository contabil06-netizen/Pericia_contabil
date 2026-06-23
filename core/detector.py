"""
Detector de layout: identifica qual sistema contábil gerou o PDF
analisando strings características nas primeiras páginas.
"""
from __future__ import annotations
import re
import pdfplumber


IDENTIFICADORES: list[tuple[str, str]] = [
    ("somar",       "SOMAR CONTABILIDADE"),
    ("contacerta",  "CONTACERTA ORGANIZACAO CONTABIL"),
    ("socium",      "CONTABILIDADE E GESTAO EMPRESARIAL SOCIUM"),
    ("ctc007",      "CTC007"),
    ("alterdata",   "ALTERDATA"),
    ("dominio",     "DOMÍNIO SISTEMAS"),
    ("prosoft",     "PROSOFT"),
    ("fortes",      "FORTES TECNOLOGIA"),
]


def detectar_layout(pdf_path: str) -> str:
    """
    Lê as primeiras 3 páginas do PDF e identifica o sistema contábil.
    Retorna o id do layout ou 'generico' se não reconhecido.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            paginas = min(3, len(pdf.pages))
            texto = ""
            for i in range(paginas):
                t = pdf.pages[i].extract_text() or ""
                texto += t.upper()

        for layout_id, identificador in IDENTIFICADORES:
            if identificador.upper() in texto:
                return layout_id

        return _detectar_por_estrutura(texto)

    except Exception:
        return "generico"


def _detectar_por_estrutura(texto: str) -> str:
    """Fallback heurístico quando o identificador de sistema não é encontrado."""
    if re.search(r'\d+\.\d+\.\d+\.\d+\.\d+', texto):
        return "hierarquico_5niveis"
    if re.search(r'\bD/C\b', texto) or re.search(r'\bD-BAN\b|\bD-REC\b|\bD-PAG\b', texto):
        return "ctc007"
    if re.search(r'\[\d+\]', texto):
        return "contacerta"
    if re.search(r'=[\w\s]+', texto):
        return "somar"
    return "generico"


def extrair_metadados_cabecalho(pdf_path: str) -> dict:
    """Extrai empresa, CNPJ e período da(s) primeira(s) página(s)."""
    meta = {"empresa": "", "cnpj": "", "periodo_inicio": "", "periodo_fim": ""}

    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto = pdf.pages[0].extract_text() or ""

        # Fallback OCR: se pdfplumber não extraiu texto (PDF escaneado),
        # usa OCR da primeira página para obter os metadados do cabeçalho.
        if not texto.strip():
            try:
                from core.ocr import extrair_texto_ocr
                paginas_ocr = extrair_texto_ocr(pdf_path)
                texto = paginas_ocr[0] if paginas_ocr else ""
            except Exception:
                pass

        cnpj_match = re.search(
            r'CNPJ[:\s]+([0-9]{2}[\.\-]?[0-9]{3}[\.\-]?[0-9]{3}[\/\-]?[0-9]{4}[\-]?[0-9]{2})',
            texto, re.IGNORECASE
        )
        if cnpj_match:
            meta["cnpj"] = re.sub(r'[^\d\/\-\.]', '', cnpj_match.group(1))

        periodo_match = re.search(
            r'(\d{2}/\d{2}/\d{4})\s*(?:até|a|-)\s*(\d{2}/\d{2}/\d{4})',
            texto, re.IGNORECASE
        )
        if periodo_match:
            meta["periodo_inicio"] = periodo_match.group(1)
            meta["periodo_fim"]    = periodo_match.group(2)

        linhas = [l.strip() for l in texto.split('\n') if l.strip()]
        for linha in linhas[:5]:
            if len(linha) > 5 and not re.search(r'CNPJ|BALANCETE|CRC|CPF', linha, re.IGNORECASE):
                meta["empresa"] = linha
                break

    except Exception:
        pass

    return meta
