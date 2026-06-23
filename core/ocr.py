"""
Módulo OCR para extração de texto em PDFs escaneados (imagem).
Utiliza pymupdf (fitz) para renderizar as páginas em imagem e
pytesseract com língua portuguesa para reconhecimento de caracteres.

Integração com parsers.py: BaseParser._extrair_texto_paginas()
chama is_scanned_pdf() e, se True, delega para extrair_texto_ocr().
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
from typing import List


# ───────────────────────────────────────────────────────────────────────────
# Constantes de configuração
# ───────────────────────────────────────────────────────────────────────────

_LIMIAR_CHARS_PAGINA    = 50      # média abaixo disto → PDF escaneado
_DPI_RENDERIZACAO       = 300     # resolução para pymupdf → imagem
_TESSERACT_WINDOWS_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
_LANG_TENTATIVAS        = ["por+eng", "por", "eng"]
_PSM                    = 6       # PSM 6: bloco de texto uniforme (balancetes)
_OEM                    = 3       # OEM 3: LSTM padrão


# ───────────────────────────────────────────────────────────────────────────
# API pública
# ───────────────────────────────────────────────────────────────────────────

def is_scanned_pdf(pdf_path: str) -> bool:
    """
    Retorna True se o PDF não contém texto extraível suficiente
    (média < _LIMIAR_CHARS_PAGINA caracteres por página).

    Detecta PDFs de imagem sem precisar das dependências de OCR —
    usa apenas pdfplumber, que já é dependência obrigatória do sistema.
    """
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return True
            total = sum(len(p.extract_text() or "") for p in pdf.pages)
            media = total / len(pdf.pages)
        return media < _LIMIAR_CHARS_PAGINA
    except Exception:
        return False


def extrair_texto_ocr(pdf_path: str) -> List[str]:
    """
    Renderiza cada página do PDF como imagem em escala de cinza (300 DPI)
    e aplica OCR via pytesseract.

    Retorna lista de strings (uma por página), equivalente ao que
    BaseParser._extrair_texto_paginas() retornaria em um PDF nativo.

    Raises:
        RuntimeError: dependência ausente (pymupdf, pytesseract ou Pillow)
                      ou Tesseract não encontrado no sistema.
    """
    _verificar_dependencias()
    _configurar_tesseract()

    import fitz                       # pymupdf
    import pytesseract
    from PIL import Image
    import io

    paginas_texto: List[str] = []
    # Matriz de escala para atingir o DPI desejado a partir de 72 DPI base
    matriz = fitz.Matrix(_DPI_RENDERIZACAO / 72, _DPI_RENDERIZACAO / 72)

    with fitz.open(pdf_path) as doc:
        for n_pag in range(len(doc)):
            pagina = doc[n_pag]
            pix    = pagina.get_pixmap(matrix=matriz, colorspace=fitz.csGRAY)
            img    = Image.open(io.BytesIO(pix.tobytes("png")))

            img   = _preprocessar(img)
            texto = _ocr_com_fallback(img, pytesseract)
            texto = _pos_processar(texto)
            paginas_texto.append(texto)

    return paginas_texto


# ───────────────────────────────────────────────────────────────────────────
# Pré-processamento de imagem
# ───────────────────────────────────────────────────────────────────────────

def _preprocessar(img):
    """
    Aplica filtro de nitidez e aumento de contraste para melhorar o OCR.
    Opera em escala de cinza (modo 'L') para reduzir ruído de cor.
    """
    from PIL import ImageFilter, ImageEnhance

    if img.mode != "L":
        img = img.convert("L")

    img = img.filter(ImageFilter.SHARPEN)
    img = ImageEnhance.Contrast(img).enhance(1.5)

    return img


# ───────────────────────────────────────────────────────────────────────────
# OCR com tentativas de fallback de idioma
# ───────────────────────────────────────────────────────────────────────────

def _ocr_com_fallback(img, pytesseract) -> str:
    """
    Tenta OCR sequencialmente com 'por+eng', 'por' e 'eng'.
    Retorna o primeiro resultado não-vazio, ou string vazia se todos falharem.
    """
    config = f"--psm {_PSM} --oem {_OEM}"
    for lang in _LANG_TENTATIVAS:
        try:
            resultado = pytesseract.image_to_string(img, lang=lang, config=config)
            if resultado.strip():
                return resultado
        except pytesseract.TesseractError:
            continue
    return ""


# ───────────────────────────────────────────────────────────────────────────
# Pós-processamento do texto extraído
# ───────────────────────────────────────────────────────────────────────────

def _pos_processar(texto: str) -> str:
    """Aplica correções linha a linha no texto bruto do OCR."""
    linhas = []
    for linha in texto.splitlines():
        linha = _corrigir_numericos(linha)
        linha = _corrigir_espacos(linha)
        linhas.append(linha)
    return "\n".join(linhas)


def _corrigir_numericos(linha: str) -> str:
    """
    Corrige erros típicos do OCR em contexto numérico de balancetes:

    - 'O' maiúsculo entre dígitos/separadores → '0'
      Ex: '1O.234,56' → '10.234,56'
    - 'l' ou 'I' entre dígitos/separadores → '1'
      Ex: '1.23l,45' → '1.231,45'
    - Remove espaço entre número e sufixo D/C
      Ex: '1.234,56 D' → '1.234,56D'
    """
    # O maiúsculo adjacente a dígito/separador numérico → 0
    linha = re.sub(r'(?<=[\d,\.])O(?=[\d,\.])', '0', linha)
    linha = re.sub(r'(?<=[\d])O(?=[\s])', '0', linha)

    # l ou I entre dígito e separador numérico → 1
    linha = re.sub(r'(?<=[\d,\.])[lI](?=[\d,\.])', '1', linha)

    # Colapsar espaço entre número e sufixo D/C no fim de token
    linha = re.sub(r'(\d)\s+([DC])\b', r'\1\2', linha)

    return linha


def _corrigir_espacos(linha: str) -> str:
    """Remove espaços duplos consecutivos (artefato comum em tabelas escaneadas)."""
    return re.sub(r'  +', ' ', linha).strip()


# ───────────────────────────────────────────────────────────────────────────
# Verificação e configuração de dependências
# ───────────────────────────────────────────────────────────────────────────

def _verificar_dependencias() -> None:
    """
    Verifica se pymupdf, pytesseract e Pillow estão instalados.
    Levanta RuntimeError descritivo com instruções de instalação se algum faltar.
    """
    faltando: List[str] = []

    try:
        import fitz  # noqa: F401
    except ImportError:
        faltando.append("pymupdf      →  pip install pymupdf")

    try:
        import pytesseract  # noqa: F401
    except ImportError:
        faltando.append("pytesseract  →  pip install pytesseract")

    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        faltando.append("Pillow       →  pip install Pillow")

    if faltando:
        instrucoes = "\n  ".join(faltando)
        raise RuntimeError(
            "Dependências de OCR não encontradas. Instale com pip:\n\n"
            f"  {instrucoes}\n\n"
            "Também é necessário o Tesseract OCR instalado no sistema:\n"
            "  Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "           (marque o pacote de idioma 'Portuguese' na instalação)\n"
            "  Linux:   sudo apt install tesseract-ocr tesseract-ocr-por\n"
            "  macOS:   brew install tesseract tesseract-lang"
        )


def _configurar_tesseract() -> None:
    """
    No Windows, aponta o pytesseract para o executável padrão do Tesseract,
    caso ainda não tenha sido configurado manualmente.
    Em outros sistemas, assume que 'tesseract' está no PATH do sistema.
    """
    if sys.platform != "win32":
        return

    try:
        import pytesseract
        cmd_atual = pytesseract.pytesseract.tesseract_cmd
        if not cmd_atual or cmd_atual == "tesseract":
            exe = Path(_TESSERACT_WINDOWS_PATH)
            if exe.exists():
                pytesseract.pytesseract.tesseract_cmd = str(exe)
            # Se não encontrado no caminho padrão, deixa o pytesseract tentar o PATH
    except ImportError:
        pass  # _verificar_dependencias() já capturou este caso
