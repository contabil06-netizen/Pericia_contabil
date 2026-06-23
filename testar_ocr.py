"""
Script de diagnóstico do pipeline OCR.
Execute no diretório do projeto:
    python testar_ocr.py
"""
import sys
import os
from pathlib import Path

os.chdir(Path(__file__).parent)

print("=" * 50)
print("  DIAGNÓSTICO OCR — PERITUS")
print("=" * 50)

# 1. Tesseract
print("\n[1] Tesseract...")
try:
    import pytesseract
    from pathlib import Path
    exe = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if exe.exists():
        pytesseract.pytesseract.tesseract_cmd = str(exe)
        print(f"  ✓ Encontrado: {exe}")
    else:
        print(f"  ⚠ Não encontrado em {exe}")
        print("    Verifique o caminho de instalação e ajuste _TESSERACT_WINDOWS_PATH em core/ocr.py")
    ver = pytesseract.get_tesseract_version()
    print(f"  ✓ Versão: {ver}")
except Exception as e:
    print(f"  ✗ Erro: {e}")

# 2. Idiomas disponíveis
print("\n[2] Idiomas instalados no Tesseract...")
try:
    langs = pytesseract.get_languages(config="")
    print(f"  Idiomas: {', '.join(langs)}")
    if "por" in langs:
        print("  ✓ Português (por) disponível")
    else:
        print("  ✗ Português (por) NÃO encontrado")
        print("    Baixe em: https://github.com/tesseract-ocr/tessdata")
        print("    Copie 'por.traineddata' para C:\\Program Files\\Tesseract-OCR\\tessdata\\")
except Exception as e:
    print(f"  ✗ Erro: {e}")

# 3. pymupdf
print("\n[3] pymupdf (fitz)...")
try:
    import fitz
    print(f"  ✓ versão {fitz.version[0]}")
except ImportError:
    print("  ✗ Não instalado: pip install pymupdf")

# 4. Pillow
print("\n[4] Pillow...")
try:
    from PIL import Image
    import PIL
    print(f"  ✓ versão {PIL.__version__}")
except ImportError:
    print("  ✗ Não instalado: pip install Pillow")

# 5. Teste rápido de OCR
print("\n[5] Teste de OCR num PDF da FT Transporte...")
pdf_dir = Path(r"C:\Users\conta\Documents\DOCUMENTOS PARA proj\FT Transportes")
pdfs = list(pdf_dir.glob("*.pdf")) if pdf_dir.exists() else []

if not pdfs:
    print("  ⚠ Pasta de PDFs não encontrada — pulando teste")
else:
    pdf = pdfs[0]
    print(f"  Testando: {pdf.name}")
    try:
        sys.path.insert(0, str(Path(__file__).parent))
        from core.ocr import is_scanned_pdf, extrair_texto_ocr
        escaneado = is_scanned_pdf(str(pdf))
        print(f"  is_scanned_pdf: {escaneado}")
        if escaneado:
            paginas = extrair_texto_ocr(str(pdf))
            chars = sum(len(p) for p in paginas)
            print(f"  ✓ OCR concluído: {len(paginas)} página(s), {chars} chars extraídos")
            print(f"\n  --- Trecho da pág. 1 ---")
            print(paginas[0][:500])
            print("  ---")
        else:
            print("  PDF tem texto nativo — OCR não necessário")
    except Exception as e:
        print(f"  ✗ Erro: {e}")
        import traceback
        traceback.print_exc()

print("\n" + "=" * 50)
