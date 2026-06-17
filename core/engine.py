"""
Engine central do sistema de perícia contábil.
Orquestra: Detecção → Parsing → Validação → Mapeamento → Exportação

v2: aceita pdf_anterior para comparativo na DRE.
    Gera apenas UMA planilha (do mês atual), com coluna do mês anterior na DRE.
    Bloqueia execução quando cliente_id não tem YAML cadastrado.
"""
from __future__ import annotations
from pathlib import Path
from typing import Optional

from core.detector import detectar_layout, extrair_metadados_cabecalho
from core.models import Balancete, RelatorioFinal
from core.parsers import get_parser
from core.validator import validar_balancete
from core.mapper import carregar_mapa_cliente, mapear
from core.exporter import exportar


def processar(
    pdf_path: str,
    cliente_id: str | None = None,
    output_dir: str = "output",
    verbose: bool = True,
    pdf_anterior: str | None = None,
    callback_log=None,
) -> dict:
    """
    Pipeline completo de processamento de um balancete PDF.

    Args:
        pdf_path:      caminho para o PDF do balancete do mês ATUAL
        cliente_id:    identificador do cliente (nome do YAML em clientes/)
        output_dir:    diretório de saída para o Excel
        verbose:       imprimir progresso no terminal
        pdf_anterior:  caminho para o PDF do mês ANTERIOR (opcional)
                       — usado apenas para coluna comparativa na DRE
        callback_log:  função(msg, tipo) para redirecionar logs à GUI

    Returns:
        dict com status, caminho do arquivo gerado e lista de avisos
    """
    resultado = {
        "status":           "ok",
        "arquivo_gerado":   None,
        "layout_detectado": None,
        "empresa":          None,
        "total_contas":     0,
        "erros_validacao":  [],
        "erros_mapeamento": [],
        "avisos":           [],
    }

    def log(msg: str, tipo: str = "normal"):
        if callback_log:
            callback_log(msg, tipo)
        elif verbose:
            print(msg)

    # ── Bloco 0: carregar YAML do cliente (opcional para modo bruto) ────────
    config = {}
    if cliente_id:
        try:
            config = carregar_mapa_cliente(cliente_id)
        except FileNotFoundError as e:
            resultado["status"] = "erro"
            resultado["avisos"].append(str(e))
            log(f"  ✗ {e}", "erro")
            return resultado
    else:
        log("  → Modo extração bruta (sem cliente/mapeamento).", "aviso")

    # ── Etapa 1: Processar PDF do mês ATUAL ──────────────────────────────
    log(f"\n{'='*60}", "secao")
    log(f"  PROCESSANDO: {Path(pdf_path).name}", "secao")
    log(f"{'='*60}", "secao")

    relatorio_atual = _processar_pdf(pdf_path, config, cliente_id, resultado, log)
    if relatorio_atual is None:
        return resultado

    # ── Etapa 2: Processar PDF do mês ANTERIOR (opcional) ─────────────────
    relatorio_anterior: Optional[RelatorioFinal] = None
    if pdf_anterior and Path(pdf_anterior).exists():
        log(f"\n  → Processando mês anterior: {Path(pdf_anterior).name}", "info")
        resultado_ant = {
            "status": "ok", "arquivo_gerado": None, "layout_detectado": None,
            "empresa": None, "total_contas": 0, "erros_validacao": [],
            "erros_mapeamento": [], "avisos": [],
        }

        def log_ant(msg, tipo="normal"):
            log(f"    [ant] {msg}", tipo)

        relatorio_anterior = _processar_pdf(
            pdf_anterior, config, cliente_id, resultado_ant, log_ant
        )
        if relatorio_anterior is None:
            log("  ⚠ Mês anterior com erro — gerado sem comparativo.", "aviso")
    else:
        log("  → Mês anterior não fornecido — DRE sem coluna comparativa.", "aviso")

    # ── Etapa 3: Gerar UMA planilha Excel ────────────────────────────────
    log(f"\n  [4/4] Gerando Excel...", "info")
    try:
        arquivo = exportar(
            relatorio_atual,
            output_dir=output_dir,
            relatorio_anterior=relatorio_anterior,
        )
        resultado["arquivo_gerado"] = arquivo
        log(f"  ✓ Arquivo gerado: {Path(arquivo).name}", "ok")
    except Exception as e:
        resultado["status"] = "erro_exportacao"
        resultado["avisos"].append(f"Erro na exportação: {e}")
        log(f"  ✗ Erro na exportação: {e}", "erro")
        import traceback
        log(traceback.format_exc(), "erro")

    log(f"\n{'='*60}", "secao")
    log(f"  STATUS FINAL: {resultado['status'].upper()}", "secao")
    log(f"{'='*60}\n", "secao")

    return resultado


def _processar_pdf(
    pdf_path: str,
    config: dict,
    cliente_id: str,
    resultado: dict,
    log,
) -> Optional[RelatorioFinal]:
    """
    Executa Detecção → Parsing → Validação → Mapeamento para um único PDF.
    Retorna RelatorioFinal ou None em caso de erro crítico.
    """
    # Detecção
    layout = detectar_layout(pdf_path)
    resultado["layout_detectado"] = layout
    meta   = extrair_metadados_cabecalho(pdf_path)
    resultado["empresa"] = meta.get("empresa", "")

    log(f"  ✓ Layout: {layout}", "ok")
    log(f"  ✓ Empresa: {meta.get('empresa', 'não identificada')}", "ok")
    log(f"  ✓ Período: {meta.get('periodo_inicio','?')} a {meta.get('periodo_fim','?')}", "ok")

    # Parsing
    log("  [1/4] Extraindo contas...", "info")
    parser = get_parser(layout)
    contas = parser.parsear(pdf_path)

    if not contas:
        resultado["status"] = "erro"
        resultado["avisos"].append(
            "Nenhuma conta extraída. Verifique se o PDF tem camada de texto."
        )
        log("  ✗ Nenhuma conta extraída.", "erro")
        return None

    log(f"  ✓ {len(contas)} contas extraídas", "ok")
    resultado["total_contas"] = len(contas)

    # Montar Balancete
    balancete = Balancete(
        empresa=meta.get("empresa", "Empresa não identificada"),
        cnpj=meta.get("cnpj", ""),
        periodo_inicio=meta.get("periodo_inicio", ""),
        periodo_fim=meta.get("periodo_fim", ""),
        sistema_contabil=layout,
        layout_detectado=layout,
        contas=contas,
    )

    # Validação
    log("  [2/4] Validando integridade contábil...", "info")
    erros = validar_balancete(balancete)
    balancete.erros_validacao = erros
    resultado["erros_validacao"] = [
        {"severidade": e.severidade.value, "mensagem": e.mensagem} for e in erros
    ]
    criticos = sum(1 for e in erros if e.severidade.value == "CRITICO")
    avisos   = sum(1 for e in erros if e.severidade.value == "AVISO")
    log(f"  ✓ Validação: {criticos} crítico(s), {avisos} aviso(s)",
        "ok" if criticos == 0 else "aviso")

    # Mapeamento
    log(f"  [3/4] Mapeando contas ({cliente_id})...", "info")
    relatorio = mapear(balancete, config)
    resultado["erros_mapeamento"] = relatorio.erros_mapeamento
    log(f"  ✓ {len(relatorio.grupos_balanco)} grupos BP, "
        f"{len(relatorio.dre_estrutura)} grupos DRE mapeados", "ok")
    if relatorio.erros_mapeamento:
        log(f"  ⚠ {len(relatorio.erros_mapeamento)} aviso(s) de mapeamento", "aviso")

    return relatorio


def processar_lote(
    pdfs: list[str],
    cliente_id: str | None = None,
    output_dir: str = "output",
) -> list[dict]:
    """Processa múltiplos PDFs em sequência."""
    resultados = []
    for i, pdf in enumerate(pdfs, 1):
        print(f"\n[{i}/{len(pdfs)}] Processando {Path(pdf).name}...")
        r = processar(pdf, cliente_id=cliente_id, output_dir=output_dir)
        resultados.append(r)

    ok = sum(1 for r in resultados if r["status"] == "ok")
    print(f"\n{'='*60}")
    print(f"  LOTE CONCLUÍDO: {ok}/{len(pdfs)} processados com sucesso")
    print(f"{'='*60}")
    return resultados
