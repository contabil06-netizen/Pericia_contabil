---
name: pericia-gui-engine
description: >
  Use este skill sempre que precisar modificar a interface gráfica (Tkinter)
  ou o engine de orquestração do projeto de perícia contábil. Triggers incluem:
  interface travando, botão não funciona, log não aparece, adicionar campo,
  nova funcionalidade na janela, threading, callback_log, janela.after,
  engine não processa, pipeline com erro, main.py, tkinter, GUI Windows,
  adicionar período, novo botão, barra de progresso, seleção de arquivo.
  SEMPRE use este skill antes de editar main.py ou engine.py.
---

# Skill: Interface Gráfica e Engine de Orquestração

## Regra de ouro do Tkinter + threads

**Nunca escrever na GUI de dentro de uma thread que não seja a principal.**
O processamento de PDF roda em thread separada (para não travar a janela).
Para atualizar o log ou qualquer widget da janela de dentro da thread:

```python
# ERRADO — trava ou corrompe a GUI
self.area_log.insert("end", "mensagem")   # de dentro da thread

# CERTO — agenda a escrita na thread principal
self.janela.after(0, lambda: self._escrever_log("mensagem", "ok"))
```

## Estrutura do main.py

```
AplicacaoPericia
├── __init__()                    → configura janela, variáveis, chama os construtores
├── _construir_cabecalho()        → barra laranja do topo
├── _construir_area_selecao()     → campos de PDF, dropdown de cliente, botão processar
├── _construir_area_log()         → área de texto rolável com tags de cor
├── _construir_rodape()           → barra de status
│
├── _selecionar_pdf()             → abre filedialog (Windows Explorer)
├── _atualizar_lista_clientes()   → lê pasta clientes/ e preenche dropdown
│
├── _iniciar_processamento()      → valida campos, inicia thread
├── _executar_processamento()     → roda em thread; chama engine para cada PDF
├── _exibir_resumo()              → mostra resultado final no log
├── _finalizar_processamento()    → reabilita botão
│
├── _abrir_pasta_output()         → os.startfile(pasta output)
├── _log()                        → escreve no log via janela.after
└── _atualizar_status()           → atualiza rodapé
```

## Como adicionar um novo campo na interface

Exemplo: adicionar campo para "Competência (mês/ano)":

```python
# Em _construir_area_selecao(), após os campos de PDF existentes:
self.competencia = tk.StringVar()

tk.Label(container, text="Competência:", font=("Segoe UI", 10, "bold"),
         bg=CORES["fundo"], fg=CORES["texto"], width=22, anchor="w"
).grid(row=8, column=0, sticky="w", pady=4)

tk.Entry(container, textvariable=self.competencia,
         font=("Segoe UI", 10), bg=CORES["fundo_input"],
         fg=CORES["texto"], relief="flat"
).grid(row=8, column=1, sticky="ew", ipady=6)

tk.Label(container, text="Formato: MM/AAAA (ex: 01/2026)",
         font=("Segoe UI", 8), bg=CORES["fundo"], fg=CORES["texto_muted"]
).grid(row=9, column=1, sticky="w")
```

## Como abrir o Explorer do Windows

```python
from tkinter import filedialog
import os

# Abrir para selecionar UM arquivo
caminho = filedialog.askopenfilename(
    title="Selecione o balancete em PDF",
    initialdir=os.path.expanduser("~\\Desktop"),   # começa na Área de Trabalho
    filetypes=[("Arquivos PDF", "*.pdf"), ("Todos", "*.*")]
)
# caminho é string com o caminho completo, ou "" se cancelou

# Abrir para selecionar UMA pasta
pasta = filedialog.askdirectory(
    title="Selecione a pasta com os balancetes",
    initialdir=os.path.expanduser("~\\Desktop")
)

# Abrir para selecionar MÚLTIPLOS arquivos
arquivos = filedialog.askopenfilenames(
    title="Selecione os balancetes",
    filetypes=[("PDF", "*.pdf")]
)
# arquivos é uma tupla de caminhos
```

## Como escrever no log com cores

```python
# Tags disponíveis na área de log:
# "ok"     → verde   (sucesso)
# "erro"   → vermelho (erro)
# "aviso"  → amarelo  (aviso)
# "info"   → azul     (informação)
# "secao"  → laranja  (título de seção)
# "normal" → branco   (texto padrão)

def _log(self, mensagem: str, tipo: str = "normal"):
    """Método thread-safe para escrever no log."""
    self.janela.after(0, lambda m=mensagem, t=tipo: self._escrever_log(m, t))

def _escrever_log(self, mensagem: str, tipo: str):
    """Escreve de fato — só chamar da thread principal."""
    from datetime import datetime
    ts = datetime.now().strftime("%H:%M:%S")
    self.area_log.config(state="normal")
    self.area_log.insert("end", f"[{ts}] ", "normal")
    self.area_log.insert("end", f"{mensagem}\n", tipo)
    self.area_log.config(state="disabled")
    self.area_log.see("end")   # rola para o final
```

## Como estruturar o processamento em thread

```python
def _iniciar_processamento(self):
    # 1. Validar campos obrigatórios
    if not self.pdf_atual.get().strip():
        messagebox.showwarning("Campo vazio", "Selecione o PDF do mês atual.")
        return

    # 2. Travar botão para evitar duplo clique
    self.btn_processar.config(state="disabled", text="⏳ Processando...")

    # 3. Iniciar thread
    threading.Thread(
        target=self._executar_processamento,
        args=(self.pdf_atual.get(), self.pdf_anterior.get(), self.cliente_id.get()),
        daemon=True   # fecha automaticamente com a janela
    ).start()

def _executar_processamento(self, pdf_atual, pdf_anterior, cliente):
    try:
        # processamento aqui — pode demorar segundos
        ...
    except Exception as e:
        self._log(f"Erro inesperado: {e}", "erro")
    finally:
        # SEMPRE reabilitar o botão ao terminar
        self.janela.after(0, self._finalizar_processamento)

def _finalizar_processamento(self):
    self.btn_processar.config(state="normal", text="▶ PROCESSAR BALANCETES")
```

## Como o callback_log funciona no engine

O engine original usa `print()`. A versão GUI sobrescreve o engine com
uma versão que aceita `callback_log`:

```python
# Em _adaptar_engine_para_gui():
def processar_com_gui(pdf_path, cliente_id=None, output_dir="output",
                      verbose=True, callback_log=None):
    def log(msg, tipo="normal"):
        if callback_log:
            callback_log(msg, tipo)   # → self._log() da GUI
        elif verbose:
            print(msg)                # → terminal (quando sem GUI)
    ...
```

## Adicionar barra de progresso

```python
from tkinter import ttk

# Em _construir_area_selecao() ou novo método:
self.progresso = ttk.Progressbar(
    container, orient="horizontal", length=400, mode="determinate"
)
self.progresso.grid(row=10, column=0, columnspan=3, sticky="ew", pady=8)

# Atualizar de dentro da thread (via janela.after):
def _atualizar_progresso(self, valor: int):
    """valor entre 0 e 100"""
    self.janela.after(0, lambda: self.progresso.config(value=valor))
```

## Pipeline do engine (fluxo para referência)

```python
# engine.py — processar()
def processar(pdf_path, cliente_id=None, output_dir="output",
              verbose=True, callback_log=None):

    # Etapa 1: detectar layout
    layout = detectar_layout(pdf_path)
    meta   = extrair_metadados_cabecalho(pdf_path)

    # Etapa 2: parsear contas
    parser = get_parser(layout)
    contas = parser.parsear(pdf_path)
    if not contas:
        return {"status": "erro", "avisos": ["Nenhuma conta extraída"]}

    # Etapa 3: montar balancete
    balancete = Balancete(empresa=meta["empresa"], ..., contas=contas)

    # Etapa 4: validar
    erros = validar_balancete(balancete)
    balancete.erros_validacao = erros

    # Etapa 5: mapear (se cliente configurado)
    if cliente_id:
        config   = carregar_mapa_cliente(cliente_id)
        relatorio = mapear(balancete, config)
    else:
        relatorio = RelatorioFinal(balancete=balancete)

    # Etapa 6: exportar
    arquivo = exportar(relatorio, output_dir=output_dir)
    return {"status": "ok", "arquivo_gerado": arquivo, ...}
```

## Erros comuns na GUI

| Problema | Causa | Solução |
|---|---|---|
| Janela trava durante processamento | Processamento na thread principal | Mover para `threading.Thread` |
| Log não atualiza em tempo real | Escrita direta sem `janela.after` | Usar `janela.after(0, callback)` |
| `RuntimeError: main thread is not in main loop` | Widget atualizado de thread secundária | Usar `janela.after(0, ...)` |
| Botão fica desabilitado permanentemente | `_finalizar_processamento` não chamado | Garantir `finally` no `_executar_processamento` |
| `FileNotFoundError: MODELO_DE_PERICIA.xlsx` | `os.chdir()` não executado antes | Adicionar `os.chdir(Path(__file__).parent)` no `main()` |
