#!/usr/bin/env python3
# coding: utf-8
#
# ╔═══════════════════════════════════════════════════╗
# ║  PERÍCIA CONTÁBIL — Extração de Balancetes        ║
# ║  Inocêncio de Paula Advogados                     ║
# ║                                                   ║
# ║  Execute: python main.py  ou  clique duas vezes   ║
# ╚═══════════════════════════════════════════════════╝

import sys
import os
import threading
from pathlib import Path
from datetime import datetime

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except ImportError:
    print("ERRO: tkinter não encontrado.")
    print("Solução: reinstale o Python marcando 'tcl/tk and IDLE'.")
    input("Pressione Enter para fechar...")
    sys.exit(1)


# ─────────────────────────────────────────────────────────────
# VERIFICAÇÃO DE DEPENDÊNCIAS
# ─────────────────────────────────────────────────────────────
def verificar_dependencias():
    faltando = []
    for nome_import, nome_pip in [
        ("pdfplumber", "pdfplumber"),
        ("openpyxl",   "openpyxl"),
        ("pydantic",   "pydantic"),
        ("yaml",       "PyYAML"),
    ]:
        try:
            __import__(nome_import)
        except ImportError:
            faltando.append(nome_pip)

    if faltando:
        root_temp = tk.Tk()
        root_temp.withdraw()
        messagebox.showerror(
            "Bibliotecas não instaladas",
            f"Instale as bibliotecas abaixo e execute novamente:\n\n"
            f"  pip install {' '.join(faltando)}"
        )
        root_temp.destroy()
        sys.exit(1)


# ─────────────────────────────────────────────────────────────
# PALETA DE CORES — Identidade Visual Inocêncio de Paula
# ─────────────────────────────────────────────────────────────
COR = {
    # Fundo geral
    "fundo":          "#0D0D0D",
    "fundo_card":     "#FFFFFF",
    "fundo_status":   "#111111",

    # Marca
    "borgonha":       "#7B1515",   # cor principal da marca
    "borgonha_hover": "#9B2020",
    "borgonha_escuro":"#5C0F0F",   # botão processar

    # Campos e bordas
    "input_bg":       "#F8F8F8",
    "borda":          "#E0E0E0",
    "borda_escura":   "#CCCCCC",

    # Textos
    "texto_escuro":   "#1A1A1A",   # texto em área branca
    "texto_claro":    "#FFFFFF",   # texto em área escura
    "texto_muted":    "#888888",   # texto secundário
    "texto_dica":     "#AAAAAA",   # dicas e paths

    # Status
    "ok":             "#2D6A4F",
    "erro":           "#C0392B",
    "aviso":          "#B7791F",
    "info":           "#2980B9",
}


# ─────────────────────────────────────────────────────────────
# APLICAÇÃO PRINCIPAL
# ─────────────────────────────────────────────────────────────
class AplicacaoPericia:

    def __init__(self, janela: tk.Tk):
        self.janela      = janela
        self.processando = False

        # Variáveis de estado
        self.pdf_atual    = tk.StringVar()
        self.pdf_anterior = tk.StringVar()
        self.cliente_id   = tk.StringVar()

        self._configurar_janela()
        self._construir_interface()
        self._atualizar_lista_clientes()

    # ──────────────────────────────────────────────────────
    # CONFIGURAÇÃO DA JANELA
    # ──────────────────────────────────────────────────────
    def _configurar_janela(self):
        self.janela.title("Perícia Contábil — Inocêncio de Paula")
        self.janela.configure(bg=COR["fundo"])
        self.janela.resizable(True, True)
        self.janela.minsize(820, 640)

        # Centraliza na tela
        self.janela.update_idletasks()
        w, h = 880, 680
        x = (self.janela.winfo_screenwidth()  // 2) - (w // 2)
        y = (self.janela.winfo_screenheight() // 2) - (h // 2)
        self.janela.geometry(f"{w}x{h}+{x}+{y}")

        # Ícone (ignora se não existir)
        try:
            self.janela.iconbitmap("icone.ico")
        except Exception:
            pass

    # ──────────────────────────────────────────────────────
    # CONSTRUÇÃO DA INTERFACE
    # ──────────────────────────────────────────────────────
    def _construir_interface(self):
        self._construir_cabecalho()
        self._construir_card_formulario()
        self._construir_painel_status()
        self._construir_rodape()

    def _construir_cabecalho(self):
        """
        Barra superior escura com logo IP + nome da empresa + título do sistema.
        """
        frame = tk.Frame(self.janela, bg=COR["borgonha_escuro"], height=72)
        frame.pack(fill="x")
        frame.pack_propagate(False)

        # Monograma IP
        frame_logo = tk.Frame(frame, bg="#4A0A0A", width=56, height=56)
        frame_logo.pack(side="left", padx=(16, 0), pady=8)
        frame_logo.pack_propagate(False)
        tk.Label(
            frame_logo, text="IP",
            font=("Georgia", 18, "bold"),
            bg="#4A0A0A", fg=COR["texto_claro"]
        ).place(relx=0.5, rely=0.5, anchor="center")

        # Nome da empresa
        frame_nomes = tk.Frame(frame, bg=COR["borgonha_escuro"])
        frame_nomes.pack(side="left", padx=12, pady=8)
        tk.Label(
            frame_nomes, text="INOCÊNCIO DE PAULA",
            font=("Segoe UI", 11, "bold"),
            bg=COR["borgonha_escuro"], fg=COR["texto_claro"]
        ).pack(anchor="w")
        tk.Label(
            frame_nomes, text="ADVOGADOS",
            font=("Segoe UI", 8),
            bg=COR["borgonha_escuro"], fg="#FFBBBB"
        ).pack(anchor="w")

        # Separador vertical
        tk.Frame(frame, bg="#FFFFFF", width=1).pack(side="left", fill="y", pady=16, padx=16)

        # Título do sistema
        frame_titulo = tk.Frame(frame, bg=COR["borgonha_escuro"])
        frame_titulo.pack(side="left", pady=8)
        tk.Label(
            frame_titulo, text="Perícia Contábil",
            font=("Segoe UI", 14, "bold"),
            bg=COR["borgonha_escuro"], fg=COR["texto_claro"]
        ).pack(anchor="w")
        tk.Label(
            frame_titulo, text="Extração e Fechamento de Balancetes",
            font=("Segoe UI", 9),
            bg=COR["borgonha_escuro"], fg="#FFBBBB"
        ).pack(anchor="w")

        # Versão (canto direito)
        tk.Label(
            frame, text="v1.1",
            font=("Segoe UI", 9, "bold"),
            bg=COR["borgonha"], fg=COR["texto_claro"],
            padx=10, pady=4
        ).pack(side="right", padx=16, pady=20)

    def _construir_card_formulario(self):
        """
        Card branco com campos de seleção de PDF e cliente.
        """
        # Sombra simulada com frame levemente cinza
        sombra = tk.Frame(self.janela, bg="#CCCCCC")
        sombra.pack(fill="x", padx=18, pady=(16, 0))

        card = tk.Frame(sombra, bg=COR["fundo_card"])
        card.pack(fill="x", padx=1, pady=1)

        inner = tk.Frame(card, bg=COR["fundo_card"])
        inner.pack(fill="x", padx=24, pady=20)

        # ── Linha: PDF Mês Atual ─────────────────────────────
        self._linha_pdf(
            parent=inner,
            label="Balancete Mês ATUAL  ★",
            variavel=self.pdf_atual,
            cor_label=COR["borgonha"],
            linha=0,
        )

        tk.Frame(inner, bg=COR["borda"], height=1).grid(
            row=1, column=0, columnspan=4, sticky="ew", pady=8
        )

        # ── Linha: PDF Mês Anterior ──────────────────────────
        self._linha_pdf(
            parent=inner,
            label="Balancete Mês ANTERIOR",
            variavel=self.pdf_anterior,
            cor_label=COR["texto_muted"],
            linha=2,
        )

        tk.Frame(inner, bg=COR["borda"], height=1).grid(
            row=3, column=0, columnspan=4, sticky="ew", pady=8
        )

        # ── Linha: Cliente ───────────────────────────────────
        tk.Label(
            inner, text="Empresa / Cliente:",
            font=("Segoe UI", 10, "bold"),
            bg=COR["fundo_card"], fg=COR["texto_escuro"],
            width=22, anchor="w"
        ).grid(row=4, column=0, sticky="w", pady=4)

        style = ttk.Style()
        style.configure("IP.TCombobox", fieldbackground=COR["input_bg"], font=("Segoe UI", 10))
        self.combo_clientes = ttk.Combobox(
            inner, textvariable=self.cliente_id,
            font=("Segoe UI", 10), state="readonly", width=42,
            style="IP.TCombobox"
        )
        self.combo_clientes.grid(row=4, column=1, columnspan=2, sticky="ew", padx=(0, 8))

        tk.Button(
            inner, text="↻ Atualizar",
            font=("Segoe UI", 9),
            bg=COR["input_bg"], fg=COR["texto_escuro"],
            activebackground=COR["borda_escura"],
            relief="flat", cursor="hand2", padx=10, pady=6,
            command=self._atualizar_lista_clientes
        ).grid(row=4, column=3)

        tk.Label(
            inner,
            text="Sem cliente selecionado → gera planilha só com a extração bruta do PDF",
            font=("Segoe UI", 8), bg=COR["fundo_card"], fg=COR["texto_muted"]
        ).grid(row=5, column=1, columnspan=3, sticky="w", pady=(0, 4))

        # ── Separador ────────────────────────────────────────
        tk.Frame(inner, bg=COR["borda"], height=1).grid(
            row=6, column=0, columnspan=4, sticky="ew", pady=12
        )

        # ── Botões ───────────────────────────────────────────
        frame_btns = tk.Frame(inner, bg=COR["fundo_card"])
        frame_btns.grid(row=7, column=0, columnspan=4)

        self.btn_processar = tk.Button(
            frame_btns,
            text="▶   PROCESSAR BALANCETES",
            font=("Segoe UI", 11, "bold"),
            bg=COR["borgonha"], fg=COR["texto_claro"],
            activebackground=COR["borgonha_hover"],
            activeforeground=COR["texto_claro"],
            relief="flat", cursor="hand2",
            padx=32, pady=12,
            command=self._iniciar_processamento
        )
        self.btn_processar.pack(side="left", padx=(0, 12))
        self.btn_processar.bind("<Enter>", lambda e: self.btn_processar.config(bg=COR["borgonha_hover"]))
        self.btn_processar.bind("<Leave>", lambda e: self.btn_processar.config(bg=COR["borgonha"]))

        tk.Button(
            frame_btns,
            text="📂  Abrir Pasta de Saída",
            font=("Segoe UI", 10),
            bg="#EFEFEF", fg=COR["texto_escuro"],
            activebackground=COR["borda_escura"],
            relief="flat", cursor="hand2",
            padx=16, pady=12,
            command=self._abrir_pasta_output
        ).pack(side="left")

        inner.columnconfigure(1, weight=1)

    def _linha_pdf(self, parent, label, variavel, cor_label, linha):
        """Cria uma linha completa de seleção de arquivo PDF."""
        tk.Label(
            parent, text=label,
            font=("Segoe UI", 10, "bold"),
            bg=COR["fundo_card"], fg=cor_label,
            width=22, anchor="w"
        ).grid(row=linha, column=0, sticky="w", pady=4)

        tk.Entry(
            parent, textvariable=variavel,
            font=("Segoe UI", 9),
            bg=COR["input_bg"], fg=COR["texto_escuro"],
            relief="flat", state="readonly",
            highlightthickness=1, highlightbackground=COR["borda"],
            highlightcolor=COR["borgonha"],
        ).grid(row=linha, column=1, columnspan=2, sticky="ew", padx=(0, 8), ipady=6)

        frame_btns = tk.Frame(parent, bg=COR["fundo_card"])
        frame_btns.grid(row=linha, column=3)

        tk.Button(
            frame_btns, text="📄 Selecionar",
            font=("Segoe UI", 9),
            bg=COR["borgonha"], fg=COR["texto_claro"],
            activebackground=COR["borgonha_hover"],
            relief="flat", cursor="hand2", padx=10, pady=5,
            command=lambda v=variavel: self._selecionar_pdf(v)
        ).pack(side="left", padx=(0, 4))

        tk.Button(
            frame_btns, text="✕",
            font=("Segoe UI", 9),
            bg="#EFEFEF", fg=COR["texto_muted"],
            activebackground=COR["borda_escura"],
            relief="flat", cursor="hand2", padx=6, pady=5,
            command=lambda v=variavel: v.set("")
        ).pack(side="left")

    def _construir_painel_status(self):
        """
        Painel de status simplificado — mostra mensagens legíveis, não código.
        """
        outer = tk.Frame(self.janela, bg=COR["fundo"], padx=18, pady=8)
        outer.pack(fill="both", expand=True)

        # Frame principal do painel
        painel = tk.Frame(outer, bg=COR["fundo_status"], bd=0)
        painel.pack(fill="both", expand=True)

        # Cabeçalho do painel
        cab = tk.Frame(painel, bg="#1A1A1A")
        cab.pack(fill="x")
        tk.Label(
            cab, text="  Progresso",
            font=("Segoe UI", 9, "bold"),
            bg="#1A1A1A", fg=COR["texto_muted"]
        ).pack(side="left", pady=6)

        # Ícone de status (grande, central)
        self.lbl_icone = tk.Label(
            painel, text="◉",
            font=("Segoe UI", 24),
            bg=COR["fundo_status"], fg="#333333"
        )
        self.lbl_icone.pack(pady=(20, 4))

        # Mensagem principal
        self.lbl_msg_principal = tk.Label(
            painel,
            text="Aguardando seleção dos arquivos...",
            font=("Segoe UI", 12),
            bg=COR["fundo_status"], fg=COR["texto_dica"]
        )
        self.lbl_msg_principal.pack()

        # Detalhe: empresa
        self.lbl_empresa = tk.Label(
            painel, text="",
            font=("Segoe UI", 10, "bold"),
            bg=COR["fundo_status"], fg=COR["texto_claro"]
        )
        self.lbl_empresa.pack(pady=(8, 0))

        # Detalhe: etapa atual
        self.lbl_etapa = tk.Label(
            painel, text="",
            font=("Segoe UI", 9),
            bg=COR["fundo_status"], fg=COR["texto_muted"]
        )
        self.lbl_etapa.pack()

        # Linha de avisos (aparece só quando há aviso)
        self.lbl_aviso = tk.Label(
            painel, text="",
            font=("Segoe UI", 9),
            bg=COR["fundo_status"], fg=COR["aviso"]
        )
        self.lbl_aviso.pack(pady=(4, 0))

    def _construir_rodape(self):
        """Barra inferior com caminho da pasta de saída."""
        frame = tk.Frame(self.janela, bg="#1A1A1A", height=26)
        frame.pack(fill="x", side="bottom")
        frame.pack_propagate(False)

        pasta_output = Path("output").resolve()
        tk.Label(
            frame,
            text=f"  Saída:  {pasta_output}",
            font=("Segoe UI", 8),
            bg="#1A1A1A", fg=COR["borgonha"]
        ).pack(side="left", padx=8)

        tk.Label(
            frame,
            text="Inocêncio de Paula Advogados  |  v1.1  ",
            font=("Segoe UI", 8),
            bg="#1A1A1A", fg="#444444"
        ).pack(side="right")

    # ──────────────────────────────────────────────────────
    # AÇÕES
    # ──────────────────────────────────────────────────────
    def _selecionar_pdf(self, variavel_destino):
        caminho = filedialog.askopenfilename(
            title="Selecione o balancete em PDF",
            initialdir=os.path.expanduser("~\\Desktop"),
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")]
        )
        if caminho:
            variavel_destino.set(caminho)
            self._set_status(f"Arquivo selecionado: {Path(caminho).name}", "info")

    def _atualizar_lista_clientes(self):
        pasta = Path("clientes")
        opcoes = ["(sem cliente — só extração bruta)"]
        if pasta.exists():
            opcoes += [f.stem for f in sorted(pasta.glob("*.yaml"))]
        self.combo_clientes["values"] = opcoes
        self.combo_clientes.current(0)
        n = len(opcoes) - 1
        if n > 0:
            self._set_status(f"{n} cliente(s) cadastrado(s): {', '.join(opcoes[1:])}", "info")
        else:
            self._set_status("Nenhum cliente cadastrado. Crie um arquivo em clientes/", "aviso")

    def _iniciar_processamento(self):
        if self.processando:
            return

        pdf_atual = self.pdf_atual.get().strip()
        if not pdf_atual:
            messagebox.showwarning("PDF não selecionado", "Selecione pelo menos o PDF do mês ATUAL.")
            return
        if not Path(pdf_atual).exists():
            messagebox.showerror("Arquivo não encontrado", f"O arquivo não foi encontrado:\n\n{pdf_atual}")
            return

        self.processando = True
        self.btn_processar.config(
            text="⏳  Processando...",
            state="disabled",
            bg=COR["texto_muted"]
        )
        self._set_status("Iniciando processamento...", "aguardando")

        pdf_anterior = self.pdf_anterior.get().strip() or None
        cliente_raw  = self.cliente_id.get()
        cliente      = None if "(sem cliente" in cliente_raw else cliente_raw

        thread = threading.Thread(
            target=self._executar_processamento,
            args=(pdf_atual, pdf_anterior, cliente),
            daemon=True
        )
        thread.start()

    def _executar_processamento(self, pdf_atual, pdf_anterior, cliente):
        try:
            from core.engine import processar
            resultado = processar(
                pdf_path=pdf_atual,
                pdf_anterior=pdf_anterior,
                cliente_id=cliente,
                output_dir="output",
                verbose=False,
                callback_log=self._log,
            )
            self._exibir_resultado(resultado)

        except ImportError as e:
            self._log(f"Biblioteca não instalada: {e}", "erro")
            self._log("Execute: pip install pdfplumber openpyxl pydantic PyYAML", "aviso")
        except Exception as e:
            import traceback
            self._log(f"Erro inesperado: {e}", "erro")
            self._log(traceback.format_exc(), "erro")
        finally:
            self.janela.after(0, self._finalizar_processamento)

    def _exibir_resultado(self, r: dict):
        status = r.get("status", "erro")
        if status == "ok":
            arquivo = Path(r.get("arquivo_gerado", "")).name
            n_contas = r.get("total_contas", 0)
            self._log(
                f"Concluído! {n_contas} contas extraídas. Arquivo: {arquivo}",
                "ok"
            )
            for aviso in r.get("avisos", []):
                self._log(aviso, "aviso")
        else:
            self._log("Processamento concluído com erros. Verifique abaixo.", "erro")
            for aviso in r.get("avisos", []):
                self._log(aviso, "aviso")

    def _finalizar_processamento(self):
        self.processando = False
        self.btn_processar.config(
            text="▶   PROCESSAR BALANCETES",
            state="normal",
            bg=COR["borgonha"]
        )

    def _abrir_pasta_output(self):
        pasta = Path("output")
        pasta.mkdir(exist_ok=True)
        os.startfile(str(pasta.resolve()))

    # ──────────────────────────────────────────────────────
    # SISTEMA DE LOG — atualiza o painel de status
    # ──────────────────────────────────────────────────────
    def _log(self, mensagem: str, tipo: str = "normal"):
        """
        Recebe mensagens do engine e atualiza o painel visual.
        Encaminha para a thread principal via janela.after().
        """
        self.janela.after(0, lambda m=mensagem, t=tipo: self._processar_log(m, t))

    def _processar_log(self, mensagem: str, tipo: str):
        """Atualiza os widgets de status conforme o tipo da mensagem."""

        # Etapas do processamento → mensagem principal
        if tipo == "status":
            self._set_status(mensagem, "processando")

        # Dados extraídos do PDF (empresa, sistema, período)
        elif tipo == "info":
            if mensagem.lower().startswith("empresa:"):
                empresa = mensagem.split(":", 1)[-1].strip()
                self.lbl_empresa.config(text=empresa)
            else:
                self.lbl_etapa.config(text=mensagem)

        # Aviso não crítico
        elif tipo == "aviso":
            self.lbl_aviso.config(text=f"⚠  {mensagem}", fg=COR["aviso"])

        # Sucesso final
        elif tipo == "ok":
            if "concluído" in mensagem.lower() or "gerada" in mensagem.lower() or "extraídas" in mensagem.lower():
                self._set_status(mensagem, "ok")

        # Erro
        elif tipo == "erro":
            self._set_status(mensagem, "erro")

    def _set_status(self, mensagem: str, tipo: str):
        """
        Atualiza o ícone e a mensagem principal do painel de status.
        Tipos: processando, ok, erro, aviso, info, aguardando
        """
        configs = {
            "processando": ("⏳", COR["borgonha"],    COR["texto_claro"]),
            "ok":          ("✓",  COR["ok"],           COR["texto_claro"]),
            "erro":        ("✗",  COR["erro"],         COR["texto_claro"]),
            "aviso":       ("⚠",  COR["aviso"],        COR["texto_claro"]),
            "info":        ("◉",  COR["info"],         COR["texto_claro"]),
            "aguardando":  ("◉",  "#333333",           COR["texto_dica"]),
        }
        icone, cor_icone, cor_msg = configs.get(tipo, configs["aguardando"])

        self.lbl_icone.config(text=icone, fg=cor_icone)
        self.lbl_msg_principal.config(text=mensagem, fg=cor_msg)

        if tipo not in ("processando", "aguardando"):
            self.lbl_etapa.config(text="")
            if tipo == "ok":
                self.lbl_aviso.config(text="")


# ─────────────────────────────────────────────────────────────
# PONTO DE ENTRADA
# ─────────────────────────────────────────────────────────────
def main():
    verificar_dependencias()
    os.chdir(Path(__file__).parent)

    janela = tk.Tk()
    AplicacaoPericia(janela)
    janela.mainloop()


if __name__ == "__main__":
    main()
