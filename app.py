"""Interface: o usuário arrasta os arquivos e recebe o Excel final."""
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import motor

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND = True
except ImportError:  # sem arrastar e soltar: os botões continuam funcionando
    DND = False

# Pasta onde o app está (funciona também depois de virar .exe)
BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
PASTA_SAIDA = BASE / "saida"

AZUL = "#1F3864"


def abrir(caminho: Path):
    if sys.platform.startswith("win"):
        os.startfile(caminho)
    elif sys.platform == "darwin":
        subprocess.run(["open", str(caminho)])
    else:
        subprocess.run(["xdg-open", str(caminho)])


class App:
    def __init__(self):
        self.root = TkinterDnD.Tk() if DND else tk.Tk()
        self.root.title("Conferência de Repasses Stripe")
        self.root.geometry("760x720")
        self.root.configure(padx=16, pady=12)

        self.repasses, self.cadastros = [], []
        self.fila = queue.Queue()
        self.ultimo_excel = None
        self.mascarar = tk.BooleanVar(value=False)

        tk.Label(self.root, text="Conferência de Repasses Stripe",
                 font=("Arial", 16, "bold"), fg=AZUL).pack(anchor="w")
        dica = ("Arraste os arquivos para as caixas abaixo (ou use o botão Adicionar)."
                if DND else "Use o botão Adicionar para escolher os arquivos.")
        tk.Label(self.root, text=dica, font=("Arial", 10)).pack(anchor="w", pady=(0, 8))

        self.lista_rep = self._zona("1. Relatórios de repasse da Stripe (obrigatório)", self.repasses)

        self.btn_gerar = tk.Button(self.root, text="GERAR EXCEL", font=("Arial", 12, "bold"),
                                   bg=AZUL, fg="white", height=2, command=self.gerar)
        self.btn_gerar.pack(fill="x", pady=8)

        self.barra = ttk.Progressbar(self.root, mode="indeterminate")
        self.barra.pack(fill="x")

        self.texto = tk.Text(self.root, height=12, font=("Consolas", 10), state="disabled")
        self.texto.pack(fill="both", expand=True, pady=8)

        rodape = tk.Frame(self.root)
        rodape.pack(fill="x")
        self.btn_excel = tk.Button(rodape, text="Abrir Excel", state="disabled",
                                   command=lambda: abrir(self.ultimo_excel))
        self.btn_excel.pack(side="left")
        self.btn_pasta = tk.Button(rodape, text="Abrir pasta de saída",
                                   command=lambda: (PASTA_SAIDA.mkdir(exist_ok=True), abrir(PASTA_SAIDA)))
        self.btn_pasta.pack(side="left", padx=8)

        self.root.after(100, self._ouvir_fila)

    # ----- zonas de arrastar -----
    def _zona(self, titulo, armazenamento):
        quadro = tk.LabelFrame(self.root, text=titulo, font=("Arial", 10, "bold"), padx=6, pady=6)
        quadro.pack(fill="x", pady=4)
        lista = tk.Listbox(quadro, height=4, font=("Arial", 9), selectmode="extended",
                           bg="#F2F5FA", relief="groove")
        lista.pack(side="left", fill="x", expand=True)

        botoes = tk.Frame(quadro)
        botoes.pack(side="right", padx=(8, 0))
        tk.Button(botoes, text="Adicionar...", width=12,
                  command=lambda: self._escolher(armazenamento, lista)).pack(pady=1)
        tk.Button(botoes, text="Remover", width=12,
                  command=lambda: self._remover(armazenamento, lista)).pack(pady=1)
        tk.Button(botoes, text="Limpar", width=12,
                  command=lambda: self._limpar(armazenamento, lista)).pack(pady=1)

        if DND:
            lista.drop_target_register(DND_FILES)
            lista.dnd_bind("<<Drop>>", lambda e: self._soltar(e, armazenamento, lista))
        return lista

    def _adicionar(self, caminhos, armazenamento, lista):
        ignorados = []
        for c in caminhos:
            p = Path(c)
            if p.suffix.lower() not in motor.EXTENSOES:
                ignorados.append(p.name)
            elif p not in armazenamento:
                armazenamento.append(p)
                lista.insert("end", p.name)
        if ignorados:
            messagebox.showwarning("Arquivo ignorado",
                                   "Só aceito .csv, .xlsx ou .xls. Ignorados:\n" + "\n".join(ignorados))

    def _soltar(self, evento, armazenamento, lista):
        self._adicionar(self.root.tk.splitlist(evento.data), armazenamento, lista)

    def _escolher(self, armazenamento, lista):
        caminhos = filedialog.askopenfilenames(
            title="Escolha as planilhas",
            filetypes=[("Planilhas", "*.csv *.xlsx *.xls"), ("Todos", "*.*")])
        self._adicionar(caminhos, armazenamento, lista)

    def _remover(self, armazenamento, lista):
        for i in reversed(lista.curselection()):
            lista.delete(i)
            armazenamento.pop(i)

    def _limpar(self, armazenamento, lista):
        lista.delete(0, "end")
        armazenamento.clear()

    # ----- processamento -----
    def _escrever(self, msg, limpar=False):
        self.texto.configure(state="normal")
        if limpar:
            self.texto.delete("1.0", "end")
        self.texto.insert("end", msg + "\n")
        self.texto.see("end")
        self.texto.configure(state="disabled")

    def gerar(self):
        if not self.repasses:
            messagebox.showinfo("Falta o relatório", "Adicione pelo menos um relatório de repasse.")
            return
        self.btn_gerar.configure(state="disabled")
        self.btn_excel.configure(state="disabled")
        self.barra.start(12)
        self._escrever("Iniciando...", limpar=True)

        def trabalho():
            try:
                res = motor.processar(
                    list(self.repasses), list(self.cadastros), PASTA_SAIDA,
                    self.mascarar.get(), log=lambda m: self.fila.put(("log", m)))
                self.fila.put(("ok", res))
            except Exception as e:  # mostra o erro para o usuário em vez de fechar
                self.fila.put(("erro", str(e)))

        threading.Thread(target=trabalho, daemon=True).start()

    def _ouvir_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._escrever(dado)
                elif tipo == "ok":
                    self._finalizar(dado)
                elif tipo == "erro":
                    self.barra.stop()
                    self.btn_gerar.configure(state="normal")
                    self._escrever(f"\nERRO: {dado}")
                    messagebox.showerror("Erro", dado)
        except queue.Empty:
            pass
        self.root.after(100, self._ouvir_fila)

    def _finalizar(self, r):
        self.barra.stop()
        self.btn_gerar.configure(state="normal")
        self.ultimo_excel = r["caminho"]
        self.btn_excel.configure(state="normal")
        self._escrever("\n===== RESULTADO =====")
        self._escrever(f"Linhas lidas: {r['qtd_lidas']}   Duplicatas removidas: {r['qtd_dup']}")
        self._escrever(f"Transações consolidadas: {r['qtd_final']}")
        self._escrever(r["totais"].to_string(index=False))
        self._escrever(f"\nExcel salvo em:\n{r['caminho']}")
        abrir(r["caminho"])

    def rodar(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().rodar()