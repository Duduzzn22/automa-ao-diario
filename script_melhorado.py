import pyautogui
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from time import sleep
import threading
import openpyxl
import os
import unicodedata

# ──────────────────────────────────────────────────────────────────
#  CONTEXTO DO SITE (Conexão Educação):
#
#  - Cada linha = um aluno
#  - Cada coluna = um dia do mês
#  - O site já vem com TODOS os dias marcados como PRESENTE (✅)
#  - Para registrar FALTA: pressionar Space (desmarca o checkbox) + Tab
#  - Para manter PRESENÇA: apenas Tab (não toca no checkbox)
#  - Dias sem aula (vazio, Cancelado, etc.): apenas Tab para pular
#  - Tab avança para o próximo DIA do mesmo aluno (esquerda → direita)
#  - Ao terminar os dias de um aluno → Enter para ir ao próximo aluno
# ──────────────────────────────────────────────────────────────────

estado = {
    "rodando":       False,
    "pausado":       False,
    "alunos":        [],   # lista de dicionários: [ {"nome": "João", "dias": ['.', 'f', ...]}, ... ]
    "total_alunos":  0,
    "aluno_atual":   0,
    "total_celulas": 0,
    "celula_atual":  0,
}


# ──────────────────────────────────────────────────────────────────
#  Leitura do Excel
# ──────────────────────────────────────────────────────────────────
def ler_excel(caminho, aba, linha_inicio, linha_fim, col_inicio, col_fim, col_nome):
    """
    Lê o bloco de presenças do diário Excel e o nome dos alunos.
    Ignora alunos com status "Rem. de sala", "Cancelado" ou "Transf. de u.e.".
    """
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.worksheets[aba] if isinstance(aba, int) else wb[aba]

    alunos_ativos = []
    contadores = {
        "Ativos": 0,
        "Rem. de sala": 0,
        "Cancelado": 0,
        "Transf. de u.e.": 0
    }

    for row in range(linha_inicio, linha_fim + 1):
        nome_val = ws.cell(row=row, column=col_nome).value
        nome = str(nome_val).strip() if nome_val else f"Aluno da Linha {row}"
        
        status_aluno = "Ativos"
        nome_lower = nome.lower()
        
        # Verifica se o status está no próprio nome do aluno
        if "rem. de sala" in nome_lower or "remanejado" in nome_lower:
            status_aluno = "Rem. de sala"
        elif "cancelado" in nome_lower:
            status_aluno = "Cancelado"
        elif "transf. de u.e." in nome_lower or "transferido" in nome_lower:
            status_aluno = "Transf. de u.e."

        dias = []
        for col in range(col_inicio, col_fim + 1):
            val = ws.cell(row=row, column=col).value
            if val is None:
                dias.append(' ')
            else:
                s = str(val).strip().lower()
                
                # Verifica se o status foi escrito em alguma das células de dias
                if "rem. de sala" in s or "remanejado" in s:
                    status_aluno = "Rem. de sala"
                elif "cancelado" in s:
                    status_aluno = "Cancelado"
                elif "transf. de u.e." in s or "transf." in s:
                    status_aluno = "Transf. de u.e."
                

                if s == 'f':
                    dias.append('f')
                elif s == '.':
                    dias.append('.')
                # else:
                #     dias.append('.')   # Texto livre → pula

        if status_aluno != "Ativos":
            contadores[status_aluno] += 1
        else:
            contadores["Ativos"] += 1
            alunos_ativos.append({"nome": nome, "dias": dias})
            
    # Função auxiliar para remover acentos na hora de ordenar
    def normalizar_nome(aluno):
        nome_str = aluno["nome"].lower()
        return ''.join(c for c in unicodedata.normalize('NFD', nome_str) if unicodedata.category(c) != 'Mn')
        
    # Ordena os alunos em ordem alfabética ignorando acentos
    alunos_ativos.sort(key=normalizar_nome)
            
    return alunos_ativos, contadores


def carregar_excel():
    caminho = filedialog.askopenfilename(
        title="Selecione o diário Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")]
    )
    if not caminho:
        return

    try:
        aba = entrada_aba.get().strip()
        try:
            aba = int(aba)
        except ValueError:
            pass

        lin_ini = int(entrada_lin_ini.get())
        lin_fim = int(entrada_lin_fim.get())
        col_ini = int(entrada_col_ini.get())
        col_fim = int(entrada_col_fim.get())
        col_nome = int(entrada_col_nome.get())

        alunos_ativos, contadores = ler_excel(caminho, aba, lin_ini, lin_fim, col_ini, col_fim, col_nome)

        # Atualiza os contadores na interface
        lbl_ativos.config(text=f"Ativos (Processar): {contadores['Ativos']}")
        lbl_rem.config(text=f"Rem. de sala: {contadores['Rem. de sala']}")
        lbl_canc.config(text=f"Cancelado: {contadores['Cancelado']}")
        lbl_transf.config(text=f"Transf. de u.e.: {contadores['Transf. de u.e.']}")

        # Preview na área de texto
        area_texto.config(state="normal")
        area_texto.delete("1.0", tk.END)
        for aluno in alunos_ativos:
            area_texto.insert(tk.END, f"{aluno['nome']}: {''.join(aluno['dias'])}\n")
        area_texto.config(state="disabled")

        total_cel = sum(len(a["dias"]) for a in alunos_ativos)
        label_arquivo.config(text=f"📄 {os.path.basename(caminho)}")
        atualizar_status(
            f"✅ {len(alunos_ativos)} alunos ativos — {total_cel} células carregadas.", "ok"
        )
        estado["alunos"] = alunos_ativos

    except Exception as e:
        messagebox.showerror("Erro ao ler Excel", str(e))


# ──────────────────────────────────────────────────────────────────
#  Automação
# ──────────────────────────────────────────────────────────────────
def executar_automacao():
    if not estado["alunos"]:
        messagebox.showwarning(
            "Atenção", "Nenhum dado ativo carregado.\nCarregue o Excel primeiro."
        )
        return

    estado["total_alunos"]  = len(estado["alunos"])
    estado["aluno_atual"]   = 0
    estado["total_celulas"] = sum(len(a["dias"]) for a in estado["alunos"])
    estado["celula_atual"]  = 0
    estado["rodando"]       = True
    estado["pausado"]       = False

    barra_progresso["maximum"] = estado["total_alunos"]
    barra_progresso["value"]   = 0
    label_progresso.config(text=f"Iniciando: 0 / {estado['total_alunos']}")

    btn_iniciar.config(state="disabled")
    btn_pausar.config(state="normal", text="⏸  Pausar")
    btn_parar.config(state="normal")

    thread = threading.Thread(target=_loop_automacao, daemon=True)
    thread.start()


def _loop_automacao():
    delay       = float(entrada_delay.get())
    delay_aluno = float(entrada_delay_aluno.get())

    atualizar_status("Aguardando 4 s — clique na janela do site...", "aviso")
    sleep(4)
    pyautogui.press('tab')
    for idx_aluno, aluno in enumerate(estado["alunos"]):

        if not estado["rodando"]:
            break

        while estado["pausado"] and estado["rodando"]:
            sleep(0.2)

        janela.after(0, lambda i=idx_aluno, n=aluno["nome"]: _atualizar_status_aluno(i, n))

        # ── dias do aluno ──
        total_dias = len(aluno["dias"])
        for i, ch in enumerate(aluno["dias"]):
            if not estado["rodando"]:
                break
            while estado["pausado"] and estado["rodando"]:
                sleep(0.2)

            if ch == 'f':
                pyautogui.press('space')
                pyautogui.press('tab')
            else:
                pyautogui.press('tab')

            estado["celula_atual"] += 1


        # ── próximo aluno ──
        if estado["rodando"]:
            estado["aluno_atual"] = idx_aluno + 1
            janela.after(0, _atualizar_barra)
            sleep(delay_aluno)

    estado["rodando"] = False
    janela.after(0, _finalizar_ui)


def _atualizar_status_aluno(idx, nome_aluno):
    total   = estado["total_alunos"]
    cel_at  = estado["celula_atual"]
    cel_tot = estado["total_celulas"]
    pct     = int(cel_at / cel_tot * 100) if cel_tot else 0
    
    texto_prog = f"{nome_aluno} ({idx + 1} / {total})"
    # Limita o tamanho do texto para não quebrar o layout
    if len(texto_prog) > 50:
        texto_prog = texto_prog[:47] + "..."
        
    label_progresso.config(text=texto_prog)
    atualizar_status(f"Processando: {nome_aluno} ({pct}% concluído)", "ok")


def _atualizar_barra():
    barra_progresso["value"] = estado["aluno_atual"]


def _finalizar_ui():
    btn_iniciar.config(state="normal")
    btn_pausar.config(state="disabled", text="⏸  Pausar")
    btn_parar.config(state="disabled")
    total = estado["total_alunos"]
    atual = estado["aluno_atual"]
    if atual == total and total > 0:
        atualizar_status(f"✅ Concluído! {total} alunos processados.", "ok")
    else:
        atualizar_status(f"⛔ Interrompido após {atual} de {total} alunos.", "aviso")


def pausar_retomar():
    if not estado["rodando"]:
        return
    estado["pausado"] = not estado["pausado"]
    if estado["pausado"]:
        btn_pausar.config(text="▶  Retomar")
        atualizar_status("⏸ Pausado.", "aviso")
    else:
        btn_pausar.config(text="⏸  Pausar")
        atualizar_status("▶ Retomando…", "ok")


def parar():
    estado["rodando"] = False
    estado["pausado"] = False


def atualizar_status(msg, tipo="ok"):
    cores = {"ok": "#2ecc71", "aviso": "#f39c12", "erro": "#e74c3c"}
    label_status.config(text=msg, foreground=cores.get(tipo, "#ecf0f1"))


# ──────────────────────────────────────────────────────────────────
#  Interface gráfica
# ──────────────────────────────────────────────────────────────────
janela = tk.Tk()
janela.title("Frequência Automática — Conexão Educação")
janela.resizable(False, False)
janela.configure(bg="#1a1a2e")

style = ttk.Style(janela)
style.theme_use("clam")
style.configure("TFrame",       background="#1a1a2e")
style.configure("TLabel",       background="#1a1a2e", foreground="#ecf0f1", font=("Segoe UI", 10))
style.configure("TEntry",       fieldbackground="#16213e", foreground="#ecf0f1", insertcolor="#ecf0f1")
style.configure("TLabelframe",  background="#1a1a2e", foreground="#a29bfe", font=("Segoe UI", 10, "bold"))
style.configure("TLabelframe.Label", background="#1a1a2e", foreground="#a29bfe")
style.configure("cyan.Horizontal.TProgressbar",
                troughcolor="#16213e", background="#00cec9", thickness=18)

BG    = "#1a1a2e"
PANEL = "#16213e"
ROXO  = "#a29bfe"
VERDE = "#00cec9"
LRNJ  = "#fd9644"
AZUL  = "#0f3460"

# ── Título ──
tk.Label(janela, text="📋  Frequência Automática — Conexão Educação",
         bg=BG, fg=ROXO, font=("Segoe UI", 14, "bold")).pack(padx=20, pady=(16, 6), anchor="w")

# ── Configuração Excel ──
frame_excel = ttk.LabelFrame(janela, text=" 📊  Configuração do Excel ", padding=12)
frame_excel.pack(fill="x", padx=20, pady=6)

cfg = ttk.Frame(frame_excel)
cfg.pack(fill="x")

# linha 1
ttk.Label(cfg, text="Aba (nome ou 0,1,2…):").grid(row=0, column=0, sticky="w", padx=(0,4), pady=3)
entrada_aba = ttk.Entry(cfg, width=10); entrada_aba.insert(0, "Abr")
entrada_aba.grid(row=0, column=1, sticky="w", padx=(0,24), pady=3)

ttk.Label(cfg, text="Linha início dos alunos:").grid(row=0, column=2, sticky="w", padx=(0,4), pady=3)
entrada_lin_ini = ttk.Entry(cfg, width=6); entrada_lin_ini.insert(0, "11")
entrada_lin_ini.grid(row=0, column=3, sticky="w", pady=3)

# linha 2
ttk.Label(cfg, text="Linha fim dos alunos:").grid(row=1, column=0, sticky="w", padx=(0,4), pady=3)
entrada_lin_fim = ttk.Entry(cfg, width=6); entrada_lin_fim.insert(0, "0")
entrada_lin_fim.grid(row=1, column=1, sticky="w", padx=(0,24), pady=3)

ttk.Label(cfg, text="Coluna do Nome:").grid(row=1, column=2, sticky="w", padx=(0,4), pady=3)
entrada_col_nome = ttk.Entry(cfg, width=6); entrada_col_nome.insert(0, "3")
entrada_col_nome.grid(row=1, column=3, sticky="w", pady=3)

# linha 3
ttk.Label(cfg, text="Coluna início dos dias:").grid(row=2, column=0, sticky="w", padx=(0,4), pady=3)
entrada_col_ini = ttk.Entry(cfg, width=6); entrada_col_ini.insert(0, "4")
entrada_col_ini.grid(row=2, column=1, sticky="w", padx=(0,24), pady=3)

ttk.Label(cfg, text="Coluna fim dos dias:").grid(row=2, column=2, sticky="w", padx=(0,4), pady=3)
entrada_col_fim = ttk.Entry(cfg, width=6); entrada_col_fim.insert(0, "0")
entrada_col_fim.grid(row=2, column=3, sticky="w", pady=3)

# botão carregar
fb = ttk.Frame(frame_excel); fb.pack(fill="x", pady=(10, 0))
tk.Button(fb, text="📂  Carregar Excel", command=carregar_excel,
          bg=AZUL, fg="white", font=("Segoe UI", 10, "bold"),
          relief="flat", padx=12, pady=5, cursor="hand2").pack(side="left")
label_arquivo = ttk.Label(fb, text="nenhum arquivo selecionado", foreground="#636e72")
label_arquivo.pack(side="left", padx=12)

# ── Contadores ──
frame_cont = ttk.LabelFrame(janela, text=" 👥  Resumo dos Alunos ", padding=10)
frame_cont.pack(fill="x", padx=20, pady=6)

lbl_ativos = ttk.Label(frame_cont, text="Ativos: 0", foreground=VERDE, font=("Segoe UI", 10, "bold"))
lbl_ativos.pack(side="left", padx=(0, 15))

lbl_rem = ttk.Label(frame_cont, text="Rem. de sala: 0", foreground=LRNJ)
lbl_rem.pack(side="left", padx=15)

lbl_canc = ttk.Label(frame_cont, text="Cancelado: 0", foreground="#ff7675")
lbl_canc.pack(side="left", padx=15)

lbl_transf = ttk.Label(frame_cont, text="Transf. de u.e.: 0", foreground="#fdcb6e")
lbl_transf.pack(side="left", padx=15)

# ── Preview ──
frame_prev = ttk.LabelFrame(
    janela, text=" 👁  Preview  ( . = presente   f = falta   espaço = pular ) ", padding=10
)
frame_prev.pack(fill="x", padx=20, pady=6)
scroll_y = tk.Scrollbar(frame_prev)
scroll_y.pack(side="right", fill="y")
area_texto = tk.Text(frame_prev, height=7, bg=PANEL, fg="#ecf0f1",
                     font=("Consolas", 10), relief="flat", wrap="none",
                     yscrollcommand=scroll_y.set, state="disabled")
area_texto.pack(fill="x")
scroll_y.config(command=area_texto.yview)

# ── Velocidade ──
frame_vel = ttk.LabelFrame(janela, text=" ⚙️  Velocidade ", padding=12)
frame_vel.pack(fill="x", padx=20, pady=6)

fv = ttk.Frame(frame_vel); fv.pack(fill="x")
ttk.Label(fv, text="Delay entre dias (seg):").grid(row=0, column=0, sticky="w", padx=(0,6))
entrada_delay = ttk.Entry(fv, width=7); entrada_delay.insert(0, "0.05")
entrada_delay.grid(row=0, column=1, sticky="w", padx=(0,30))

ttk.Label(fv, text="Delay entre alunos (seg):").grid(row=0, column=2, sticky="w", padx=(0,6))
entrada_delay_aluno = ttk.Entry(fv, width=7); entrada_delay_aluno.insert(0, "0.3")
entrada_delay_aluno.grid(row=0, column=3, sticky="w")

# Botões
fb2 = ttk.Frame(frame_vel); fb2.pack(fill="x", pady=(12, 0))
btn_iniciar = tk.Button(fb2, text="▶  Iniciar", command=executar_automacao,
                        bg=VERDE, fg="#1a1a2e", font=("Segoe UI", 11, "bold"),
                        relief="flat", padx=16, pady=7, cursor="hand2")
btn_iniciar.pack(side="left", padx=(0, 8))

btn_pausar = tk.Button(fb2, text="⏸  Pausar", command=pausar_retomar, state="disabled",
                       bg=LRNJ, fg="#1a1a2e", font=("Segoe UI", 11, "bold"),
                       relief="flat", padx=16, pady=7, cursor="hand2")
btn_pausar.pack(side="left", padx=(0, 8))

btn_parar = tk.Button(fb2, text="⏹  Parar", command=parar, state="disabled",
                      bg="#d63031", fg="white", font=("Segoe UI", 11, "bold"),
                      relief="flat", padx=16, pady=7, cursor="hand2")
btn_parar.pack(side="left")

# ── Progresso ──
frame_prog = ttk.LabelFrame(janela, text=" 📈  Progresso ", padding=10)
frame_prog.pack(fill="x", padx=20, pady=6)
barra_progresso = ttk.Progressbar(frame_prog, style="cyan.Horizontal.TProgressbar",
                                  orient="horizontal", mode="determinate")
barra_progresso.pack(fill="x")
label_progresso = ttk.Label(frame_prog, text="Aluno 0 / 0")
label_progresso.pack(pady=(4, 0))

# ── Status ──
label_status = ttk.Label(janela,
                         text="Pronto. Configure os campos e carregue o Excel.",
                         foreground=VERDE, font=("Segoe UI", 10))
label_status.pack(pady=(6, 18))

janela.mainloop()