import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pandas as pd
import os

# Nome do arquivo Excel
ARQUIVO_EXCEL = "movimentos.xlsx"

# =========================
# Funções auxiliares
# =========================

def carregar_dados():
    if os.path.exists(ARQUIVO_EXCEL):
        return pd.read_excel(ARQUIVO_EXCEL)
    else:
        return pd.DataFrame(columns=["Valor", "Data Movimento", "Data Lançamento", "Nota", "Tipo"])

def salvar_dados(df):
    df.to_excel(ARQUIVO_EXCEL, index=False)

def atualizar_historico(filtro_inicio=None, filtro_fim=None):
    historico.delete(*historico.get_children())
    df = carregar_dados()
    saldo = 0

    if filtro_inicio and filtro_fim:
        try:
            inicio = datetime.strptime(filtro_inicio, "%d/%m/%Y")
            fim = datetime.strptime(filtro_fim, "%d/%m/%Y")
            df["Data Movimento"] = pd.to_datetime(df["Data Movimento"], dayfirst=True, errors="coerce")
            df = df[(df["Data Movimento"] >= inicio) & (df["Data Movimento"] <= fim)]
        except Exception:
            messagebox.showerror("Erro", "Datas de filtro inválidas. Use formato DD/MM/AAAA.")
            return

    for _, row in df.iterrows():
        valor = float(row["Valor"])
        tipo = row["Tipo"]

        if tipo == "Entrada":
            saldo += valor
        else:
            saldo -= valor

        historico.insert("", "end", values=(
            f"R$ {valor:.2f}",
            row["Data Movimento"].strftime("%d/%m/%Y") if isinstance(row["Data Movimento"], pd.Timestamp) else row["Data Movimento"],
            row["Data Lançamento"],
            row["Nota"],
            tipo,
            f"R$ {saldo:.2f}"
        ))

    saldo_label.config(text=f"💰 Saldo Total: R$ {saldo:.2f}")

def adicionar_movimento():
    try:
        valor = float(entry_valor.get())
        data_movimento = entry_data_movimento.get().strip()
        nota = entry_nota.get().strip()
        tipo = tipo_var.get()

        if not data_movimento:
            messagebox.showerror("Erro", "Preencha a data de movimento!")
            return

        try:
            datetime.strptime(data_movimento, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Erro", "A data de movimento deve estar no formato DD/MM/AAAA.")
            return

        data_lancamento = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        df = carregar_dados()
        novo = pd.DataFrame([{
            "Valor": round(valor, 2),
            "Data Movimento": data_movimento,
            "Data Lançamento": data_lancamento,
            "Nota": nota,
            "Tipo": tipo
        }])
        df = pd.concat([df, novo], ignore_index=True)
        salvar_dados(df)
        atualizar_historico()

        entry_valor.delete(0, tk.END)
        entry_data_movimento.delete(0, tk.END)
        entry_nota.delete(0, tk.END)
        tipo_var.set("Entrada")

    except ValueError:
        messagebox.showerror("Erro", "Digite um valor numérico válido.")

def aplicar_filtro():
    inicio = entry_filtro_inicio.get().strip()
    fim = entry_filtro_fim.get().strip()
    if not inicio or not fim:
        messagebox.showerror("Erro", "Preencha as duas datas do filtro.")
        return
    atualizar_historico(inicio, fim)

def limpar_filtro():
    entry_filtro_inicio.delete(0, tk.END)
    entry_filtro_fim.delete(0, tk.END)
    atualizar_historico()


def carregar_historico():
    global ARQUIVO

    arquivo = filedialog.askopenfilename(
        title="Selecione a planilha de histórico",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if not arquivo:
        return

    ARQUIVO = arquivo
    try:
        df = pd.read_excel(ARQUIVO)
        atualizar_historico(df)
        messagebox.showinfo("Sucesso", f"Histórico carregado de:\n{ARQUIVO}")
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível carregar o arquivo:\n{e}")

# =========================
# Interface Tkinter
# =========================
root = tk.Tk()
root.title("Fluxo de Caixa Simples")
root.geometry("900x800")

# Valor
tk.Label(root, text="Valor (R$):").pack()
entry_valor = tk.Entry(root)
entry_valor.pack()

# Data de Movimento
tk.Label(root, text="Data de Movimento (DD/MM/AAAA):").pack()
entry_data_movimento = tk.Entry(root)
entry_data_movimento.pack()

# Nota
tk.Label(root, text="Nota:").pack()
entry_nota = tk.Entry(root)
entry_nota.pack()

# Tipo (Entrada ou Saída)
tk.Label(root, text="Tipo:").pack()
tipo_var = tk.StringVar(value="Entrada")
tipo_dropdown = ttk.Combobox(root, textvariable=tipo_var, values=["Entrada", "Saída"], state="readonly")
tipo_dropdown.pack()

# Botão de adicionar (mais amigável)
btn_add = tk.Button(root, text="➕ Adicionar Movimento", command=adicionar_movimento,
                    bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
btn_add.pack(pady=10)

# Botão de carregar histórico (corrigido)
btn_carregar = tk.Button(root, text="📂 Carregar Histórico", command=carregar_historico,
                         bg="#2196F3", fg="white", font=("Arial", 12, "bold"))
btn_carregar.pack(pady=10)

# Filtros
filtro_frame = tk.Frame(root)
filtro_frame.pack(pady=10)

tk.Label(filtro_frame, text="Filtro por Data de Movimento:").grid(row=0, column=0, columnspan=2)

tk.Label(filtro_frame, text="Início (DD/MM/AAAA):").grid(row=1, column=0)
entry_filtro_inicio = tk.Entry(filtro_frame)
entry_filtro_inicio.grid(row=1, column=1)

tk.Label(filtro_frame, text="Fim (DD/MM/AAAA):").grid(row=2, column=0)
entry_filtro_fim = tk.Entry(filtro_frame)
entry_filtro_fim.grid(row=2, column=1)

btn_filtrar = tk.Button(filtro_frame, text="🔎 Aplicar Filtro", command=aplicar_filtro,
                        bg="#2196F3", fg="white", font=("Arial", 10, "bold"))
btn_filtrar.grid(row=3, column=0, pady=5, padx=5)

btn_limpar = tk.Button(filtro_frame, text="🧹 Limpar Filtro", command=limpar_filtro,
                       bg="#f44336", fg="white", font=("Arial", 10, "bold"))
btn_limpar.grid(row=3, column=1, pady=5, padx=5)

# Histórico
tk.Label(root, text="Histórico de Movimentos:").pack()
historico = ttk.Treeview(root, columns=("Valor", "DataMov", "DataLanc", "Nota", "Tipo", "Saldo"), show="headings", height=15)
historico.pack(fill="both", expand=True)

historico.heading("Valor", text="Valor")
historico.heading("DataMov", text="Data Movimento")
historico.heading("DataLanc", text="Data Lançamento")
historico.heading("Nota", text="Nota")
historico.heading("Tipo", text="Tipo")
historico.heading("Saldo", text="Saldo Acumulado")

# Saldo total
saldo_label = tk.Label(root, text="💰 Saldo Total: R$ 0.00", font=("Arial", 14, "bold"))
saldo_label.pack(pady=10)

# Inicializa histórico
atualizar_historico()

root.mainloop()
