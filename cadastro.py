import tkinter as tk
from tkinter import messagebox, Toplevel, filedialog, colorchooser
from tkinter import ttk
from tkcalendar import DateEntry
import pandas as pd
import os
import re
from PIL import Image, ImageTk

class CadastroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cadastro de Cliente")
        self.root.geometry("400x520")
        self.root.configure(bg="#e6f2ff")

        # Adiciona o logo centralizado
        logo_img = Image.open("C:\\Users\\Ping13\\Desktop\\cadastroCliente\\images.png")
        logo_img = logo_img.resize((80, 80))
        logo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(root, image=logo, bg="#e6f2ff")
        logo_label.image = logo
        logo_label.grid(row=0, column=0, columnspan=4, pady=10, sticky="nsew")
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=1)
        root.grid_columnconfigure(2, weight=1)
        root.grid_columnconfigure(3, weight=1)

        campos = [
            ("Nome Completo:", "nome"),
            ("Email:", "email"),
            ("CPF:", "cpf"),
            ("Telefone:", "telefone"),
            ("Data de Nascimento:", "nascimento"),
            ("Endereço:", "endereco"),
            ("Número da Residência:", "numero"),
            ("CEP:", "cep"),
            ("Cidade:", "cidade")
        ]

        self.entries = {}
        for i, (label_text, key) in enumerate(campos):
            tk.Label(root, text=label_text, bg="#e6f2ff").grid(row=i+1, column=0, padx=10, pady=5, sticky="e")
            if key == "nascimento":
                entry = DateEntry(root, width=27, date_pattern='dd/mm/yyyy')
            else:
                entry = tk.Entry(root, width=30, bg="#ffffff")
                if key == "cpf":
                    entry.bind("<KeyRelease>", self.mascara_cpf)
                elif key == "telefone":
                    entry.bind("<KeyRelease>", self.mascara_telefone)
                elif key == "cep":
                    entry.bind("<KeyRelease>", self.mascara_cep)
            entry.grid(row=i+1, column=1, padx=10, pady=5)
            self.entries[key] = entry

        tk.Button(root, text="Cadastrar", command=self.cadastrar_usuario, bg="#e6f2ff").grid(row=len(campos)+1, column=0, columnspan=2, pady=10)
        tk.Button(root, text="Visualizar Clientes", command=self.visualizar_clientes, bg="#e6f2ff").grid(row=len(campos)+2, column=0, columnspan=2, pady=10)

        # Botão para alterar cor do sistema (embaixo)
        def escolher_cor():
            cor = colorchooser.askcolor(title="Escolha a cor do fundo")[1]
            if cor:
                # Detecta se o fundo é escuro (média dos valores RGB)
                rgb = self.root.winfo_rgb(cor)
                media = sum([v//256 for v in rgb]) / 3
                cor_fonte = "white" if media < 128 else "black"

                self.root.configure(bg=cor)
                logo_label.configure(bg=cor)
                for widget in root.winfo_children():
                    if isinstance(widget, (tk.Label, tk.Entry, tk.Button)) and not isinstance(widget, DateEntry):
                        widget.configure(bg=cor, fg=cor_fonte)
                for widget in root.winfo_children():
                    if isinstance(widget, DateEntry):
                        widget.configure(background=cor, foreground=cor_fonte)

        tk.Button(root, text="Alterar cor do sistema", command=escolher_cor, bg="#e6f2ff").grid(row=len(campos)+3, column=0, columnspan=2, pady=10)

    # Máscaras:
    def mascara_cpf(self, event):
        texto = re.sub(r'\D', '', self.entries["cpf"].get())[:11]
        novo_texto = ""
        for i in range(len(texto)):
            if i in [3,6]:
                novo_texto += '.'
            if i == 9:
                novo_texto += '-'
            novo_texto += texto[i]
        self.entries["cpf"].delete(0, tk.END)
        self.entries["cpf"].insert(0, novo_texto)

    def mascara_telefone(self, event):
        texto = re.sub(r'\D', '', self.entries["telefone"].get())[:11]
        novo_texto = ""
        if len(texto) > 0:
            novo_texto += "("
        if len(texto) >= 2:
            novo_texto += texto[:2] + ") "
            if len(texto) > 6:
                novo_texto += texto[2:7] + "-"
                novo_texto += texto[7:]
            elif len(texto) > 2:
                novo_texto += texto[2:]
        else:
            novo_texto += texto
        self.entries["telefone"].delete(0, tk.END)
        self.entries["telefone"].insert(0, novo_texto)

    def mascara_cep(self, event):
        texto = re.sub(r'\D', '', self.entries["cep"].get())[:8]
        novo_texto = ""
        for i in range(len(texto)):
            if i == 5:
                novo_texto += "-"
            novo_texto += texto[i]
        self.entries["cep"].delete(0, tk.END)
        self.entries["cep"].insert(0, novo_texto)

    def cadastrar_usuario(self):
        dados = {key: self.entries[key].get().strip() for key in self.entries}

        if not self.validar_nome(dados["nome"]):
            messagebox.showerror("Erro", "Nome completo inválido.")
            return
        if not self.validar_email(dados["email"]):
            messagebox.showerror("Erro", "Email inválido.")
            return
        if not self.validar_cpf(dados["cpf"]):
            messagebox.showerror("Erro", "CPF inválido.")
            return
        if not self.validar_telefone(dados["telefone"]):
            messagebox.showerror("Erro", "Telefone inválido.")
            return

        dados["cpf"] = re.sub(r'\D', '', dados["cpf"])
        dados["telefone"] = re.sub(r'\D', '', dados["telefone"])
        dados["cep"] = re.sub(r'\D', '', dados["cep"])

        df_novo = pd.DataFrame([{
            'Nome Completo': dados["nome"],
            'Email': dados["email"],
            'CPF': dados["cpf"],
            'Telefone': dados["telefone"],
            'Data de Nascimento': self.entries["nascimento"].get_date(),
            'Endereço': dados["endereco"],
            'Número da Residência': dados["numero"],
            'CEP': dados["cep"],
            'Cidade': dados["cidade"]
        }])

        if os.path.isfile("cadastros.xlsx"):
            try:
                df_existente = pd.read_excel("cadastros.xlsx")
                df = pd.concat([df_existente, df_novo], ignore_index=True)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler arquivo: {e}")
                return
        else:
            df = df_novo

        try:
            df.to_excel("cadastros.xlsx", index=False)
            messagebox.showinfo("Sucesso", "Cadastro salvo com sucesso!")
            self.limpar_campos()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {e}")

    def visualizar_clientes(self):
        janela = Toplevel(self.root)
        janela.title("Clientes Cadastrados")
        janela.geometry("1100x500")

        frame = tk.Frame(janela)
        frame.pack(fill="both", expand=True)

        search_label = tk.Label(frame, text="Buscar CPF:")
        search_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        cpf_entry = tk.Entry(frame)
        cpf_entry.grid(row=0, column=1, padx=5, pady=5)

        col_titles = ["Nome Completo", "Email", "CPF", "Telefone", "Data de Nascimento",
                      "Endereço", "Número da Residência", "CEP", "Cidade"]
        tree = ttk.Treeview(frame, columns=col_titles, show='headings')
        for col in col_titles:
            tree.heading(col, text=col)
            tree.column(col, width=120)

        tree.grid(row=1, column=0, columnspan=6, sticky='nsew')

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=1, column=6, sticky='ns')

        frame.grid_columnconfigure(1, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        def carregar_dados(filtro_cpf=None):
            for item in tree.get_children():
                tree.delete(item)
            if os.path.exists("cadastros.xlsx"):
                try:
                    df = pd.read_excel("cadastros.xlsx")

                    if 'Data de Nascimento' in df.columns:
                        coluna_nascimento = 'Data de Nascimento'
                    elif 'Nascimento' in df.columns:
                        coluna_nascimento = 'Nascimento'
                    else:
                        coluna_nascimento = None

                    for col in col_titles:
                        if col not in df.columns and col != "Data de Nascimento":
                            df[col] = ""

                    if coluna_nascimento and coluna_nascimento in df.columns:
                        df[coluna_nascimento] = pd.to_datetime(df[coluna_nascimento], errors='coerce').dt.strftime('%d/%m/%Y').fillna('')

                    if filtro_cpf:
                        filtro_cpf_numeros = re.sub(r'\D', '', filtro_cpf)
                        df = df[df["CPF"].astype(str).str.contains(filtro_cpf_numeros)]

                    if coluna_nascimento and coluna_nascimento != "Data de Nascimento":
                        df = df.rename(columns={coluna_nascimento: "Data de Nascimento"})
                    elif not coluna_nascimento:
                        df["Data de Nascimento"] = ""

                    for _, row in df.iterrows():
                        tree.insert("", "end", values=(
                            row.get('Nome Completo', ''),
                            row.get('Email', ''),
                            self.formatar_cpf(str(row.get('CPF', ''))),
                            self.formatar_telefone(str(row.get('Telefone', ''))),
                            row.get('Data de Nascimento', ''),
                            row.get('Endereço', ''),
                            row.get('Número da Residência', ''),
                            self.formatar_cep(str(row.get('CEP', ''))),
                            row.get('Cidade', '')
                        ))
                except Exception as e:
                    messagebox.showerror("Erro ao carregar dados", str(e))

        def buscar():
            cpf = cpf_entry.get().strip()
            carregar_dados(cpf)

        def editar():
            item = tree.selection()
            if not item:
                messagebox.showwarning("Aviso", "Selecione um cliente para editar.")
                return

            valores = tree.item(item, "values")

            edit_win = Toplevel(janela)
            edit_win.title("Editar Cliente")

            labels = col_titles
            fields = []
            for i, label in enumerate(labels):
                tk.Label(edit_win, text=label).grid(row=i, column=0, padx=10, pady=5)
                if label == "Data de Nascimento":
                    entry = DateEntry(edit_win, width=37, date_pattern='dd/mm/yyyy')
                    try:
                        entry.set_date(pd.to_datetime(valores[i], dayfirst=True))
                    except Exception:
                        entry.set_date(pd.Timestamp.today())
                else:
                    entry = tk.Entry(edit_win, width=40)
                    entry.insert(0, valores[i])
                    if label == "CPF":
                        entry.bind("<KeyRelease>", lambda e, ent=entry: self.mascara_cpf_edit(ent))
                    elif label == "Telefone":
                        entry.bind("<KeyRelease>", lambda e, ent=entry: self.mascara_telefone_edit(ent))
                    elif label == "CEP":
                        entry.bind("<KeyRelease>", lambda e, ent=entry: self.mascara_cep_edit(ent))
                entry.grid(row=i, column=1, padx=10, pady=5)
                fields.append(entry)

            def salvar_edicao():
                dados_editados = {}
                for i, col in enumerate(col_titles):
                    if col == "Data de Nascimento":
                        dados_editados[col] = fields[i].get_date()
                    else:
                        dados_editados[col] = fields[i].get().strip()

                # Validações
                if not self.validar_nome(dados_editados["Nome Completo"]):
                    messagebox.showerror("Erro", "Nome completo inválido.")
                    return
                if not self.validar_email(dados_editados["Email"]):
                    messagebox.showerror("Erro", "Email inválido.")
                    return
                if not self.validar_cpf(dados_editados["CPF"]):
                    messagebox.showerror("Erro", "CPF inválido.")
                    return
                if not self.validar_telefone(dados_editados["Telefone"]):
                    messagebox.showerror("Erro", "Telefone inválido.")
                    return

                dados_editados["CPF"] = re.sub(r'\D', '', dados_editados["CPF"])
                dados_editados["Telefone"] = re.sub(r'\D', '', dados_editados["Telefone"])
                dados_editados["CEP"] = re.sub(r'\D', '', dados_editados["CEP"])

                if os.path.exists("cadastros.xlsx"):
                    try:
                        df = pd.read_excel("cadastros.xlsx")
                        cpf_original = re.sub(r'\D', '', valores[2])
                        idx = df.index[df['CPF'].astype(str) == cpf_original]
                        if len(idx) == 0:
                            messagebox.showerror("Erro", "Cliente não encontrado para editar.")
                            return
                        idx = idx[0]

                        df.at[idx, 'Nome Completo'] = dados_editados["Nome Completo"]
                        df.at[idx, 'Email'] = dados_editados["Email"]
                        df.at[idx, 'CPF'] = dados_editados["CPF"]
                        df.at[idx, 'Telefone'] = dados_editados["Telefone"]
                        df.at[idx, 'Data de Nascimento'] = dados_editados["Data de Nascimento"]
                        df.at[idx, 'Endereço'] = dados_editados["Endereço"]
                        df.at[idx, 'Número da Residência'] = dados_editados["Número da Residência"]
                        df.at[idx, 'CEP'] = dados_editados["CEP"]
                        df.at[idx, 'Cidade'] = dados_editados["Cidade"]

                        df.to_excel("cadastros.xlsx", index=False)
                        messagebox.showinfo("Sucesso", "Cadastro editado com sucesso!")
                        carregar_dados()
                        edit_win.destroy()
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao editar cadastro: {e}")
                else:
                    messagebox.showerror("Erro", "Arquivo de cadastro não encontrado.")

            btn_salvar = tk.Button(edit_win, text="Salvar", command=salvar_edicao)
            btn_salvar.grid(row=len(col_titles), column=0, columnspan=2, pady=10)

        def exportar_excel():
            if os.path.exists("cadastros.xlsx"):
                caminho = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
                    title="Salvar planilha como"
                )
                if caminho:
                    try:
                        df = pd.read_excel("cadastros.xlsx")
                        df.to_excel(caminho, index=False)
                        messagebox.showinfo("Sucesso", "Arquivo exportado com sucesso!")
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao exportar arquivo: {e}")
            else:
                messagebox.showerror("Erro", "Nenhum cadastro para exportar.")

        tk.Button(frame, text="Buscar", command=buscar).grid(row=0, column=2, padx=5, pady=5)
        tk.Button(frame, text="Editar Selecionado", command=editar).grid(row=2, column=0, pady=10)
        tk.Button(frame, text="Exportar para Excel", command=exportar_excel).grid(row=2, column=1, pady=10)

        carregar_dados()

    # Máscaras para edição (usadas na tela de editar)
    def mascara_cpf_edit(self, entry):
        texto = re.sub(r'\D', '', entry.get())[:11]
        novo_texto = ""
        for i in range(len(texto)):
            if i in [3,6]:
                novo_texto += '.'
            if i == 9:
                novo_texto += '-'
            novo_texto += texto[i]
        entry.delete(0, tk.END)
        entry.insert(0, novo_texto)

    def mascara_telefone_edit(self, entry):
        texto = re.sub(r'\D', '', entry.get())[:11]
        novo_texto = ""
        if len(texto) > 0:
            novo_texto += "("
        if len(texto) >= 2:
            novo_texto += texto[:2] + ") "
            if len(texto) > 6:
                novo_texto += texto[2:7] + "-"
                novo_texto += texto[7:]
            elif len(texto) > 2:
                novo_texto += texto[2:]
        else:
            novo_texto += texto
        entry.delete(0, tk.END)
        entry.insert(0, novo_texto)

    def mascara_cep_edit(self, entry):
        texto = re.sub(r'\D', '', entry.get())[:8]
        novo_texto = ""
        for i in range(len(texto)):
            if i == 5:
                novo_texto += "-"
            novo_texto += texto[i]
        entry.delete(0, tk.END)
        entry.insert(0, novo_texto)

    # Formatações para visualização
    def formatar_cpf(self, cpf):
        cpf = re.sub(r'\D', '', cpf)[:11]
        if len(cpf) == 11:
            return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        return cpf

    def formatar_telefone(self, telefone):
        telefone = re.sub(r'\D', '', telefone)[:11]
        if len(telefone) == 11:
            return f"({telefone[:2]}) {telefone[2:7]}-{telefone[7:]}"
        elif len(telefone) == 10:
            return f"({telefone[:2]}) {telefone[2:6]}-{telefone[6:]}"
        return telefone

    def formatar_cep(self, cep):
        cep = re.sub(r'\D', '', cep)[:8]
        if len(cep) == 8:
            return f"{cep[:5]}-{cep[5:]}"
        return cep

    # Validações básicas
    def validar_nome(self, nome):
        return len(nome) > 2

    def validar_email(self, email):
        return "@" in email and "." in email

    def validar_cpf(self, cpf):
        cpf = re.sub(r'\D', '', cpf)
        return len(cpf) == 11 and cpf.isdigit()

    def validar_telefone(self, telefone):
        telefone = re.sub(r'\D', '', telefone)
        return len(telefone) >= 10 and telefone.isdigit()

    def limpar_campos(self):
        for key in self.entries:
            if key == "nascimento":
                self.entries[key].set_date(pd.Timestamp.today())
            else:
                self.entries[key].delete(0, tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = CadastroApp(root)
    root.mainloop()