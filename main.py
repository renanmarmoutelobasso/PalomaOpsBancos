import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
from datetime import datetime
import subprocess
import random
from tkinter import Tk, PhotoImage

# Dados iniciais para simulação
CONFIG_FILE = "config.json"
STATUS_FILE = "status.json"


# Funções auxiliares
def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as file:
                return json.load(file)
        return {"selected_filial": None, "filiais": {}}
    except json.JSONDecodeError:
        return {"selected_filial": None, "filiais": {}}

def save_config(config):
    with open(CONFIG_FILE, "w") as file:
        json.dump(config, file, indent=4)

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

from datetime import datetime

def read_files_in_path(path, file_type, filial_cnpj):
    """
    Lê todos os arquivos na pasta especificada e filtra apenas arquivos de remessa ou retorno relacionados à filial selecionada.
    """
    print(f"Lendo arquivos do caminho: {path}")
    files = []

    for file_name in os.listdir(path):
        file_path = os.path.join(path, file_name)

        if not os.path.isfile(file_path):
            continue

        try:
            with open(file_path, "r", encoding="latin-1") as file:
                content = file.readlines()
                if not content:
                    continue

                first_line = content[0]

                if file_type == "remessa" and "REMESSA" in first_line[2:10]:
                    arquivo_cnpj = first_line[31:46].strip()
                    if arquivo_cnpj != filial_cnpj:
                        print(f"Arquivo {file_name} ignorado: CNPJ {arquivo_cnpj} não corresponde à filial selecionada.")
                        continue
                    raw_date = first_line[94:102].strip()
                    formatted_date = datetime.strptime(raw_date, "%Y%m%d").strftime("%d/%m/%Y")
                    files.append({
                        "Numero": first_line[110:118].strip(),
                        "Data": formatted_date,
                        "Banco": first_line[79:87].strip(),
                        "CNPJ": arquivo_cnpj,
                        "Status": "Pendente",
                        "DataAcao": "",
                        "Caminho": file_path,
                    })
                elif file_type == "retorno" and "RETORNO" in first_line[2:10]:
                    arquivo_cnpj = first_line[31:46].strip()
                    if arquivo_cnpj != filial_cnpj:
                        print(f"Arquivo {file_name} ignorado: CNPJ {arquivo_cnpj} não corresponde à filial selecionada.")
                        continue
                    raw_date = first_line[94:102].strip()
                    formatted_date = datetime.strptime(raw_date, "%Y%m%d").strftime("%d/%m/%Y")
                    files.append({
                        "Numero": first_line[110:118].strip(),
                        "Data": formatted_date,
                        "Banco": first_line[82:90].strip(),
                        "CNPJ": arquivo_cnpj,
                        "Status": "Pendente",
                        "DataAcao": "",
                        "Caminho": file_path,
                    })
                else:
                    print(f"Arquivo {file_name} ignorado: Não é {file_type.upper()}.")
        except Exception as e:
            print(f"Erro ao processar arquivo {file_name}: {e}")

    return files





# Classe principal

class MainApp:
    def __init__(self, root):
        self.root = root
        self.config = load_config()
        self.selected_filial = self.config.get("selected_filial")

        # Carregar a imagem do ícone
        icone = PhotoImage(file="logo.png")

        # Definir o ícone da janela
        self.root.iconphoto(False, icone)

        self.root.title("Paloma Ops - Gerenciador de Arquivos Bancários")
        self.root.geometry("1000x900")
        self.root.resizable(True, True)
        center_window(self.root)

        if not self.config["filiais"]:
            messagebox.showerror("Erro", "Nenhuma filial configurada. Configure pelo menu Configurações.")

        self.create_widgets()

        # Configurar ação ao fechar a janela
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Salva a filial selecionada e fecha o programa."""
        self.config["selected_filial"] = self.selected_filial
        save_config(self.config)
        self.root.destroy()



    def create_widgets(self):
        # Frame do cabeçalho
        header_frame = tk.Frame(self.root)
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)

        title_label = tk.Label(header_frame, text="Gerenciador de Arquivos Bancários", font=("Arial", 22))
        title_label.pack(side=tk.LEFT, padx=5)

        self.filial_combobox = ttk.Combobox(
            header_frame,
            values=list(self.config["filiais"].keys()),
            state="readonly",
            width=30
        )
        self.filial_combobox.pack(side=tk.LEFT, padx=10)
        self.filial_combobox.set(self.selected_filial if self.selected_filial else "Selecionar")
        self.filial_combobox.bind("<<ComboboxSelected>>", self.change_filial)

        config_button = tk.Button(header_frame, text="Configurações", command=self.open_config_window)
        config_button.pack(side=tk.RIGHT, padx=5)

        # Frame do conteúdo principal
        self.content_frame = tk.Frame(self.root)
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        # Configuração para expandir dinamicamente
        self.root.grid_rowconfigure(1, weight=1)  # Permitir expansão vertical
        self.root.grid_columnconfigure(0, weight=1)  # Permitir expansão horizontal

        # Criar visualizações dos arquivos
        self.create_file_views()

    def change_filial(self, event):
        """Atualiza a filial selecionada no ComboBox."""
        selected_filial = self.filial_combobox.get()
        if selected_filial in self.config["filiais"]:
            self.selected_filial = selected_filial
            print(f"Filial alterada para: {self.selected_filial}")
            self.reload_view()
        else:
            print("Filial selecionada não encontrada nos dados de configuração.")


    def open_with_program(self, tree, file_type, script_path):
        """Abre o programa relacionado (executável ou script)."""
        # Obtém o item selecionado no TreeView
        selected_item = tree.focus()
        if not selected_item:
            messagebox.showerror("Erro", "Nenhum item selecionado!")
            return

        # Obtém os valores da linha selecionada
        values = tree.item(selected_item, "values")
        if not values:
            messagebox.showerror("Erro", "Nenhuma informação válida selecionada!")
            return

        # Obtém o número do arquivo e busca o caminho completo
        numero_arquivo = values[0]
        filial_cnpj = self.config["filiais"].get(self.selected_filial, {}).get("cnpj", "").strip()
        path = self.config["filiais"].get(self.selected_filial, {}).get("path", "")

        arquivos = read_files_in_path(path, file_type, filial_cnpj)
        arquivo_selecionado = next((f for f in arquivos if f["Numero"] == numero_arquivo), None)

        if not arquivo_selecionado:
            messagebox.showerror("Erro", "Caminho do arquivo não encontrado!")
            return

        caminho_completo = arquivo_selecionado["Caminho"]

        # Executa o programa com o caminho completo
        try:
            print(f"Executando {script_path} com arquivo {caminho_completo}")
            subprocess.Popen([script_path, caminho_completo], shell=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir o programa: {e}")











    def create_file_views(self):
        # Adiciona os frames para Remessa e Retorno
        remessa_frame = tk.Frame(self.content_frame)
        remessa_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        retorno_frame = tk.Frame(self.content_frame)
        retorno_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        self.content_frame.grid_rowconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(1, weight=1)
        self.content_frame.grid_columnconfigure(0, weight=1)

        self.create_file_view(remessa_frame, "remessa")
        self.create_file_view(retorno_frame, "retorno")
       


    def create_file_view(self, parent, file_type):
        """
        Cria a visualização dos arquivos, incluindo botões de ações e caminho para retornos.
        """
        print(f"Criando visualização para: {file_type.upper()}")

        # Configurar layout responsivo
        parent.grid_rowconfigure(0, weight=0)  # Título (não expande)
        parent.grid_rowconfigure(1, weight=1)  # Treeview (expande)
        parent.grid_rowconfigure(2, weight=0)  # Frame de ações (não expande)
        parent.grid_columnconfigure(0, weight=1)  # Expande horizontalmente

        # Título
        title_label = tk.Label(parent, text=file_type.upper(), font=("Arial", 18))
        title_label.grid(row=0, column=0, pady=5, sticky="n")  # Alinhado ao topo

        # Treeview
        columns = ("Numero", "Data", "Banco", "Status", "DataAcao")
        tree = ttk.Treeview(parent, columns=columns, show="headings", style="Treeview")
        tree.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)  # Expande em todas as direções

        # Configuração das colunas
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self.sort_tree(tree, c, False))
            tree.column(col, width=120, anchor=tk.CENTER)

        # Adicionar cores às tags
        tree.tag_configure("red", foreground="red")
        tree.tag_configure("blue", foreground="blue")

        # Frame de ações
        action_frame = tk.Frame(parent)
        action_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=5)  # Alinhado horizontalmente

        # Botão para abrir arquivos de remessa ou retorno

        if file_type == "remessa":
            remessa_button = tk.Button(
                action_frame,
                text="Abrir Remessa",
                command=lambda: self.open_with_program(tree, file_type, "./remessa.exe"),  # Caminho do executável
                font=("Arial", 12)
            )
            remessa_button.pack(side=tk.RIGHT, padx=5)
        elif file_type == "retorno":
            retorno_button = tk.Button(
                action_frame,
                text="Abrir Retorno",
                command=lambda: self.open_with_program(tree, file_type, "./retorno.exe"),  # Caminho do executável
                font=("Arial", 12)
            )
            retorno_button.pack(side=tk.RIGHT, padx=5)



        # Botão Copiar
        copy_button = tk.Button(
            action_frame,
            text="Copiar",
            command=lambda: self.handle_copy(tree, file_type),
            font=("Arial", 12)
        )
        copy_button.pack(side=tk.RIGHT, padx=5)

        # Adicionando instruções e botão para arquivos de retorno
        if file_type == "retorno":
            path_frame = tk.Frame(action_frame)
            path_frame.pack(side=tk.LEFT, fill=tk.X, padx=5)

            lbl_instructions = tk.Label(
                path_frame,
                text="Clique em copiar para salvar o caminho do Retorno",
                font=("Arial", 12, "bold"),
                anchor="w"
            )
            lbl_instructions.pack(side=tk.LEFT, padx=5)

            btn_copy_footer = tk.Button(
                path_frame,
                text="Copiar Caminho do Retorno",
                font=("Arial", 12),
                command=self.copy_return_path
            )
            btn_copy_footer.pack(side=tk.LEFT, padx=5)

        # Carregar arquivos e atualizar a visualização do Treeview
        if not self.selected_filial or self.selected_filial not in self.config["filiais"]:
            print("Nenhuma filial selecionada ou inválida!")
            return

        path = self.config["filiais"].get(self.selected_filial, {}).get("path")
        filial_cnpj = self.config["filiais"].get(self.selected_filial, {}).get("cnpj", "").strip()

        if not path or not os.path.exists(path):
            print(f"Caminho inválido ou inexistente: {path}")
            return

        files = read_files_in_path(path, file_type, filial_cnpj)
        status_data = self.load_status().get(self.selected_filial, {}).get(file_type, {})

        for file in files:
            file_status = status_data.get(file["Numero"], {}).get("Status", "Pendente")
            file_data_acao = status_data.get(file["Numero"], {}).get("DataAcao", "")
            file["Status"] = file_status
            file["DataAcao"] = file_data_acao

            color = "red" if file_status == "Pendente" else "blue"
            tree.insert("", tk.END, values=(
                file["Numero"],
                file["Data"],
                file["Banco"],
                file_status,
                file_data_acao,
            ), tags=(color,))


            
    def sort_tree(self, tree, col, reverse):
        if col == "Numero":  # Certifique-se de que "Numero" é tratado como número
            items = [(int(tree.set(k, col)), k) for k in tree.get_children("") if tree.set(k, col).isdigit()]
        else:
            items = [(tree.set(k, col), k) for k in tree.get_children("")]
        items.sort(reverse=reverse)
        for index, (val, k) in enumerate(items):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_tree(tree, col, not reverse))

    
    def copy_path_with_filename(self):
        """Copia o caminho configurado com o próximo nome de arquivo para a área de transferência."""
        if not self.selected_filial or self.selected_filial not in self.config["filiais"]:
            messagebox.showerror("Erro", "Nenhuma filial selecionada ou caminho configurado.")
            return

        # Obter o caminho configurado
        base_path = self.config["filiais"][self.selected_filial].get("path")
        if not base_path or not os.path.exists(base_path):
            messagebox.showerror("Erro", "Caminho configurado é inválido ou inexistente.")
            return

        try:
            # Gerar um número sequencial baseado no timestamp
            from datetime import datetime
            proximo_numero = datetime.now().strftime("%Y%m%d%H%M%S")  # Sequência baseada no timestamp
            proximo_arquivo = f"{base_path}\\Ret{proximo_numero}"  # Nome sem extensão

            # Copiar para o clipboard
            self.root.clipboard_clear()
            self.root.clipboard_append(proximo_arquivo)
            self.root.update()
            messagebox.showinfo("Caminho Copiado", f"Caminho copiado:\n{proximo_arquivo}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar o próximo arquivo: {e}")


   
    def handle_copy(self, tree, file_type):
        """Copia o caminho do arquivo selecionado e altera o status."""
        selected_item = tree.focus()
        if not selected_item:
            messagebox.showerror("Erro", "Nenhum item selecionado para copiar!")
            return

        # Obtém os valores da linha selecionada
        values = tree.item(selected_item, "values")
        numero_arquivo = values[0]

        # Busca o caminho do arquivo correspondente
        path = self.config["filiais"].get(self.selected_filial, {}).get("path")
        filial_cnpj = self.config["filiais"].get(self.selected_filial, {}).get("cnpj", "").strip()
        arquivos = read_files_in_path(path, file_type, filial_cnpj)

        arquivo_selecionado = next((f for f in arquivos if f["Numero"] == numero_arquivo), None)

        if not arquivo_selecionado:
            messagebox.showerror("Erro", f"Arquivo correspondente não encontrado para o número {numero_arquivo}.")
            return

        caminho_corrigido = arquivo_selecionado["Caminho"].replace("/", "\\")
        self.root.clipboard_clear()
        self.root.clipboard_append(caminho_corrigido)
        self.root.update()

        # Define o novo status
        novo_status = "Enviado" if file_type == "remessa" else "Lido"

        # Atualiza o status no JSON e na interface
        self.update_status(numero_arquivo, novo_status, file_type)

        messagebox.showinfo("Sucesso", f"Caminho copiado:\n{caminho_corrigido}")


    def update_status(self, numero, new_status, file_type):
        """Atualiza o status do arquivo no JSON e no TreeView."""
        # Carrega o arquivo de status
        status_data = self.load_status()

        # Atualiza os dados no arquivo JSON
        if not status_data.get(self.selected_filial):
            status_data[self.selected_filial] = {}
        if not status_data[self.selected_filial].get(file_type):
            status_data[self.selected_filial][file_type] = {}

        status_data[self.selected_filial][file_type][numero] = {
            "Status": new_status,
            "DataAcao": datetime.now().strftime("%d/%m/%Y %H:%M:%S")  # Formato DD/MM/AAAA
        }

        self.save_status(status_data)

        # Recarregar a visualização da interface
        self.reload_view()




        
    def load_status(self):
        """Carrega o arquivo status.json."""
        try:
            with open(STATUS_FILE, "r") as file:
                return json.load(file)
        except FileNotFoundError:
            return {}
        except json.JSONDecodeError:
            return {}

    def save_status(self, status_data):
        """Salva os dados no arquivo status.json."""
        try:
            with open(STATUS_FILE, "w") as file:
                json.dump(status_data, file, indent=4)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo status.json: {e}")


            
    def copy_return_path(self):
        """Gera um nome único para o arquivo de retorno e copia o caminho para a área de transferência."""
        if not self.selected_filial:
            messagebox.showerror("Erro", "Nenhuma filial selecionada.")
            return

        # Obter o caminho configurado para a filial
        base_path = self.config["filiais"].get(self.selected_filial, {}).get("path")
        if not base_path or not os.path.exists(base_path):
            messagebox.showerror("Erro", "Caminho para arquivos de retorno não configurado ou inexistente.")
            return

        try:
            # Gera um número aleatório para o nome do arquivo
            while True:
                random_number = random.randint(10000, 99999)
                file_name = f"Ret{random_number}.txt"
                file_path = os.path.join(base_path, file_name)
                if not os.path.exists(file_path):  # Verifica se o nome já existe
                    break

            # Copia o caminho gerado para a área de transferência
            self.root.clipboard_clear()
            self.root.clipboard_append(file_path)
            self.root.update()
            messagebox.showinfo("Caminho Copiado", f"Caminho copiado:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o caminho de retorno: {e}")


    def copy_file_path(self, file_type):
        selected_item = self.tree.focus()
        if not selected_item:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
            return

        values = self.tree.item(selected_item, "values")
        file_name, file_path, status = values

        # Atualiza status e cor conforme o tipo
        new_status = "Enviado" if file_type == "remessa" else "Lido"
        self.tree.item(selected_item, values=(file_name, file_path, new_status))
        
        # Atualiza os dados do status.json
        if self.selected_filial not in self.status_data:
            self.status_data[self.selected_filial] = {}
        if file_type not in self.status_data[self.selected_filial]:
            self.status_data[self.selected_filial][file_type] = {}

        self.status_data[self.selected_filial][file_type][file_name] = {
            "Status": new_status,
            "DataAcao": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }

        # Salva o status no arquivo status.json
        try:
            with open(STATUS_FILE, "w") as file:
                json.dump(self.status_data, file, indent=4)
            print(f"Status salvo com sucesso no status.json para o arquivo: {file_name}")
        except Exception as e:
            print(f"Erro ao salvar status.json: {e}")
            messagebox.showerror("Erro", "Erro ao salvar o status no arquivo status.json!")
            return

        # Copia para a área de transferência
        self.root.clipboard_clear()
        self.root.clipboard_append(file_path)
        self.root.update()

        # Mostra mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Caminho copiado: {file_path}")



    def reload_view(self):
        """Recarrega a visualização principal."""
        print(f"Recarregando visualização para filial: {self.selected_filial}")
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        self.create_file_views()


    def handle_read(self, tree, file_type):
        selected_item = tree.focus()
        if not selected_item:
            messagebox.showerror("Erro", "Nenhum item selecionado para ler!")
            return

        values = tree.item(selected_item, "values")
        caminho = next((f["Caminho"] for f in read_files_in_path(self.config["filiais"].get(self.selected_filial, {}).get("path"), file_type) if f["Numero"] == values[0]), None)

        if not caminho:
            messagebox.showerror("Erro", "Caminho não encontrado para o arquivo selecionado.")
            return

        corrected_path = caminho.replace("/", "\\")  # Corrige as barras para o formato Windows
        print(f"Lendo arquivo: {corrected_path}")
        try:
            with open(corrected_path, "r") as file:
                content = file.read()
                print(f"Conteúdo do arquivo: {content[:100]}...")  # Exibe os 100 primeiros caracteres
        except Exception as e:
            print(f"Erro ao ler arquivo: {e}")
            
    def reload_view(self):
        """Recarrega a visualização principal."""
        print(f"Recarregando visualização para filial: {self.selected_filial}")
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        self.create_file_views()



    def open_config_window(self):
        ConfigWindow(self)

    def update_selected_filial(self, filial):
        self.selected_filial = filial
        self.config["selected_filial"] = filial
        save_config(self.config)
        self.reload_view()

    def on_tree_action(self, tree, file_type):
        selected_item = tree.focus()
        if not selected_item:
            return

        values = tree.item(selected_item, "values")
        action = values[5]
        caminho = next((f["Caminho"] for f in read_files_in_path(self.config["filiais"].get(self.selected_filial, {}).get("path"), file_type) if f["Numero"] == values[0]), None)

        if not caminho:
            print("Caminho não encontrado para o arquivo selecionado.")
            return

        if "Copiar" in action:
            self.copy_to_clipboard(caminho)
        elif "Ler" in action:
            self.read_file(caminho)

    def copy_to_clipboard(self, text):
        corrected_text = text.replace("/", "\\")  # Corrige as barras para o formato Windows
        self.root.clipboard_clear()
        self.root.clipboard_append(corrected_text)
        self.root.update()  # Mantém o texto no clipboard
        print(f"Texto copiado para o clipboard: {corrected_text}")

    def read_file(self, file_path):
        corrected_path = file_path.replace("/", "\\")  # Corrige as barras para o formato Windows
        print(f"Lendo arquivo: {corrected_path}")
        try:
            with open(corrected_path, "r") as file:
                content = file.read()
                print(f"Conteúdo do arquivo: {content[:100]}...")  # Exibe os 100 primeiros caracteres
        except Exception as e:
            print(f"Erro ao ler arquivo: {e}")
    
    def copy_return_path(self):
        """Gera um nome único para o arquivo de retorno e copia o caminho para a área de transferência."""
        if not self.selected_filial:
            messagebox.showerror("Erro", "Nenhuma filial selecionada.")
            return

        # Obter o caminho configurado para a filial
        base_path = self.config["filiais"].get(self.selected_filial, {}).get("path")
        if not base_path or not os.path.exists(base_path):
            messagebox.showerror("Erro", "Caminho para arquivos de retorno não configurado ou inexistente.")
            return

        try:
            # Gera um número aleatório para o nome do arquivo
            while True:
                random_number = random.randint(10000, 99999)
                file_name = f"Ret{random_number}"  # Sem extensão
                file_path = os.path.join(base_path, file_name)
                if not os.path.exists(file_path):  # Verifica se o nome já existe
                    break

            # Copia o caminho gerado para a área de transferência
            self.root.clipboard_clear()
            self.root.clipboard_append(file_path)
            self.root.update()

            # Mensagem de sucesso
            messagebox.showinfo("Caminho Copiado", f"Caminho copiado para a área de transferência:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar o caminho de retorno: {e}")

    def load_initial_file(self, filepath):
        try:
            self.filepath = filepath
            filial_name = self.data.get("filial", "Desconhecida")
            filial_config = self.config.get("filiais", {}).get(filial_name, {})
            base_path = filial_config.get("path", "")

            self.lbl_file_path_val_footer.config(text=base_path or "Caminho não configurado")
            self.data = self.read_file(filepath)
            self.update_header()
            self.display_data()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo inicial: {e}")

    def create_footer(self):
        # Frame para o rodapé (caminho do arquivo e botão Copiar)
        footer_frame = tk.Frame(self.root)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        # Label para o caminho configurado
        lbl_file_path = tk.Label(
            footer_frame,
            text="Caminho configurado:",
            font=("Arial", 12, "bold")
        )
        lbl_file_path.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Label para exibir o caminho
        self.lbl_file_path_val = tk.Label(
            footer_frame,
            text=self.get_filial_path(),
            font=("Arial", 12),
            relief=tk.SUNKEN,
            anchor="w",
            width=50  # Ajustar largura
        )
        self.lbl_file_path_val.grid(row=0, column=1, padx=5, pady=5)

        # Botão Copiar
        btn_copy_footer = tk.Button(
            self.footer_frame,
            text="Copiar",
            font=("Arial", 12),
            command=self.copy_path_with_filename
        )
        btn_copy_footer.grid(row=0, column=2, padx=5, pady=5)



        # Ajustar o layout
        footer_frame.grid_columnconfigure(1, weight=1)

    def get_filial_path(self):
        """Retorna o caminho configurado para a filial selecionada."""
        if not self.selected_filial or self.selected_filial not in self.config["filiais"]:
            return "Nenhuma filial selecionada ou caminho não configurado."
        return self.config["filiais"][self.selected_filial].get("path", "Caminho não encontrado.")


    def open_config_window(self):
        ConfigWindow(self)


# Janela de Configuração
class ConfigWindow:
    def __init__(self, main_app):
        self.main_app = main_app
        self.config = main_app.config
        self.window = tk.Toplevel()
        self.window.title("Paloma Ops - Selecionar Filial")
        self.window.geometry("400x350")
        self.window.resizable(False, False)
        self.window.iconbitmap("ico.ico")
        center_window(self.window)

        # Torna a janela modal
        self.window.grab_set()  # Impede que outras janelas sejam interativas enquanto essa está aberta

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self.window, text="Selecionar Filial", font=("Arial", 14)).pack(pady=10)

        tk.Label(self.window, text="Filial Selecionada:", font=("Arial", 10)).pack()
        self.filial_combobox = ttk.Combobox(self.window, values=list(self.config["filiais"].keys()))
        self.filial_combobox.pack(pady=5)
        self.filial_combobox.set(self.main_app.selected_filial if self.main_app.selected_filial else "Selecionar")
        self.filial_combobox.bind("<<ComboboxSelected>>", self.load_filial_details)

        tk.Label(self.window, text="Nome da Filial:", font=("Arial", 10)).pack()
        self.filial_name_entry = tk.Entry(self.window, width=50, state="disabled")
        self.filial_name_entry.pack(pady=5)

        tk.Label(self.window, text="CNPJ:", font=("Arial", 10)).pack()
        self.cnpj_entry = tk.Entry(self.window, width=50, state="disabled")
        self.cnpj_entry.pack(pady=5)

        tk.Label(self.window, text="Caminho dos Arquivos:", font=("Arial", 10)).pack()

        path_frame = tk.Frame(self.window)
        path_frame.pack(pady=10)
        self.path_entry = tk.Entry(path_frame, width=40, state="disabled")
        self.path_entry.grid(row=0, column=0, padx=5)

        self.procurar_button = tk.Button(path_frame, text="Procurar", command=self.select_path, state="disabled")
        self.procurar_button.grid(row=0, column=1, padx=5)

        button_frame = tk.Frame(self.window)
        button_frame.pack(pady=20)

        self.new_button = tk.Button(button_frame, text="Novo", command=self.create_new_filial)
        self.new_button.pack(side=tk.LEFT, padx=10)

        self.edit_button = tk.Button(button_frame, text="Editar", command=self.enable_edit)
        self.edit_button.pack(side=tk.LEFT, padx=10)

        self.save_button = tk.Button(button_frame, text="Salvar", command=self.save_filial, state="disabled")
        self.save_button.pack(side=tk.LEFT, padx=10)

        self.cancel_button = tk.Button(button_frame, text="Cancelar", command=self.cancel_action, state="disabled")
        self.cancel_button.pack(side=tk.LEFT, padx=10)

        self.delete_button = tk.Button(button_frame, text="Excluir", command=self.delete_filial)
        self.delete_button.pack(side=tk.LEFT, padx=10)

        tk.Button(self.window, text="Fechar", command=self.window.destroy).pack(pady=10)

        # Configura validação e máscara para o campo CNPJ
        self.setup_cnpj_field()

        # Carrega automaticamente os dados da filial selecionada
        if self.filial_combobox.get() != "Selecionar":
            self.load_filial_details(None)




    def create_new_filial(self):
        """Prepara o formulário para criar uma nova filial."""
        self.filial_combobox.set("")  # Limpa o combobox
        self.filial_name_entry.config(state="normal")
        self.cnpj_entry.config(state="normal")
        self.path_entry.config(state="normal")
        self.filial_name_entry.delete(0, tk.END)
        self.cnpj_entry.delete(0, tk.END)
        self.path_entry.delete(0, tk.END)
        self.procurar_button.config(state="normal")
        self.save_button.config(state="normal")
        self.cancel_button.config(state="normal")
        self.edit_button.config(state="disabled")
        self.new_button.config(state="disabled")
        self.delete_button.config(state="disabled")

        # Foca no campo de Nome da Filial
        self.filial_name_entry.focus_set()



    def load_filial_details(self, event):
        """Carrega os detalhes da filial selecionada no combobox."""
        selected_filial = self.filial_combobox.get()
        if selected_filial in self.config["filiais"]:
            filial_data = self.config["filiais"][selected_filial]
            self.filial_name_entry.config(state="normal")
            self.cnpj_entry.config(state="normal")
            self.path_entry.config(state="normal")
            self.filial_name_entry.delete(0, tk.END)
            self.filial_name_entry.insert(0, selected_filial)
            self.cnpj_entry.delete(0, tk.END)
            self.cnpj_entry.insert(0, filial_data.get("cnpj", ""))
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, filial_data.get("path", ""))
            self.filial_name_entry.config(state="disabled")
            self.cnpj_entry.config(state="disabled")
            self.path_entry.config(state="disabled")
            self.edit_button.config(state="normal")
            self.delete_button.config(state="normal")
        else:
            self.clear_fields()


    def save_filial(self):
        """Salva as alterações ou criação de nova filial."""
        name = self.filial_name_entry.get().strip()
        cnpj = self.cnpj_entry.get().strip()
        path = self.path_entry.get().strip()

        # Validação do nome da filial
        if not name:
            messagebox.showerror("Erro", "O campo 'Nome da Filial' não pode estar vazio.")
            self.filial_name_entry.focus()
            return

        # Validação do CNPJ usando self.is_valid_cnpj
        if not self.is_valid_cnpj(cnpj):
            messagebox.showerror("Erro", "CNPJ inválido. Verifique e tente novamente.")
            self.cnpj_entry.focus()
            return

        # Validação do caminho
        if not path:
            messagebox.showerror("Erro", "O campo 'Caminho dos Arquivos' não pode estar vazio.")
            self.path_entry.focus()
            return

        # Salvando a filial
        self.config["filiais"][name] = {"cnpj": cnpj, "path": path}
        save_config(self.config)
        messagebox.showinfo("Sucesso", f"Filial '{name}' salva com sucesso!")

        # Atualiza o ComboBox da janela de configurações
        self.filial_combobox['values'] = list(self.config["filiais"].keys())
        self.filial_combobox.set(name)  # Seleciona automaticamente a filial recém-salva

        # Atualiza a aplicação principal
        self.main_app.update_selected_filial(name)
        self.main_app.reload_view()

        # Fecha a janela de configurações
        self.window.destroy()

    def setup_cnpj_field(self):


        def limit_cnpj_length(event):
            # Verifica o texto atual no campo
            current_value = self.cnpj_entry.get()

            # Permite backspace, delete, setas e entrada de texto enquanto limita a 14 caracteres
            if len(current_value) >= 14 and event.keysym not in ("BackSpace", "Delete", "Left", "Right"):
                return "break"  # Impede a entrada de caracteres adicionais

        # Vincula o evento ao pressionar teclas no campo
        self.cnpj_entry.bind("<KeyPress>", limit_cnpj_length)

    
    def is_valid_cnpj(self, cnpj):
        """Valida o CNPJ com base nos dígitos verificadores."""
        cnpj = ''.join(filter(str.isdigit, cnpj))  # Remove caracteres não numéricos

        # Verifica se o CNPJ tem exatamente 14 dígitos
        if len(cnpj) != 14:
            return False

        # Lista de CNPJs inválidos conhecidos
        invalid_cnpjs = [
            "00000000000000", "11111111111111", "22222222222222", "33333333333333",
            "44444444444444", "55555555555555", "66666666666666", "77777777777777",
            "88888888888888", "99999999999999"
        ]
        if cnpj in invalid_cnpjs:
            return False

        def calculate_digit(cnpj, weights):
            """Calcula um dígito verificador do CNPJ."""
            soma = sum(int(cnpj[i]) * weights[i] for i in range(len(weights)))
            resto = soma % 11
            return "0" if resto < 2 else str(11 - resto)

        weights1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        weights2 = [6] + weights1

        digit1 = calculate_digit(cnpj[:12], weights1)
        digit2 = calculate_digit(cnpj[:12] + digit1, weights2)

        return cnpj[-2:] == digit1 + digit2







    def delete_filial(self):
        """Remove a filial selecionada."""
        selected_filial = self.filial_combobox.get()
        if selected_filial in self.config["filiais"]:
            confirm = messagebox.askyesno(
                "Confirmar Exclusão",
                f"Tem certeza de que deseja excluir a filial '{selected_filial}'?"
            )
            if confirm:
                del self.config["filiais"][selected_filial]
                save_config(self.config)
                messagebox.showinfo("Sucesso", f"Filial '{selected_filial}' excluída com sucesso.")
                self.filial_combobox['values'] = list(self.config["filiais"].keys())
                self.filial_combobox.set("Selecionar")
                self.clear_fields()
                self.main_app.update_selected_filial(None)
        else:
            messagebox.showerror("Erro", "Nenhuma filial válida selecionada.")


    def clear_fields(self):
        """Limpa os campos do formulário."""
        self.filial_name_entry.config(state="normal")
        self.cnpj_entry.config(state="normal")
        self.path_entry.config(state="normal")
        self.filial_name_entry.delete(0, tk.END)
        self.cnpj_entry.delete(0, tk.END)
        self.path_entry.delete(0, tk.END)
        self.filial_name_entry.config(state="disabled")
        self.cnpj_entry.config(state="disabled")
        self.path_entry.config(state="disabled")
        self.procurar_button.config(state="disabled")
        self.save_button.config(state="disabled")
        self.cancel_button.config(state="disabled")
        self.edit_button.config(state="disabled")
        self.delete_button.config(state="disabled")


    def enable_edit(self):
        """Habilita os campos para edição."""
        self.filial_name_entry.config(state="normal")
        self.cnpj_entry.config(state="normal")
        self.path_entry.config(state="normal")
        self.procurar_button.config(state="normal")
        self.save_button.config(state="normal")
        self.cancel_button.config(state="normal")
        self.edit_button.config(state="disabled")
        self.new_button.config(state="disabled")

    def cancel_action(self):
        """Cancela a ação atual e retorna ao estado inicial."""
        self.filial_name_entry.config(state="disabled")
        self.cnpj_entry.config(state="disabled")
        self.path_entry.config(state="disabled")
        self.procurar_button.config(state="disabled")
        self.save_button.config(state="disabled")
        self.cancel_button.config(state="disabled")
        self.edit_button.config(state="normal")
        self.new_button.config(state="normal")
        self.load_filial_details(None)



    def select_path(self):
        """Abre uma janela para selecionar a pasta e atualiza o campo de caminho."""
        path = filedialog.askdirectory(parent=self.window)  # Define a janela pai para evitar que a janela vá para trás
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path)


# Inicialização do aplicativo
def main():
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()



if __name__ == "__main__":
    main()