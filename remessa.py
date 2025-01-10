import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import datetime
import os
from fpdf import FPDF
import json
import sys

if len(sys.argv) < 2:
    print("Erro: Caminho do arquivo não fornecido!")
    sys.exit(1)

caminho_arquivo = sys.argv[1]

try:
    with open(caminho_arquivo, "r") as arquivo:
        conteudo = arquivo.read()
        print(f"Conteúdo do arquivo:\n{conteudo}")
except Exception as e:
    print(f"Erro ao abrir o arquivo {caminho_arquivo}: {e}")


class RemessaReaderApp:
    def __init__(self, master, initial_file=None):
        self.master = master
        self.master.title("Paloma Ops - Leitor de Remessa Bancária")
        self.master.iconbitmap("ico.ico")

        # Centralizar a janela
        window_width, window_height = 1200, 700
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        self.master.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")

        # Inicialização de atributos
        self.filepath = None
        self.data = {"header": None, "detalhes": []}
        self.config = self.load_config()

        # Botões
        self.button_frame = tk.Frame(master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        self.btn_open = tk.Button(self.button_frame, text="Abrir Arquivo", command=self.open_file)
        self.btn_open.pack(side=tk.LEFT, padx=5)

        self.btn_export_excel = tk.Button(self.button_frame, text="Exportar para Excel", command=self.export_to_excel)
        self.btn_export_excel.pack(side=tk.LEFT, padx=5)

        self.btn_export_pdf = tk.Button(self.button_frame, text="Exportar para PDF", command=self.export_to_pdf)
        self.btn_export_pdf.pack(side=tk.LEFT, padx=5)

        self.btn_print = tk.Button(self.button_frame, text="Imprimir", command=self.print_to_printer)
        self.btn_print.pack(side=tk.LEFT, padx=5)

        # Exibição do cabeçalho do arquivo
        self.header_frame = tk.LabelFrame(master, text="Cabeçalho do Arquivo")
        self.header_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        self.lbl_remessa = tk.Label(self.header_frame, text="Número da Remessa: ")
        self.lbl_remessa.pack(side=tk.LEFT, padx=5)
        self.lbl_remessa_val = tk.Label(self.header_frame, text="")
        self.lbl_remessa_val.pack(side=tk.LEFT, padx=5)

        self.lbl_data_remessa = tk.Label(self.header_frame, text="Data da Remessa: ")
        self.lbl_data_remessa.pack(side=tk.LEFT, padx=5)
        self.lbl_data_remessa_val = tk.Label(self.header_frame, text="")
        self.lbl_data_remessa_val.pack(side=tk.LEFT, padx=5)

        self.lbl_banco = tk.Label(self.header_frame, text="Banco: ")
        self.lbl_banco.pack(side=tk.LEFT, padx=5)
        self.lbl_banco_val = tk.Label(self.header_frame, text="")
        self.lbl_banco_val.pack(side=tk.LEFT, padx=5)

        self.lbl_filial = tk.Label(self.header_frame, text="Filial: ")
        self.lbl_filial.pack(side=tk.LEFT, padx=5)
        self.lbl_filial_val = tk.Label(self.header_frame, text="")
        self.lbl_filial_val.pack(side=tk.LEFT, padx=5)

        # Exibição da tabela com barra de rolagem
        table_frame = tk.Frame(master)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = [
            "Instrução", "Emissão", "Nome", "Valor do Título", "Vencimento",
            "% Multa", "Juros (R$)", "Nosso Número", "Seu Número"
        ]

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
        for col in columns:
            if col in ["Emissão", "Vencimento", "Juros (R$)"]:
                self.tree.column(col, width=120, anchor="center")
            elif col == "% Multa":
                self.tree.column(col, width=80, anchor="center")
            elif col == "Nome":
                self.tree.column(col, width=250, anchor="w")  # Alinhar à esquerda
            else:
                self.tree.column(col, width=100, anchor="center")
            self.tree.heading(col, text=col)

        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scroll_y.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill="both", expand=True)

        # Se um arquivo inicial foi passado, carregá-lo automaticamente
        if initial_file:
            self.filepath = initial_file
            self.data = self.read_remessa(self.filepath)
            self.update_header()
            self.display_data()


    def load_config(self):
        try:
            with open("config.json", "r") as config_file:
                return json.load(config_file)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar config.json: {e}")
            return None

    def open_file(self):
        self.filepath = filedialog.askopenfilename()
        if self.filepath:
            self.data = self.read_remessa(self.filepath)
            self.update_header()
            self.display_data()

    def read_remessa(self, filepath):
        with open(filepath, "r", encoding="latin-1") as file:
            lines = file.readlines()

        registros = {"header": None, "detalhes": []}

        # Cabeçalho
        cnpj = lines[0][31:45].strip()
        registros["header"] = {
            "Número da Remessa": lines[0][110:117].strip(),
            "Data da Remessa": self.format_date(lines[0][94:102].strip()),
            "Banco": lines[0][79:87].strip(),
            "CNPJ": cnpj,
            "Filial": self.get_filial_by_cnpj(cnpj)
        }

        # Detalhes
        for line in lines:
            if line.startswith("1"):
                registros["detalhes"].append(self.parse_detalhe(line))

        return registros

    def parse_detalhe(self, line):
        def safe_float(value):
            try:
                return float(value.strip()) / 100 if value.strip().isdigit() else 0.0
            except ValueError:
                return 0.0

        instrucoes_legenda = {
            "10": "Cadastro de Títulos",
            "20": "Pedido de Baixa",
            "40": "Concessão de Abatimento",
        }

        instrucao = line[109:111].strip()
        emissao = line[150:156].strip()
        vencimento = line[120:126].strip()

        return {
            "Instrução": instrucoes_legenda.get(instrucao, "Desconhecida"),
            "Emissão": self.format_date(emissao),
            "Nome": line[234:274].strip(),
            "Valor do Título": f"R$ {safe_float(line[126:139]):,.2f}",
            "Vencimento": self.format_date(vencimento),
            "% Multa": f"{safe_float(line[93:96]):.2f}%",
            "Juros (R$)": f"R$ {safe_float(line[161:173]):,.2f}",
            "Nosso Número": line[48:57].strip(),
            "Seu Número": line[111:121].strip(),
        }

    def format_date(self, date_str):
        try:
            return datetime.datetime.strptime(date_str, "%d%m%y").strftime("%d/%m/%Y")
        except ValueError:
            return date_str

    def get_filial_by_cnpj(self, cnpj):
        for filial, info in self.config["filiais"].items():
            if info["cnpj"] == cnpj:
                return filial
        return "Desconhecida"

    def update_header(self):
        header = self.data["header"]
        self.lbl_remessa_val.config(text=header["Número da Remessa"])
        self.lbl_data_remessa_val.config(text=header["Data da Remessa"])
        self.lbl_banco_val.config(text=header["Banco"])
        self.lbl_filial_val.config(text=header["Filial"])

    def display_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Agrupando por instrução e ordenando
        detalhes_ordenados = sorted(
            self.data["detalhes"],
            key=lambda x: (x["Instrução"], x["Vencimento"], x["Nome"])
        )

        for detalhe in detalhes_ordenados:
            valores = [
                detalhe.get("Instrução", ""),
                detalhe.get("Emissão", ""),
                detalhe.get("Nome", ""),
                detalhe.get("Valor do Título", ""),
                detalhe.get("Vencimento", ""),
                detalhe.get("% Multa", ""),
                detalhe.get("Juros (R$)", ""),
                detalhe.get("Nosso Número", ""),
                detalhe.get("Seu Número", ""),
            ]
            self.tree.insert("", "end", values=valores)

    def export_to_pdf(self, filename=None):
        if not filename:
            filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not filename:
                return  # Se o usuário cancelar, não faz nada

        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        # Cabeçalho
        pdf.cell(275, 10, txt="Relatório de Remessa Bancária", ln=True, align="C")

        # Colunas
        columns = [
            "Instrução", "Emissão", "Nome", "Valor do Título", "Vencimento",
            "% Multa", "Juros (R$)", "Nosso Número", "Seu Número"
        ]
        col_widths = [30, 25, 60, 30, 30, 25, 30, 30, 30]

        for col in columns:
            pdf.cell(col_widths[columns.index(col)], 10, col, border=1, align="C")
        pdf.ln()

        # Linhas
        for detalhe in self.data["detalhes"]:
            for col in columns:
                value = str(detalhe[col])
                if col == "Nome" and len(value) > 25:
                    value = value[:25] + "..."  # Limitar o tamanho do nome
                pdf.cell(col_widths[columns.index(col)], 10, value, border=1)
            pdf.ln()

        pdf.output(filename)
        messagebox.showinfo("Exportação", "Exportado para PDF com sucesso!")

    def export_to_excel(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            pd.DataFrame(self.data["detalhes"]).to_excel(save_path, index=False)
            messagebox.showinfo("Exportação", "Exportado para Excel com sucesso!")

    def print_to_printer(self):
        temp_pdf_path = "temp_remessa.pdf"
        self.export_to_pdf(temp_pdf_path)
        if os.path.exists(temp_pdf_path):
            os.startfile(temp_pdf_path, "print")
        else:
            messagebox.showerror("Erro", "O arquivo temporário PDF não foi encontrado.")


if __name__ == "__main__":
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    root = tk.Tk()
    app = RemessaReaderApp(root, initial_file)
    root.mainloop()

