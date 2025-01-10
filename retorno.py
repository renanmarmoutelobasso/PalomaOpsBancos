import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import datetime
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

class RetornoReaderApp:
    def __init__(self, master, initial_file=None):
        self.master = master
        self.master.title("Paloma Ops - Leitor de Retorno Bancário")
        self.center_window(1200, 700)
        self.master.iconbitmap("ico.ico")

        # Inicializar atributos
        self.data = {}
        self.config = self.load_config()
        self.tabelas = self.load_tables()

        # Configurar a interface
        self.create_buttons()
        self.create_header()
        self.create_table()

        # Se um arquivo inicial foi fornecido, carregá-lo automaticamente
        if initial_file:
            self.load_initial_file(initial_file)
    def center_window(self, width, height):
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.master.geometry(f"{width}x{height}+{x}+{y}")

    def load_config(self):
        try:
            with open("config.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo config.json não encontrado.")
            return {}

    def load_tables(self):
        try:
            with open("tabelas.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo tabelas.json não encontrado.")
            return {}

    def create_buttons(self):
        self.button_frame = tk.Frame(self.master)
        self.button_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5, anchor="nw")

        self.btn_open = tk.Button(self.button_frame, text="Abrir Arquivo", command=self.open_file)
        self.btn_open.pack(side=tk.LEFT, padx=5)

        self.btn_export_excel = tk.Button(self.button_frame, text="Exportar para Excel", command=self.export_to_excel)
        self.btn_export_excel.pack(side=tk.LEFT, padx=5)

        self.btn_export_pdf = tk.Button(self.button_frame, text="Exportar para PDF", command=self.export_to_pdf)
        self.btn_export_pdf.pack(side=tk.LEFT, padx=5)

    def create_header(self):
        self.header_frame = tk.Frame(self.master)
        self.header_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5, anchor="nw")

        self.lbl_numero_retorno = tk.Label(self.header_frame, text="Número do Retorno: ")
        self.lbl_numero_retorno.grid(row=0, column=0, padx=5, sticky="w")
        self.lbl_numero_retorno_val = tk.Label(self.header_frame, text="")
        self.lbl_numero_retorno_val.grid(row=0, column=1, padx=5, sticky="w")

        self.lbl_data_retorno = tk.Label(self.header_frame, text="Data do Retorno: ")
        self.lbl_data_retorno.grid(row=0, column=2, padx=5, sticky="w")
        self.lbl_data_retorno_val = tk.Label(self.header_frame, text="")
        self.lbl_data_retorno_val.grid(row=0, column=3, padx=5, sticky="w")

        self.lbl_banco = tk.Label(self.header_frame, text="Banco: ")
        self.lbl_banco.grid(row=0, column=4, padx=5, sticky="w")
        self.lbl_banco_val = tk.Label(self.header_frame, text="")
        self.lbl_banco_val.grid(row=0, column=5, padx=5, sticky="w")

        self.lbl_filial = tk.Label(self.header_frame, text="Filial: ")
        self.lbl_filial.grid(row=0, column=6, padx=5, sticky="w")
        self.lbl_filial_val = tk.Label(self.header_frame, text="")
        self.lbl_filial_val.grid(row=0, column=7, padx=5, sticky="w")

    def create_table(self):
        table_frame = tk.Frame(self.master)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        columns = [
            "Seu Número", "Vencimento", "Descrição da Ocorrência", "Descrição do Motivo",
            "Valor do Título", "Multa", "Juros", "Desconto", "Valor Pago", "Nosso Número",
            "Despesa Cobrança", "Data Ocorrência"
        ]

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings")

        col_widths = {
            "Seu Número": 80, "Vencimento": 80, "Descrição da Ocorrência": 150,
            "Descrição do Motivo": 150, "Valor do Título": 80, "Multa": 80, "Juros": 80,
            "Desconto": 80, "Valor Pago": 80, "Nosso Número": 100, "Despesa Cobrança": 100,
            "Data Ocorrência": 80
        }

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 80), anchor="center")

        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscroll=scroll_y.set, xscroll=scroll_x.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill="both", expand=True)

    def open_file(self):
        filepath = filedialog.askopenfilename(title="Selecione um Arquivo de Retorno", filetypes=[("Todos os Arquivos", "*.*")])
        if filepath:
            self.data = self.read_file(filepath)
            self.update_header()
            self.display_data()

    def read_file(self, filepath):
        with open(filepath, "r", encoding="latin-1") as file:
            lines = file.readlines()

        data = {
            "numero_retorno": lines[0][110:118].strip(),
            "data_retorno": self.format_date(lines[0][94:103].strip()),
            "banco": lines[0][82:90].strip(),
            "filial": self.get_filial_by_cnpj(lines[0][31:46].strip()),
            "detalhes": []
        }

        for line in lines:
            if line.startswith("1"):
                data["detalhes"].append(self.parse_detail(line))

        return data
        
    def load_initial_file(self, filepath):
        try:
            self.data = self.read_file(filepath)
            self.update_header()
            self.display_data()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo inicial: {e}")
                
    def open_file(self):
        filepath = filedialog.askopenfilename(title="Selecione um Arquivo de Retorno", filetypes=[("Todos os Arquivos", "*.*")])
        if filepath:
            self.data = self.read_file(filepath)
            self.update_header()
            self.display_data()                
                    
    def format_date(self, date_str):
        try:
            return datetime.datetime.strptime(date_str, "%Y/%m/%d").strftime("%d/%m/%Y")
        except ValueError:
            return date_str

    def get_filial_by_cnpj(self, cnpj):
        for filial, info in self.config.get("filiais", {}).items():
            if info.get("cnpj") == cnpj:
                return filial
        return "Desconhecida"

    def parse_detail(self, line):
        def safe_float(value):
            try:
                return float(value.strip()) / 1000
            except ValueError:
                return 0.0

        def safe_date(value):
            try:
                return datetime.datetime.strptime(value.strip(), "%d%m%y").strftime("%d/%m/%Y")
            except ValueError:
                return value.strip()

        descricao_ocorrencia = self.tabelas["ocorrencias"].get(line[108:110].strip(), "Desconhecida")
        motivo = line[318:320].strip()

        if descricao_ocorrencia == "Tarifa":
            descricao_motivo = self.tabelas["tarifas"].get(motivo, "Desconhecida")
        else:
            descricao_motivo = self.tabelas["motivos"].get(motivo, "Desconhecido")

        return {
            "Seu Número": line[116:126].strip(),
            "Vencimento": safe_date(line[146:152].strip()),
            "Descrição da Ocorrência": descricao_ocorrencia,
            "Descrição do Motivo": descricao_motivo,
            "Valor do Título": f"R$ {safe_float(line[152:165]):,.2f}",
            "Multa": f"R$ {safe_float(line[279:292]):,.2f}",
            "Juros": f"R$ {safe_float(line[266:279]):,.2f}",
            "Desconto": f"R$ {safe_float(line[241:254]):,.2f}",
            "Valor Pago": f"R$ {safe_float(line[253:266]):,.2f}",
            "Nosso Número": line[47:62].strip(),
            "Despesa Cobrança": f"R$ {safe_float(line[176:189]):,.2f}",
            "Data Ocorrência": safe_date(line[110:116].strip()),
        }

    def update_header(self):
        self.lbl_numero_retorno_val.config(text=self.data.get("numero_retorno", ""))
        self.lbl_data_retorno_val.config(text=self.data.get("data_retorno", ""))
        self.lbl_banco_val.config(text=self.data.get("banco", ""))
        self.lbl_filial_val.config(text=self.data.get("filial", ""))

    def display_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        detalhes_ordenados = sorted(
            self.data["detalhes"],
            key=lambda x: (x["Vencimento"], x["Seu Número"])
        )

        for detalhe in detalhes_ordenados:
            self.tree.insert("", "end", values=tuple(detalhe.values()))

    def export_to_excel(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                import pandas as pd
                pd.DataFrame(self.data["detalhes"]).to_excel(save_path, index=False)
                messagebox.showinfo("Exportação", "Arquivo exportado para Excel com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar para Excel: {e}")

    def export_to_pdf(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if save_path:
            try:
                from fpdf import FPDF

                pdf = FPDF()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                pdf.set_font("Arial", size=10)

                pdf.cell(200, 10, txt="Relatório de Retorno Bancário", ln=True, align="C")
                pdf.ln(10)

                for header in self.tree["columns"]:
                    pdf.cell(30, 10, txt=header, border=1)
                pdf.ln()

                for item in self.data["detalhes"]:
                    for value in item.values():
                        pdf.cell(30, 10, txt=str(value), border=1)
                    pdf.ln()

                pdf.output(save_path)
                messagebox.showinfo("Exportação", "Arquivo exportado para PDF com sucesso.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar para PDF: {e}")

if __name__ == "__main__":
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    root = tk.Tk()
    app = RetornoReaderApp(root, initial_file)
    root.mainloop()