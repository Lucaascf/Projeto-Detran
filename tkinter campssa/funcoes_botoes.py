import logging
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
import tkinter as tk
from tkinter import *
from tkinter import messagebox, filedialog, Frame, Label, Entry, Button, simpledialog, ttk
from planilhas import Planilhas
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl
from openpyxl.styles import Font
import subprocess
from datetime import datetime
from tkcalendar import DateEntry


# Configurando logs
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def role_e_click(driver, xpath):
    """Rola a página até um elemento e clica nele.

    Args:
        driver: Instância do WebDriver do Selenium.
        xpath: XPath do elemento a ser clicado.
    """
    element = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, xpath))
    )
    # Rola para o elemento
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    element.click()


class FuncoesBotoes:
    """Classe que encapsula as funções relacionadas aos botões da interface."""

    def __init__(
        self, master: tk, planilhas: Planilhas, file_path: str, app, current_user=None
    ):
        """Inicializa a classe FuncoesBotoes.

        Args:
            master: Janela pai do Tkinter.
            planilhas: Instância da classe Planilhas.
            file_path: Caminho do arquivo de planilha.
            app: Instância da aplicação principal.
        """
        self.master = master
        self.planilhas = planilhas
        self.wb = self.planilhas.wb if self.planilhas else None
        self.file_path = file_path
        self.app = app
        self.current_user = current_user
        self.login_frame = None
        self.criar_conta_frame = None
        self.login_frame = None

        # Variáveis para opções de pagamento
        self.forma_pagamento_var = tk.StringVar(value="")
        self.radio_var = tk.StringVar(value="")

        # Variáveis de controle para checkbuttons
        self.d_var = tk.IntVar()
        self.c_var = tk.IntVar()
        self.e_var = tk.IntVar()
        self.p_var = tk.IntVar()

        # Entradas para os campos de pagamento
        self.entry_d = tk.Entry(master)
        self.entry_c = tk.Entry(master)
        self.entry_e = tk.Entry(master)
        self.entry_p = tk.Entry(master)

        # Entradas para os valores associados
        self.entry_valor_d = tk.Entry(master)  # Entrada para valor de D
        self.entry_valor_c = tk.Entry(master)  # Entrada para valor de C
        self.entry_valor_e = tk.Entry(master)  # Entrada para valor de E
        self.entry_valor_p = tk.Entry(master)  # Entrada para valor de P

    def center(self, window):
        """Centraliza a janela na tela.

        Args:
            window: A instância da janela Tkinter que deve ser centralizada.
        """
        # Atualiza o tamanho solicitado pela janela
        window.update_idletasks()

        # Obtém as dimensões atuais da janela
        width = window.winfo_width()
        height = window.winfo_height()

        # Obtém as dimensões da tela
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        # Calcula as coordenadas x e y para centralizar a janela
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)

        # Aplica a nova geometria à janela
        window.geometry(f"{width}x{height}+{x}+{y}")

        # Mostra a janela (se estiver oculta)
        window.deiconify()

    def adicionar_informacao(self):
        """Cria uma nova janela para adicionar informações de pacientes."""
        # Criação da nova janela
        self.adicionar_window = tk.Toplevel(self.master)
        self.adicionar_window.title("Adicionar Paciente")
        self.adicionar_window.geometry("400x380")
        self.adicionar_window.minsize(width=400, height=380)
        self.adicionar_window.maxsize(width=400, height=380)

        # Configuração das cores
        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"
        cor_selecionado = "#2C3E50"

        self.adicionar_window.configure(bg=cor_fundo)

        # Centraliza a nova janela
        self.center(self.adicionar_window)

        # Adicionando os componentes da interface
        tk.Label(
            self.adicionar_window,
            text="Preencha as informações:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 16, "bold"),
        ).pack(pady=(15, 5))

        # Frame para os RadioButtons
        frame_radios = tk.Frame(self.adicionar_window, bg=cor_fundo)
        frame_radios.pack(pady=5)

        # RadioButtons para seleção de tipo
        tipos = [("Médico", "medico"), ("Psicólogo", "psicologo"), ("Ambos", "ambos")]
        for tipo, valor in tipos:
            tk.Radiobutton(
                frame_radios,
                text=tipo,
                variable=self.radio_var,
                value=valor,
                bg=cor_fundo,
                fg=cor_texto,
                selectcolor=cor_selecionado,
                activebackground=cor_fundo,
                activeforeground=cor_texto,
                highlightthickness=0,
                font=("Arial", 12),
            ).pack(side=tk.LEFT, padx=2)

        # Frame para entrada de nome
        self.criar_entry(
            frame_nome="Nome:", var_name="nome_entry", parent=self.adicionar_window
        )

        # Frame para entrada de Renach
        self.criar_entry(
            frame_nome="Renach:", var_name="renach_entry", parent=self.adicionar_window
        )

        tk.Label(
            self.adicionar_window,
            text="Forma de Pagamento:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 16, "bold"),
        ).pack(pady=(15, 5))

        # Frame para opções de pagamento
        frame_pagamento = tk.Frame(self.adicionar_window, bg=cor_fundo)
        frame_pagamento.pack(pady=2)

        # Variáveis para checkbuttons
        self.d_var = tk.IntVar()
        self.c_var = tk.IntVar()
        self.e_var = tk.IntVar()
        self.p_var = tk.IntVar()

        # Lista de checkbuttons
        checkbuttons = [
            ("D", self.d_var),
            ("C", self.c_var),
            ("E", self.e_var),
            ("P", self.p_var),
        ]

        # Campos de entrada ao lado de cada checkbutton
        entry_widgets = [tk.Entry(frame_pagamento, width=10) for _ in range(4)]

        for i, (text, var) in enumerate(checkbuttons):
            tk.Checkbutton(
                frame_pagamento,
                text=text,
                variable=var,
                bg=cor_fundo,
                fg=cor_texto,
                selectcolor=cor_selecionado,
                activebackground=cor_fundo,
                activeforeground=cor_texto,
                highlightthickness=0,
            ).grid(row=i, column=0, padx=2, pady=2, sticky="w")
            entry_widgets[i].grid(row=i, column=1, padx=2, pady=2)

        # Frame para botões
        frame_botoes = tk.Frame(self.adicionar_window, bg=cor_fundo)
        frame_botoes.pack(pady=10)

        tk.Button(
            frame_botoes,
            text="Adicionar",
            command=self.salvar_informacao,
            highlightthickness=0,
            activebackground="#2C3E50",
            activeforeground="#ECF0F1",
        ).pack(side=tk.LEFT, padx=10, pady=10)

        tk.Button(
            frame_botoes,
            text="Voltar",
            command=self.adicionar_window.destroy,
            activebackground="#2C3E50",
            activeforeground="#ECF0F1",
        ).pack(side=tk.LEFT, padx=10, pady=10)

    def criar_entry(self, frame_nome, var_name, parent):
        """Cria um frame com label e entry para entradas de texto.

        Args:
            frame_nome: O texto do label.
            var_name: O nome da variável de entrada a ser criada.
            parent: O widget pai onde o frame será adicionado.
        """
        frame = tk.Frame(parent, bg=parent.cget("bg"))
        frame.pack(pady=2)

        tk.Label(
            frame,
            text=frame_nome,
            bg=parent.cget("bg"),
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(side=tk.LEFT, anchor="w", padx=5)

        entry = tk.Entry(frame)
        entry.pack(side=tk.LEFT, padx=2)

        # Armazena a entrada na instância da classe
        setattr(self, var_name, entry)

    def salvar_informacao(self):
        # Obter dados dos campos de entrada
        nome = self.nome_entry.get().strip().upper()
        renach = self.renach_entry.get().strip()

        # Validar preenchimento do nome e RENACH
        if not nome or not renach:
            messagebox.showerror(
                "Erro", "Por favor, preencha os campos de nome e RENACH."
            )
            return

        # Validar se RENACH é um número inteiro
        if not renach.isdigit():
            messagebox.showerror("Erro", "O RENACH deve ser um número inteiro.")
            return

        # Verificar se o RENACH já existe na planilha
        ws = self.wb.active
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
            if str(row[2].value) == renach or str(row[8].value) == renach:
                messagebox.showerror("Erro", "Este RENACH já está registrado.")
                return

        # Obter dados dos checkbuttons de pagamento
        pagamentos_selecionados = [
            (
                "D",
                self.d_var.get(),
                self.entry_d.get().strip(),
                self.entry_valor_d.get().strip(),
            ),
            (
                "C",
                self.c_var.get(),
                self.entry_c.get().strip(),
                self.entry_valor_c.get().strip(),
            ),
            (
                "E",
                self.e_var.get(),
                self.entry_e.get().strip(),
                self.entry_valor_e.get().strip(),
            ),
            (
                "P",
                self.p_var.get(),
                self.entry_p.get().strip(),
                self.entry_valor_p.get().strip(),
            ),
        ]

        # Filtrar formas de pagamento selecionadas
        selecionados = [p for p in pagamentos_selecionados if p[1] == 1]

        # Verificar se ao menos uma forma de pagamento foi selecionada
        if not selecionados:
            messagebox.showerror("Erro", "Selecione pelo menos uma forma de pagamento.")
            return

        # Verificar escolha entre médico, psicólogo ou ambos
        escolha = self.radio_var.get()
        if escolha not in ["medico", "psicologo", "ambos"]:
            messagebox.showerror("Erro", "Selecione Médico, Psicólogo ou Ambos.")
            return

        # Encontrar a próxima linha vazia para médicos e psicólogos
        nova_linha_medico = next(
            (row for row in range(3, ws.max_row + 2) if not ws[f"B{row}"].value), None
        )
        nova_linha_psicologo = next(
            (row for row in range(3, ws.max_row + 2) if not ws[f"H{row}"].value), None
        )

        # Inserir informações na planilha com base na escolha
        if escolha == "medico":
            ws[f"B{nova_linha_medico}"] = nome
            ws[f"C{nova_linha_medico}"] = renach
            ws[f"F{nova_linha_medico}"] = ", ".join(
                [f"{p[0]}: {p[2]} - {p[3]}" for p in selecionados]
            )
            messagebox.showinfo("Paciente adicionado para se consultar com médico!")

        elif escolha == "psicologo":
            ws[f"H{nova_linha_psicologo}"] = nome
            ws[f"I{nova_linha_psicologo}"] = renach
            ws[f"L{nova_linha_psicologo}"] = ", ".join(
                [f"{p[0]}: {p[2]} - {p[3]}" for p in selecionados]
            )
            messagebox.showinfo("Paciente adicionado para se consultar com psicólogo!")

        elif escolha == "ambos":
            ws[f"B{nova_linha_medico}"] = nome
            ws[f"C{nova_linha_medico}"] = renach
            ws[f"F{nova_linha_medico}"] = ", ".join(
                [f"{p[0]}: {p[2]} - {p[3]}" for p in selecionados]
            )
            ws[f"H{nova_linha_psicologo}"] = nome
            ws[f"I{nova_linha_psicologo}"] = renach
            ws[f"L{nova_linha_psicologo}"] = ", ".join(
                [f"{p[0]}: {p[2]} - {p[3]}" for p in selecionados]
            )
            messagebox.showinfo(
                "Paciente adicionado para se consultar com médico e psicólogo!"
            )

        # Salvar na planilha
        self.wb.save(self.planilhas.file_path)

        # Limpar os campos de entrada
        self.nome_entry.delete(0, tk.END)
        self.renach_entry.delete(0, tk.END)
        self.radio_var.set("")
        self.d_var.set(0)
        self.c_var.set(0)
        self.e_var.set(0)
        self.p_var.set(0)
        self.entry_d.delete(0, tk.END)
        self.entry_c.delete(0, tk.END)
        self.entry_e.delete(0, tk.END)
        self.entry_p.delete(0, tk.END)
        self.entry_valor_d.delete(0, tk.END)
        self.entry_valor_c.delete(0, tk.END)
        self.entry_valor_e.delete(0, tk.END)
        self.entry_valor_p.delete(0, tk.END)

    def excluir(self):
        """Remove informações de pacientes da planilha com base no RENACH fornecido pelo usuário e reorganiza as linhas."""
        ws = self.wb.active
        pacientes_medicos = {}
        pacientes_psicologos = {}

        # Armazenar pacientes de médicos
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
            if row[1].value and row[2].value:
                try:
                    renach_medico = int(row[2].value)
                    pacientes_medicos.setdefault(renach_medico, []).append(row[0].row)
                except ValueError:
                    print(f"RENACH inválido na linha {row[0].row}: {row[2].value}")

        # Armazenar pacientes psicólogos
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
            if row[7].value and row[8].value:
                try:
                    renach_psicologo = int(row[8].value)
                    pacientes_psicologos.setdefault(renach_psicologo, []).append(
                        row[0].row
                    )
                except ValueError:
                    print(f"RENACH inválido na linha {row[0].row}: {row[8].value}")

        # Janela de exclusão
        excluir_window = tk.Toplevel(self.master)
        excluir_window.title("Excluir Paciente")
        excluir_window.geometry("400x150")
        excluir_window.configure(bg=self.master.cget("bg"))

        tk.Label(
            excluir_window,
            text="Informe o RENACH:",
            bg=self.master.cget("bg"),
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=10)
        renach_entry = tk.Entry(excluir_window)
        renach_entry.pack(pady=5)

        def excluir_paciente():
            """Função para excluir o paciente com o RENACH fornecido"""
            try:
                renach = int(renach_entry.get())

                def reorganizar_linhas(linha_excluida):
                    """Função auxiliar para mover os dados de cada linha uma posição para cima"""
                    for row in range(linha_excluida, ws.max_row):
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=row, column=col).value = ws.cell(
                                row=row + 1, column=col
                            ).value
                    # Limpar a última linha
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=ws.max_row, column=col).value = None

                # Excluir paciente de médico
                if renach in pacientes_medicos:
                    for linha in pacientes_medicos[renach]:
                        reorganizar_linhas(linha)

                # Excluir paciente de psicólogo
                if renach in pacientes_psicologos:
                    for linha in pacientes_psicologos[renach]:
                        reorganizar_linhas(linha)

                self.wb.save("CAMPSSA.xlsx")
                print("Paciente foi excluído com sucesso!")
            except ValueError:
                print("RENACH inválido. Por favor, insira um número válido.")

        tk.Button(excluir_window, text="Excluir", command=excluir_paciente).pack(
            pady=10
        )

        self.center(excluir_window)

    def exibir_informacao(self):
        """Exibe informações dos pacientes em uma nova janela com barra de rolagem."""
        # Carrega o workbook e seleciona a sheet correta
        wb = load_workbook(self.file_path)
        if hasattr(self.planilhas, 'sheet_name') and self.planilhas.sheet_name:
            ws = wb[self.planilhas.sheet_name]
        else:  
            ws = wb.active
        
        medico, psi = [], []

        # Informações de médicos
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=6):
            linha = [cell.value for cell in row if isinstance(cell.value, (str, int)) and str(cell.value).strip()]
            if linha:
                medico.append(linha)

        # Informações de psicólogos
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=8, max_col=12):
            linha = [cell.value for cell in row if isinstance(cell.value, (str, int)) and str(cell.value).strip()]
            if linha:
                psi.append(linha)

        wb.close()  # Fecha o workbook após usar

        # Verificação de dados coletados
        if not medico and not psi:
            messagebox.showinfo("Aviso", "Nenhuma informação encontrada!")
            return

        # Criando uma nova janela
        janela_informacao = tk.Toplevel(self.master)
        janela_informacao.title("Informação dos Pacientes")
        janela_informacao.geometry("600x600")
        cor_fundo = self.master.cget("bg")
        janela_informacao.configure(bg=cor_fundo)

        # Canvas e scrollbar
        canvas = tk.Canvas(janela_informacao, bg=cor_fundo)
        scrollbar = tk.Scrollbar(janela_informacao, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=cor_fundo)
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Bind para rolagem com o mouse
        def scroll(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Adicionar informações de médicos
        if medico:
            tk.Label(scrollable_frame, text="MÉDICO:", font=("Arial", 16, "bold"), bg=cor_fundo, fg="#ECF0F1").pack(pady=(10, 0), anchor="center")
            for i, paciente in enumerate(medico, start=1):
                tk.Label(scrollable_frame, text=f"{i} - {paciente}", bg=cor_fundo, fg="#ECF0F1", font=("Arial", 12)).pack(anchor="w", padx=10, pady=5)

        # Adicionar informações de psicólogos
        if psi:
            tk.Label(scrollable_frame, text="PSICÓLOGO:", font=("Arial", 16, "bold"), bg=cor_fundo, fg="#ECF0F1").pack(pady=(10, 0), anchor="center")
            for i, paciente in enumerate(psi, start=1):
                tk.Label(scrollable_frame, text=f"{i} - {paciente}", bg=cor_fundo, fg="#ECF0F1", font=("Arial", 12)).pack(anchor="w", padx=10, pady=5)

        # Adicionar widgets à janela
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Detectar o sistema operacional para ajustar rolagem com mouse
        import sys
        if sys.platform.startswith("win") or sys.platform == "darwin":  # Windows e MacOS
            janela_informacao.bind_all("<MouseWheel>", scroll)
        else:  # Linux
            janela_informacao.bind_all("<Button-4>", lambda event: canvas.yview_scroll(-1, "units"))
            janela_informacao.bind_all("<Button-5>", lambda event: canvas.yview_scroll(1, "units"))

        # Remover bindings do mouse ao fechar a janela
        def on_closing():
            janela_informacao.unbind_all("<MouseWheel>")
            janela_informacao.unbind_all("<Button-4>")
            janela_informacao.unbind_all("<Button-5>")
            janela_informacao.destroy()

        janela_informacao.protocol("WM_DELETE_WINDOW", on_closing)

        # Configurar região de rolagem
        scrollable_frame.update_idletasks()  # Atualiza o frame antes de definir o scrollregion
        canvas.configure(scrollregion=canvas.bbox("all"))

    def valores_totais(self):
        n_medico, pag_medico = self.planilhas.contar_medico()
        n_psicologo, pag_psicologo = self.planilhas.contar_psi()

        total_medico = n_medico * 148.65
        total_psicologo = n_psicologo * 192.61

        valor_medico = n_medico * 49
        valor_psicologo = n_psicologo * 63.50

        janela_contas = tk.Toplevel(self.master)
        janela_contas.geometry("250x220")
        janela_contas.maxsize(width=250, height=220)
        janela_contas.minsize(width=250, height=220)
        janela_contas.configure(bg="#2C3E50")

        tk.Label(
            janela_contas,
            text="Contas Médico:",
            font=("Arial", 16),
            bg="#2C3E50",
            fg="#ECF0F1",
        ).pack(pady=5)
        tk.Label(
            janela_contas,
            text=f"Valor Total: {total_medico:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)
        tk.Label(
            janela_contas,
            text=f"Pagar: {valor_medico:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)
        tk.Label(janela_contas, text="", bg="#2C3E50").pack()

        tk.Label(
            janela_contas,
            text="Contas Psicólogo:",
            font=("Arial", 16),
            bg="#2C3E50",
            fg="#ECF0F1",
        ).pack(pady=5)
        tk.Label(
            janela_contas,
            text=f"Valor Total: {total_psicologo:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)
        tk.Label(
            janela_contas,
            text=f"Pagar: {valor_psicologo:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        self.center(janela_contas)

    def processar_notas_fiscais(self):
        driver = webdriver.Chrome()
        cpfs = {"medico": [], "psicologo": [], "ambos": []}

        try:
            # Ler o arquivo Excel
            logging.info("Lendo o arquivo Excel")
            df = pd.read_excel(
                self.file_path, skiprows=1, sheet_name="17.10", dtype={"Renach": str}
            )
        except Exception as e:
            logging.error(f"Erro ao ler o arquivo Excel: {e}")
            messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
            return

        # logging.info(f'DataFrame lido: {df.head()}')
        logging.info("DataFrame lido!")

        try:
            renach_c = df.iloc[:, 2].dropna().tolist()
            renach_i = df.iloc[:, 8].dropna().tolist()

            # Acessando o site e fazendo login
            logging.info("Acessando o site do DETRAN e fazendo login")
            driver.get("https://clinicas.detran.ba.gov.br/")
            usuario = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="documento"]'))
            )
            doc = "11599160000115"
            for numero in doc:
                usuario.send_keys(numero)

            actions = ActionChains(driver)
            actions.send_keys(Keys.TAB).perform()
            time.sleep(1)

            senha = "475869"
            campo_senha = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="senha"]'))
            )
            for numero in senha:
                campo_senha.send_keys(numero)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="acessar"]'))
            ).click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/aside/section/ul/li[2]/a/span[1]")
                )
            ).click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/aside/section/ul/li[2]/ul/li/a/span")
                )
            ).click()

            barra_pesquisa = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="list_items_filter"]/label/input')
                )
            )

            # coletando informações do cliente
            logging.info("Coletando informações de CPFs")

            def coletar_cpf(dados, tipo):
                for dado in dados:
                    dado = str(dado).strip()
                    barra_pesquisa.clear()
                    barra_pesquisa.send_keys(dado)
                    time.sleep(2)
                    try:
                        paciente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, '//*[@id="list_items"]/tbody/tr/td[3]')
                            )
                        )
                        cpf = paciente.text

                        print(f"coletado dado: {dado}, tipo: {tipo}")

                        if tipo == "medico" and dado in renach_i:
                            cpfs["ambos"].append(cpf)
                        elif tipo == "medico":
                            cpfs["medico"].append(cpf)
                        elif tipo == "psicologo" and cpf not in cpfs["ambos"]:
                            cpfs["psicologo"].append(cpf)
                    except Exception as e:
                        logging.error(f"Error ao coletar CPF: {e}")

            coletar_cpf(renach_c, "medico")
            coletar_cpf(renach_i, "psicologo")

            cpfs["medico"] = [cpf for cpf in cpfs["medico"] if cpf not in cpfs["ambos"]]
            cpfs["psicologo"] = [
                cpf for cpf in cpfs["psicologo"] if cpf not in cpfs["ambos"]
            ]

            logging.info("Acessando site para emissão de NTFS-e")
            driver.get("https://nfse.salvador.ba.gov.br/")

            # usuario
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="txtLogin"]'))
            ).send_keys("11599160000115")
            # senha
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="txtSenha"]'))
            ).send_keys("486258camp")
            # esperar resolver o captcha
            WebDriverWait(driver, 30).until(
                EC.invisibility_of_element_located((By.XPATH, '//*[@id="img1"]'))
            )
            # emissao NFS-e
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="menu-lateral"]/li[1]/a')
                )
            ).click()

            def emitir_nota(cpf, tipo):
                try:
                    barra_pesquisa = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, '//*[@id="tbCPFCNPJTomador"]')
                        )
                    )
                    barra_pesquisa.clear()
                    barra_pesquisa.send_keys(cpf)
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btAvancar"]'))
                    ).click()

                    # Complete as informações da nota fiscal
                    role_e_click(driver, '//*[@id="ddlCNAE_chosen"]/a')
                    print("cnae clicada")
                    # opcao cane
                    WebDriverWait(driver, 30).until(
                        EC.visibility_of_element_located(
                            (By.XPATH, '//*[@id="ddlCNAE_chosen"]/div/ul/li[2]')
                        )
                    ).click()
                    print("opcao cnae visivel")
                    # aliq %
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//*[@id="tbAliquota"]')
                        )
                    ).send_keys("2,5")

                    servicos = {
                        "ambos": "Exame de sanidade física e mental",
                        "psicologo": "Exame de sanidade mental",
                        "medico": "Exame de sanidade física",
                    }
                    # preenchendo o tipo de serviço
                    tipo_servico = servicos.get(
                        tipo, "Exame de sanidade física"
                    )  # valor padrao, caso 'tipo' nao esteja no dicionario

                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//*[@id="tbDiscriminacao"]')
                        )
                    ).send_keys(tipo_servico)

                    valor_nota = (
                        "148,65"
                        if tipo == "medico"
                        else "192,61" if tipo == "psicolgo" else "341,26"
                    )
                    # valor pago na consulta
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="tbValor"]'))
                    ).send_keys(valor_nota)
                    # emitindo nota
                    WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btEmitir"]'))
                    ).click()
                    # aceitando o alerta
                    WebDriverWait(driver, 20).until(EC.alert_is_present())
                    alert = Alert(driver)
                    alert.accept()
                    # botao voltar - voltando para emissao de nota fiscal por cpf
                    WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btVoltar"]'))
                    ).click()
                    logging.info(f"Nota emitida para o CPF: {cpf}, Valor: {valor_nota}")

                except Exception as e:
                    logging.error("Erro ao emitir nota: {e}")

            # emitir notas
            try:
                for cpf in cpfs["medico"]:
                    emitir_nota(cpf, "medico")
                for cpf in cpfs["psicologo"]:
                    emitir_nota(cpf, "psicologo")
                for cpf in cpfs["ambos"]:
                    emitir_nota(cpf, "ambos")
            except Exception as e:
                logging.error(f"Erro na emissao das notas: {e}")
        finally:
            driver.quit()
            logging.info("Processo finalizado")
            return cpfs

    def exibir_resultado(self):
        """Exibe os resultados de contagem para médicos e psicólogos."""
        janela_exibir_resultado = tk.Toplevel(self.master)
        janela_exibir_resultado.geometry("300x210")
        janela_exibir_resultado.maxsize(width=300, height=210)
        janela_exibir_resultado.minsize(width=300, height=210)

        # usando a cor de fundo da janela principal
        cor_fundo = self.master.cget("bg")
        janela_exibir_resultado.configure(bg=cor_fundo)

        n_medico, pag_medico = self.planilhas.contar_medico()
        n_psicologo, pag_psicologo = self.planilhas.contar_psi()

        # Criando rótulos (Labels) e adicionando à janela
        tk.Label(
            janela_exibir_resultado,
            text="MÉDICO:",
            font=("Arial", 16, "bold"),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack(pady=(15, 0))
        tk.Label(
            janela_exibir_resultado,
            text=f"Pacientes: {n_medico}",
            font=("Arial", 12),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack()

        # Formatando as formas de pagamento
        texto_med = "  ".join(
            [
                f"{tipo_pagamento}: {quantidade}"
                for tipo_pagamento, quantidade in pag_medico.items()
            ]
        )
        label = tk.Label(
            janela_exibir_resultado,
            text=texto_med,
            font=("Arial", 12),
            bg=cor_fundo,
            fg="#ECF0F1",
        )
        label.pack()  # Adicionando o label com as formas de pagamento

        # Se você também quiser exibir informações sobre psicólogos, adicione aqui:
        tk.Label(
            janela_exibir_resultado,
            text="PSICÓLOGO:",
            font=("Arial", 16, "bold"),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack(pady=(20, 0))
        tk.Label(
            janela_exibir_resultado,
            text=f"Pacientes: {n_psicologo}",
            font=("Arial", 12),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack()

        texto_psic = "  ".join(
            [
                f"{tipo_pagamento}: {quantidade}"
                for tipo_pagamento, quantidade in pag_psicologo.items()
            ]
        )
        label_psic = tk.Label(
            janela_exibir_resultado,
            text=texto_psic,
            font=("Arial", 12),
            bg=cor_fundo,
            fg="#ECF0F1",
        )
        label_psic.pack()  # Adicionando o label com as formas de pagamento dos psicólogos

        self.center(janela_exibir_resultado)

    def enviar_whatsapp(self):
        # Janela número ou nome do grupo
        janela_wpp = tk.Toplevel(self.master)
        janela_wpp.geometry("300x210")
        cor_fundo = self.master.cget("bg")
        janela_wpp.configure(bg=cor_fundo)
        self.center(janela_wpp)

        tk.Label(
            janela_wpp,
            text="Enviar para:",
            font=("Arial", 16, "bold"),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack(anchor="center", padx=5, pady=5)

        self.wpp_entry = tk.Entry(janela_wpp)
        self.wpp_entry.pack(padx=5, pady=5)

        # Checkbutton para salvar as informações
        tk.Button(
            janela_wpp, text="Enviar", command=self.processar_envio_whatsapp
        ).pack(pady=10)

    def processar_envio_whatsapp(self):
        # Captura o valor do campo de entrada
        group_name = self.wpp_entry.get().strip()

        if not group_name:
            messagebox.showerror("Erro", "Insira um número, grupo ou nome")
            return

        # Preparar as informações para enviar a mensagem
        n_medico, pag_medico = self.planilhas.contar_medico()
        n_psicologo, pag_psicologo = self.planilhas.contar_psi()

        valor_medico = n_medico * 49
        valor_psicologo = n_psicologo * 63.50

        message_medico = f"Valor medico: {valor_medico}"
        message_psicologo = f"Valor psicologo: {valor_psicologo}"

        dir_path = os.getcwd()
        profile = os.path.join(dir_path, "profile", "wpp")

        # Configurar opções do Chrome
        logging.info("Configurando Chrome...")
        options = Options()
        options.add_argument(r"user-data-dir={}".format(profile))
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        # Inicializar o WebDriver
        logging.info("Inicializando o WebDriver...")
        service = Service(executable_path="/usr/local/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=options)

        # Acessar o WhatsApp Web
        logging.info("Abrindo WhatsApp...")
        driver.get("https://web.whatsapp.com/")

        # Aguardar até que a página esteja totalmente carregada
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"]'))
            )
        except Exception as e:
            messagebox.showerror("Erro" f"Erro ao esnaear QR Code: {e}")
            logging.error(f"Erro ao escanear QR Code: {e}")
            driver.quit()
            return

        # Selecionar o grupo
        try:
            logging.info("Enviando mensagem...")
            barra_pesquisa = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located(
                    (By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div/p')
                )
            )
            barra_pesquisa.send_keys(group_name)
            time.sleep(1)
            barra_pesquisa.send_keys(Keys.ENTER)
        except Exception as e:
            messagebox.showerror(f"Erro ao selecionar grupo: {e}")
            logging.error(f"Erro ao seleconar o grupo: {e}")
            driver.quit()
            return

        # Enviar a mensagem
        try:
            logging.info("Enviando mensagem...")
            enviar_mensagem = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p',
                    )
                )
            )
            enviar_mensagem.send_keys(message_medico)
            time.sleep(1)
            enviar_mensagem.send_keys(Keys.ENTER)

            enviar_mensagem = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p',
                    )
                )
            )
            enviar_mensagem.send_keys(message_psicologo)
            time.sleep(1)
            enviar_mensagem.send_keys(Keys.ENTER)
            messagebox.showinfo("Mensagens enviadas")

            time.sleep(7)
        except Exception as e:
            messagebox.showerror(f"Erro ao enviar as mensagens: {e}")
        finally:
            driver.quit()

    def enviar_email(self):
        janela_email = tk.Toplevel(self.master)
        janela_email.geometry("300x400")
        cor_fundo = self.master.cget("bg")
        janela_email.configure(bg=cor_fundo)
        self.center(janela_email)

        tk.Label(
            janela_email,
            text="Email:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_email = tk.Entry(janela_email)
        entry_email.pack(pady=5)

        tk.Label(
            janela_email,
            text="Senha:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_senha = tk.Entry(janela_email, show="*")  # Ocultar senha
        entry_senha.pack(pady=5)

        tk.Label(
            janela_email,
            text="Destinatário:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_destinatario = tk.Entry(janela_email)
        entry_destinatario.pack(pady=5)

        tk.Label(
            janela_email,
            text="Assunto:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_assunto = tk.Entry(janela_email)
        entry_assunto.pack(pady=5)

        tk.Button(
            janela_email,
            text="Selecionar XLSX",
            command=lambda: self.selecionar_xlsx(
                entry_email.get(),
                entry_senha.get(),
                entry_destinatario.get(),
                entry_assunto.get(),
            ),
        ).pack(pady=20)

    def selecionar_xlsx(self, email, senha, destinatario, assunto):
        """
        Abre diálogo para selecionar arquivo XLSX
        """
        if not all([email, senha, destinatario, assunto]):
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        arquivo_xlsx = filedialog.askopenfilename(
            title="Selecione o arquivo XLSX",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        )

        if arquivo_xlsx:
            self.enviar(email, senha, destinatario, assunto, arquivo_xlsx)

    def enviar(self, email, senha, destinatario, assunto, caminho_xlsx):
        """
        Envia e-mail com arquivo XLSX anexado
        """
        smtp_server = "smtp.gmail.com"  # Para Gmail
        smtp_port = 587

        try:
            # Criando a mensagem
            msg = MIMEMultipart()
            msg["From"] = email
            msg["To"] = destinatario
            msg["Subject"] = assunto

            # Corpo do e-mail padrão
            corpo = "Segue em anexo o arquivo XLSX conforme solicitado."
            msg.attach(MIMEText(corpo, "plain"))

            # Anexar arquivo XLSX
            with open(caminho_xlsx, "rb") as arquivo:
                parte_xlsx = MIMEApplication(arquivo.read(), _subtype="xlsx")
                parte_xlsx.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=os.path.basename(caminho_xlsx),
                )
                msg.attach(parte_xlsx)

            # Contexto SSL para conexão segura
            context = ssl.create_default_context()

            # Enviando o e-mail
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls(context=context)  # Inicia a segurança TLS
                server.login(email, senha)  # Faz login no servidor
                server.send_message(msg)  # Envia a mensagem

            # Mensagem de sucesso
            messagebox.showinfo(
                "Sucesso",
                f"E-mail enviado com sucesso para {destinatario}!\nAnexo: {os.path.basename(caminho_xlsx)}",
            )

        except smtplib.SMTPAuthenticationError:
            messagebox.showerror(
                "Erro de Autenticação",
                "Verifique seu email e senha. Use uma senha de aplicativo para o Gmail.",
            )
        except Exception as e:
            messagebox.showerror("Erro ao Enviar", f"Ocorreu um erro: {str(e)}")

    def configurar_frames(self, login_frame, criar_conta_frame):
        self.login_frame = login_frame
        self.criar_conta_frame = criar_conta_frame

    def mostrar_criar_conta(self):
        self.login_frame.grid_forget()
        self.criar_conta_frame.grid()

    def voltar_para_login(self):
        self.criar_conta_frame.grid_forget()
        self.login_frame.grid()

    def formatar_planilha(self):
        """Formata a planilha com os dados do usuário e data atual."""
        wb = load_workbook(self.file_path)
        ws = wb.active

        # Define a borda para as células
        borda = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        # Define estilos de fonte
        font_bold = Font(
            name="Arial", bold=True, size=11, color="000000"
        )  # Fonte em negrito
        font_regular = Font(name="Arial", size=11)

        # Define um alinhamento
        alignment_center = Alignment(horizontal="center", vertical="center")

        # Definindo largura das colunas
        ws.column_dimensions["I"].width = 13
        ws.column_dimensions["C"].width = 13
        ws.column_dimensions["B"].width = 55
        ws.column_dimensions["H"].width = 55

        # Preenchendo o topo da planilha com usuário e data atual
        usuario = (
            "Nome do Usuário"  # Substitua pelo método que você usa para obter o usuário
        )
        data_atual = datetime.now().strftime("%d/%m/%Y")
        ws["A1"] = f"({usuario}) Atendimento Médico {data_atual}"
        ws["G1"] = f"({usuario}) Atendimento Psicológico {data_atual}"

        # Aplicando a formatação ao cabeçalho
        ws["A1"].font = font_bold
        ws["A1"].alignment = alignment_center
        ws["G1"].font = font_bold
        ws["G1"].alignment = alignment_center

        # Valores fixos na planilha com formatação, Medico e Psicologo
        cabeçalhos = ["Ordem", "Nome", "Renach", "Reexames", "Valor"]
        for col, valor in enumerate(
            cabeçalhos, start=1
        ):  # start=1 para começar na coluna A
            cell = ws.cell(row=2, column=col)
            cell.value = valor
            cell.font = font_bold  # Aplica a formatação de fonte
            cell.alignment = alignment_center  # Aplica o alinhamento

        for col, valor in enumerate(
            cabeçalhos, start=7
        ):  # start=7 para começar na coluna G
            cell = ws.cell(row=2, column=col)
            cell.value = valor
            cell.font = font_bold  # Aplica a formatação de fonte
            cell.alignment = alignment_center  # Aplica o alinhamento

        # Mesclando células cabeçalho planilha
        ws.merge_cells("A1:E1")
        ws.merge_cells("G1:K1")

        # Primeiro, encontra onde termina a seção do médico e preenche os valores
        ultima_linha_nome_medico = None
        numero_pacientes_medico = 0
        for row in range(3, ws.max_row + 1):
            if ws[f"B{row}"].value is not None:
                ultima_linha_nome_medico = row
                numero_pacientes_medico += 1
                # Preenche o valor fixo de 148.65 na coluna E
                ws[f"E{row}"].value = 148.65
                ws[f"E{row}"].alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                ws[f"E{row}"].border = borda

        # Depois, encontra onde termina a seção do psicólogo e preenche os valores
        ultima_linha_nome_psicologo = None
        numero_pacientes_psicologo = 0
        for row in range(3, ws.max_row + 1):
            if ws[f"H{row}"].value is not None:
                ultima_linha_nome_psicologo = row
                numero_pacientes_psicologo += 1
                # Preenche o valor fixo de 192.61 na coluna K
                ws[f"K{row}"].value = 192.61
                ws[f"K{row}"].alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                ws[f"K{row}"].border = borda

        # Calcula as somas
        soma_medico = numero_pacientes_medico * 148.65
        soma_psicologo = numero_pacientes_psicologo * 192.61

        # Adiciona soma e valores adicionais para médico
        if ultima_linha_nome_medico is not None:
            # Linha da soma
            ws[f"D{ultima_linha_nome_medico + 1}"] = "Soma"
            ws[f"D{ultima_linha_nome_medico + 1}"].font = Font(bold=True)
            ws[f"D{ultima_linha_nome_medico + 1}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws[f"D{ultima_linha_nome_medico + 1}"].border = borda
            ws[f"E{ultima_linha_nome_medico + 1}"] = soma_medico
            ws[f"E{ultima_linha_nome_medico + 1}"].border = borda
            ws[f"E{ultima_linha_nome_medico + 1}"].font = Font(bold=True)
            ws[f"E{ultima_linha_nome_medico + 1}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            # Linha do valor por paciente
            ws[f"D{ultima_linha_nome_medico + 2}"] = "Médico"
            ws[f"D{ultima_linha_nome_medico + 2}"].font = Font(bold=True)
            ws[f"D{ultima_linha_nome_medico + 2}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws[f"D{ultima_linha_nome_medico + 2}"].border = borda
            valor_medico = numero_pacientes_medico * 49
            ws[f"E{ultima_linha_nome_medico + 2}"] = valor_medico
            ws[f"E{ultima_linha_nome_medico + 2}"].border = borda
            ws[f"E{ultima_linha_nome_medico + 2}"].font = Font(bold=True)
            ws[f"E{ultima_linha_nome_medico + 2}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            # Linha do total
            ws[f"D{ultima_linha_nome_medico + 3}"] = "Total"
            ws[f"D{ultima_linha_nome_medico + 3}"].font = Font(bold=True)
            ws[f"D{ultima_linha_nome_medico + 3}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws[f"D{ultima_linha_nome_medico + 3}"].border = borda
            total_medico = numero_pacientes_medico * (148.65 - 49)
            ws[f"E{ultima_linha_nome_medico + 3}"] = total_medico
            ws[f"E{ultima_linha_nome_medico + 3}"].border = borda
            ws[f"E{ultima_linha_nome_medico + 3}"].font = Font(bold=True)
            ws[f"E{ultima_linha_nome_medico + 3}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )

        # Adiciona soma e valores adicionais para psicólogo
        if ultima_linha_nome_psicologo is not None:
            # Linha da soma
            ws[f"J{ultima_linha_nome_psicologo + 1}"] = "Soma"
            ws[f"J{ultima_linha_nome_psicologo + 1}"].font = Font(bold=True)
            ws[f"J{ultima_linha_nome_psicologo + 1}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws[f"J{ultima_linha_nome_psicologo + 1}"].border = borda
            ws[f"K{ultima_linha_nome_psicologo + 1}"] = soma_psicologo
            ws[f"K{ultima_linha_nome_psicologo + 1}"].border = borda
            ws[f"K{ultima_linha_nome_psicologo + 1}"].font = Font(bold=True)
            ws[f"K{ultima_linha_nome_psicologo + 1}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            # Linha do valor por paciente
            ws[f"J{ultima_linha_nome_psicologo + 2}"] = "Psicólogo"
            ws[f"J{ultima_linha_nome_psicologo + 2}"].font = Font(bold=True)
            ws[f"J{ultima_linha_nome_psicologo + 2}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws[f"J{ultima_linha_nome_psicologo + 2}"].border = borda
            valor_psicologo = numero_pacientes_psicologo * 63.50
            ws[f"K{ultima_linha_nome_psicologo + 2}"] = valor_psicologo
            ws[f"K{ultima_linha_nome_psicologo + 2}"].border = borda
            ws[f"K{ultima_linha_nome_psicologo + 2}"].font = Font(bold=True)
            ws[f"K{ultima_linha_nome_psicologo + 2}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            # Linha do total
            ws[f"J{ultima_linha_nome_psicologo + 3}"] = "Total"
            ws[f"J{ultima_linha_nome_psicologo + 3}"].font = Font(bold=True)
            ws[f"J{ultima_linha_nome_psicologo + 3}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ws[f"J{ultima_linha_nome_psicologo + 3}"].border = borda
            total_psicologo = numero_pacientes_psicologo * (192.61 - 63.50)
            ws[f"K{ultima_linha_nome_psicologo + 3}"] = total_psicologo
            ws[f"K{ultima_linha_nome_psicologo + 3}"].border = borda
            ws[f"K{ultima_linha_nome_psicologo + 3}"].font = Font(bold=True)
            ws[f"K{ultima_linha_nome_psicologo + 3}"].alignment = Alignment(
                horizontal="center", vertical="center"
            )

        # Informações gerais do atendimento
        if ultima_linha_nome_psicologo is not None:
            medico = 49
            psicologo = 63.50
            total_clinica = (soma_medico + soma_psicologo) - (
                (numero_pacientes_medico * medico)
                + (numero_pacientes_psicologo * psicologo)
            )

            # Lista de células e valores a serem preenchidos
            cells_to_fill = [
                (
                    f"I{ultima_linha_nome_psicologo+8}",
                    "Atendimento Médico",
                    f"K{ultima_linha_nome_psicologo+8}",
                    soma_medico,
                ),
                (
                    f"I{ultima_linha_nome_psicologo+9}",
                    "Atendimento Psicológico",
                    f"K{ultima_linha_nome_psicologo+9}",
                    soma_psicologo,
                ),
                (
                    f"I{ultima_linha_nome_psicologo+10}",
                    "Total",
                    f"K{ultima_linha_nome_psicologo+10}",
                    soma_medico + soma_psicologo,
                ),
                (
                    f"I{ultima_linha_nome_psicologo+12}",
                    "Pagamento Médico",
                    f"K{ultima_linha_nome_psicologo+12}",
                    numero_pacientes_medico * medico,
                ),
                (
                    f"I{ultima_linha_nome_psicologo+13}",
                    "Pagamento Psicológico",
                    f"K{ultima_linha_nome_psicologo+13}",
                    numero_pacientes_psicologo * psicologo,
                ),
                (
                    f"I{ultima_linha_nome_psicologo+14}",
                    "Soma",
                    f"K{ultima_linha_nome_psicologo+14}",
                    total_clinica,
                ),
            ]

            # Loop para aplicar os valores e a formatação
            for left_cell, left_value, right_cell, right_value in cells_to_fill:
                ws[left_cell] = left_value
                ws[left_cell].font = font_bold  # Aplica fonte em negrito
                ws[right_cell] = right_value
                ws[right_cell].font = font_bold  # Aplica fonte em negrito

                # Alinhamento e bordas
                ws[left_cell].alignment = alignment_center
                ws[right_cell].alignment = alignment_center
                ws[left_cell].border = borda
                ws[right_cell].border = borda

            # Mescla as células das colunas I e J para cada linha
            for left_cell, _, _, _ in cells_to_fill:
                ws.merge_cells(f"{left_cell}:J{left_cell[1:]}")

        # Aplica bordas nas células preenchidas
        for row in ws.iter_rows(
            min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
        ):
            for cell in row:
                if cell.value is not None:
                    cell.border = borda

        # Salva e abre o arquivo no Linux
        wb.save(self.file_path)
        try:
            subprocess.run(["xdg-open", self.file_path])
        except Exception as e:
            print("Erro ao abrir o arquivo:", e)



class SistemaContas:
    def __init__(self, file_path: str, current_user=None):
        self.file_path = file_path
        self.current_user = current_user
        self.sheet_name = "Contas Fechamento"
        self.criar_sheet_se_nao_existir()

    def abrir_janela(self):
        """Cria uma nova janela para o sistema de contas"""
        self.window = tk.Toplevel()
        self.window.title("Sistema de Gerenciamento de Contas")
        self.window.geometry("500x400")
        self.criar_interface()

        # Configurar a janela como modal
        self.window.transient(self.window.master)
        self.window.grab_set()
        self.window.focus_set()

    def criar_sheet_se_nao_existir(self):
        """Cria a planilha e a aba (sheet) se não existirem."""
        if os.path.exists(self.file_path):
            wb = load_workbook(self.file_path)
            if self.sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(title=self.sheet_name)
                ws.append(["DATA", "CONTAS", "VALOR"])
                wb.save(self.file_path)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = self.sheet_name
            ws.append(["DATA", "CONTAS", "VALOR"])
            wb.save(self.file_path)

    def criar_interface(self):
        """Cria a interface gráfica usando grid layout"""
        # Configurando o frame principal
        main_frame = tk.Frame(self.window, padx=20, pady=20)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Configurando expansão do grid
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # Título
        title_label = tk.Label(
            main_frame, text="Gerenciamento de Contas", font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Data
        tk.Label(main_frame, text="Data:", font=("Arial", 10, "bold")).grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.date_entry = DateEntry(
            main_frame,
            width=20,
            date_pattern="dd/mm/yyyy",
            background="darkblue",
            foreground="white",
            borderwidth=2,
        )
        self.date_entry.grid(row=1, column=1, sticky="we", padx=(5, 0), pady=5)

        # Descrição
        tk.Label(main_frame, text="Descrição:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky="w", pady=5
        )
        self.info_entry = tk.Entry(main_frame)
        self.info_entry.grid(row=2, column=1, sticky="we", padx=(5, 0), pady=5)

        # Valor
        tk.Label(main_frame, text="Valor (R$):", font=("Arial", 10, "bold")).grid(
            row=3, column=0, sticky="w", pady=5
        )
        self.valor_entry = tk.Entry(main_frame)
        self.valor_entry.grid(row=3, column=1, sticky="we", padx=(5, 0), pady=5)

        # Frame para botões
        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        button_frame.grid_columnconfigure((0, 1), weight=1)

        # Botões
        save_button = tk.Button(
            button_frame,
            text="Salvar",
            command=self.capturar_dados,
            width=20,
            bg="#4CAF50",
            fg="white",
        )
        save_button.grid(row=0, column=0, padx=5)

        clear_button = tk.Button(
            button_frame, text="Limpar", command=self.limpar_campos, width=20
        )
        clear_button.grid(row=0, column=1, padx=5)

        # Botão Fechar
        close_button = tk.Button(
            button_frame, text="Fechar", command=self.window.destroy, width=20
        )
        close_button.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        # Frame para mensagens de status
        self.status_frame = tk.Frame(main_frame)
        self.status_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=(10, 0))

        self.status_label = tk.Label(self.status_frame, text="", foreground="green")
        self.status_label.grid(row=0, column=0, sticky="we")

        # Configurar foco inicial
        self.info_entry.focus()

    def salvar_informacoes(self, data_escolhida, info, valor):
        """Salva as informações na planilha, agrupando por data e colocando informações na mesma célula."""
        try:
            wb = load_workbook(self.file_path)
            ws = wb[self.sheet_name]

            try:
                data_formatada = datetime.strptime(data_escolhida, "%d/%m/%Y").date()
            except ValueError:
                messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA")
                return False

            dados = []
            for row in ws.iter_rows(min_row=2):
                if row[0].value:
                    dados.append(
                        {
                            "data": row[0].value.date(),
                            "info": row[1].value,
                            "valor": row[2].value,
                            "linha": row[0].row,
                        }
                    )

            data_existe = False
            for i, dado in enumerate(dados):
                if dado["data"] == data_formatada:
                    data_existe = True
                    dados[i]["info"] = (
                        f"{dado['info']}\n{info}" if dado["info"] else info
                    )
                    dados[i]["valor"] = (
                        f"{dado['valor']}\n{valor}" if dado["valor"] else valor
                    )
                    break

            if not data_existe:
                dados.append(
                    {
                        "data": data_formatada,
                        "info": info,
                        "valor": valor,
                        "linha": None,
                    }
                )

            dados_ordenados = sorted(dados, key=lambda x: x["data"])

            # Limpa os dados existentes
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.value = None

            for i, dado in enumerate(dados_ordenados, start=2):
                ws.cell(row=i, column=1).value = dado["data"]
                ws.cell(row=i, column=2).value = dado["info"]

                # Atribui o valor à célula e formata como moeda
                cell_valor = ws.cell(row=i, column=3)
                if dado["valor"] is not None:
                    cell_valor.value = dado["valor"]  # Aqui você armazena o valor
                    cell_valor.number_format = '"R$"#,##0.00'  # Formato de moeda
                else:
                    cell_valor.value = valor  # Caso não tenha dado anterior
                    cell_valor.number_format = '"R$"#,##0.00'  # Formato de moeda

                # Centraliza a data
                ws.cell(row=i, column=1).alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            # Ajusta a formatação de texto e alinhamento
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="center")

            # Ajusta a largura das colunas
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = max_length + 2
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(self.file_path)
            messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")
            return True

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar informações: {str(e)}")
            return False

    def validar_campos(self):
        """Valida os campos antes de salvar"""
        info = self.info_entry.get().strip()
        valor = self.valor_entry.get().strip()
        data = self.date_entry.get().strip()

        if not all([data, info, valor]):
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return False

        try:
            float(valor.replace(",", "."))
            return True
        except ValueError:
            messagebox.showerror("Erro", "O valor deve ser um número válido!")
            return False

    def limpar_campos(self):
        """Limpa os campos após salvar"""
        self.info_entry.delete(0, tk.END)
        self.valor_entry.delete(0, tk.END)

    def capturar_dados(self):
        """Captura e processa os dados do formulário"""
        if self.validar_campos():
            data = self.date_entry.get()
            info = self.info_entry.get()
            valor = self.valor_entry.get().replace(",", ".")

            try:
                # Converte o valor para float e formata como moeda
                valor_float = float(valor)
                valor_formatado = f"R$ {valor_float:,.2f}"  # Formatação para moeda

                # Chama a função de salvar com o valor formatado
                if self.salvar_informacoes(data, info, valor_formatado):
                    self.limpar_campos()
            except ValueError:
                messagebox.showerror(
                    "Erro", "Por favor, insira um valor numérico válido."
                )


class GerenciadorPlanilhas:
    def __init__(self, master, sistema_contas):
        self.master = master
        self.sistema_contas = sistema_contas
        self.file_path = None
        self.sheet_name = None
        self.active_window = None
        
    def abrir_gerenciador(self):
        """Abre a janela de gerenciamento de planilhas"""
        if self.active_window:
            self.active_window.lift()
            return
            
        self.active_window = Toplevel(self.master)
        self.active_window.title("Gerenciador de Planilhas")
        self.active_window.geometry('600x700')
        self.active_window.resizable(False, False)
        
        # Centralizar a janela
        window_width = 600
        window_height = 700
        screen_width = self.active_window.winfo_screenwidth()
        screen_height = self.active_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.active_window.geometry(f'{window_width}x{window_height}+{x}+{y}')

        # Configurar grid da janela
        self.active_window.grid_columnconfigure(0, weight=1)
        self.active_window.grid_rowconfigure(0, weight=1)

        self._setup_interface()
        
        # Cleanup quando a janela for fechada
        self.active_window.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Tornar a janela modal
        self.active_window.transient(self.master)
        self.active_window.grab_set()

    def _setup_interface(self):
        """Configura a interface do gerenciador"""
        # Frame principal com padding
        main_frame = ttk.Frame(self.active_window, padding="20 20 20 20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        # Título
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        title_frame.grid_columnconfigure(0, weight=1)

        title_label = ttk.Label(
            title_frame,
            text="Gerenciador de Planilhas Excel",
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0)

        # Frame para arquivo atual
        file_frame = ttk.LabelFrame(main_frame, text="Arquivo Atual", padding="10")
        file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        file_frame.grid_columnconfigure(0, weight=1)

        self.lbl_arquivo = ttk.Label(
            file_frame,
            text=self.sistema_contas.file_path if hasattr(self.sistema_contas, 'file_path') else "Nenhum arquivo selecionado",
            wraplength=500
        )
        self.lbl_arquivo.grid(row=0, column=0, sticky="ew", padx=5)

        # Frame para lista de sheets
        list_frame = ttk.LabelFrame(main_frame, text="Planilhas Disponíveis", padding="10")
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        # Container para lista e scrollbars
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky="nsew")
        list_container.grid_columnconfigure(0, weight=1)
        list_container.grid_rowconfigure(0, weight=1)

        self.listbox = Listbox(
            list_container,
            font=('Arial', 10),
            selectmode=SINGLE,
            height=10,
            borderwidth=1,
            relief="solid"
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(list_container, orient=VERTICAL, command=self.listbox.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        
        scrollbar_x = ttk.Scrollbar(list_container, orient=HORIZONTAL, command=self.listbox.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self.listbox.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        # Frame para criar nova sheet
        create_frame = ttk.LabelFrame(main_frame, text="Criar Nova Planilha", padding="10")
        create_frame.grid(row=3, column=0, sticky="ew", pady=(0, 20))
        create_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(
            create_frame,
            text="Nome:",
            font=('Arial', 10)
        ).grid(row=0, column=0, padx=(0, 10), sticky="w")

        self.nova_sheet_entry = ttk.Entry(create_frame)
        self.nova_sheet_entry.grid(row=0, column=1, sticky="ew")

        # Frame para botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, sticky="ew")
        for i in range(2):
            button_frame.grid_columnconfigure(i, weight=1)

        # Primeira linha de botões
        ttk.Button(
            button_frame,
            text="Nova Planilha Excel",
            command=self.criar_nova_planilha
        ).grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        ttk.Button(
            button_frame,
            text="Abrir Planilha Existente",
            command=self.abrir_planilha
        ).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Segunda linha de botões
        ttk.Button(
            button_frame,
            text="Selecionar Sheet",
            command=self.selecionar_sheet
        ).grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        ttk.Button(
            button_frame,
            text="Criar Nova Sheet",
            command=self.criar_nova_sheet
        ).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.atualizar_lista_sheets()

    def criar_nova_planilha(self):
        """Cria um novo arquivo Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                wb = Workbook()
                wb.save(file_path)
                self.sistema_contas.file_path = file_path
                self.lbl_arquivo.config(text=file_path)
                self.atualizar_lista_sheets()
                messagebox.showinfo("Sucesso", "Nova planilha Excel criada com sucesso!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao criar planilha: {str(e)}")

    def abrir_planilha(self):
        """Abre uma planilha Excel existente"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                wb = load_workbook(file_path)
                self.sistema_contas.file_path = file_path
                
                # Pega a sheet ativa atual
                active_sheet = wb.active
                self.sistema_contas.sheet_name = active_sheet.title
                
                wb.close()
                self.lbl_arquivo.config(text=file_path)
                self.atualizar_lista_sheets()
                messagebox.showinfo("Sucesso", f"Planilha aberta com sucesso! Sheet ativa: {active_sheet.title}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir planilha: {str(e)}")

    def atualizar_lista_sheets(self):
        """Atualiza a lista de sheets disponíveis"""
        self.listbox.delete(0, END)
        if hasattr(self.sistema_contas, 'file_path') and self.sistema_contas.file_path and os.path.exists(self.sistema_contas.file_path):
            try:
                wb = load_workbook(self.sistema_contas.file_path)
                for sheet in wb.sheetnames:
                    self.listbox.insert(END, sheet)
                wb.close()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao listar planilhas: {str(e)}")

    def selecionar_sheet(self):
        """Seleciona uma sheet existente e a torna ativa"""
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showerror("Erro", "Selecione uma planilha!")
            return
            
        nome_sheet = self.listbox.get(selection[0])
        try:
            wb = load_workbook(self.sistema_contas.file_path)
            if nome_sheet in wb.sheetnames:
                # Define a sheet selecionada como ativa
                wb.active = wb[nome_sheet]
                wb.save(self.sistema_contas.file_path)
                
                # Atualiza o nome da sheet no sistema_contas
                self.sistema_contas.sheet_name = nome_sheet
                
                wb.close()
                messagebox.showinfo("Sucesso", f"Planilha '{nome_sheet}' selecionada e ativada!")
                self.active_window.destroy()
                self.active_window = None
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar planilha: {str(e)}")

    def criar_nova_sheet(self):
        """Cria uma nova sheet e a torna ativa"""
        nome_sheet = self.nova_sheet_entry.get().strip()
        if not nome_sheet:
            messagebox.showerror("Erro", "Digite um nome para a nova planilha!")
            return

        if not hasattr(self.sistema_contas, 'file_path') or not self.sistema_contas.file_path:
            messagebox.showerror("Erro", "Primeiro abra ou crie uma planilha Excel!")
            return

        try:
            wb = load_workbook(self.sistema_contas.file_path)
            if nome_sheet in wb.sheetnames:
                messagebox.showerror("Erro", "Já existe uma planilha com este nome!")
                wb.close()
                return

            # Cria nova sheet e a torna ativa
            new_sheet = wb.create_sheet(title=nome_sheet)
            wb.active = new_sheet
            wb.save(self.sistema_contas.file_path)
            wb.close()
            
            # Atualiza o nome da sheet no sistema_contas
            self.sistema_contas.sheet_name = nome_sheet
            
            self.atualizar_lista_sheets()
            messagebox.showinfo("Sucesso", f"Planilha '{nome_sheet}' criada e ativada com sucesso!")
            self.active_window.destroy()
            self.active_window = None
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar planilha: {str(e)}")

    def _on_closing(self):
        """Handler para quando a janela for fechada"""
        self.active_window.destroy()
        self.active_window = None