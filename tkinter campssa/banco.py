import sqlite3
import tkinter as tk
from tkinter import messagebox, Toplevel, Frame, Label
from tkcalendar import DateEntry
from funcoes_botoes import FuncoesBotoes
from planilhas import Planilhas


'''def create_db():
    conn = sqlite3.connect("db_marcacao.db")
    cursor = conn.cursor()

    # Criação da tabela de usuários
    cursor.execute(
        """CREATE TABLE IF NOT EXISTS patients (
                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                   name TEXT NOT NULL,
                   renach TEXT NOT NULL,
                   phone TEXT,
                   appointment_date TEXT NOT NULL)"""
    )

    conn.commit()
    conn.close()'''


# Função CRUD
class DataBaseLogin:
    def __init__(self, db_name="login.db"):
        self.db_name = db_name
        self.create_db()

    # Função para criar o banco de dados e a tabela de usuários
    def create_db(self):
        conn = sqlite3.connect("login.db")
        cursor = conn.cursor()

        # Criação da tabela de usuários
        cursor.execute(
            """CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user TEXT NOT NULL,
                    password TEXT NOT NULL)"""
        )

        conn.commit()
        conn.close()

    # Função para criar novo usuário
    def create_user(self, user, password):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO users (user, password) VALUES (?, ?)", (user, password)
        )
        conn.commit()
        conn.close()
        print(f"usuario {user} criado com sucesso")

    # Função para ser um usuário com base no user
    def read_user(self, user):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE user = ?", (user,))
        usuario = cursor.fetchone()
        conn.close()
        return usuario

    # Função para atualizar a senha de um usuário
    def update_user(self, user, new_password):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE users SET password = ? WHERE user =?", (new_password, user)
        )

    # Função para deletar um usuário com base no user
    def delete_user(self, user):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE user = ?", (user,))
        conn.comit()
        conn.close()


class DataBaseMarcacao:
    def __init__(self, master, planilhas:Planilhas, file_path:str, app, db_name="db_marcacao.db"):
        self.db_name = db_name
        self.master = master
        self.create_db() # Cria o banco de dados e a tabela
        self.funcoes_botoes = FuncoesBotoes(self.master, planilhas, file_path, app)

        # Inicializar os campos de entrada como atributos da classe
        self.name_entry = None
        self.renach_entry = None
        self.phone_entry = None
        self.appointment_entry = None

    def create_db(self):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        cursor.execute(
        """CREATE TABLE IF NOT EXISTS patients (
                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                   name TEXT NOT NULL,
                   renach TEXT NOT NULL,
                   phone TEXT,
                   appointment_date TEXT NOT NULL)"""
    )
        conn.commit()
        conn.close()

    def add_user(self):
        # Função para formatar o número de telefone
        def format_phone(phone):
            phone = ''.join(filter(str.isdigit, phone))  # Remove tudo exceto os dígitos
            if len(phone) == 11:
                return f"({phone[:2]}) {phone[2:7]}-{phone[7:]}"
            elif len(phone) == 10:
                return f"({phone[:2]}) {phone[2:6]}-{phone[6:]}"
            else:
                return phone  # Retorna sem formatação se não estiver no tamanho esperado
            
        # Funcao para inserir o paciente no banco de dados
        def submit_patient():
            name = self.name_entry.get()
            renach = self.renach_entry.get()
            phone = format_phone(self.phone_entry.get())
            appointment_date = self.appointment_entry.get_date().strftime("%Y-%m-%d")

            if not name or not renach or not phone or not appointment_date:
                messagebox.showerror("Preencha todos os campos!")
                return

            # Adiciona o paciente no banco de dados
            self.add_patient(name, renach, phone, appointment_date)
            messagebox.showinfo("Paciente adicionado com sucesso!")

            # Limpa os campos
            self.name_entry.delete(0, tk.END)
            self.renach_entry.delete(0, tk.END)
            self.phone_entry.delete(0, tk.END)
            self.appointment_entry.delete(0, tk.END)

        # Criar a janela Tkinter
        self.window = tk.Toplevel()
        self.window.geometry("300x300")
        self.window.minsize(width=300, height=300)
        self.window.maxsize(width=300, height=300)
        self.window.title("Marcar Paciente")

        # Configurando cores
        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"
        self.window.configure(bg=cor_fundo)

        # Frame para os campos de entrada
        frame_principal = tk.Frame(self.window, bg=cor_fundo)
        frame_principal.pack(pady=10)

        # Campos de entrada
        tk.Label(frame_principal, text="Nome:", bg=cor_fundo, fg=cor_texto).pack(anchor='w', pady=5)  # Alinha à esquerda
        self.name_entry = tk.Entry(frame_principal)
        self.name_entry.pack(fill='x', padx=5)  # Preenche horizontalmente

        tk.Label(frame_principal, text="Renach:", bg=cor_fundo, fg=cor_texto).pack(anchor='w', pady=5)  # Alinha à esquerda
        self.renach_entry = tk.Entry(frame_principal)
        self.renach_entry.pack(fill='x', padx=5)  # Preenche horizontalmente

        tk.Label(frame_principal, text="Phone:", bg=cor_fundo, fg=cor_texto).pack(anchor='w', pady=5)  # Alinha à esquerda
        self.phone_entry = tk.Entry(frame_principal)
        self.phone_entry.pack(fill='x', padx=5)  # Preenche horizontalmente

        tk.Label(frame_principal, text="Data:", bg=cor_fundo, fg=cor_texto).pack(anchor='w', pady=5)  # Alinha à esquerda
        self.appointment_entry = DateEntry(
            frame_principal,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
        )
        self.appointment_entry.pack(pady=5)

        # Botão para adicionar paciente
        add_button = tk.Button(
            self.window, text="Adicionar Paciente", command=submit_patient
        )
        add_button.pack(pady=5)

        self.funcoes_botoes.center(self.window)

    def add_patient(self, name, renach, phone, appointment_date):
        """Adiciona um novo paciente ao banco de dados."""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        # Inserir os dados na tabela de pacientes
        cursor.execute(
                """
                INSERT INTO patients (name, renach, phone, appointment_date) VALUES (?, ?, ?, ?)
                """,
                (name, renach, phone, appointment_date),
            )

        conn.commit()
        conn.close()

    def view_marcacoes(self):
        self.marcacoes_window = Toplevel(self.master)
        self.marcacoes_window.title("Visualizar Marcações")
        self.marcacoes_window.geometry("400x400")
        self.marcacoes_window.configure(bg=self.master.cget('bg'))

        Label(self.marcacoes_window, text="Selecione uma data para ver os pacientes:", bg=self.master.cget('bg'), fg='#ECF0F1').pack(pady=10)

        self.date_entry = DateEntry(self.marcacoes_window, width=12, background="darkblue", foreground="white", borderwidth=2)
        self.date_entry.pack(pady=10)

        self.date_entry.bind("<<DateEntrySelected>>", self.update_patient_list)

        # Frame para resultados formatados como uma tabela
        self.results_frame = Frame(self.marcacoes_window, bg=self.master.cget('bg'))
        self.results_frame.pack(fill="both", expand=True, pady=10)

    def get_patients_by_date(self, date):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("SELECT name, phone, renach FROM patients WHERE appointment_date = ?", (date,))
        patients = cursor.fetchall()
        conn.close()
        return patients

    def update_patient_list(self, event):
        selected_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        patients = self.get_patients_by_date(selected_date)
        
        # Limpa a tabela anterior (caso exista)
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        # Adiciona cabeçalho
        headers = ["Nome", "Telefone", "RENACH"]
        for j, header in enumerate(headers):
            header_cell = Label(self.results_frame, text=header, font=("Arial", 12, "bold"),
                                bg=self.master.cget('bg'), fg='#ECF0F1', width=15, anchor="w", borderwidth=1, relief="solid")
            header_cell.grid(row=0, column=j, padx=5, pady=2)
        
        # Cria a tabela com os dados dos pacientes
        for i, paciente in enumerate(patients, start=1):  # Começa na linha 1 para dar espaço para o cabeçalho
            for j, info in enumerate(paciente):
                cell = Label(self.results_frame, text=info, font=("Arial", 12),
                            bg=self.master.cget('bg'), fg='#ECF0F1', width=15, anchor="w", borderwidth=1, relief="solid")
                cell.grid(row=i, column=j, padx=5, pady=2)

"""def list_tables():
    # Conectar ao banco de dados
    connection = sqlite3.connect('db_marcacao.db')
    cursor = connection.cursor()
    
    # Consultar o sqlite_master para obter o nome das tabelas
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()
    
    # Imprimir o nome das tabelas
    for table in tables:
        print(table[0])  # Cada item é uma tupla, pegamos o primeiro elemento
    
    # Fechar a conexão
    connection.close()

# Chamar a função para listar as tabelas
list_tables()
"""
