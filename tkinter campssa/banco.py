import sqlite3
import tkinter as tk
import tkinter as ttk
from tkinter import messagebox, Toplevel, Frame, Label
from tkcalendar import DateEntry
from funcoes_botoes import FuncoesBotoes
from planilhas import Planilhas
import datetime


# Função CRUD
class DataBaseLogin:
    def __init__(self, db_name="login.db"):
        self.db_name = db_name
        self.conn = sqlite3.connect(db_name)
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

        # Verifica se o usuario já existe
        cursor.execute("SELECT * FROM users WHERE user = ?", (user,))
        resultado = cursor.fetchone()

        if resultado:
            conn.close()
            return False 

        # Se nao existir, cria um novo usuário
        cursor.execute(
            "INSERT INTO users (user, password) VALUES (?, ?)", (user, password)
        )
        conn.commit()
        conn.close()
        return True

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

    def validate_user(self, user, password):
        """Verifica se o nome de usuário e a senha são válidos no banco de dados."""
        cursor = self.conn.cursor()
        cursor.execute('SELECT * FROM users WHERE user = ? AND password = ?', (user, password))
        result = cursor.fetchone()
        return result is not None

class DataBaseMarcacao:
    def __init__(self, master, planilhas: Planilhas, file_path: str, app, db_name="db_marcacao.db"):
        self.db_name = db_name
        self.master = master
        self.create_db()
        self.funcoes_botoes = FuncoesBotoes(self.master, planilhas, file_path, app)

        # Inicializar os campos de entrada como atributos da classe
        self.name_entry = None
        self.renach_entry = None
        self.phone_entry = None
        self.appointment_entry = None
        self.observation_text = None
        self.window = None
        self.marcacoes_window = None
        self.results_frame = None
        self.date_entry = None

    def create_db(self):
        """Cria o banco de dados e a tabela de pacientes com o novo campo de observação."""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS patients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                renach TEXT NOT NULL,
                phone TEXT,
                appointment_date TEXT NOT NULL,
                observation TEXT
            )
        """)
        conn.commit()
        conn.close()

    def format_phone(self, phone):
        """Formata o número de telefone no padrão (XX) XXXXX-XXXX ou (XX) XXXX-XXXX."""
        phone = ''.join(filter(str.isdigit, phone))
        if len(phone) == 11:
            return f"({phone[:2]}) {phone[2:7]}-{phone[7:]}"
        elif len(phone) == 10:
            return f"({phone[:2]}) {phone[2:6]}-{phone[6:]}"
        return phone

    def validate_fields(self):
        """Valida os campos obrigatórios do formulário."""
        name = self.name_entry.get().strip()
        renach = self.renach_entry.get().strip()
        
        if not all([name, renach]):
            messagebox.showerror("Erro", "Por favor, preencha todos os campos obrigatórios!")
            return False
        
        if not renach.isdigit():
            messagebox.showerror("Erro", "O RENACH deve conter apenas números!")
            return False
        
        return True

    def clear_fields(self):
        """Limpa todos os campos do formulário."""
        self.name_entry.delete(0, tk.END)
        self.renach_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)
        self.observation_text.delete('1.0', tk.END)

    def submit_patient(self):
        """Processa o envio do formulário de paciente."""
        if not self.validate_fields():
            return

        name = self.name_entry.get().strip().upper()
        renach = self.renach_entry.get().strip()
        phone = self.format_phone(self.phone_entry.get().strip())
        appointment_date = self.appointment_entry.get_date().strftime("%Y-%m-%d")
        observation = self.observation_text.get('1.0', tk.END).strip()

        # Verifica se o RENACH já existe
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            cursor.execute("SELECT id, appointment_date FROM patients WHERE renach = ?", (renach,))
            existing_patient = cursor.fetchone()

            if existing_patient:
                # Pergunta se deseja atualizar a data do paciente existente
                if messagebox.askyesno("Paciente Existente", 
                                     "Este RENACH já está cadastrado. Deseja atualizar a data da consulta?"):
                    cursor.execute("""
                        UPDATE patients 
                        SET appointment_date = ?, observation = ?
                        WHERE renach = ?
                    """, (appointment_date, observation, renach))
                    messagebox.showinfo("Sucesso", "Data da consulta atualizada com sucesso!")
                else:
                    return
            else:
                # Adiciona novo paciente
                cursor.execute("""
                    INSERT INTO patients (name, renach, phone, appointment_date, observation)
                    VALUES (?, ?, ?, ?, ?)
                """, (name, renach, phone, appointment_date, observation))
                messagebox.showinfo("Sucesso", "Paciente cadastrado com sucesso!")

            conn.commit()
            self.clear_fields()
            
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao processar operação: {str(e)}")
        finally:
            conn.close()

    def add_user(self):
        """Cria a interface para adicionar/atualizar paciente."""
        self.window = tk.Toplevel(self.master)
        self.window.title("Gerenciar Paciente")
        self.window.geometry("400x500")
        self.window.minsize(width=400, height=500)
        self.window.maxsize(width=400, height=500)

        # Configuração visual
        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"
        self.window.configure(bg=cor_fundo)

        # Frame principal
        main_frame = tk.Frame(self.window, bg=cor_fundo)
        main_frame.pack(expand=True, fill='both', padx=20, pady=10)

        # Título
        tk.Label(
            main_frame,
            text="Cadastro de Paciente",
            font=("Arial", 14, "bold"),
            bg=cor_fundo,
            fg=cor_texto
        ).pack(pady=(0, 15))

        # Campos de entrada
        campos = [
            ("Nome:", "name_entry"),
            ("Renach:", "renach_entry"),
            ("Telefone:", "phone_entry")
        ]

        for label_text, entry_name in campos:
            frame = tk.Frame(main_frame, bg=cor_fundo)
            frame.pack(fill='x', pady=5)
            
            tk.Label(
                frame,
                text=label_text,
                bg=cor_fundo,
                fg=cor_texto,
                width=10,
                anchor='w'
            ).pack(side='left')
            
            entry = tk.Entry(frame)
            entry.pack(side='left', expand=True, fill='x', padx=(0, 10))
            setattr(self, entry_name, entry)

        # Campo de data
        date_frame = tk.Frame(main_frame, bg=cor_fundo)
        date_frame.pack(fill='x', pady=5)
        
        tk.Label(
            date_frame,
            text="Data:",
            bg=cor_fundo,
            fg=cor_texto,
            width=10,
            anchor='w'
        ).pack(side='left')

        self.appointment_entry = DateEntry(
            date_frame,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2
        )
        self.appointment_entry.pack(side='left')

        # Campo de observação
        tk.Label(
            main_frame,
            text="Observações:",
            bg=cor_fundo,
            fg=cor_texto,
            anchor='w'
        ).pack(fill='x', pady=(10, 5))

        self.observation_text = tk.Text(
            main_frame,
            height=4,
            wrap=tk.WORD,
            font=("Arial", 10)
        )
        self.observation_text.pack(fill='x', pady=(0, 10))

        # Botões
        button_frame = tk.Frame(main_frame, bg=cor_fundo)
        button_frame.pack(fill='x', pady=10)

        tk.Button(
            button_frame,
            text="Salvar",
            command=self.submit_patient,
            width=15
        ).pack(side='left', padx=5)

        tk.Button(
            button_frame,
            text="Limpar",
            command=self.clear_fields,
            width=15
        ).pack(side='left', padx=5)

        tk.Button(
            button_frame,
            text="Fechar",
            command=self.window.destroy,
            width=15
        ).pack(side='left', padx=5)

        self.funcoes_botoes.center(self.window)

    def get_patients_by_date(self, date):
        """Obtém os pacientes para uma data específica."""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT name, phone, renach, observation 
            FROM patients 
            WHERE appointment_date = ?
            ORDER BY name
        """, (date,))
        patients = cursor.fetchall()
        conn.close()
        return patients

    def update_patient_list(self, event=None):
        """Atualiza a lista de pacientes na interface."""
        selected_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        patients = self.get_patients_by_date(selected_date)
        
        # Limpa a tabela anterior
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        # Adiciona cabeçalho
        headers = ["Nome", "Telefone", "RENACH", "Observações"]
        for j, header in enumerate(headers):
            header_cell = tk.Label(
                self.results_frame,
                text=header,
                font=("Arial", 12, "bold"),
                bg=self.master.cget('bg'),
                fg='#ECF0F1',
                width=15,
                anchor="w",
                borderwidth=1,
                relief="solid"
            )
            header_cell.grid(row=0, column=j, padx=5, pady=2, sticky="ew")
        
        # Popula a tabela com os dados dos pacientes
        for i, paciente in enumerate(patients, start=1):
            for j, info in enumerate(paciente):
                # Ajusta a largura da coluna de observações
                width = 30 if j == 3 else 15
                cell = tk.Label(
                    self.results_frame,
                    text=info if info else "",
                    font=("Arial", 11),
                    bg=self.master.cget('bg'),
                    fg='#ECF0F1',
                    width=width,
                    anchor="w",
                    borderwidth=1,
                    relief="solid",
                    wraplength=300 if j == 3 else None  # Quebra de linha para observações
                )
                cell.grid(row=i, column=j, padx=5, pady=2, sticky="ew")

        # Configura o redimensionamento das colunas
        self.results_frame.grid_columnconfigure(3, weight=1)

    def view_marcacoes(self):
        """Cria a interface para visualização e gestão das marcações."""
        # Configuração da janela principal
        self.marcacoes_window = tk.Toplevel(self.master)
        self.marcacoes_window.title("Gerenciador de Marcações")
        self.marcacoes_window.geometry("1000x700")
        cor_fundo = self.master.cget('bg')
        cor_texto = '#ECF0F1'
        self.marcacoes_window.configure(bg=cor_fundo)

        # Frame principal
        main_frame = tk.Frame(self.marcacoes_window, bg=cor_fundo)
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # Título
        title_frame = tk.Frame(main_frame, bg=cor_fundo)
        title_frame.pack(fill='x', pady=(0, 20))

        tk.Label(
            title_frame,
            text="Gerenciador de Marcações",
            font=("Arial", 18, "bold"),
            bg=cor_fundo,
            fg=cor_texto
        ).pack(side='left')

        # Frame de controles
        control_frame = tk.Frame(main_frame, bg=cor_fundo)
        control_frame.pack(fill='x', pady=(0, 20))

        # Frame para seleção de data
        date_frame = tk.Frame(control_frame, bg=cor_fundo)
        date_frame.pack(side='left')

        tk.Label(
            date_frame,
            text="Data:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 12, "bold")
        ).pack(side='left', padx=(0, 10))

        self.date_entry = DateEntry(
            date_frame,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
            font=("Arial", 10)
        )
        self.date_entry.pack(side='left')

        # Frame para busca
        search_frame = tk.Frame(control_frame, bg=cor_fundo)
        search_frame.pack(side='right')

        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_marcacoes)

        tk.Label(
            search_frame,
            text="Buscar:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 12, "bold")
        ).pack(side='left', padx=(0, 10))

        tk.Entry(
            search_frame,
            textvariable=self.search_var,
            width=30,
            font=("Arial", 10)
        ).pack(side='left')

        # Frame para a tabela
        table_container = tk.Frame(main_frame)
        table_container.pack(fill='both', expand=True)

        # Canvas e scrollbars
        canvas = tk.Canvas(table_container, bg=cor_fundo)
        scrollbar_y = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        scrollbar_x = ttk.Scrollbar(table_container, orient="horizontal", command=canvas.xview)

        self.table_frame = tk.Frame(canvas, bg=cor_fundo)
        canvas.create_window((0, 0), window=self.table_frame, anchor="nw")

        # Configuração do canvas
        canvas.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        # Cabeçalho da tabela
        headers = ["Nome", "RENACH", "Telefone", "Observações", "Ações"]
        for col, header in enumerate(headers):
            tk.Label(
                self.table_frame,
                text=header,
                font=("Arial", 12, "bold"),
                bg="#2C3E50",
                fg=cor_texto,
                padx=10,
                pady=5,
                relief="raised",
                width=25 if col in [0, 3] else 15
            ).grid(row=0, column=col, sticky="nsew", padx=1, pady=1)

        # Configurar as colunas para expandir corretamente
        for i in range(len(headers)):
            self.table_frame.grid_columnconfigure(i, weight=1)

        # Empacotamento dos componentes de rolagem
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        # Frame de estatísticas
        stats_frame = tk.Frame(main_frame, bg=cor_fundo)
        stats_frame.pack(fill='x', pady=(20, 0))

        self.stats_label = tk.Label(
            stats_frame,
            text="",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 10)
        )
        self.stats_label.pack(side='left')

        # Configuração de eventos
        self.date_entry.bind("<<DateEntrySelected>>", self.update_patient_list)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        self.table_frame.bind("<Configure>", _on_frame_configure)

        # Centraliza a janela
        self.funcoes_botoes.center(self.marcacoes_window)
        
        # Atualiza a lista inicial
        self.update_patient_list()

    def filter_marcacoes(self, *args):
        """Filtra as marcações com base no termo de busca."""
        self.update_patient_list()

    def update_patient_list(self, event=None):
        """Atualiza a lista de pacientes com base na data selecionada e filtro de busca."""
        selected_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        search_term = self.search_var.get().lower()
        
        # Limpa a tabela atual (mantém o cabeçalho)
        for widget in self.table_frame.grid_slaves():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()

        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()
            
            # Consulta SQL com filtro de busca
            cursor.execute("""
                SELECT name, renach, phone, observation 
                FROM patients 
                WHERE appointment_date = ? 
                AND (LOWER(name) LIKE ? OR LOWER(renach) LIKE ? OR LOWER(phone) LIKE ?)
                ORDER BY name
            """, (selected_date, f"%{search_term}%", f"%{search_term}%", f"%{search_term}%"))
            
            patients = cursor.fetchall()
            
            if not patients:
                no_results = tk.Label(
                    self.table_frame,
                    text="Nenhuma marcação encontrada para esta data.",
                    font=("Arial", 11),
                    bg=self.master.cget('bg'),
                    fg='#ECF0F1',
                    pady=20
                )
                no_results.grid(row=1, column=0, columnspan=5)
            else:
                for i, patient in enumerate(patients, start=1):
                    row = i
                    
                    # Dados do paciente
                    for col, value in enumerate(patient):
                        tk.Label(
                            self.table_frame,
                            text=str(value) if value else "",
                            font=("Arial", 10),
                            bg=self.master.cget('bg'),
                            fg='#ECF0F1',
                            padx=5,
                            pady=3,
                            wraplength=300 if col == 3 else None  # Quebra de linha para observações
                        ).grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
                    
                    # Botões de ação
                    action_frame = tk.Frame(self.table_frame, bg=self.master.cget('bg'))
                    action_frame.grid(row=row, column=4, padx=5, pady=1)
                    
                    def create_edit_callback(p):
                        return lambda: self.edit_marcacao(p)
                    
                    def create_delete_callback(p):
                        return lambda: self.delete_marcacao(p)
                    
                    tk.Button(
                        action_frame,
                        text="Editar",
                        command=create_edit_callback(patient),
                        width=8
                    ).pack(side='left', padx=2)
                    
                    tk.Button(
                        action_frame,
                        text="Excluir",
                        command=create_delete_callback(patient),
                        width=8
                    ).pack(side='left', padx=2)

            # Atualiza estatísticas
            cursor.execute("""
                SELECT COUNT(*) from patients WHERE appointment_date = ?
            """, (selected_date,))
            total_pacientes = cursor.fetchone()[0]
            
            self.stats_label.config(
                text=f"Total de marcações para {self.date_entry.get()}: {total_pacientes}"
            )

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao buscar marcações: {str(e)}")
        finally:
            conn.close()

    def edit_marcacao(self, patient):
        """Abre janela para edição de marcação."""
        edit_window = tk.Toplevel(self.marcacoes_window)
        edit_window.title("Editar Marcação")
        edit_window.geometry("400x500")
        edit_window.configure(bg=self.master.cget('bg'))
        
        # Configurações de cores
        cor_fundo = self.master.cget('bg')
        cor_texto = '#ECF0F1'
        
        # Frame principal
        main_frame = tk.Frame(edit_window, bg=cor_fundo)
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Título
        tk.Label(
            main_frame,
            text="Editar Dados da Marcação",
            font=("Arial", 14, "bold"),
            bg=cor_fundo,
            fg=cor_texto
        ).pack(pady=(0, 20))
        
        # Frame para os campos
        fields_frame = tk.Frame(main_frame, bg=cor_fundo)
        fields_frame.pack(fill='x', pady=10)
        
        # Função para criar campos de entrada
        def create_field(parent, label_text, default_value=""):
            frame = tk.Frame(parent, bg=cor_fundo)
            frame.pack(fill='x', pady=5)
            
            tk.Label(
                frame,
                text=label_text,
                font=("Arial", 10, "bold"),
                bg=cor_fundo,
                fg=cor_texto,
                width=12,
                anchor='w'
            ).pack(side='left')
            
            entry = tk.Entry(frame, font=("Arial", 10))
            entry.pack(side='left', fill='x', expand=True)
            entry.insert(0, default_value)
            return entry
        
        # Criação dos campos
        nome_entry = create_field(fields_frame, "Nome:", patient[0])
        renach_entry = create_field(fields_frame, "RENACH:", patient[1])
        telefone_entry = create_field(fields_frame, "Telefone:", patient[2])
        
        # Campo de data
        date_frame = tk.Frame(fields_frame, bg=cor_fundo)
        date_frame.pack(fill='x', pady=5)
        
        tk.Label(
            date_frame,
            text="Data:",
            font=("Arial", 10, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
            width=12,
            anchor='w'
        ).pack(side='left')
        
        # Buscar a data atual do paciente
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("SELECT appointment_date FROM patients WHERE renach = ?", (patient[1],))
        current_date = cursor.fetchone()[0]
        conn.close()
        
        date_entry = DateEntry(
            date_frame,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
            font=("Arial", 10)
        )
        date_entry.pack(side='left')
        
        # Definir a data atual do paciente
        try:
            current_date_obj = datetime.strptime(current_date, "%Y-%m-%d").date()
            date_entry.set_date(current_date_obj)
        except:
            pass
        
        # Campo de observações
        obs_frame = tk.Frame(fields_frame, bg=cor_fundo)
        obs_frame.pack(fill='x', pady=5)
        
        tk.Label(
            obs_frame,
            text="Observações:",
            font=("Arial", 10, "bold"),
            bg=cor_fundo,
            fg=cor_texto
        ).pack(anchor='w')
        
        obs_text = tk.Text(
            obs_frame,
            height=4,
            font=("Arial", 10),
            wrap=tk.WORD
        )
        obs_text.pack(fill='x', pady=(5, 0))
        obs_text.insert('1.0', patient[3] if patient[3] else "")
        
        # Frame para botões
        button_frame = tk.Frame(main_frame, bg=cor_fundo)
        button_frame.pack(pady=20)
        
        def save_changes():
            """Salva as alterações no banco de dados."""
            nome = nome_entry.get().strip()
            renach = renach_entry.get().strip()
            telefone = telefone_entry.get().strip()
            nova_data = date_entry.get_date().strftime("%Y-%m-%d")
            observacoes = obs_text.get('1.0', tk.END).strip()
            
            # Validações
            if not nome or not renach or not telefone:
                messagebox.showerror("Erro", "Por favor, preencha todos os campos obrigatórios!")
                return
                
            try:
                conn = sqlite3.connect(self.db_name)
                cursor = conn.cursor()
                
                # Atualiza os dados do paciente
                cursor.execute("""
                    UPDATE patients 
                    SET name = ?, phone = ?, appointment_date = ?, observation = ?
                    WHERE renach = ?
                """, (nome, telefone, nova_data, observacoes, patient[1]))
                
                conn.commit()
                messagebox.showinfo("Sucesso", "Dados atualizados com sucesso!")
                edit_window.destroy()
                self.update_patient_list()  # Atualiza a lista de pacientes
                
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao atualizar dados: {str(e)}")
            finally:
                conn.close()
        
        def cancel_edit():
            """Cancela a edição."""
            if messagebox.askyesno("Confirmar", "Deseja realmente cancelar a edição? As alterações serão perdidas."):
                edit_window.destroy()
        
        # Botões
        tk.Button(
            button_frame,
            text="Salvar",
            command=save_changes,
            width=15,
            font=("Arial", 10),
            bg="#2ecc71",
            fg="white",
            activebackground="#27ae60"
        ).pack(side='left', padx=5)
        
        tk.Button(
            button_frame,
            text="Cancelar",
            command=cancel_edit,
            width=15,
            font=("Arial", 10),
            bg="#e74c3c",
            fg="white",
            activebackground="#c0392b"
        ).pack(side='left', padx=5)
        
        # Centraliza a janela
        self.funcoes_botoes.center(edit_window)
        
        # Torna a janela modal
        edit_window.transient(self.marcacoes_window)
        edit_window.grab_set()
        
        # Prevent window resize
        edit_window.resizable(False, False)

    def delete_marcacao(self, patient):
        """Remove uma marcação após confirmação."""
        if messagebox.askyesno("Confirmar Exclusão", 
                            f"Deseja realmente excluir a marcação de {patient[0]}?"):
            try:
                conn = sqlite3.connect(self.db_name)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM patients WHERE renach = ?", (patient[1],))
                conn.commit()
                messagebox.showinfo("Sucesso", "Marcação excluída com sucesso!")
                self.update_patient_list()
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao excluir marcação: {str(e)}")
            finally:
                conn.close()

