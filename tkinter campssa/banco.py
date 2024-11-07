import sqlite3
import tkinter as tk
import tkinter as ttk
from tkcalendar import DateEntry
from datetime import datetime
from tkinter import messagebox, Toplevel, Frame, Label
from tkcalendar import DateEntry
from funcoes_botoes import FuncoesBotoes
from planilhas import Planilhas
import json
from datetime import datetime


class DataBaseLogin:
    """Função CRUD"""

    def __init__(self, db_name="login.db"):
        self.db_name = db_name
        self.conn = sqlite3.connect(db_name)
        self.create_db()

    """Função para criar o banco de dados e a tabela de usuários"""

    def create_db(self):
        """Função para criar o banco de dados e a tabela de usuários"""
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

    """Função para criar novo usuário"""

    def create_user(self, user, password):
        """Função para criar novo usuário"""
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

    """Função para ser um usuário com base no user"""

    def read_user(self, user):
        """Função para ser um usuário com base no user"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE user = ?", (user,))
        usuario = cursor.fetchone()
        conn.close()
        return usuario

    """Função para atualizar a senha de um usuário"""

    def update_user(self, user, new_password):
        """Função para atualizar a senha de um usuário"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute(
            "UPDATE users SET password = ? WHERE user =?", (new_password, user)
        )

    """Função para deletar um usuário com base no user"""

    def delete_user(self, user):
        """Função para deletar um usuário com base no user"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM users WHERE user = ?", (user,))
        conn.comit()
        conn.close()

    """Verifica se o nome de usuário e a senha são válidos no banco de dados."""

    def validate_user(self, user, password):
        """Verifica se o nome de usuário e a senha são válidos no banco de dados."""
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT * FROM users WHERE user = ? AND password = ?", (user, password)
        )
        result = cursor.fetchone()
        return result is not None


class DataBaseMarcacao:
    def __init__(
        self,
        master,
        planilhas: Planilhas,
        file_path: str,
        app,
        db_name="db_marcacao.db",
    ):
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
        self.search_window = None
        self.search_var = None
        self.table_frame = None

    """Cria o banco de dados e atualiza a estrutura se necessário."""

    def create_db(self):
        """Cria o banco de dados e atualiza a estrutura se necessário."""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        try:
            # Primeiro cria a tabela se não existir
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS patients (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    renach TEXT NOT NULL,
                    phone TEXT,
                    appointment_date TEXT NOT NULL,
                    observation TEXT
                )
            """
            )

            # Verifica se as colunas novas existem
            cursor.execute("PRAGMA table_info(patients)")
            columns = [column[1] for column in cursor.fetchall()]

            # Adiciona a coluna attendance_status se não existir
            if "attendance_status" not in columns:
                cursor.execute(
                    """
                    ALTER TABLE patients 
                    ADD COLUMN attendance_status TEXT DEFAULT 'pending'
                """
                )

            # Adiciona a coluna attendance_history se não existir
            if "attendance_history" not in columns:
                cursor.execute(
                    """
                    ALTER TABLE patients 
                    ADD COLUMN attendance_history TEXT DEFAULT '[]'
                """
                )

            conn.commit()

        except sqlite3.Error as e:
            print(f"Erro ao criar/atualizar banco de dados: {e}")
        finally:
            conn.close()

    """Formata o número de telefone no padrão (XX) XXXXX-XXXX ou (XX) XXXX-XXXX."""

    def format_phone(self, phone):
        """Formata o número de telefone no padrão (XX) XXXXX-XXXX ou (XX) XXXX-XXXX."""
        phone = "".join(filter(str.isdigit, phone))
        if len(phone) == 11:
            return f"({phone[:2]}) {phone[2:7]}-{phone[7:]}"
        elif len(phone) == 10:
            return f"({phone[:2]}) {phone[2:6]}-{phone[6:]}"
        return phone

    """Valida os campos obrigatórios do formulário."""

    def validate_fields(self):
        """Valida os campos obrigatórios do formulário."""
        name = self.name_entry.get().strip()
        renach = self.renach_entry.get().strip()

        if not all([name, renach]):
            messagebox.showerror(
                "Erro", "Por favor, preencha todos os campos obrigatórios!"
            )
            return False

        if not renach.isdigit():
            messagebox.showerror("Erro", "O RENACH deve conter apenas números!")
            return False

        return True

    """Limpa todos os campos do formulário."""

    def clear_fields(self):
        """Limpa todos os campos do formulário."""
        self.name_entry.delete(0, tk.END)
        self.renach_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)
        self.observation_text.delete("1.0", tk.END)

    """Processa o envio do formulário de paciente."""

    def submit_patient(self):
        """Processa o envio do formulário de paciente."""
        if not self.validate_fields():
            return

        name = self.name_entry.get().strip().upper()
        renach = self.renach_entry.get().strip()
        phone = self.format_phone(self.phone_entry.get().strip())
        appointment_date = self.appointment_entry.get_date().strftime("%Y-%m-%d")
        observation = self.observation_text.get("1.0", tk.END).strip()

        # Verifica se o RENACH já existe
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        try:
            cursor.execute(
                "SELECT id, appointment_date FROM patients WHERE renach = ?", (renach,)
            )
            existing_patient = cursor.fetchone()

            if existing_patient:
                # Pergunta se deseja atualizar a data do paciente existente
                if messagebox.askyesno(
                    "Paciente Existente",
                    "Este RENACH já está cadastrado. Deseja atualizar a data da consulta?",
                ):
                    cursor.execute(
                        """
                        UPDATE patients 
                        SET appointment_date = ?, observation = ?
                        WHERE renach = ?
                    """,
                        (appointment_date, observation, renach),
                    )
                    messagebox.showinfo(
                        "Sucesso", "Data da consulta atualizada com sucesso!"
                    )
                else:
                    return
            else:
                # Adiciona novo paciente
                cursor.execute(
                    """
                    INSERT INTO patients (name, renach, phone, appointment_date, observation)
                    VALUES (?, ?, ?, ?, ?)
                """,
                    (name, renach, phone, appointment_date, observation),
                )
                messagebox.showinfo("Sucesso", "Paciente cadastrado com sucesso!")

            conn.commit()
            self.clear_fields()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao processar operação: {str(e)}")
        finally:
            conn.close()

    """Cria a interface para adicionar/atualizar paciente."""

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
        main_frame.pack(expand=True, fill="both", padx=20, pady=10)

        # Título
        tk.Label(
            main_frame,
            text="Cadastro de Paciente",
            font=("Arial", 14, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
        ).pack(pady=(0, 15))

        # Campos de entrada
        campos = [
            ("Nome:", "name_entry"),
            ("Renach:", "renach_entry"),
            ("Telefone:", "phone_entry"),
        ]

        for label_text, entry_name in campos:
            frame = tk.Frame(main_frame, bg=cor_fundo)
            frame.pack(fill="x", pady=5)

            tk.Label(
                frame, text=label_text, bg=cor_fundo, fg=cor_texto, width=10, anchor="w"
            ).pack(side="left")

            entry = tk.Entry(frame)
            entry.pack(side="left", expand=True, fill="x", padx=(0, 10))
            setattr(self, entry_name, entry)

        # Campo de data
        date_frame = tk.Frame(main_frame, bg=cor_fundo)
        date_frame.pack(fill="x", pady=5)

        tk.Label(
            date_frame, text="Data:", bg=cor_fundo, fg=cor_texto, width=10, anchor="w"
        ).pack(side="left")

        self.appointment_entry = DateEntry(
            date_frame,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
        )
        self.appointment_entry.pack(side="left")

        # Campo de observação
        tk.Label(
            main_frame, text="Observações:", bg=cor_fundo, fg=cor_texto, anchor="w"
        ).pack(fill="x", pady=(10, 5))

        self.observation_text = tk.Text(
            main_frame, height=4, wrap=tk.WORD, font=("Arial", 10)
        )
        self.observation_text.pack(fill="x", pady=(0, 10))

        # Botões
        button_frame = tk.Frame(main_frame, bg=cor_fundo)
        button_frame.pack(fill="x", pady=10)

        tk.Button(
            button_frame, text="Salvar", command=self.submit_patient, width=15
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame, text="Limpar", command=self.clear_fields, width=15
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame, text="Fechar", command=self.window.destroy, width=15
        ).pack(side="left", padx=5)

        self.funcoes_botoes.center(self.window)

    """Cria a interface para visualização e gestão das marcações."""

    def view_marcacoes(self):
        """Cria a interface para visualização e gestão das marcações."""
        # Configuração da janela principal
        self.marcacoes_window = tk.Toplevel(self.master)
        self.marcacoes_window.title("Gerenciador de Marcações")
        self.marcacoes_window.geometry("1000x700")
        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"
        self.marcacoes_window.configure(bg=cor_fundo)

        # Frame principal
        main_frame = tk.Frame(self.marcacoes_window, bg=cor_fundo)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Título
        title_frame = tk.Frame(main_frame, bg=cor_fundo)
        title_frame.pack(fill="x", pady=(0, 20))

        tk.Label(
            title_frame,
            text="Gerenciador de Marcações",
            font=("Arial", 18, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
        ).pack(side="left")

        # Frame de controles
        control_frame = tk.Frame(main_frame, bg=cor_fundo)
        control_frame.pack(fill="x", pady=(0, 20))

        # Frame para seleção de data
        date_frame = tk.Frame(control_frame, bg=cor_fundo)
        date_frame.pack(side="left")

        tk.Label(
            date_frame,
            text="Data:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 12, "bold"),
        ).pack(side="left", padx=(0, 10))

        self.date_entry = DateEntry(
            date_frame,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
            font=("Arial", 10),
        )
        self.date_entry.set_date(datetime.now().date())
        self.date_entry.pack(side="left")

        # Frame para busca
        search_frame = tk.Frame(control_frame, bg=cor_fundo)
        search_frame.pack(side="right")

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_marcacoes)

        tk.Label(
            search_frame,
            text="Buscar:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 12, "bold"),
        ).pack(side="left", padx=(0, 10))

        tk.Entry(
            search_frame, textvariable=self.search_var, width=30, font=("Arial", 10)
        ).pack(side="left")

        # Botão para ver histórico
        tk.Button(
            control_frame,
            text="Ver Histórico",
            command=self.view_patient_history,
            bg="#3498db",
            fg="white",
            font=("Arial", 10),
        ).pack(side="right", padx=10)

        # Frame para resultados com scroll
        table_container = tk.Frame(main_frame, bg=cor_fundo)
        table_container.pack(fill="both", expand=True)

        # Canvas e scrollbars
        canvas = tk.Canvas(table_container, bg=cor_fundo)
        scrollbar_y = ttk.Scrollbar(
            table_container, orient="vertical", command=canvas.yview
        )
        scrollbar_x = ttk.Scrollbar(
            table_container, orient="horizontal", command=canvas.xview
        )

        self.results_frame = tk.Frame(canvas, bg=cor_fundo)
        canvas.create_window((0, 0), window=self.results_frame, anchor="nw")
        self.update_patient_list()

        # Configuração do canvas
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Empacotamento dos componentes de rolagem
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        # Atualiza a lista inicial
        self.update_patient_list()

        # Configuração de eventos
        self.date_entry.bind("<<DateEntrySelected>>", self.update_patient_list)
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        self.results_frame.bind("<Configure>", _on_frame_configure)

        # Centraliza a janela
        self.funcoes_botoes.center(self.marcacoes_window)

    """Filtra as marcações com base no termo de busca."""

    def filter_marcacoes(self, *args):
        """Filtra as marcações com base no termo de busca."""
        search_term = self.search_var.get().strip()
        if search_term:
            self.update_patient_list()
        else:
            self.update_patient_list(None)

    """Obtém os pacientes por nome ou renach, independentemente da data."""

    def get_patients_by_name_or_renach(self, search_term, selected_date=None):
        """Obtém os pacientes por nome ou renach, filtrando pela data selecionada, se fornecida."""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        search_term = search_term.lower() if search_term else ""

        try:
            query = """
                SELECT name, phone, renach, 
                    COALESCE(attendance_status, 'pending') as attendance_status, 
                    observation, appointment_date
                FROM patients 
                WHERE (LOWER(name) LIKE ? OR LOWER(renach) LIKE ?)
            """
            params = [f"%{search_term}%", f"%{search_term}%"]

            if selected_date and not search_term:
                query += " AND appointment_date = ?"
                params.append(selected_date)

            query += " ORDER BY name"

            cursor.execute(query, params)

            patients = cursor.fetchall()
            return patients
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao buscar pacientes: {str(e)}")
            return []
        finally:
            conn.close()

    """Atualiza a lista de pacientes com status de comparecimento."""

    def update_patient_list(self, event=None):
        """Atualiza a lista de pacientes com status de comparecimento."""
        search_term = self.search_var.get().strip()
        selected_date = self.date_entry.get_date().strftime("%Y-%m-%d")
        patients = self.get_patients_by_name_or_renach(search_term, selected_date)

        # Limpa a tabela anterior
        for widget in self.results_frame.winfo_children():
            widget.destroy()

        # Adiciona cabeçalho atualizado
        headers = [
            "Nome",
            "Telefone",
            "RENACH",
            "Status",
            "Observações",
            "Data",
            "Ações",
        ]
        for j, header in enumerate(headers):
            header_cell = tk.Label(
                self.results_frame,
                text=header,
                font=("Arial", 12, "bold"),
                bg=self.master.cget("bg"),
                fg="#ECF0F1",
                width=20,
                anchor="w",
            )
            header_cell.grid(row=0, column=j, padx=5, pady=2, sticky="ew")

        if not patients:
            tk.Label(
                self.results_frame,
                text="Nenhum paciente encontrado.",
                font=("Arial", 11),
                bg=self.master.cget("bg"),
                fg="#ECF0F1",
                pady=20,
            ).grid(row=1, column=0, columnspan=len(headers))
            return

        # Popula a tabela com os dados atualizados
        for i, patient in enumerate(patients, start=1):
            for j, info in enumerate(patient):
                # Ajusta o texto do status
                if j == 3:  # Coluna de status
                    status_text = {
                        "attended": "Compareceu",
                        "missed": "Não Compareceu",
                        "pending": "Pendente",
                    }.get(info, "Pendente")

                    # Cores diferentes para cada status
                    status_colors = {
                        "attended": "#2ecc71",
                        "missed": "#e74c3c",
                        "pending": "#f1c40f",
                    }
                    bg_color = status_colors.get(info, "#f1c40f")
                elif j == 5:  # Coluna de data
                    date_text = info if info else ""
                    if date_text:
                        date_text = datetime.strptime(date_text, "%Y-%m-%d").strftime(
                            "%d/%m/%Y"
                        )
                    bg_color = self.master.cget("bg")
                else:
                    date_text = info if info else ""
                    bg_color = self.master.cget("bg")

                cell = tk.Label(
                    self.results_frame,
                    text=date_text if j == 5 else status_text if j == 3 else info,
                    font=("Arial", 11),
                    bg=bg_color,
                    fg="#ECF0F1" if j != 3 else "#000000",
                    width=25 if j in [0, 4] else 15,
                    anchor="w",
                    wraplength=300 if j == 4 else None,
                )
                cell.grid(row=i, column=j, padx=5, pady=2, sticky="ew")

            # Frame para botões de ação
            action_frame = tk.Frame(self.results_frame, bg=self.master.cget("bg"))
            action_frame.grid(row=i, column=len(headers) - 1, padx=5, pady=2)

            # Botões de ação
            def create_edit_callback(p):
                return lambda: self.edit_marcacao(p)

            def create_delete_callback(p):
                return lambda: self.delete_marcacao(p)

            tk.Button(
                action_frame,
                text="Editar",
                command=create_edit_callback(patient),
                width=8,
            ).pack(side="left", padx=2)

            tk.Button(
                action_frame,
                text="Excluir",
                command=create_delete_callback(patient),
                width=8,
            ).pack(side="left", padx=2)

            # Botões de status
            tk.Button(
                action_frame,
                text="✓",
                command=lambda p=patient: self.update_attendance_status(
                    p[2], "attended"
                ),
                bg="#2ecc71",
                fg="white",
                width=3,
            ).pack(side="left", padx=2)

            tk.Button(
                action_frame,
                text="✗",
                command=lambda p=patient: self.update_attendance_status(p[2], "missed"),
                bg="#e74c3c",
                fg="white",
                width=3,
            ).pack(side="left", padx=2)

            tk.Button(
                action_frame,
                text="⟲",
                command=lambda p=patient: self.update_attendance_status(
                    p[2], "pending"
                ),
                bg="#f1c40f",
                fg="white",
                width=3,
            ).pack(side="left", padx=2)

        # Estatísticas
        stats_frame = tk.Frame(self.results_frame, bg=self.master.cget("bg"))
        stats_frame.grid(
            row=len(patients) + 1, column=0, columnspan=len(headers), pady=10
        )

        # Contagem de status
        attended_count = sum(1 for p in patients if p[3] == "attended")
        missed_count = sum(1 for p in patients if p[3] == "missed")
        pending_count = sum(1 for p in patients if p[3] == "pending")

        stats_text = f"Total: {len(patients)} | Compareceram: {attended_count} | Não Compareceram: {missed_count} | Pendentes: {pending_count}"
        tk.Label(
            stats_frame,
            text=stats_text,
            font=("Arial", 10),
            bg=self.master.cget("bg"),
            fg="#ECF0F1",
        ).pack()
        """Abre janela para edição de marcação."""

    """Abre janela para edição de marcação."""

    def edit_marcacao(self, patient):
        """Abre janela para edição de marcação."""
        edit_window = tk.Toplevel(self.marcacoes_window)
        edit_window.title("Editar Marcação")
        edit_window.geometry("400x500")
        edit_window.configure(bg=self.master.cget("bg"))

        # Configurações de cores
        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"

        # Frame principal
        main_frame = tk.Frame(edit_window, bg=cor_fundo)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Título
        tk.Label(
            main_frame,
            text="Editar Dados da Marcação",
            font=("Arial", 14, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
        ).pack(pady=(0, 20))

        # Frame para os campos
        fields_frame = tk.Frame(main_frame, bg=cor_fundo)
        fields_frame.pack(fill="x", pady=10)

        # Função para criar campos de entrada
        def create_field(parent, label_text, default_value=""):
            frame = tk.Frame(parent, bg=cor_fundo)
            frame.pack(fill="x", pady=5)

            tk.Label(
                frame,
                text=label_text,
                font=("Arial", 10, "bold"),
                bg=cor_fundo,
                fg=cor_texto,
                width=12,
                anchor="w",
            ).pack(side="left")

            entry = tk.Entry(frame, font=("Arial", 10))
            entry.pack(side="left", fill="x", expand=True)
            entry.insert(0, default_value)
            return entry

        # Criação dos campos
        nome_entry = create_field(fields_frame, "Nome:", patient[0])
        renach_entry = create_field(fields_frame, "RENACH:", patient[1])
        telefone_entry = create_field(fields_frame, "Telefone:", patient[2])

        # Campo de data
        date_frame = tk.Frame(fields_frame, bg=cor_fundo)
        date_frame.pack(fill="x", pady=5)

        tk.Label(
            date_frame,
            text="Data:",
            font=("Arial", 10, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
            width=12,
            anchor="w",
        ).pack(side="left")

        # Buscar a data atual do paciente
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT appointment_date FROM patients WHERE renach = ?", (patient[2],)
        )
        current_date = cursor.fetchone()[0]
        conn.close()

        date_entry = DateEntry(
            date_frame,
            width=12,
            background="darkblue",
            foreground="white",
            borderwidth=2,
            font=("Arial", 10),
        )
        date_entry.pack(side="left")

        # Definir a data atual do paciente
        try:
            current_date_obj = datetime.strptime(current_date, "%Y-%m-%d").date()
            date_entry.set_date(current_date_obj)
        except:
            pass

        # Campo de observações
        tk.Label(
            fields_frame,
            text="Observações:",
            font=("Arial", 10, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
        ).pack(anchor="w", pady=(10, 5))

        obs_text = tk.Text(fields_frame, height=4, font=("Arial", 10), wrap=tk.WORD)
        obs_text.pack(fill="x")
        obs_text.insert("1.0", patient[4] if patient[4] else "")

        def save_changes():
            """Salva as alterações no banco de dados."""
            nome = nome_entry.get().strip()
            renach = renach_entry.get().strip()
            telefone = telefone_entry.get().strip()
            nova_data = date_entry.get_date().strftime("%Y-%m-%d")
            observacoes = obs_text.get("1.0", tk.END).strip()

            if not nome or not renach:
                messagebox.showerror("Erro", "Nome e RENACH são campos obrigatórios!")
                return

            try:
                conn = sqlite3.connect(self.db_name)
                cursor = conn.cursor()

                cursor.execute(
                    """
                    UPDATE patients 
                    SET name = ?, phone = ?, appointment_date = ?, observation = ?
                    WHERE renach = ?
                """,
                    (nome, telefone, nova_data, observacoes, patient[2]),
                )

                conn.commit()
                messagebox.showinfo("Sucesso", "Dados atualizados com sucesso!")
                edit_window.destroy()
                self.update_patient_list()

            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao atualizar dados: {str(e)}")
            finally:
                conn.close()

        # Frame para botões
        button_frame = tk.Frame(main_frame, bg=cor_fundo)
        button_frame.pack(pady=20)

        tk.Button(
            button_frame,
            text="Salvar",
            command=save_changes,
            width=15,
            bg="#2ecc71",
            fg="white",
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame,
            text="Cancelar",
            command=edit_window.destroy,
            width=15,
            bg="#e74c3c",
            fg="white",
        ).pack(side="left", padx=5)

        # Centraliza a janela
        self.funcoes_botoes.center(edit_window)

    """Remove uma marcação do banco de dados."""

    def delete_marcacao(self, patient):
        """Remove uma marcação do banco de dados."""
        if messagebox.askyesno(
            "Confirmar Exclusão",
            f"Deseja realmente excluir a marcação de {patient[0]}?",
        ):
            try:
                conn = sqlite3.connect(self.db_name)
                cursor = conn.cursor()

                cursor.execute("DELETE FROM patients WHERE renach = ?", (patient[2],))
                conn.commit()

                messagebox.showinfo("Sucesso", "Marcação excluída com sucesso!")
                self.update_patient_list()

            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao excluir marcação: {str(e)}")
            finally:
                conn.close()

    """Interface para visualizar histórico completo de pacientes."""

    def view_patient_history(self):
        """Interface para visualizar histórico completo de pacientes."""
        history_window = tk.Toplevel(self.master)
        history_window.title("Histórico de Pacientes")
        history_window.geometry("900x600")
        history_window.configure(bg=self.master.cget("bg"))

        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"

        # Frame principal
        main_frame = tk.Frame(history_window, bg=cor_fundo)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # Frame de busca
        search_frame = tk.Frame(main_frame, bg=cor_fundo)
        search_frame.pack(fill="x", pady=(0, 20))

        tk.Label(
            search_frame,
            text="Buscar por nome ou RENACH:",
            font=("Arial", 12, "bold"),
            bg=cor_fundo,
            fg=cor_texto,
        ).pack(side="left", padx=(0, 10))

        search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_frame, textvariable=search_var, width=40, font=("Arial", 11)
        )
        search_entry.pack(side="left")

        # Frame para a tabela
        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill="both", expand=True)

        # Treeview com scrollbar
        tree_scroll = ttk.Scrollbar(table_frame)
        tree_scroll.pack(side="right", fill="y")

        style = ttk.Style()
        style.configure(
            "Treeview",
            background=cor_fundo,
            fieldbackground=cor_fundo,
            foreground=cor_texto,
        )

        tree = ttk.Treeview(
            table_frame,
            columns=("Data", "Nome", "RENACH", "Status", "Observação"),
            show="headings",
            yscrollcommand=tree_scroll.set,
        )
        tree.pack(fill="both", expand=True)

        tree_scroll.config(command=tree.yview)

        # Configurar colunas
        tree.heading("Data", text="Data")
        tree.heading("Nome", text="Nome")
        tree.heading("RENACH", text="RENACH")
        tree.heading("Status", text="Status")
        tree.heading("Observação", text="Observação")

        for col in ("Data", "Nome", "RENACH", "Status", "Observação"):
            tree.column(col, width=150)

        def search_history(*args):
            search_term = search_var.get().strip().lower()
            if len(search_term) < 3:
                return

            # Limpa a tabela
            for item in tree.get_children():
                tree.delete(item)

            try:
                conn = sqlite3.connect(self.db_name)
                cursor = conn.cursor()

                # Busca pacientes que correspondam ao termo de busca
                cursor.execute(
                    """
                    SELECT name, renach, appointment_date, attendance_status, observation, attendance_history
                    FROM patients
                    WHERE LOWER(name) LIKE ? OR LOWER(renach) LIKE ?
                    ORDER BY appointment_date DESC
                """,
                    (f"%{search_term}%", f"%{search_term}%"),
                )

                results = cursor.fetchall()

                for row in results:
                    nome, renach, data, status, obs, history = row

                    # Formata o status
                    status_text = {
                        "attended": "Compareceu",
                        "missed": "Não Compareceu",
                        "pending": "Pendente",
                    }.get(status, "Desconhecido")

                    # Insere a marcação atual
                    tree.insert(
                        "",
                        "end",
                        values=(
                            datetime.strptime(data, "%Y-%m-%d").strftime("%d/%m/%Y"),
                            nome,
                            renach,
                            status_text,
                            obs or "",
                        ),
                    )

                    # Insere o histórico se existir
                    if history:
                        try:
                            hist_data = json.loads(history)
                            for entry in hist_data:
                                if isinstance(entry, dict):
                                    tree.insert(
                                        "",
                                        "end",
                                        values=(
                                            datetime.strptime(
                                                entry["date"], "%Y-%m-%d"
                                            ).strftime("%d/%m/%Y"),
                                            nome,
                                            renach,
                                            entry["status"],
                                            f"Histórico: {entry['updated_at']}",
                                        ),
                                    )
                        except json.JSONDecodeError:
                            pass

            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao buscar histórico: {str(e)}")
            finally:
                conn.close()

        search_var.trace("w", search_history)

        # Centraliza a janela
        self.funcoes_botoes.center(history_window)

    """Atualiza o status de comparecimento do paciente."""

    def update_attendance_status(self, renach: str, status: str):
        """Atualiza o status de comparecimento do paciente."""
        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()

            # Recupera informações atuais do paciente
            cursor.execute(
                """
                SELECT attendance_history, appointment_date
                FROM patients 
                WHERE renach = ?
            """,
                (renach,),
            )

            result = cursor.fetchone()
            if not result:
                return

            current_history, appointment_date = result

            # Carrega o histórico existente ou cria novo
            try:
                history = json.loads(current_history)
            except:
                history = []

            # Adiciona nova entrada ao histórico
            history.append(
                {
                    "date": appointment_date,
                    "status": status,
                    "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            )

            # Atualiza o banco de dados
            cursor.execute(
                """
                UPDATE patients 
                SET attendance_status = ?, 
                    attendance_history = ?
                WHERE renach = ?
            """,
                (status, json.dumps(history), renach),
            )

            conn.commit()
            self.update_patient_list()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao atualizar status: {str(e)}")
        finally:
            conn.close()
