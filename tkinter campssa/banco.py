import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, Frame, Label
from tkcalendar import DateEntry
from datetime import datetime
import json
from funcoes_botoes import FuncoesBotoes
from planilhas import Planilhas
from typing import Optional, List, Dict, Any, Tuple
import logging
from database_connection import DatabaseConnection 


# Configuração de logging
logging.basicConfig(
    filename="database_operations.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


class DataBaseLogin:
    """Gerenciamento de autenticação de usuários"""

    def __init__(self, db_name: str = "login.db"):
        self.db_name = db_name
        self.create_db()

    """Cria o banco de dados de usuários se não existir"""

    def create_db(self) -> None:
        """Cria o banco de dados de usuários se não existir"""
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user TEXT NOT NULL UNIQUE,
                        password TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """
                )
                conn.commit()
                logger.info("Tabela de usuários criada/verificada com sucesso")
        except sqlite3.Error as e:
            logger.error(f"Erro ao criar banco de dados de usuários: {e}")
            raise

    """Cria um novo usuário"""

    def create_user(self, user: str, password: str) -> bool:
        """
        Cria um novo usuário
        Returns: True se criado com sucesso, False se já existe
        """
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO users (user, password) VALUES (?, ?)", (user, password)
                )
                conn.commit()
                logger.info(f"Novo usuário criado: {user}")
                return True
        except sqlite3.IntegrityError:
            logger.warning(f"Tentativa de criar usuário duplicado: {user}")
            return False
        except sqlite3.Error as e:
            logger.error(f"Erro ao criar usuário: {e}")
            raise

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

    """Valida credenciais do usuário"""

    def validate_user(self, user: str, password: str) -> bool:
        """Valida credenciais do usuário"""
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT id FROM users WHERE user = ? AND password = ?",
                    (user, password),
                )
                is_valid = cursor.fetchone() is not None
                logger.info(
                    f"Tentativa de login para usuário {user}: {'sucesso' if is_valid else 'falha'}"
                )
                return is_valid
        except sqlite3.Error as e:
            logger.error(f"Erro ao validar usuário: {e}")
            return False


class DataBaseMarcacao:
    """Gerenciamento de marcações de pacientes"""

    def __init__(self, master: tk.Tk, planilhas: Planilhas, file_path: str, app: Any, db_name: str = "db_marcacao.db"):
        # Configuração do logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        # Verifica se já existe um handler para evitar duplicação de logs
        if not self.logger.handlers:
            # Handler para arquivo
            fh = logging.FileHandler('database_operations.log')
            fh.setLevel(logging.INFO)
            
            # Handler para console
            ch = logging.StreamHandler()
            ch.setLevel(logging.INFO)
            
            # Formato do log
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            fh.setFormatter(formatter)
            ch.setFormatter(formatter)
            
            # Adiciona os handlers ao logger
            self.logger.addHandler(fh)
            self.logger.addHandler(ch)

        self.db_name = db_name
        self.master = master
        self.create_db()
        self.funcoes_botoes = FuncoesBotoes(self.master, planilhas, file_path, app)
        
        # UI Components
        self.window: Optional[Toplevel] = None
        self.marcacoes_window: Optional[Toplevel] = None
        self.results_frame: Optional[Frame] = None
        self.date_entry: Optional[DateEntry] = None
        self.search_window: Optional[Toplevel] = None
        self.search_var: Optional[tk.StringVar] = None
        self.table_frame: Optional[Frame] = None
        
        # Form fields
        self.name_entry: Optional[tk.Entry] = None
        self.renach_entry: Optional[tk.Entry] = None
        self.phone_entry: Optional[tk.Entry] = None
        self.appointment_entry: Optional[DateEntry] = None
        self.observation_text: Optional[tk.Text] = None

    """Cria e atualiza a estrutura do banco de dados"""

    def create_db(self) -> None:
        """Cria e atualiza a estrutura do banco de dados"""
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()

                # Verifica se a tabela existe
                cursor.execute(
                    """
                    SELECT name FROM sqlite_master 
                    WHERE type='table' AND name='marcacoes'
                """
                )

                if cursor.fetchone() is None:
                    # Criação da tabela principal com nomes de colunas corrigidos
                    cursor.execute(
                        """
                        CREATE TABLE IF NOT EXISTS marcacoes (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            nome TEXT NOT NULL,
                            telefone TEXT,
                            renach TEXT NOT NULL UNIQUE,
                            data_agendamento TEXT NOT NULL,
                            observacao TEXT,
                            status_comparecimento TEXT DEFAULT 'pending',
                            historico_comparecimento TEXT DEFAULT '[]',
                            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                            atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                        )
                    """
                    )

                    # Trigger para atualizar atualizado_em
                    cursor.execute(
                        """
                        CREATE TRIGGER IF NOT EXISTS update_marcacoes_timestamp 
                        AFTER UPDATE ON marcacoes
                        BEGIN
                            UPDATE marcacoes SET atualizado_em = CURRENT_TIMESTAMP 
                            WHERE id = NEW.id;
                        END;
                    """
                    )

                    conn.commit()
                    logger.info("Tabela de marcações criada com sucesso")
                else:
                    # Verificar e atualizar estrutura da tabela existente
                    cursor.execute("PRAGMA table_info(marcacoes)")
                    columns = {column[1] for column in cursor.fetchall()}

                    # Mapeamento de colunas antigas para novas
                    needed_columns = {
                        "nome": "TEXT NOT NULL",
                        "telefone": "TEXT",
                        "renach": "TEXT NOT NULL UNIQUE",
                        "data_agendamento": "TEXT NOT NULL",
                        "observacao": "TEXT",
                        "status_comparecimento": 'TEXT DEFAULT "pending"',
                        "historico_comparecimento": 'TEXT DEFAULT "[]"',
                        "criado_em": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
                        "atualizado_em": "TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
                    }

                    # Adicionar colunas faltantes
                    for col_name, col_type in needed_columns.items():
                        if col_name not in columns:
                            try:
                                cursor.execute(
                                    f"ALTER TABLE marcacoes ADD COLUMN {col_name} {col_type}"
                                )
                                logger.info(
                                    f"Coluna {col_name} adicionada à tabela marcacoes"
                                )
                            except sqlite3.Error as e:
                                logger.error(
                                    f"Erro ao adicionar coluna {col_name}: {e}"
                                )

                    conn.commit()

        except sqlite3.Error as e:
            logger.error(f"Erro ao criar/atualizar banco de dados de marcações: {e}")
            raise

    """Realiza a migração do banco de dados para a estrutura mais recente"""
    def migrate_database(self) -> None:
        """Realiza a migração do banco de dados para a estrutura mais recente"""
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()
                
                # Primeiro fazemos backup da tabela atual
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS marcacoes_backup AS 
                    SELECT * FROM marcacoes
                """)
                
                # Removemos a tabela antiga
                cursor.execute("DROP TABLE IF EXISTS marcacoes")
                
                # Criamos a nova tabela com a estrutura correta
                cursor.execute("""
                    CREATE TABLE marcacoes (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        nome TEXT NOT NULL,
                        renach TEXT NOT NULL UNIQUE,
                        telefone TEXT,
                        data_agendamento TEXT NOT NULL,
                        observacao TEXT,
                        status_comparecimento TEXT DEFAULT 'pending',
                        historico_comparecimento TEXT DEFAULT '[]',
                        criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        atualizado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)
                
                # Tentamos migrar os dados antigos
                try:
                    cursor.execute("""
                        INSERT INTO marcacoes (
                            nome, renach, telefone, data_agendamento, 
                            observacao, status_comparecimento, historico_comparecimento
                        )
                        SELECT 
                            nome, renach, telefone, data_agendamento,
                            observacao, status_comparecimento, historico_comparecimento
                        FROM marcacoes_backup
                    """)
                except sqlite3.Error as e:
                    logger.error(f"Erro ao migrar dados antigos: {e}")
                
                # Criamos o trigger para atualização de timestamp
                cursor.execute("""
                    CREATE TRIGGER IF NOT EXISTS update_marcacoes_timestamp 
                    AFTER UPDATE ON marcacoes
                    BEGIN
                        UPDATE marcacoes SET atualizado_em = CURRENT_TIMESTAMP 
                        WHERE id = NEW.id;
                    END;
                """)
                
                conn.commit()
                logger.info("Migração do banco de dados concluída com sucesso")
                
        except sqlite3.Error as e:
            logger.error(f"Erro durante a migração do banco: {e}")
            raise
    
    """Formata o número de telefone no padrão (XX) XXXXX-XXXX ou (XX) XXXX-XXXX."""

    @staticmethod
    def format_phone(phone: str) -> str:
        """Formata número de telefone para (XX) XXXXX-XXXX ou (XX) XXXX-XXXX"""
        phone = "".join(filter(str.isdigit, phone))
        if len(phone) == 11:
            return f"({phone[:2]}) {phone[2:7]}-{phone[7:]}"
        elif len(phone) == 10:
            return f"({phone[:2]}) {phone[2:6]}-{phone[6:]}"
        return phone

    """Valida os campos obrigatórios do formulário."""

    def validate_fields(self) -> bool:
        """Valida os campos obrigatórios do formulário"""
        if not all([self.name_entry, self.renach_entry]):
            logger.error("Campos de formulário não inicializados")
            messagebox.showerror("Erro", "Erro de inicialização do formulário")
            return False

        name = self.name_entry.get().strip()
        renach = self.renach_entry.get().strip()

        if not all([name, renach]):
            logger.warning("Tentativa de submissão com campos obrigatórios vazios")
            messagebox.showerror("Erro", "Nome e RENACH são obrigatórios!")
            return False

        if not renach.isdigit():
            logger.warning(f"RENACH inválido fornecido: {renach}")
            messagebox.showerror("Erro", "RENACH deve conter apenas números!")
            return False

        return True

    """Limpa todos os campos do formulário."""

    def clear_fields(self) -> None:
        """Limpa todos os campos do formulário"""
        if all(
            [
                self.name_entry,
                self.renach_entry,
                self.phone_entry,
                self.observation_text,
            ]
        ):
            self.name_entry.delete(0, tk.END)
            self.renach_entry.delete(0, tk.END)
            self.phone_entry.delete(0, tk.END)
            self.observation_text.delete("1.0", tk.END)
            logger.info("Campos do formulário limpos")
        else:
            logger.warning("Tentativa de limpar campos não inicializados")

    """Processa o envio do formulário de paciente."""

    def submit_patient(self) -> None:
        """Processa o envio do formulário de paciente"""
        if not self.validate_fields():
            return

        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()

                nome = self.name_entry.get().strip().upper()
                renach = self.renach_entry.get().strip()
                telefone = self.format_phone(self.phone_entry.get().strip())
                data_agendamento = self.appointment_entry.get_date().strftime(
                    "%Y-%m-%d"
                )
                observacao = self.observation_text.get("1.0", tk.END).strip()

                # Verifica existência do RENACH
                cursor.execute("SELECT id FROM marcacoes WHERE renach = ?", (renach,))

                if cursor.fetchone():
                    if messagebox.askyesno(
                        "Paciente Existente",
                        "Este RENACH já está cadastrado. Deseja atualizar a data da consulta?",
                    ):
                        cursor.execute(
                            """
                            UPDATE marcacoes 
                            SET data_agendamento = ?, 
                                observacao = ?
                            WHERE renach = ?
                        """,
                            (data_agendamento, observacao, renach),
                        )
                        logger.info(f"Marcação atualizada para RENACH: {renach}")
                        messagebox.showinfo("Sucesso", "Data da consulta atualizada!")
                else:
                    cursor.execute(
                        """
                        INSERT INTO marcacoes (
                            nome, renach, telefone, data_agendamento, observacao
                        ) VALUES (?, ?, ?, ?, ?)
                    """,
                        (nome, renach, telefone, data_agendamento, observacao),
                    )
                    logger.info(f"Nova marcação criada para RENACH: {renach}")
                    messagebox.showinfo("Sucesso", "Paciente cadastrado com sucesso!")

                conn.commit()
                self.clear_fields()

        except sqlite3.Error as e:
            logger.error(f"Erro ao submeter paciente: {e}")
            messagebox.showerror("Erro", "Erro ao processar operação. Verifique o log.")

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

    def get_patients_by_name_or_renach(self, search_term: str, selected_date: Optional[str] = None) -> List[Tuple]:
        """Busca pacientes por nome ou RENACH"""
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()
                search_term = search_term.lower() if search_term else ""

                query = """
                    SELECT nome, telefone, renach, 
                        COALESCE(status_comparecimento, 'pending') as status_comparecimento, 
                        observacao, data_agendamento
                    FROM marcacoes 
                    WHERE (LOWER(nome) LIKE ? OR LOWER(renach) LIKE ?)
                """
                params = [f"%{search_term}%", f"%{search_term}%"]

                if selected_date and not search_term:
                    query += " AND data_agendamento = ?"
                    params.append(selected_date)

                query += " ORDER BY nome"
                
                cursor.execute(query, params)
                return cursor.fetchall()
                
        except sqlite3.Error as e:
            self.logger.error(f"Erro ao buscar pacientes: {e}")
            return []

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

    def update_attendance_status(self, renach: str, status: str) -> None:
        """Atualiza o status de comparecimento do paciente"""
        try:
            with DatabaseConnection(self.db_name) as conn:
                cursor = conn.cursor()

                cursor.execute(
                    """
                    SELECT historico_comparecimento, data_agendamento 
                    FROM marcacoes 
                    WHERE renach = ?
                    """,
                    (renach,)
                )
                result = cursor.fetchone()
                
                if not result:
                    logger.warning(f"RENACH não encontrado: {renach}")
                    return

                historico_atual, data_agendamento = result

                try:
                    historico = json.loads(historico_atual) if historico_atual else []
                except json.JSONDecodeError:
                    logger.warning(f"Histórico inválido para RENACH {renach}, iniciando novo")
                    historico = []

                historico.append({
                    "data": data_agendamento,
                    "status": status,
                    "atualizado_em": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

                cursor.execute(
                    """
                    UPDATE marcacoes 
                    SET status_comparecimento = ?, 
                        historico_comparecimento = ?
                    WHERE renach = ?
                    """, 
                    (status, json.dumps(historico), renach)
                )

                conn.commit()
                logger.info(f"Status atualizado para RENACH {renach}: {status}")
                self.update_patient_list()

        except sqlite3.Error as e:
            logger.error(f"Erro ao atualizar status: {e}")
            messagebox.showerror("Erro", "Erro ao atualizar status. Verifique o log.")
