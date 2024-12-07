# /home/lusca/py_excel/tkinter campssa/frames/main_frame.py
from tkinter import *
import tkinter as tk
from tkinter import ttk, messagebox
from funcoes_botoes import FuncoesBotoes, GerenciadorPlanilhas
from banco import SistemaContas
from planilhas import Planilhas
from tkcalendar import DateEntry
from banco import DataBaseMarcacao
from config import config_manager
import json
import hashlib
from frames.ntfs_frame import EmitirNota
from graficos import GraficoMarcacoes
from auth.user_manager import UserManager
import sqlite3

class AdminUserManager:
    def __init__(self, master, user_manager):
        self.master = master
        self.user_manager = user_manager
        self.window = None

    def show(self):
        self.window = tk.Toplevel(self.master)
        self.window.title("Gerenciamento de Usuários")
        self.window.geometry("800x600")
        self.window.configure(bg='#2c3e50')

        # Frame principal
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill='both', expand=True)

        # Lista de usuários
        list_frame = ttk.LabelFrame(main_frame, text="Usuários", padding="10")
        list_frame.pack(fill='both', expand=True, pady=(0, 10))

        # Treeview para listar usuários
        columns = ('Username', 'Role', 'Status', 'Criado Por', 'Último Login')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings')

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Botões de ação
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=5)

        ttk.Button(
            btn_frame,
            text="Criar Subusuário",
            command=self.show_create_user
        ).pack(side='left', padx=5)

        ttk.Button(
            btn_frame,
            text="Editar Permissões",
            command=self.show_edit_permissions
        ).pack(side='left', padx=5)

        ttk.Button(
            btn_frame,
            text="Desativar/Ativar Usuário",
            command=self.toggle_user_status
        ).pack(side='left', padx=5)

        self.load_users()

    def load_users(self):
        """Carrega a lista de usuários do admin atual"""
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)

            with sqlite3.connect('login.db') as conn:
                cursor = conn.cursor()
                
                # Busca apenas os usuários criados pelo admin atual
                cursor.execute("""
                    SELECT user, role, is_active, created_by, last_login
                    FROM users
                    WHERE created_by = ?
                    ORDER BY user
                """, (self.user_manager.current_user.username,))

                for user in cursor.fetchall():
                    username, role, is_active, created_by, last_login = user
                    status = "Ativo" if is_active else "Inativo"
                    
                    self.tree.insert('', 'end', values=(
                        username,
                        role,
                        status,
                        created_by or "Sistema",
                        last_login or "Nunca"
                    ))

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar usuários: {str(e)}")

    def show_create_user(self):
        """Mostra janela para criar novo subusuário"""
        window = tk.Toplevel(self.window)
        window.title("Criar Subusuário")
        window.geometry("500x600")
        window.configure(bg='#2c3e50')

        frame = ttk.Frame(window, padding="20")
        frame.pack(fill='both', expand=True)

        # Campos do formulário
        ttk.Label(frame, text="Username:").pack(anchor='w', pady=2)
        username_entry = ttk.Entry(frame, width=40)
        username_entry.pack(fill='x', pady=2)

        ttk.Label(frame, text="Password:").pack(anchor='w', pady=2)
        password_entry = ttk.Entry(frame, show="*", width=40)
        password_entry.pack(fill='x', pady=2)

        # Frame de permissões
        perm_frame = ttk.LabelFrame(frame, text="Permissões", padding=10)
        perm_frame.pack(fill='x', pady=10)

        permission_vars = {}
        for perm_key, perm_name in self.user_manager.PERMISSIONS.items():
            var = tk.BooleanVar(value=False)
            permission_vars[perm_key] = var
            ttk.Checkbutton(
                perm_frame,
                text=perm_name,
                variable=var
            ).pack(anchor='w')

        def create_subuser():
            username = username_entry.get().strip()
            password = password_entry.get().strip()

            if not username or not password:
                messagebox.showerror("Erro", "Username e password são obrigatórios")
                return

            selected_permissions = [
                perm for perm, var in permission_vars.items()
                if var.get()
            ]

            try:
                with sqlite3.connect('login.db') as conn:
                    cursor = conn.cursor()
                    
                    # Verifica se já existe
                    cursor.execute("SELECT 1 FROM users WHERE user = ?", (username,))
                    if cursor.fetchone():
                        messagebox.showerror("Erro", "Username já existe")
                        return

                    # Cria o subusuário
                    hashed_password = hashlib.sha256(password.encode()).hexdigest()
                    cursor.execute("""
                        INSERT INTO users (
                            user, password, role, permissions, is_active, 
                            created_by, parent_admin
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (
                        username, hashed_password, 'subuser',
                        json.dumps(selected_permissions), 1,
                        self.user_manager.current_user.username,
                        self.user_manager.current_user.username
                    ))
                    
                    conn.commit()
                    messagebox.showinfo("Sucesso", "Subusuário criado com sucesso!")
                    window.destroy()
                    self.load_users()

            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao criar subusuário: {str(e)}")

        ttk.Button(
            frame,
            text="Criar Subusuário",
            command=create_subuser
        ).pack(pady=20)

    def show_edit_permissions(self):
        """Mostra janela para editar permissões de um subusuário"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um usuário")
            return

        username = self.tree.item(selection[0])['values'][0]

        window = tk.Toplevel(self.window)
        window.title(f"Editar Permissões - {username}")
        window.geometry("400x500")
        window.configure(bg='#2c3e50')

        frame = ttk.Frame(window, padding="20")
        frame.pack(fill='both', expand=True)

        # Carrega permissões atuais
        current_permissions = []
        try:
            with sqlite3.connect('login.db') as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT permissions FROM users WHERE user = ?",
                    (username,)
                )
                result = cursor.fetchone()
                if result and result[0]:
                    current_permissions = json.loads(result[0])
        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao carregar permissões: {str(e)}")
            return

        # Frame de permissões
        perm_frame = ttk.LabelFrame(frame, text="Permissões", padding=10)
        perm_frame.pack(fill='x', pady=10)

        permission_vars = {}
        for perm_key, perm_name in self.user_manager.PERMISSIONS.items():
            var = tk.BooleanVar(value=perm_key in current_permissions)
            permission_vars[perm_key] = var
            ttk.Checkbutton(
                perm_frame,
                text=perm_name,
                variable=var
            ).pack(anchor='w')

        def save_permissions():
            selected_permissions = [
                perm for perm, var in permission_vars.items()
                if var.get()
            ]

            try:
                with sqlite3.connect('login.db') as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        UPDATE users 
                        SET permissions = ?
                        WHERE user = ?
                    """, (json.dumps(selected_permissions), username))
                    conn.commit()

                messagebox.showinfo("Sucesso", "Permissões atualizadas!")
                window.destroy()
                self.load_users()

            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao atualizar permissões: {str(e)}")

        ttk.Button(
            frame,
            text="Salvar Permissões",
            command=save_permissions
        ).pack(pady=20)

    def toggle_user_status(self):
        """Ativa/Desativa um subusuário"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um usuário")
            return

        username = self.tree.item(selection[0])['values'][0]
        current_status = self.tree.item(selection[0])['values'][2]
        new_status = 0 if current_status == "Ativo" else 1

        try:
            with sqlite3.connect('login.db') as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    UPDATE users 
                    SET is_active = ?
                    WHERE user = ? AND created_by = ?
                """, (new_status, username, self.user_manager.current_user.username))
                conn.commit()

            messagebox.showinfo(
                "Sucesso",
                f"Usuário {'ativado' if new_status else 'desativado'} com sucesso!"
            )
            self.load_users()

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao alterar status: {str(e)}")




class MainFrame(Frame):
    """Frame principal da aplicação que gerencia a interface do usuário e suas interações."""

    def __init__(self, master, planilhas: Planilhas, file_path: str, app, user_manager: UserManager):
        self.ui_config = config_manager.get_config("UI_CONFIG")
        self.app_config = config_manager.get_config("APP_CONFIG")
        super().__init__(master, bg=self.ui_config["colors"]["background"])
        self.user_manager = user_manager
        self.configure_window()
        self._init_attributes(master, planilhas, file_path, app)
        self._setup_styles()
        self.create_widgets()
        self.grafico_marcacoes = GraficoMarcacoes(master, planilhas, file_path, app)


    def configure_window(self):
        """Configura as dimensões e posicionamento da janela."""
        window_config = self.app_config["window"]
        self.master.geometry(self.app_config["main_geometry"])
        self.master.minsize(window_config["min_width"], window_config["min_height"])
        self.master.maxsize(window_config["max_width"], window_config["max_height"])
        self.center_window()
        self.master.resizable(True, True)

    def center_window(self):
        """Centraliza a janela na tela."""
        self.master.update_idletasks()
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        window_width = self.master.winfo_width()
        window_height = self.master.winfo_height()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.master.geometry(f"+{x}+{y}")

    def _init_attributes(self, master, planilhas, file_path, app):
        """Inicializa atributos e dependências da classe."""
        self.current_user = getattr(app, "get_current_user", lambda: None)()
        self.funcoes_botoes = FuncoesBotoes(master, planilhas, file_path, app)
        self.emitir_nota = EmitirNota(master)
        self.banco = DataBaseMarcacao(master, planilhas, file_path, app)
        self.sistema_contas = SistemaContas(file_path, current_user=self.current_user)
        self.gerenciador_planilhas = GerenciadorPlanilhas(master, self.sistema_contas)
        self.master = master
        self.file_path = file_path
        self.app = app

        # Atalho para maximizar
        self.master.bind("<F10>", lambda e: self.master.state('zoomed'))

    def _setup_styles(self):
        """Configura estilos dos widgets."""
        style = ttk.Style()
        frame_style = self.ui_config["styles"]["frame"]
        
        style.configure(
            "Custom.TLabelframe",
            background=self.ui_config["colors"]["frame"],
            padding=frame_style["padding"],
            relief=frame_style["relief"],
            borderwidth=frame_style["borderwidth"]
        )
        
        style.configure(
            "Custom.TLabelframe.Label",
            background=self.ui_config["colors"]["frame"],
            foreground=self.ui_config["colors"]["text"],
            font=self.ui_config["fonts"]["header"]
        )

    def create_button(self, parent, text, command, width=20):
        """Cria botão com estilo consistente."""
        button_style = self.ui_config["styles"]["button"]
        colors = self.ui_config["colors"]
        
        btn = Button(
            parent,
            text=text,
            command=command,
            bg=colors["button"],
            fg=colors["text"],
            font=self.ui_config["fonts"]["button"],
            **button_style
        )

        btn.bind("<Enter>", lambda e: btn.config(bg=colors["button_hover"]))
        btn.bind("<Leave>", lambda e: btn.config(bg=colors["button"]))

        return btn

    def create_widgets(self):
        """Cria e organiza os widgets da interface."""
        main_container = Frame(self, bg=self.ui_config["colors"]["background"])
        main_container.pack(expand=True, fill="both")

        # Título
        title_text = self.app_config["title"]
        if self.current_user:
            title_text += f" - {self.current_user}"

        Label(
            main_container,
            text=title_text,
            font=self.ui_config["fonts"]["title"],
            bg=self.ui_config["colors"]["background"],
            fg=self.ui_config["colors"]["title"]
        ).pack(pady=self.ui_config["padding"]["title"])

        # Frame para organização em grid
        grid_frame = Frame(main_container, bg=self.ui_config["colors"]["background"])
        grid_frame.pack(expand=True, fill="both", padx=self.ui_config["padding"]["large"])
        grid_frame.grid_columnconfigure((0, 1), weight=1)

        self._create_section_frames(grid_frame)

    def _create_section_frames(self, grid_frame):
        """Cria as seções principais da interface."""
        sections = [
            ("Cadastro e Gestão", 0, 0, [
                ("Adicionar Paciente", self.adicionar_informacao, "add_paciente"),
                ("Excluir Paciente", self.excluir_informacao, "delet_paciente"),
                ("Informações do Atendimento", self.exibir, "information_service")
            ]),
            ("Agenda e Marcações", 0, 1, [
                ("Marcar Paciente", self.marcar_paciente, "marcar_paciente"),
                ("Visualizar Marcações", self.visu_marcacoes, "vizu_marcacoes")
            ]),
            ("Gestão Financeira", 1, 0, [
                ("Relatorio de Pagamentos", self.resultados_consulta, "relatorio_pag"),
                ("Valores Atendimento", self.exibir_contas, "valores_atend"),
                ("Gastos da Clinica", self.fechamento_contas, "gastos_clinica")
            ]),
            ("Documentos e Relatórios", 1, 1, [
                ("Emitir NTFS-e", self.emitir_notas, "emitir_ntfs"),
                ("Enviar Relatório WhatsApp", self.relatorio_wpp, "enviar_wpp"),
                ("Enviar Relatório Email", self.relatorio_email, "enviar_email")
            ]),
            ("Ferramentas", 2, 0, [
                ("Gerenciar Planilhas/Sheets", self.planilha_sheet, "gerenciar_planilha"),
                ("Gráficos Gerais", self.abrir_grafico, "graficos_gerais")
            ], 2)
        ]

        # Adiciona seção de administração apenas se o usuário for admin
        is_dev_tools_admin = False
        if self.user_manager.current_user and self.user_manager.current_user.role == 'admin':
            try:
                with sqlite3.connect('login.db') as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        SELECT created_by 
                        FROM users 
                        WHERE user = ? AND role = 'admin'
                    """, (self.user_manager.current_user.username,))
                    result = cursor.fetchone()
                    if result and result[0] == 'dev_tools':
                        is_dev_tools_admin = True
            except sqlite3.Error as e:
                print(f"Erro ao verificar admin: {e}")

        if (self.user_manager.current_user and 
            self.user_manager.current_user.role == 'admin' and
            self.user_manager.current_user.created_by == 'dev_tools'):
            sections.append(
                ("Administração", 3, 0, [
                    ("Gerenciar Usuários", self.open_user_management, "manage_users")
                ], 2)
            )

        padding = self.ui_config["padding"]
        
        for section_info in sections:
            title, row, col, buttons, *extra = section_info
            colspan = extra[0] if extra else 1
            
            frame = ttk.LabelFrame(
                grid_frame,
                text=f" {title} ",
                style="Custom.TLabelframe"
            )
            frame.grid(
                row=row, column=col, columnspan=colspan,
                padx=padding["section"], 
                pady=padding["section"], 
                sticky="nsew"
            )

            for btn_text, btn_command, permission in buttons:
                if self.user_manager.verificar_permissao(permission):
                    self.create_button(frame, btn_text, btn_command).pack(
                        pady=padding["button"],
                        padx=padding["button"],
                        fill="x"
                    )

    def _create_frame(self, parent, title, row, column, buttons, columnspan=1):
        """Cria um frame com título e botões."""
        frame = ttk.LabelFrame(parent, text=title, style="Custom.TLabelframe")
        frame.grid(row=row, column=column, columnspan=columnspan, padx=8, pady=4, sticky="nsew")

        for btn_text, btn_command in buttons:
            self.create_button(frame, btn_text, btn_command).pack(pady=4, padx=8, fill="x")

        return frame

    def get_current_user(self):
        return self.user_manager.current_user



    # Adicione o método para abrir o gerenciador na classe MainFrame
    def open_user_management(self):
        """Abre o gerenciador de usuários"""
        if hasattr(self, 'user_manager') and self.user_manager.current_user:
            try:
                with sqlite3.connect('login.db') as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        SELECT created_by 
                        FROM users 
                        WHERE user = ? AND role = 'admin'
                    """, (self.user_manager.current_user.username,))
                    result = cursor.fetchone()
                    
                    if result and result[0] == 'dev_tools':
                        manager = AdminUserManager(self.master, self.user_manager)
                        manager.show()
                    else:
                        messagebox.showwarning("Acesso Negado", 
                            "Apenas administradores criados pelo Developer Tools podem acessar esta função")
            except sqlite3.Error as e:
                messagebox.showerror("Erro", f"Erro ao verificar permissões: {str(e)}")
        else:
            messagebox.showwarning("Acesso Negado", "Acesso não autorizado")



    def adicionar_informacao(self): 
        self.funcoes_botoes.adicionar_informacao()
    
    def excluir_informacao(self): 
        self.funcoes_botoes.excluir()
    
    def exibir(self): 
        self.funcoes_botoes.exibir_informacao()
    
    def exibir_contas(self): 
        self.funcoes_botoes.valores_totais()
    
    def emitir_notas(self): 
        self.funcoes_botoes.processar_notas_fiscais()
    
    def resultados_consulta(self): 
        self.funcoes_botoes.mostrar_valores_atendimentos()
    
    def relatorio_wpp(self): 
        self.funcoes_botoes.enviar_whatsapp()
    
    def relatorio_email(self): 
        self.funcoes_botoes.enviar_email()
    
    def marcar_paciente(self): 
        self.banco.add_user()
    
    def visu_marcacoes(self): 
        self.banco.view_marcacoes()
    
    def fechamento_contas(self): 
        self.sistema_contas.abrir_janela()
    
    def planilha_sheet(self): 
        self.gerenciador_planilhas.abrir_gerenciador()
    
    def abrir_grafico(self):
        self.grafico_marcacoes.gerar_grafico()