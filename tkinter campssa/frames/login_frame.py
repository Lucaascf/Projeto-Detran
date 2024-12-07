# /home/lusca/py_excel/tkinter campssa/frames/login_frame.py
import tkinter as tk
import re
import logging
from tkinter import messagebox, ttk
from typing import Callable
from auth.user_manager import UserManager, User
from config import ConfigManager



class BaseFrame(tk.Frame):
    def __init__(self, master, config_manager):
        self.config_manager = config_manager
        ui_config = self.config_manager.get_config('UI_CONFIG')
        super().__init__(master, bg=ui_config['colors']['background'])
        self._configure_window()
        self._setup_base_grid()

    def _configure_window(self):
        width, height = 400, 300
        screen_w = self.master.winfo_screenwidth()
        screen_h = self.master.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.master.geometry(f"{width}x{height}+{x}+{y}")
        self.master.minsize(width, height)
        self.master.maxsize(width, height)
        self.master.resizable(False, False)

    def _setup_base_grid(self):
        self.grid(row=0, column=0, sticky="nsew")
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def show(self): self.grid()
    def hide(self): self.grid_remove()

class LoginFrame(BaseFrame):
    def __init__(self, master, config_manager, user_manager: UserManager, on_login_success: Callable):
        self.user_manager = user_manager
        self.on_login_success = on_login_success
        self.funcoes_botoes = None
        super().__init__(master, config_manager)
        self._create_widgets()

    def set_funcoes_botoes(self, funcoes_botoes):
        self.funcoes_botoes = funcoes_botoes

    @property
    def current_user(self):
        return self.user_manager.current_user

    def _create_widgets(self):
        ui_config = self.config_manager.get_config('UI_CONFIG')
        colors = ui_config['colors']
        fonts = ui_config['fonts']

        container = tk.Frame(self, bg=colors['background'])
        container.grid(row=1, column=0)

        tk.Label(
            container,
            text="Faça Login",
            font=fonts['title'],
            bg=colors['background'],
            fg=colors['title']
        ).grid(row=0, column=0, columnspan=2, pady=(0, 30))

        fields = [("Usuário:", "entry_user", ""), ("Senha:", "entry_password", "*")]
        for i, (label_text, attr_name, show_char) in enumerate(fields, 1):
            tk.Label(
                container,
                text=label_text,
                font=fonts['normal'],
                bg=colors['background'],
                fg=colors['text']
            ).grid(row=i, column=0, padx=(0, 10), pady=5, sticky="e")
            
            entry = tk.Entry(
                container,
                font=fonts['normal'],
                show=show_char,
                width=20,
                bg=colors['frame'],
                fg=colors['text'],
                insertbackground=colors['text']
            )
            entry.grid(row=i, column=1, pady=5, sticky="w")
            setattr(self, attr_name, entry)

        button_frame = tk.Frame(container, bg=colors['background'])
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        buttons = [
            ("Login", self.perform_login),
            ("Criar Conta", self.funcoes_botoes.mostrar_criar_conta if self.funcoes_botoes else None)
        ]

        for i, (text, command) in enumerate(buttons):
            if command:
                tk.Button(
                    button_frame,
                    text=text,
                    command=command,
                    bg=colors['button'],
                    fg=colors['text'],
                    activebackground=colors['button_hover'],
                    activeforeground=colors['text'],
                    font=fonts['button'],
                    width=12
                ).grid(row=0, column=i, padx=5)

    def perform_login(self):
        user = self.entry_user.get().strip()
        password = self.entry_password.get().strip()
        
        if not all([user, password]):
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        if self.user_manager.authenticate(user, password):
            self.on_login_success()
        else:
            messagebox.showerror("Erro", "Usuário ou senha incorretos")
            self.entry_user.delete(0, tk.END)
            self.entry_password.delete(0, tk.END)
            self.entry_user.focus()


class CriarContaFrame(BaseFrame):
    """Frame para criação de conta com validação"""
    def __init__(self, master, config_manager, user_manager: UserManager, on_account_created: Callable):
        self.user_manager = user_manager
        self.on_account_created = on_account_created
        self.funcoes_botoes = None
        super().__init__(master, config_manager)
        self._create_widgets()

    def _create_widgets(self):
        ui_config = self.config_manager.get_config('UI_CONFIG')
        colors = ui_config['colors']
        fonts = ui_config['fonts']

        container = tk.Frame(self, bg=colors['background'])
        container.grid(row=1, column=0)

        # Título
        tk.Label(
            container,
            text="Criar Nova Conta",
            font=fonts['title'],
            bg=colors['background'],
            fg=colors['title']
        ).grid(row=0, column=0, columnspan=2, pady=(0, 30))

        # Fields
        self.entry_widgets = {}
        fields = [
            ("username", "Usuário:", ""),
            ("password", "Senha:", "*"),
            ("confirm_password", "Confirmar Senha:", "*"),
        ]
        
        # Cria campos
        for i, (field_name, label_text, show_char) in enumerate(fields, 1):
            tk.Label(
                container,
                text=label_text,
                font=fonts['normal'],
                bg=colors['background'],
                fg=colors['text']
            ).grid(row=i, column=0, padx=(0, 10), pady=5, sticky="e")
            
            entry = tk.Entry(
                container,
                font=fonts['normal'],
                show=show_char,
                width=20,
                bg=colors['frame'],
                fg=colors['text'],
                insertbackground=colors['text']
            )
            entry.grid(row=i, column=1, pady=5, sticky="w")
            self.entry_widgets[field_name] = entry

        # Buttons
        button_frame = tk.Frame(container, bg=colors['background'])
        button_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=20)

        buttons = [("Criar Conta", self.create_account), ("Voltar", self.voltar_login)]
        for i, (text, command) in enumerate(buttons):
            tk.Button(
                button_frame,
                text=text,
                command=command,
                bg=colors['button'],
                fg=colors['text'],
                activebackground=colors['button_hover'],
                activeforeground=colors['text'],
                font=fonts['button'],
                width=12
            ).grid(row=0, column=i, padx=5)

    def create_account(self):
        """Cria uma nova conta com validação adequada"""
        try:
            username = self.entry_widgets["username"].get().strip()
            password = self.entry_widgets["password"].get().strip()
            confirm_password = self.entry_widgets["confirm_password"].get().strip()

            if not all([username, password, confirm_password]):
                messagebox.showerror("Erro", "Preencha todos os campos!")
                return

            if password != confirm_password:
                messagebox.showerror("Erro", "As senhas não coincidem!")
                return

            # Verifica se tem permissão para criar usuários
            current_user = self.user_manager.current_user
            if current_user and 'manage_users' not in current_user.permissions:
                messagebox.showerror("Erro", "Sem permissão para criar novos usuários")
                return
            
            # Define papel e permissões padrão
            role = 'employee'
            permissions = UserManager.DEFAULT_PERMISSIONS['employee']

            # Cria o usuário
            if self.user_manager.create_user(username, password, role, permissions):
                messagebox.showinfo("Sucesso", "Conta criada com sucesso!")
                self.clear_fields()
                self.on_account_created()
            else:
                messagebox.showerror("Erro", "Usuário já existe!")
                self.entry_widgets["username"].focus()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar conta: {str(e)}")
            logging.error(f"Erro na criação de conta: {str(e)}")

    def clear_fields(self):
        """Limpa todos os campos do formulário"""
        for entry in self.entry_widgets.values():
            entry.delete(0, tk.END)

    def voltar_login(self):
        """Volta para a tela de login"""
        self.clear_fields()
        if self.funcoes_botoes:
            self.funcoes_botoes.voltar_para_login()

    def set_funcoes_botoes(self, funcoes_botoes):
        """Define as funções dos botões"""
        self.funcoes_botoes = funcoes_botoes
    


class UserInterface:
    """Interface gráfica para gerenciamento de usuários"""
    
    def __init__(self, master: tk.Tk, user_manager: UserManager):
        self.master = master
        self.user_manager = user_manager
        self.setup_styles()
        
    def setup_styles(self):
        """Configura estilos da interface"""
        style = ttk.Style()
        style.configure('Header.TLabel', font=('Arial', 12, 'bold'))
        style.configure('Normal.TLabel', font=('Arial', 10))
        style.configure('Success.TLabel', font=('Arial', 10), foreground='green')
        style.configure('Error.TLabel', font=('Arial', 10), foreground='red')

    def show_login(self):
        """Exibe tela de login"""
        login_window = tk.Toplevel(self.master)
        login_window.title("Login")
        login_window.geometry("300x200")
        
        frame = ttk.Frame(login_window, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(frame, text="Usuário:", style='Normal.TLabel').grid(row=0, column=0, pady=5)
        username_entry = ttk.Entry(frame)
        username_entry.grid(row=0, column=1, pady=5)
        
        ttk.Label(frame, text="Senha:", style='Normal.TLabel').grid(row=1, column=0, pady=5)
        password_entry = ttk.Entry(frame, show="*")
        password_entry.grid(row=1, column=1, pady=5)
        
        def try_login():
            username = username_entry.get()
            password = password_entry.get()
            
            user = self.user_manager.authenticate(username, password)
            if user:
                login_window.destroy()
                self.show_main_interface(user)
            else:
                messagebox.showerror("Erro", "Usuário ou senha inválidos")
        
        ttk.Button(frame, text="Login", command=try_login).grid(row=2, column=0, columnspan=2, pady=20)
        
        # Centraliza a janela
        login_window.update_idletasks()
        width = login_window.winfo_width()
        height = login_window.winfo_height()
        x = (login_window.winfo_screenwidth() // 2) - (width // 2)
        y = (login_window.winfo_screenheight() // 2) - (height // 2)
        login_window.geometry(f'{width}x{height}+{x}+{y}')

    def show_main_interface(self, user: User):
        """Exibe interface principal com base nas permissões do usuário"""
        window = tk.Toplevel(self.master)
        window.title(f"Sistema - {user.username} ({UserManager.ROLES[user.role]})")
        window.geometry("800x600")
        
        # Menu principal
        menubar = tk.Menu(window)
        window.config(menu=menubar)
        
        # Menu de usuários (apenas para quem tem permissão)
        if 'manage_users' in user.permissions:
            user_menu = tk.Menu(menubar, tearoff=0)
            menubar.add_cascade(label="Usuários", menu=user_menu)
            user_menu.add_command(label="Criar Usuário", command=lambda: self.show_create_user())
            user_menu.add_command(label="Gerenciar Usuários", command=lambda: self.show_manage_users())
        
        # Notebook para diferentes seções
        notebook = ttk.Notebook(window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Adiciona abas baseadas nas permissões
        if 'view_reports' in user.permissions:
            reports_frame = ttk.Frame(notebook)
            notebook.add(reports_frame, text='Relatórios')
            self.setup_reports_tab(reports_frame)
        
        if any(perm in user.permissions for perm in ['add_patients', 'edit_patients']):
            patients_frame = ttk.Frame(notebook)
            notebook.add(patients_frame, text='Pacientes')
            self.setup_patients_tab(patients_frame, user.permissions)
        
        if 'financial_access' in user.permissions:
            financial_frame = ttk.Frame(notebook)
            notebook.add(financial_frame, text='Financeiro')
            self.setup_financial_tab(financial_frame)
        
        # Barra de status
        status_frame = ttk.Frame(window)
        status_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(
            status_frame, 
            text=f"Logado como: {user.username} | Último acesso: {user.last_login}",
            style='Normal.TLabel'
        ).pack(side='left')
        
        # Botão de logout
        ttk.Button(
            status_frame,
            text="Logout",
            command=lambda: self.logout(window)
        ).pack(side='right')
        
        # Centraliza a janela
        self._center_window(window)
        
    def show_create_user(self):
        """Interface para criação de novo usuário"""
        window = tk.Toplevel(self.master)
        window.title("Criar Novo Usuário")
        window.geometry("400x500")
        
        frame = ttk.Frame(window, padding="20")
        frame.pack(fill='both', expand=True)
        
        # Campos do formulário
        ttk.Label(frame, text="Nome de Usuário:", style='Normal.TLabel').pack(anchor='w', pady=5)
        username_entry = ttk.Entry(frame, width=40)
        username_entry.pack(fill='x', pady=5)
        
        ttk.Label(frame, text="Senha:", style='Normal.TLabel').pack(anchor='w', pady=5)
        password_entry = ttk.Entry(frame, show="*", width=40)
        password_entry.pack(fill='x', pady=5)
        
        ttk.Label(frame, text="Confirmar Senha:", style='Normal.TLabel').pack(anchor='w', pady=5)
        confirm_entry = ttk.Entry(frame, show="*", width=40)
        confirm_entry.pack(fill='x', pady=5)
        
        # Seleção de cargo
        ttk.Label(frame, text="Cargo:", style='Normal.TLabel').pack(anchor='w', pady=5)
        role_var = tk.StringVar()
        role_combo = ttk.Combobox(frame, textvariable=role_var, state='readonly')
        role_combo['values'] = list(UserManager.ROLES.values())
        role_combo.pack(fill='x', pady=5)
        
        # Frame para permissões
        perm_frame = ttk.LabelFrame(frame, text="Permissões", padding="10")
        perm_frame.pack(fill='x', pady=10)
        
        # Variáveis para checkbuttons de permissões
        perm_vars = {}
        for perm_key, perm_label in UserManager.PERMISSIONS.items():
            var = tk.BooleanVar()
            perm_vars[perm_key] = var
            ttk.Checkbutton(
                perm_frame,
                text=perm_label,
                variable=var
            ).pack(anchor='w')
        
        def on_role_change(*args):
            """Atualiza permissões padrão ao mudar o cargo"""
            selected_role = [k for k, v in UserManager.ROLES.items() if v == role_var.get()][0]
            default_perms = UserManager.DEFAULT_PERMISSIONS.get(selected_role, [])
            
            for perm_key, var in perm_vars.items():
                var.set(perm_key in default_perms)
        
        role_var.trace('w', on_role_change)
        
        def validate_and_create():
            """Valida entradas e cria novo usuário"""
            username = username_entry.get().strip()
            password = password_entry.get()
            confirm = confirm_entry.get()
            
            if not username or not password:
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios")
                return
                
            if password != confirm:
                messagebox.showerror("Erro", "As senhas não coincidem")
                return
                
            if not role_var.get():
                messagebox.showerror("Erro", "Selecione um cargo")
                return
            
            # Obtém role_key do valor selecionado
            role_key = [k for k, v in UserManager.ROLES.items() if v == role_var.get()][0]
            
            # Coleta permissões selecionadas
            selected_permissions = [
                perm for perm, var in perm_vars.items() 
                if var.get()
            ]
            
            # Tenta criar o usuário
            if self.user_manager.create_user(username, password, role_key, selected_permissions):
                messagebox.showinfo("Sucesso", "Usuário criado com sucesso!")
                window.destroy()
            else:
                messagebox.showerror("Erro", "Não foi possível criar o usuário")
        
        # Botões de ação
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill='x', pady=20)
        
        ttk.Button(
            button_frame,
            text="Criar Usuário",
            command=validate_and_create
        ).pack(side='left', padx=5)
        
        ttk.Button(
            button_frame,
            text="Cancelar",
            command=window.destroy
        ).pack(side='right', padx=5)
        
        self._center_window(window)

    def show_manage_users(self):
        """Interface para gerenciamento de usuários"""
        window = tk.Toplevel(self.master)
        window.title("Gerenciar Usuários")
        window.geometry("800x600")
        
        # Frame principal
        main_frame = ttk.Frame(window, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        # Treeview para lista de usuários
        tree = ttk.Treeview(main_frame, columns=('Username', 'Role', 'Status', 'Last Login'))
        tree.heading('Username', text='Usuário')
        tree.heading('Role', text='Cargo')
        tree.heading('Status', text='Status')
        tree.heading('Last Login', text='Último Acesso')
        
        tree.column('#0', width=0, stretch=False)
        tree.column('Username', width=150)
        tree.column('Role', width=100)
        tree.column('Status', width=100)
        tree.column('Last Login', width=150)
        
        tree.pack(fill='both', expand=True, pady=10)
        
        def load_users():
            """Carrega lista de usuários"""
            for item in tree.get_children():
                tree.delete(item)
                
            for user in self.user_manager.get_users():
                status = "Ativo" if user.is_active else "Inativo"
                last_login = user.last_login or "Nunca"
                tree.insert('', 'end', values=(
                    user.username,
                    UserManager.ROLES[user.role],
                    status,
                    last_login
                ))
        
        def edit_user():
            """Abre janela para edição do usuário selecionado"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Aviso", "Selecione um usuário para editar")
                return
                
            username = tree.item(selection[0])['values'][0]
            selected_user = next((u for u in self.user_manager.get_users() 
                                if u.username == username), None)
                                
            if selected_user:
                self.show_edit_user(selected_user, after_edit=load_users)
        
        def toggle_user_status():
            """Ativa/desativa o usuário selecionado"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Aviso", "Selecione um usuário")
                return
                
            username = tree.item(selection[0])['values'][0]
            selected_user = next((u for u in self.user_manager.get_users() 
                                if u.username == username), None)
                                
            if selected_user:
                if selected_user.username == 'admin':
                    messagebox.showwarning("Aviso", "Não é possível desativar o usuário admin")
                    return
                    
                new_status = not selected_user.is_active
                if self.user_manager.update_user(selected_user.id, {'is_active': new_status}):
                    load_users()
                    messagebox.showinfo(
                        "Sucesso", 
                        f"Usuário {'ativado' if new_status else 'desativado'} com sucesso"
                    )
        
        # Botões de ação
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        ttk.Button(
            button_frame,
            text="Editar Usuário",
            command=edit_user
        ).pack(side='left', padx=5)
        
        ttk.Button(
            button_frame,
            text="Ativar/Desativar",
            command=toggle_user_status
        ).pack(side='left', padx=5)
        
        ttk.Button(
            button_frame,
            text="Atualizar Lista",
            command=load_users
        ).pack(side='left', padx=5)
        
        ttk.Button(
            button_frame,
            text="Fechar",
            command=window.destroy
        ).pack(side='right', padx=5)
        
        # Carrega usuários inicialmente
        load_users()
        
        self._center_window(window)

    def show_edit_user(self, user: User, after_edit=None):
        """Interface para edição de usuário existente"""
        window = tk.Toplevel(self.master)
        window.title(f"Editar Usuário - {user.username}")
        window.geometry("400x500")
        
        frame = ttk.Frame(window, padding="20")
        frame.pack(fill='both', expand=True)
        
        # Informações do usuário
        ttk.Label(frame, text=f"Usuário: {user.username}", style='Header.TLabel').pack(anchor='w', pady=5)
        
        # Campos de edição
        ttk.Label(frame, text="Nova Senha (deixe em branco para manter):", style='Normal.TLabel').pack(anchor='w', pady=5)
        password_entry = ttk.Entry(frame, show="*", width=40)
        password_entry.pack(fill='x', pady=5)
        
        ttk.Label(frame, text="Confirmar Nova Senha:", style='Normal.TLabel').pack(anchor='w', pady=5)
        confirm_entry = ttk.Entry(frame, show="*", width=40)
        confirm_entry.pack(fill='x', pady=5)
        
        # Seleção de cargo
        ttk.Label(frame, text="Cargo:", style='Normal.TLabel').pack(anchor='w', pady=5)
        role_var = tk.StringVar(value=UserManager.ROLES[user.role])
        role_combo = ttk.Combobox(frame, textvariable=role_var, state='readonly')
        role_combo['values'] = list(UserManager.ROLES.values())
        role_combo.pack(fill='x', pady=5)
        
        # Frame para permissões
        perm_frame = ttk.LabelFrame(frame, text="Permissões", padding="10")
        perm_frame.pack(fill='x', pady=10)
        
        # Variáveis para checkbuttons de permissões
        perm_vars = {}
        for perm_key, perm_label in UserManager.PERMISSIONS.items():
            var = tk.BooleanVar(value=perm_key in user.permissions)
            perm_vars[perm_key] = var
            ttk.Checkbutton(
                perm_frame,
                text=perm_label,
                variable=var
            ).pack(anchor='w')

        def validate_and_update():
            """Valida entradas e atualiza usuário"""
            updates = {}
            
            # Verifica senha se fornecida
            password = password_entry.get()
            if password:
                if password != confirm_entry.get():
                    messagebox.showerror("Erro", "As senhas não coincidem")
                    return
                updates['password'] = password
            
            # Obtém role_key do valor selecionado
            role_key = [k for k, v in UserManager.ROLES.items() if v == role_var.get()][0]
            if role_key != user.role:
                # Verifica se não está tentando alterar o admin
                if user.username == 'admin' and role_key != 'admin':
                    messagebox.showerror("Erro", "Não é possível alterar o cargo do usuário admin")
                    return
                updates['role'] = role_key
            
            # Coleta permissões selecionadas
            selected_permissions = [
                perm for perm, var in perm_vars.items() 
                if var.get()
            ]
            if set(selected_permissions) != set(user.permissions):
                # Verifica permissões mínimas para admin
                if user.username == 'admin' and not all(perm in selected_permissions for perm in ['manage_users']):
                    messagebox.showerror("Erro", "O usuário admin deve manter a permissão de gerenciar usuários")
                    return
                updates['permissions'] = selected_permissions
            
            if not updates:
                messagebox.showinfo("Aviso", "Nenhuma alteração realizada")
                return
            
            # Tenta atualizar o usuário
            if self.user_manager.update_user(user.id, updates):
                messagebox.showinfo("Sucesso", "Usuário atualizado com sucesso!")
                if after_edit:
                    after_edit()
                window.destroy()
            else:
                messagebox.showerror("Erro", "Não foi possível atualizar o usuário")
        
        # Botões de ação
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill='x', pady=20)
        
        ttk.Button(
            button_frame,
            text="Salvar Alterações",
            command=validate_and_update
        ).pack(side='left', padx=5)
        
        ttk.Button(
            button_frame,
            text="Cancelar",
            command=window.destroy
        ).pack(side='right', padx=5)
        
        self._center_window(window)

    def logout(self, window):
        """Realiza logout do usuário"""
        if messagebox.askyesno("Logout", "Deseja realmente sair?"):
            self.user_manager.current_user = None
            window.destroy()
            self.show_login()

    def _center_window(self, window):
        """Centraliza uma janela na tela"""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    # Métodos auxiliares para configuração das abas
    def setup_reports_tab(self, frame):
        """Configura a aba de relatórios"""
        ttk.Label(
            frame,
            text="Relatórios disponíveis",
            style='Header.TLabel'
        ).pack(pady=20)
        
        # Aqui você pode adicionar a interface específica para relatórios
        ttk.Label(
            frame,
            text="Interface de relatórios em desenvolvimento",
            style='Normal.TLabel'
        ).pack()

    def setup_patients_tab(self, frame, permissions):
        """Configura a aba de pacientes"""
        ttk.Label(
            frame,
            text="Gerenciamento de Pacientes",
            style='Header.TLabel'
        ).pack(pady=20)
        
        # Interface específica para gerenciamento de pacientes
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10)
        
        if 'add_patients' in permissions:
            ttk.Button(
                button_frame,
                text="Adicionar Paciente",
                command=lambda: messagebox.showinfo("Info", "Função em desenvolvimento")
            ).pack(side='left', padx=5)
        
        if 'edit_patients' in permissions:
            ttk.Button(
                button_frame,
                text="Editar Paciente",
                command=lambda: messagebox.showinfo("Info", "Função em desenvolvimento")
            ).pack(side='left', padx=5)
        
        if 'delete_patients' in permissions:
            ttk.Button(
                button_frame,
                text="Excluir Paciente",
                command=lambda: messagebox.showinfo("Info", "Função em desenvolvimento")
            ).pack(side='left', padx=5)

    def setup_financial_tab(self, frame):
        """Configura a aba financeira"""
        ttk.Label(
            frame,
            text="Controle Financeiro",
            style='Header.TLabel'
        ).pack(pady=20)
        
        # Interface específica para controle financeiro
        ttk.Label(
            frame,
            text="Interface financeira em desenvolvimento",
            style='Normal.TLabel'
        ).pack()