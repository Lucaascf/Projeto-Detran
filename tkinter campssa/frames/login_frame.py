import tkinter as tk
import re
import logging
from tkinter import messagebox
from banco import DataBaseLogin
from config import ConfigManager

class BaseFrame(tk.Frame):
    def __init__(self, master):
        self.config_manager = ConfigManager()
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
    def __init__(self, master, on_login_success, funcoes_botoes=None):
        super().__init__(master)
        self.on_login_success = on_login_success
        self.funcoes_botoes = funcoes_botoes
        self.db = DataBaseLogin()
        self.current_user = None
        self._create_widgets()

    def _create_widgets(self):
        ui_config = self.config_manager.get_config('UI_CONFIG')
        colors = ui_config['colors']
        fonts = ui_config['fonts']

        container = tk.Frame(self, bg=colors['background'])
        container.grid(row=1, column=0)

        # Title
        tk.Label(
            container,
            text="Faça Login",
            font=fonts['title'],
            bg=colors['background'],
            fg=colors['title']
        ).grid(row=0, column=0, columnspan=2, pady=(0, 30))

        # Fields
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

        # Buttons
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
        if self.db.validate_user(user, password):
            self.current_user = user
            self.on_login_success()
        else:
            messagebox.showerror("Erro", "Usuário ou senha incorretos")
            self.entry_user.delete(0, tk.END)
            self.entry_password.delete(0, tk.END)
            self.entry_user.focus()

class CriarContaFrame(BaseFrame):
    def __init__(self, master, db, funcoes_botoes=None):
        super().__init__(master)
        self.db = db
        self.funcoes_botoes = funcoes_botoes
        self.current_user = None
        self.validation_rules = {
            "user": {"min_length": 3, "max_length": 20, "pattern": r"^[a-zA-Z0-9_]+$"},
            "password": {"min_length": 6, "max_length": 20, "pattern": r"^[a-zA-Z0-9@#$%^&+=]+$"}
        }
        self._create_widgets()

    def _create_widgets(self):
        ui_config = self.config_manager.get_config('UI_CONFIG')
        colors = ui_config['colors']
        fonts = ui_config['fonts']

        container = tk.Frame(self, bg=colors['background'])
        container.grid(row=1, column=0)

        # Title
        tk.Label(
            container,
            text="Criar Nova Conta",
            font=fonts['title'],
            bg=colors['background'],
            fg=colors['title']
        ).grid(row=0, column=0, columnspan=2, pady=(0, 30))

        # Fields
        self.entry_widgets = {}
        fields = [("user", "Usuário:"), ("password", "Senha:")]
        for i, (field_name, label_text) in enumerate(fields, 1):
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
                show="*" if field_name == "password" else "",
                width=20,
                bg=colors['frame'],
                fg=colors['text'],
                insertbackground=colors['text']
            )
            entry.grid(row=i, column=1, pady=5, sticky="w")
            self.entry_widgets[field_name] = entry
            entry.bind("<Return>", lambda e: self.create_account())

        # Buttons
        button_frame = tk.Frame(container, bg=colors['background'])
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

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

        # Requirements
        tk.Label(
            container,
            text="Requisitos:\n• Usuário: 3-20 caracteres\n• Senha: 6-20 caracteres",
            font=fonts['small'],
            bg=colors['background'],
            fg=colors['text'],
            justify="left"
        ).grid(row=4, column=0, columnspan=2, pady=20)

    def _validate_fields(self):
        for field_name, entry in self.entry_widgets.items():
            value = entry.get().strip()
            rules = self.validation_rules[field_name]
            if len(value) < rules["min_length"]:
                raise ValueError(f"{field_name.title()} deve ter pelo menos {rules['min_length']} caracteres")
            if len(value) > rules["max_length"]:
                raise ValueError(f"{field_name.title()} deve ter no máximo {rules['max_length']} caracteres")
            if not re.match(rules["pattern"], value):
                raise ValueError(f"{field_name.title()} contém caracteres inválidos")
        return True

    def create_account(self):
        try:
            if self._validate_fields():
                user = self.entry_widgets["user"].get().strip()
                password = self.entry_widgets["password"].get().strip()
                if self.db.create_user(user, password):
                    self.current_user = user
                    messagebox.showinfo("Sucesso", f"Conta criada com sucesso para {user}!")
                    self.clear_fields()
                    self.voltar_login()
                else:
                    messagebox.showerror("Erro", "Nome de usuário já existe.")
                    self.entry_widgets["user"].focus()
        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar conta: {str(e)}")
            logging.error(f"Erro na criação de conta: {str(e)}")

    def clear_fields(self):
        for entry in self.entry_widgets.values():
            entry.delete(0, tk.END)
        self.current_user = None

    def voltar_login(self):
        self.clear_fields()
        if self.funcoes_botoes:
            self.funcoes_botoes.voltar_para_login()

    def get_current_user(self):
        return self.current_user