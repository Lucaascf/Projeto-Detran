import tkinter as tk
import re
import logging
from tkinter import messagebox
from banco import DataBaseLogin

class BaseFrame(tk.Frame):
    """Classe base com funcionalidades comuns."""
    
    def __init__(self, master, bg_color="#2C3E50"):
        """Inicializa o frame base."""
        super().__init__(master, bg=bg_color)
        self._configure_window()
        self._setup_base_grid()

    def _configure_window(self):
        """Configura dimensões e posição da janela."""
        width, height = 400, 500
        screen_w = self.master.winfo_screenwidth()
        screen_h = self.master.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        
        self.master.geometry(f"{width}x{height}+{x}+{y}")
        self.master.minsize(width, height)
        self.master.maxsize(width, height)
        self.master.resizable(False, False)
        
    def _setup_base_grid(self):
        """Configura o grid básico do frame."""
        # Configura o frame principal para ocupar todo o espaço
        self.grid(row=0, column=0, sticky="nsew")
        
        # Configura pesos do grid no master
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)
        
        # Configura pesos internos do frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
    def show(self):
        """Exibe o frame."""
        self.grid()
        
    def hide(self):
        """Esconde o frame."""
        self.grid_remove()

class LoginFrame(BaseFrame):
    """Frame de login."""
    
    def __init__(self, master, on_login_success, funcoes_botoes=None):
        super().__init__(master)
        self.on_login_success = on_login_success
        self.funcoes_botoes = funcoes_botoes
        self.db = DataBaseLogin()
        self.current_user = None
        self._create_widgets()

    def _create_widgets(self):  # Para LoginFrame e CriarContaFrame
        # Container principal com configuração de grid
        container = tk.Frame(self, bg=self["bg"])
        container.grid(row=0, column=0)
        
        # Centraliza o container no frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Configura o grid do container
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)
        
        # Título
        tk.Label(
            container,
            text="Faça Login",
            font=("Arial", 24, "bold"),
            bg=self["bg"],
            fg="#ECF0F1"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 30))

        # Campos de entrada
        fields = [
            ("Usuário:", "entry_user", ""),
            ("Senha:", "entry_password", "*")
        ]
        
        for i, (label_text, attr_name, show_char) in enumerate(fields, start=1):
            tk.Label(
                container,
                text=label_text,
                bg=self["bg"],
                fg="#ECF0F1",
                font=("Arial", 12),
                width=8,
                anchor='e'
            ).grid(row=i, column=0, padx=(0, 10), pady=10, sticky='e')
            
            entry = tk.Entry(
                container,
                font=("Arial", 12),
                show=show_char,
                width=20,
                relief=tk.FLAT
            )
            entry.grid(row=i, column=1, pady=10, sticky='w')
            setattr(self, attr_name, entry)

        # Frame para botões
        button_frame = tk.Frame(container, bg=self["bg"])
        button_frame.grid(row=3, column=0, columnspan=2, pady=30)

        button_style = {
            "font": ("Arial", 12),
            "relief": tk.FLAT,
            "width": 12,
            "height": 2
        }

        # Botões
        tk.Button(
            button_frame,
            text="Login",
            command=self.perform_login,
            bg="#2ecc71",
            fg="white",
            activebackground="#27ae60",
            activeforeground="white",
            **button_style
        ).grid(row=0, column=0, padx=5)

        if self.funcoes_botoes:
            tk.Button(
                button_frame,
                text="Criar Conta",
                command=self.funcoes_botoes.mostrar_criar_conta,
                bg="#3498db",
                fg="white",
                activebackground="#2980b9",
                activeforeground="white",
                **button_style
            ).grid(row=0, column=1, padx=5)

    def perform_login(self):
        """Executa o processo de login."""
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
    """Frame de criação de conta."""
    
    def __init__(self, master, db, funcoes_botoes=None):
        super().__init__(master)
        self.db = db
        self.funcoes_botoes = funcoes_botoes
        self.current_user = None
        self.validation_rules = {
            'user': {'min_length': 3, 'max_length': 20, 'pattern': r'^[a-zA-Z0-9_]+$'},
            'password': {'min_length': 6, 'max_length': 20, 'pattern': r'^[a-zA-Z0-9@#$%^&+=]+$'}
        }
        self._create_widgets()

    def _create_widgets(self):  # Para LoginFrame e CriarContaFrame
        # Container principal com configuração de grid
        container = tk.Frame(self, bg=self["bg"])
        container.grid(row=0, column=0)
        
        # Centraliza o container no frame
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # Configura o grid do container
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=1)

        # Título
        tk.Label(
            container,
            text="Criar Nova Conta",
            font=("Arial", 24, "bold"),
            bg=self["bg"],
            fg="#ECF0F1"
        ).grid(row=0, column=0, columnspan=2, pady=(0, 30))

        # Campos de entrada
        self.entry_widgets = {}
        fields = [
            ("user", "Usuário:"),
            ("password", "Senha:")
        ]

        for i, (field_name, label_text) in enumerate(fields, start=1):
            tk.Label(
                container,
                text=label_text,
                bg=self["bg"],
                fg="#ECF0F1",
                font=("Arial", 12),
                width=8,
                anchor='e'
            ).grid(row=i, column=0, padx=(0, 10), pady=10, sticky='e')

            entry = tk.Entry(
                container,
                font=("Arial", 12),
                show="*" if field_name == "password" else "",
                width=20,
                relief=tk.FLAT
            )
            entry.grid(row=i, column=1, pady=10, sticky='w')
            self.entry_widgets[field_name] = entry

        # Frame para botões
        button_frame = tk.Frame(container, bg=self["bg"])
        button_frame.grid(row=3, column=0, columnspan=2, pady=30)

        button_style = {
            "font": ("Arial", 12),
            "relief": tk.FLAT,
            "width": 12,
            "height": 2
        }

        tk.Button(
            button_frame,
            text="Criar Conta",
            command=self.create_account,
            bg="#2ecc71",
            fg="white",
            activebackground="#27ae60",
            activeforeground="white",
            **button_style
        ).grid(row=0, column=0, padx=5)

        tk.Button(
            button_frame,
            text="Voltar",
            command=self.voltar_login,
            bg="#e74c3c",
            fg="white",
            activebackground="#c0392b",
            activeforeground="white",
            **button_style
        ).grid(row=0, column=1, padx=5)

        # Texto de ajuda
        tk.Label(
            container,
            text="Requisitos:\n• Usuário: 3-20 caracteres\n• Senha: 6-20 caracteres",
            bg=self["bg"],
            fg="#95a5a6",
            font=("Arial", 10, "italic"),
            justify="left"
        ).grid(row=4, column=0, columnspan=2, pady=20)

        # Bindings
        for entry in self.entry_widgets.values():
            entry.bind('<Return>', lambda e: self.create_account())
            
    def _validate_fields(self):
        """Valida os campos do formulário."""
        for field_name, entry in self.entry_widgets.items():
            value = entry.get().strip()
            rules = self.validation_rules[field_name]
            
            if len(value) < rules['min_length']:
                raise ValueError(f"{field_name.title()} deve ter pelo menos {rules['min_length']} caracteres")
            if len(value) > rules['max_length']:
                raise ValueError(f"{field_name.title()} deve ter no máximo {rules['max_length']} caracteres")
            if not re.match(rules['pattern'], value):
                raise ValueError(f"{field_name.title()} contém caracteres inválidos")
        
        return True

    def create_account(self):
        """Processa a criação de conta."""
        try:
            if self._validate_fields():
                user = self.entry_widgets['user'].get().strip()
                password = self.entry_widgets['password'].get().strip()
                
                if self.db.create_user(user, password):
                    self.current_user = user
                    messagebox.showinfo("Sucesso", f"Conta criada com sucesso para {user}!")
                    self.clear_fields()
                    self.voltar_login()
                else:
                    messagebox.showerror("Erro", "Nome de usuário já existe.")
                    self.entry_widgets['user'].focus()
                    
        except ValueError as e:
            messagebox.showerror("Erro de Validação", str(e))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar conta: {str(e)}")
            logging.error(f"Erro na criação de conta: {str(e)}")

    def clear_fields(self):
        """Limpa os campos do formulário."""
        for entry in self.entry_widgets.values():
            entry.delete(0, tk.END)
        self.current_user = None

    def voltar_login(self):
        """Retorna à tela de login."""
        self.clear_fields()
        if self.funcoes_botoes:
            self.funcoes_botoes.voltar_para_login()

    def get_current_user(self):
        """Retorna o usuário atual."""
        return self.current_user
