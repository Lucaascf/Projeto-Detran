import tkinter as tk

from tkinter import ttk, Label, Button, Entry
from banco import DataBaseLogin
from tkinter import messagebox

class LoginFrame(tk.Frame):
    def __init__(self, master, on_login_success, funcoes_botoes=None):
        super().__init__(master, bg="#2C3E50")  # Definindo o fundo do frame principal
        self.on_login_success = on_login_success
        self.funcoes_botoes = funcoes_botoes
        self.db = DataBaseLogin()
        self.current_user = None
        self.create_widgets()

    def create_widgets(self):
        self.master.title("Login")
        self.master.configure(bg="#2C3E50")
        self.grid_configure(padx=20, pady=20)

        # Título
        title_label = tk.Label(self, text="Faça Login", font=("Arial", 18, "bold"), bg="#2C3E50", fg="#ECF0F1")
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Campos de entrada
        user_label = tk.Label(self, text="Usuário:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 12))
        user_label.grid(row=1, column=0, pady=10, padx=5, sticky="e")
        self.entry_user = tk.Entry(self, font=("Arial", 12), relief=tk.FLAT)
        self.entry_user.grid(row=1, column=1, pady=10, padx=5)

        password_label = tk.Label(self, text="Senha:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 12))
        password_label.grid(row=2, column=0, pady=10, padx=5, sticky="e")
        self.entry_password = tk.Entry(self, show="*", font=("Arial", 12), relief=tk.FLAT)
        self.entry_password.grid(row=2, column=1, pady=10, padx=5)

        # Botões
        button_frame = tk.Frame(self, bg="#2C3E50")  # Fundo consistente no button frame
        button_frame.grid(row=3, columnspan=2, pady=20)

        login_button = tk.Button(button_frame, text="Login", command=self.login, font=("Arial", 12), bg="#34495E", fg="#ECF0F1", relief=tk.FLAT, activebackground="#485460")
        login_button.pack(side="left", padx=10)

        if self.funcoes_botoes:
            create_account_button = tk.Button(button_frame, text="Criar Conta", command=self.funcoes_botoes.mostrar_criar_conta, font=("Arial", 12), bg="#34495E", fg="#ECF0F1", relief=tk.FLAT, activebackground="#485460")
            create_account_button.pack(side="left", padx=10)

    def login(self):
        user = self.entry_user.get()
        password = self.entry_password.get()

        if user and password:
            if self.db.validate_user(user, password):
                self.current_user = user
                self.on_login_success()
            else:
                messagebox.showerror("Erro", "Usuário ou senha incorretos")
                self.entry_user.delete(0, 'end')
                self.entry_password.delete(0, 'end')

class CriarContaFrame(tk.Frame):
    def __init__(self, master, db, funcoes_botoes=None):
        super().__init__(master, bg="#2C3E50")  # Define fundo do frame principal
        self.db = db
        self.funcoes_botoes = funcoes_botoes
        self.create_widgets()
        self.current_user = None

    def create_widgets(self):
        self.master.title("Criar Conta")
        self.master.configure(bg="#2C3E50")
        self.grid_configure(padx=20, pady=20)

        title_label = tk.Label(self, text="Criar Conta", font=("Arial", 18, "bold"), bg="#2C3E50", fg="#ECF0F1")
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        user_label = tk.Label(self, text="Novo Usuário:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 12))
        user_label.grid(row=1, column=0, pady=10, padx=5, sticky="e")
        self.entry_user = tk.Entry(self, font=("Arial", 12), relief=tk.FLAT)
        self.entry_user.grid(row=1, column=1, pady=10, padx=5)

        password_label = tk.Label(self, text="Senha:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 12))
        password_label.grid(row=2, column=0, pady=10, padx=5, sticky="e")
        self.entry_password = tk.Entry(self, show="*", font=("Arial", 12), relief=tk.FLAT)
        self.entry_password.grid(row=2, column=1, pady=10, padx=5)

        button_frame = tk.Frame(self, bg="#2C3E50")
        button_frame.grid(row=3, columnspan=2, pady=20)

        create_button = tk.Button(button_frame, text="Criar Conta", command=self.criar_conta, font=("Arial", 12), bg="#34495E", fg="#ECF0F1", relief=tk.FLAT, activebackground="#485460")
        create_button.pack(padx=10)

    def create_account(self):
        user = self.entry_user.get()
        password = self.entry_password.get()

        if user and password:
            if self.db.create_user(user, password):
                messagebox.showinfo("Sucesso", "Conta criada com sucesso!")
                if self.funcoes_botoes:
                    self.funcoes_botoes.voltar_para_login()
            else:
                messagebox.showerror("Erro", "Usuário já existe")
                self.entry_user.delete(0, 'end')
                self.entry_password.delete(0, 'end')
        else:
            self.current_user = user
            messagebox.showinfo("Sucesso", f"Usuário {user} criado com sucesso.")

    def get_current_user(self):
        return self.current_user

    def voltar_login(self):
        self.funcoes_botoes.voltar_para_login()




class CriarContaFrame(tk.Frame):
    """Classe que representa o frame de criação de conta da aplicação."""

    def __init__(self, master, db, funcoes_botoes=None):
        """Inicializa a classe CriarContaFrame.

        Args:
            master: Janela pai do Tkinter.
            db: Instância da classe DataBase para gerenciar o banco de dados.
            funcoes_botoes: Instância da classe FuncoesBotoes para acessar suas funções.
        """
        super().__init__(master)  # Chama o construtor da classe Frame
        self.db = db  # Armazena a referência do banco de dados
        self.funcoes_botoes = funcoes_botoes  # Armazena a referência para FuncoesBotoes
        self.create_widgets()  # Cria os widgets da interface
        self.current_user = None

    def create_widgets(self):
        """Cria e organiza os widgets do frame de criação de conta."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Título do frame de criação de conta
        Label(self, text="Criar Conta", font=("Arial", 18, "bold"), bg="#2C3E50", fg="#ECF0F1").grid(row=0, column=0, columnspan=2, pady=10)
        
        # Campo para entrada do novo nome de usuário
        Label(self, text="Novo Usuário:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 11, "bold")).grid(row=1, column=0, pady=10, padx=5, sticky="e")
        self.entry_user = Entry(self)
        self.entry_user.grid(row=1, column=1, pady=10, padx=5)

        # Campo para entrada da nova senha
        Label(self, text="Senha:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 11, "bold")).grid(row=2, column=0, pady=10, padx=5, sticky="e")
        self.entry_password = Entry(self, show="*")
        self.entry_password.grid(row=2, column=1, pady=10, padx=5)

        # Botão para criar a conta
        Button(self, text="Criar Conta", command=self.create_account).grid(row=3, column=0, pady=10)
        Button(self, text='Voltar', command=self.voltar_login).grid(row=3, column=1, pady=10)

    def create_account(self):
        """Cria uma nova conta de usuário no banco de dados."""
        user = self.entry_user.get().strip()  # Obtém o nome de usuário
        password = self.entry_password.get().strip()  # Obtém a senha

    # Verifica se o nome de usuário e a senha foram fornecidos
        if not user or not password:
            messagebox.showerror("Erro", "Preencha todos os campos.")
            self.entry_user.delete(0, 'end')  # Limpa o campo de usuário
            self.entry_password.delete(0, 'end')  # Limpa o campo de senha
            return

        if not self.db.create_user(user, password):  # Tenta criar o usuário
            messagebox.showerror("Erro", "Usuário já existente.")
            self.entry_user.delete(0, 'end')  # Limpa o campo de usuário
            self.entry_password.delete(0, 'end')  # Limpa o campo de senha
        else:
            self.current_user = user # Armazena o nome do usuário criado
            messagebox.showinfo("Sucesso", f"Usuário {user} criado com sucesso.")
    
    def get_current_user(self):
        """Retorna o nome do usuário atual."""
        return self.current_user

    def voltar_login(self):
        self.funcoes_botoes.voltar_para_login()
