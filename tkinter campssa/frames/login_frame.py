# frames/login_frame.py
from tkinter import *
from banco import DataBaseLogin
from tkinter import messagebox

class LoginFrame(Frame):
    """Classe que representa o frame de login da aplicação."""

    def __init__(self, master, on_login_success, funcoes_botoes=None):
        """Inicializa a classe LoginFrame.

        Args:
            master: Janela pai do Tkinter.
            on_login_success: Função a ser chamada quando o login for bem-sucedido.
            funcoes_botoes: Instância da classe FuncoesBotoes para acessar suas funções.
        """
        super().__init__(master)  # Chama o construtor da classe Frame
        self.on_login_success = on_login_success  # Armazena a função de sucesso do login
        self.funcoes_botoes = funcoes_botoes  # Armazena a referência para FuncoesBotoes
        self.db = DataBaseLogin()
        self.current_user = None
        self.create_widgets()  # Cria os widgets da interface

    def create_widgets(self):
        """Cria e organiza os widgets do frame de login."""
        # Configura as colunas da grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # Título do frame de login
        title_label = Label(self, text="Faça Login", font=("Arial", 18, "bold"),
                            bg="#2C3E50", fg="#ECF0F1")
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Campo para entrada do nome de usuário
        Label(self, text="Usuário:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 11, "bold")).grid(row=1, column=0, pady=10, padx=5, sticky="e")
        self.entry_user = Entry(self)
        self.entry_user.grid(row=1, column=1, pady=10, padx=5)

        # Campo para entrada da senha
        Label(self, text="Senha:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 11, "bold")).grid(row=2, column=0, pady=10, padx=5, sticky="e")
        self.entry_password = Entry(self, show="*")
        self.entry_password.grid(row=2, column=1, pady=10, padx=5)

        # Frame para os botões de login e criação de conta
        button_frame = Frame(self, bg="#2C3E50")
        button_frame.grid(row=3, columnspan=2, pady=20)
        
        # Botão de login
        Button(button_frame, text="Login", command=self.login).pack(side="left", padx=10)

        # Botão para criar conta, se as funções de botões estiverem disponíveis
        Button(button_frame, text="Criar Conta", command=self.funcoes_botoes.mostrar_criar_conta if self.funcoes_botoes else None).pack(side="left", padx=10)

    def login(self):
        """Realiza o processo de login."""
        user = self.entry_user.get()  # Obtém o nome de usuário
        password = self.entry_password.get()  # Obtém a senha

        # Verifica se o nome de usuário e a senha foram fornecidos
        if user and password:
            if self.db.validate_user(user, password):
                self.current_user = user
                self.on_login_success()
            else:
                messagebox.showerror("Erro", "Usuário ou senha incorretos")  # Mensagem de erro para campos vazios
                self.entry_user.delete(0, 'end')
                self.entry_password.delete(0, 'end')


class CriarContaFrame(Frame):
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
