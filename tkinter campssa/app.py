# Importando as bibliotecas necessárias do Tkinter
from tkinter import *
from frames.login_frame import LoginFrame, CriarContaFrame
from frames.main_frame import MainFrame
from tkinter import filedialog, messagebox
from funcoes_botoes import FuncoesBotoes
from planilhas import Planilhas
from banco import DataBase

class App(Frame):
    """Classe principal do aplicativo, responsável por gerenciar a interface gráfica e a lógica do aplicativo."""

    def __init__(self, master=None):
        """Inicializa a classe App.

        Args:
            master: Janela principal do Tkinter.
        """
        super().__init__(master)  # Chama o construtor da classe Frame
        self.master = master
        self.planilhas = None  # Inicializa a variável para armazenar planilhas
        self.file_path = None  # Inicializa a variável para armazenar o caminho do arquivo
        self.db = DataBase()  # Inicializa a conexão com o banco de dados
        self.funcoes_botoes = FuncoesBotoes(master, self.planilhas, self.file_path, self)
        
        self.setup_ui()  # Configura a interface gráfica
        self.frames_da_tela()  # Cria e organiza os frames na tela
        self.grid()  # Adiciona o frame principal à grid

    def setup_ui(self):
        """Configura as propriedades da janela principal do aplicativo."""
        self.master.title('CAMPSSA')  # Título da janela
        self.master.configure(background='#2C3E50')  # Cor de fundo da janela
        self.master.geometry('350x250')  # Tamanho inicial da janela
        self.master.maxsize(width=350, height=250)  # Tamanho máximo da janela
        self.master.minsize(width=350, height=250)  # Tamanho mínimo da janela
        self.center()  # Centraliza a janela na tela

    def frames_da_tela(self):
        """Cria e organiza os frames de login e de criação de conta."""
        self.master.grid_rowconfigure(0, weight=1)  # Configura a row para expandir
        self.master.grid_columnconfigure(0, weight=1)  # Configura a column para expandir

        # Inicializa o frame de login
        self.login_frame = LoginFrame(self.master, self.login_success, self.funcoes_botoes)
        self.login_frame.configure(bg='#2C3E50')  # Define a cor de fundo do frame de login
        self.login_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=20)  # Adiciona o frame à grid

        # Inicializa o frame de criação de conta
        self.criar_conta_frame = CriarContaFrame(self.master, self.db, self.funcoes_botoes)
        self.criar_conta_frame.configure(bg='#2C3E50')  # Define a cor de fundo do frame de criação de conta
        self.criar_conta_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=20)  # Adiciona o frame à grid
        self.criar_conta_frame.grid_forget()  # Esconde o frame de criação de conta inicialmente

        # Configura as funções dos botões para gerenciar os frames
        self.funcoes_botoes.configurar_frames(self.login_frame, self.criar_conta_frame)

    def login_success(self):
        """Método chamado quando o login é bem-sucedido."""
        self.login_frame.grid_forget()  # Esconde o frame de login
        self.master.geometry('700x300')  # Redimensiona a janela
        self.master.maxsize(width=700, height=300)  # Tamanho máximo da janela
        self.master.minsize(width=700, height=300)  # Tamanho mínimo da janela
        self.open_file()  # Abre a função para selecionar um arquivo

    def open_file(self):
        """Abre um diálogo para seleção de arquivos e carrega as planilhas."""
        try:
            self.center()  # Centraliza a janela novamente
            file_path = filedialog.askopenfilename(
                title='Selecionar Planilha',
                filetypes=[('Arquivos Excel', '*.xlsx'), ('Todos os Arquivos', '*.*')]
            )
            if file_path:
                self.file_path = file_path  # Atualiza o caminho do arquivo
                self.planilhas = Planilhas(self.file_path)  # Carrega as planilhas
                
                self.funcoes_botoes.planilhas = self.planilhas  # Passa as planilhas para as funções dos botões
                print("Arquivo selecionado com sucesso")

                # Instancia e mostra o MainFrame
                self.main_frame = MainFrame(self.master, self.planilhas, self.file_path, self)  
                self.main_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=20)  # Adiciona o frame à grid
                self.login_frame.grid_forget()  # Esconde o frame de login
            else:
                messagebox.showwarning("Nenhum Arquivo", "Nenhum arquivo foi selecionado.")
                self.login_frame.grid(row=0, column=0, pady=150)  # Reexibe o frame de login
        except Exception as e:
            print(f"Erro: {e}")
            messagebox.showerror('Erro ao abrir o arquivo', str(e))  # Exibe erro em caso de falha

    def center(self):
        """Centraliza a janela principal na tela."""
        self.update_idletasks()  # Atualiza as tarefas pendentes
        width = self.master.winfo_width()  # Obtém a largura da janela
        height = self.master.winfo_height()  # Obtém a altura da janela
        screen_width = self.master.winfo_screenwidth()  # Obtém a largura da tela
        screen_height = self.master.winfo_screenheight()  # Obtém a altura da tela
        x = (screen_width // 2) - (width // 2)  # Calcula a posição x para centralizar
        y = (screen_height // 2) - (height // 2)  # Calcula a posição y para centralizar
        self.master.geometry(f'{width}x{height}+{x}+{y}')  # Define a nova geometria da janela
        self.master.deiconify()  # Exibe a janela

    def mostrar_criar_conta(self):
        """Mostra o frame de criação de conta e esconde o frame de login."""
        self.login_frame.grid_forget()  # Esconde o frame de login
        self.criar_conta_frame.grid()  # Exibe o frame de criação de conta

    def voltar_para_login(self):
        """Retorna ao frame de login e esconde o frame de criação de conta."""
        self.criar_conta_frame.grid_forget()  # Esconde o frame de criação de conta
        self.login_frame.grid()  # Exibe o frame de login

if __name__ == '__main__':
    root = Tk()  # Cria a janela principal
    app = App(master=root)  # Inicializa a aplicação
    app.mainloop()  # Inicia o loop principal do Tkinter
