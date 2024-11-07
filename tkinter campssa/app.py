# app.py
from tkinter import *
from tkinter import filedialog, messagebox
import logging
from frames.login_frame import LoginFrame, CriarContaFrame
from frames.main_frame import MainFrame
from funcoes_botoes import FuncoesBotoes
from planilhas import Planilhas
from banco import DataBaseLogin
from config import config_manager
from datetime import datetime


class App(Frame):
    """Classe principal do aplicativo, responsável por gerenciar a interface gráfica e a lógica do aplicativo."""

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self._init_attributes()
        self._setup_logging()

        # Obtém as configurações do config_manager
        self.app_config = config_manager.get_config("APP_CONFIG")
        self.ui_config = config_manager.get_config("UI_CONFIG")

        self.setup_ui()
        self.frames_da_tela()
        self.grid()

    """Inicializa os atributos da classe."""

    def _init_attributes(self):
        """Inicializa os atributos da classe."""
        self.planilhas = None
        self.file_path = None
        self.db = DataBaseLogin()
        # Remova o current_user da inicialização
        self.funcoes_botoes = FuncoesBotoes(
            master=self.master,
            planilhas=self.planilhas,
            file_path=self.file_path,
            app=self,
        )
        self.current_user = None
        self.login_frame = None
        self.criar_conta_frame = None
        self.main_frame = None

    """Configura o sistema de logging."""

    def _setup_logging(self):
        """Configura o sistema de logging."""
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[logging.FileHandler("app.log"), logging.StreamHandler()],
        )
        self.logger = logging.getLogger(__name__)

    """Configura as propriedades da janela principal do aplicativo."""

    def setup_ui(self):
        """Configura as propriedades da janela principal do aplicativo."""
        self.master.title(self.app_config["title"])
        self.master.configure(background=self.ui_config["colors"]["background"])
        self.master.geometry(self.app_config["initial_geometry"])
        self.center()

    """Cria e organiza os frames de login e de criação de conta."""

    def frames_da_tela(self):
        """Cria e organiza os frames de login e de criação de conta."""
        self._configure_grid()
        self._create_login_frame()
        self._create_signup_frame()
        self.funcoes_botoes.configurar_frames(self.login_frame, self.criar_conta_frame)

    """Configura o grid da janela principal."""

    def _configure_grid(self):
        """Configura o grid da janela principal."""
        self.master.grid_rowconfigure(0, weight=1)
        self.master.grid_columnconfigure(0, weight=1)

    """Cria e configura o frame de login."""

    def _create_login_frame(self):
        """Cria e configura o frame de login."""
        self.login_frame = LoginFrame(
            self.master, self.login_success, self.funcoes_botoes
        )
        # Usar self.app_config em vez de APP_CONFIG
        self.login_frame.configure(bg=self.app_config["background_color"])
        self.login_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)

    """Cria e configura o frame de criação de conta."""

    def _create_signup_frame(self):
        """Cria e configura o frame de criação de conta."""
        self.criar_conta_frame = CriarContaFrame(
            self.master, self.db, self.funcoes_botoes
        )
        # Usar self.app_config em vez de APP_CONFIG
        self.criar_conta_frame.configure(bg=self.app_config["background_color"])
        self.criar_conta_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.criar_conta_frame.grid_forget()

    """Gerencia o processo após um login bem-sucedido."""

    def login_success(self):
        """Gerencia o processo após um login bem-sucedido."""
        try:
            self.current_user = self.login_frame.current_user
            # Atualiza o current_user nas funcoes_botoes
            self.funcoes_botoes.set_current_user(self.current_user)
            self.login_frame.grid_forget()
            self.master.geometry(self.app_config["main_geometry"])
            self.center()
            self.open_file()
            self.logger.info(f"Login bem-sucedido para o usuário: {self.current_user}")
        except Exception as e:
            self.logger.error(f"Erro no processo de login: {str(e)}")
            messagebox.showerror("Erro", f"Erro no processo de login: {str(e)}")

    """Retorna o usuário atual em maiúsculas."""

    def get_current_user(self):
        """Retorna o usuário atual em maiúsculas."""
        return self.current_user.upper() if self.current_user else None

    """Abre um diálogo para seleção de arquivos e carrega as planilhas."""

    def open_file(self):
        """Abre um diálogo para seleção de arquivos e carrega as planilhas."""
        try:
            self.center()
            file_path = filedialog.askopenfilename(
                title="Selecionar Planilha",
                # Usar self.app_config em vez de APP_CONFIG
                filetypes=self.app_config["file_types"],
            )

            if file_path:
                self._handle_file_selection(file_path)
            else:
                self._handle_no_file_selected()

        except Exception as e:
            self.logger.error(f"Erro ao abrir arquivo: {str(e)}")
            messagebox.showerror("Erro ao abrir o arquivo", str(e))

    """Processa a seleção de arquivo."""

    def _handle_file_selection(self, file_path):
        """Processa a seleção de arquivo."""
        try:
            self.file_path = file_path
            self.planilhas = Planilhas(self.file_path)
            self.funcoes_botoes.planilhas = self.planilhas
            self.logger.info("Arquivo selecionado com sucesso")
            self._create_main_frame()
        except Exception as e:
            self.logger.error(f"Erro ao processar arquivo: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao processar arquivo: {str(e)}")

    """Gerencia o caso onde nenhum arquivo foi selecionado."""

    def _handle_no_file_selected(self):
        """Gerencia o caso onde nenhum arquivo foi selecionado."""
        if messagebox.askyesno(
            "Confirmação", "Você não selecionou nenhum arquivo. Deseja realmente sair?"
        ):
            self.master.quit()
        else:
            self.login_frame.grid(row=0, column=0, pady=150)
            self.open_file()

    """Cria e configura o frame principal."""

    def _create_main_frame(self):
        """Cria e configura o frame principal."""
        self.main_frame = MainFrame(self.master, self.planilhas, self.file_path, self)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.login_frame.grid_forget()

    """Centraliza a janela na tela."""

    def center(self):
        """Centraliza a janela na tela."""
        self.update_idletasks()
        width = self.master.winfo_width()
        height = self.master.winfo_height()
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.master.geometry(f"{width}x{height}+{x}+{y}")
        self.master.deiconify()

    """Mostra o frame de criação de conta."""

    def mostrar_criar_conta(self):
        """Mostra o frame de criação de conta."""
        self.login_frame.grid_forget()
        self.criar_conta_frame.grid()

    """Retorna ao frame de login."""

    def voltar_para_login(self):
        """Retorna ao frame de login."""
        self.criar_conta_frame.grid_forget()
        self.login_frame.grid()


if __name__ == "__main__":
    try:
        root = Tk()
        app = App(master=root)
        app.mainloop()
    except Exception as e:
        logging.error(f"Erro fatal na aplicação: {str(e)}")
        messagebox.showerror("Erro Fatal", f"Um erro fatal ocorreu: {str(e)}")
