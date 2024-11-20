from tkinter import *
from tkinter import ttk
from funcoes_botoes import FuncoesBotoes, SistemaContas, GerenciadorPlanilhas
from planilhas import Planilhas
from tkcalendar import DateEntry
from banco import DataBaseMarcacao
from config import config_manager
from frames.ntfs_fame import EmitirNota

class MainFrame(Frame):
    """Frame principal da aplicação que gerencia a interface do usuário e suas interações."""

    def __init__(self, master, planilhas: Planilhas, file_path: str, app):
        self.ui_config = config_manager.get_config("UI_CONFIG")
        self.app_config = config_manager.get_config("APP_CONFIG")
        super().__init__(master, bg=self.ui_config["colors"]["background"])
        self.configure_window()
        self._init_attributes(master, planilhas, file_path, app)
        self._setup_styles()
        self.create_widgets()


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
                ("Adicionar Informações", self.adicionar_informacao),
                ("Excluir Informação", self.excluir_informacao),
                ("Exibir Informações", self.exibir),
                ('nota', self.mostrar)
            ]),
            ("Agenda e Marcações", 0, 1, [
                ("Nova Marcação", self.marcar_paciente),
                ("Visualizar Marcações", self.visu_marcacoes)
            ]),
            ("Gestão Financeira", 1, 0, [
                ("Valores Atendimento", self.resultados_consulta),
                ("Exibir Contas", self.exibir_contas),
                ("Fechamento Contas", self.fechamento_contas)
            ]),
            ("Documentos e Relatórios", 1, 1, [
                ("Emitir NTFS-e", self.emitir_notas),
                ("Enviar Relatório WhatsApp", self.relatorio_wpp),
                ("Enviar Relatório Email", self.relatorio_email)
            ]),
            ("Ferramentas", 2, 0, [
                ("Gerenciar Planilhas/Sheets", self.planilha_sheet)
            ], 2)
        ]

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

            for btn_text, btn_command in buttons:
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

    # Métodos de ação permanecem os mesmos
    def adicionar_informacao(self): self.funcoes_botoes.adicionar_informacao()
    def excluir_informacao(self): self.funcoes_botoes.excluir()
    def exibir(self): self.funcoes_botoes.exibir_informacao()
    def exibir_contas(self): self.funcoes_botoes.valores_totais()
    def emitir_notas(self): self.funcoes_botoes.processar_notas_fiscais()
    def resultados_consulta(self): self.funcoes_botoes.mostrar_valores_atendimentos()
    def relatorio_wpp(self): self.funcoes_botoes.enviar_whatsapp()
    def relatorio_email(self): self.funcoes_botoes.enviar_email()
    def marcar_paciente(self): self.banco.add_user()
    def visu_marcacoes(self): self.banco.view_marcacoes()
    def fechamento_contas(self): self.sistema_contas.abrir_janela()
    def planilha_sheet(self): self.gerenciador_planilhas.abrir_gerenciador()
    def mostrar(self): self.emitir_nota.show()