from tkinter import *
from tkinter import ttk
from funcoes_botoes import FuncoesBotoes, SistemaContas, GerenciadorPlanilhas
from planilhas import Planilhas
from tkcalendar import DateEntry
from banco import DataBaseMarcacao

class MainFrame(Frame):
    """Classe que representa o frame principal da aplicação, responsável por gerenciar as interações do usuário."""

    def __init__(self, master, planilhas: Planilhas, file_path: str, app):
        super().__init__(master, bg=master.cget('bg'))
        # Configurar tamanho e posição da janela principal
        self.configure_window()

        self._init_attributes(master, planilhas, file_path, app)
        self._setup_styles()
        self.create_widgets()

    def configure_window(self):
        """Configura o tamanho e posição inicial da janela."""
        # Definir tamanho inicial
        self.master.geometry("1200x600")
        
        # Definir tamanhos mínimo e máximo
        self.master.minsize(1000, 600)
        self.master.maxsize(1600, 1000)
        
        # Centralizar a janela
        self.center_window()
        
        # Configurar comportamento de redimensionamento
        self.master.resizable(True, True)

    def center_window(self):
        """Centraliza a janela na tela."""
        # Atualiza a janela para garantir dimensões corretas
        self.master.update_idletasks()
        
        # Obtém dimensões da tela
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        
        # Obtém dimensões da janela
        window_width = self.master.winfo_width()
        window_height = self.master.winfo_height()
        
        # Calcula posição para centralizar
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # Define a posição da janela
        self.master.geometry(f"+{x}+{y}")
    
    def toggle_fullscreen(self, event=None):
        """Alterna entre tela cheia e tamanho normal."""
        is_fullscreen = self.master.attributes('-fullscreen')
        self.master.attributes('-fullscreen', not is_fullscreen)
    
    def maximize_window(self, event=None):
        """Maximiza a janela usando dimensões da tela."""
        try:
            # Primeiro, tenta usar o método específico do sistema
            import sys
            if sys.platform == "win32":
                self.master.state('zoomed')
            else:
                self.master.attributes('-zoomed', True)
        except:
            # Se falhar, usa o método universal
            w = self.master.winfo_screenwidth()
            h = self.master.winfo_screenheight()
            # Compensa a barra de título e bordas
            h_adjusted = h - 60  # Ajuste para barra de título e barra de tarefas
            self.master.geometry(f"{w}x{h_adjusted}+0+0")

    def restore_window(self, event=None):
        """Restaura a janela ao tamanho normal."""
        try:
            # Primeiro, tenta usar o método específico do sistema
            import sys
            if sys.platform == "win32":
                self.master.state('normal')
            else:
                self.master.attributes('-zoomed', False)
        except:
            # Se falhar, usa o método universal
            self.master.geometry("1200x800")
        
        # Em qualquer caso, centraliza a janela
        self.center_window()

    def _init_attributes(self, master, planilhas, file_path, app):
        """Inicializa os atributos da classe."""
        self.current_user = app.get_current_user() if hasattr(app, 'get_current_user') else None
        self.funcoes_botoes = FuncoesBotoes(master, planilhas, file_path, app)
        self.banco = DataBaseMarcacao(master, planilhas, file_path, app)
        self.sistema_contas = SistemaContas(file_path, current_user=self.current_user)
        self.gerenciador_planilhas = GerenciadorPlanilhas(master, self.sistema_contas)
        self.master = master
        self.file_path = file_path
        self.app = app

        # Cores temáticas
        self.colors = {
            'bg': master.cget('bg'),
            'fg': '#ECF0F1',
            'button_primary': '#2980b9',
            'button_secondary': '#27ae60',
            'button_warning': '#e67e22',
            'button_danger': '#c0392b',
            'button_hover': '#3498db',
            'frame_bg': '#34495e'
        }

        # Adiciona apenas o atalho para maximizar
        self.master.bind('<F10>', self.maximize_window)

    def _setup_styles(self):
        """Configura os estilos dos widgets."""
        style = ttk.Style()
        
        # Estilo para LabelFrame
        style.configure(
            'Custom.TLabelframe',
            background=self.colors['frame_bg'],
            foreground=self.colors['fg'],
            padding=10
        )
        
        # Estilo para Label do LabelFrame
        style.configure(
            'Custom.TLabelframe.Label',
            background=self.colors['frame_bg'],
            foreground=self.colors['fg'],
            font=('Arial', 12, 'bold')
        )

    def create_button(self, parent, text, command, color=None, width=20):
        """Cria um botão customizado com hover effect."""
        if color is None:
            color = self.colors['button_primary']
            
        btn = Button(
            parent,
            text=text,
            command=command,
            width=width,
            bg=color,
            fg=self.colors['fg'],
            font=('Arial', 10),
            relief='raised',
            borderwidth=1,
            cursor='hand2'
        )
        
        # Hover effects
        btn.bind('<Enter>', lambda e, b=btn: self._on_enter(b))
        btn.bind('<Leave>', lambda e, b=btn, c=color: self._on_leave(b, c))
        
        return btn

    def _on_enter(self, btn):
        """Efeito quando o mouse passa sobre o botão."""
        btn.config(bg=self.colors['button_hover'])

    def _on_leave(self, btn, original_color):
        """Efeito quando o mouse sai do botão."""
        btn.config(bg=original_color)

    def create_widgets(self):
        """Cria e organiza os widgets na interface principal."""
        # Container principal com padding
        main_container = Frame(self, bg=self.colors['bg'])
        main_container.pack(expand=True, fill='both', padx=20, pady=20)
        
        # Título
        title_text = "Gerenciamento de Pacientes"
        if self.current_user:
            title_text += f" - {self.current_user}"
            
        title_label = Label(
            main_container,
            text=title_text,
            font=('Arial', 18, 'bold'),
            bg=self.colors['bg'],
            fg=self.colors['fg']
        )
        title_label.pack(pady=(0, 20))

        # Frame para organizar as seções em grid
        grid_frame = Frame(main_container, bg=self.colors['bg'])
        grid_frame.pack(expand=True, fill='both')
        
        # Configurar grid
        for i in range(2):  # 2 colunas
            grid_frame.grid_columnconfigure(i, weight=1)

        # 1. Seção de Cadastro e Gestão
        cadastro_frame = ttk.LabelFrame(
            grid_frame,
            text=" Cadastro e Gestão ",
            style='Custom.TLabelframe'
        )
        cadastro_frame.grid(row=0, column=0, padx=10, pady=5, sticky='nsew')
        
        self.create_button(
            cadastro_frame,
            'Adicionar Informações',
            self.adicionar_informacao,
            self.colors['button_primary']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            cadastro_frame,
            'Excluir Informação',
            self.excluir_informacao,
            self.colors['button_danger']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            cadastro_frame,
            'Exibir Informações',
            self.exibir,
            self.colors['button_secondary']
        ).pack(pady=5, padx=10, fill='x')

        # 2. Seção de Marcações
        marcacao_frame = ttk.LabelFrame(
            grid_frame,
            text=" Agenda e Marcações ",
            style='Custom.TLabelframe'
        )
        marcacao_frame.grid(row=0, column=1, padx=10, pady=5, sticky='nsew')
        
        self.create_button(
            marcacao_frame,
            'Nova Marcação',
            self.marcar_paciente,
            self.colors['button_primary']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            marcacao_frame,
            'Visualizar Marcações',
            self.visu_marcacoes,
            self.colors['button_secondary']
        ).pack(pady=5, padx=10, fill='x')

        # 3. Seção Financeira
        financeiro_frame = ttk.LabelFrame(
            grid_frame,
            text=" Gestão Financeira ",
            style='Custom.TLabelframe'
        )
        financeiro_frame.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')
        
        self.create_button(
            financeiro_frame,
            'Valores Atendimento',
            self.resultados_consulta,
            self.colors['button_primary']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            financeiro_frame,
            'Exibir Contas',
            self.exibir_contas,
            self.colors['button_secondary']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            financeiro_frame,
            'Fechamento Contas',
            self.fechamento_contas,
            self.colors['button_warning']
        ).pack(pady=5, padx=10, fill='x')

        # 4. Seção de Documentos e Relatórios
        documentos_frame = ttk.LabelFrame(
            grid_frame,
            text=" Documentos e Relatórios ",
            style='Custom.TLabelframe'
        )
        documentos_frame.grid(row=1, column=1, padx=10, pady=5, sticky='nsew')
        
        self.create_button(
            documentos_frame,
            'Emitir NTFS-e',
            self.emitir_notas,
            self.colors['button_primary']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            documentos_frame,
            'Enviar Relatório WhatsApp',
            self.relatorio_wpp,
            self.colors['button_secondary']
        ).pack(pady=5, padx=10, fill='x')
        
        self.create_button(
            documentos_frame,
            'Enviar Relatório Email',
            self.relatorio_email,
            self.colors['button_warning']
        ).pack(pady=5, padx=10, fill='x')

        # 5. Seção de Ferramentas
        tools_frame = ttk.LabelFrame(
        grid_frame,
        text=" Ferramentas ",
        style='Custom.TLabelframe'
    )
        tools_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky='nsew')
        
        tools_button_frame = Frame(tools_frame, bg=self.colors['frame_bg'])
        tools_button_frame.pack(fill='x', expand=True)
        
        self.create_button(
            tools_button_frame,
            'Formatar Planilha',
            self.format_planilha,
            self.colors['button_primary']
        ).pack(side='left', pady=5, padx=10, expand=True)
        
        self.create_button(
            tools_button_frame,
            'Gerenciar Planilhas/Sheets',
            self.planilha_sheet,
            self.colors['button_secondary']
        ).pack(side='left', pady=5, padx=10, expand=True)

    # Os métodos de ação permanecem os mesmos
    def adicionar_informacao(self): self.funcoes_botoes.adicionar_informacao()
    def excluir_informacao(self): self.funcoes_botoes.excluir()
    def exibir(self): self.funcoes_botoes.exibir_informacao()
    def exibir_contas(self): self.funcoes_botoes.valores_totais()
    def emitir_notas(self): self.funcoes_botoes.processar_notas_fiscais()
    def resultados_consulta(self): self.funcoes_botoes.exibir_resultado()
    def relatorio_wpp(self): self.funcoes_botoes.enviar_whatsapp()
    def relatorio_email(self): self.funcoes_botoes.enviar_email()
    def marcar_paciente(self): self.banco.add_user()
    def visu_marcacoes(self): self.banco.view_marcacoes()
    def format_planilha(self): self.funcoes_botoes.formatar_planilha()
    def fechamento_contas(self): self.sistema_contas.abrir_janela()
    def planilha_sheet(self): self.gerenciador_planilhas.abrir_gerenciador()