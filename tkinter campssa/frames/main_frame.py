# Importando as bibliotecas necessárias do Tkinter
from tkinter import *
from funcoes_botoes import FuncoesBotoes, SistemaContas, GerenciadorPlanilhas
from planilhas import Planilhas
from tkcalendar import DateEntry
from banco import DataBaseMarcacao


class MainFrame(Frame):
    """Classe que representa o frame principal da aplicação, responsável por gerenciar as interações do usuário."""

    def __init__(self, master, planilhas: Planilhas, file_path: str, app, current_user=None):
        """Inicializa a classe MainFrame.

        Args:
            master: Janela pai do Tkinter.
            planilhas: Instância da classe Planilhas para gerenciar as planilhas.
            file_path: Caminho do arquivo que contém os dados.
            app: Instância da aplicação principal.
        """
        super().__init__(master, bg=master.cget('bg'))  # Chama o construtor da classe Frame
        self.current_user = current_user
        self.funcoes_botoes = FuncoesBotoes(master, planilhas, file_path, app, current_user=self.current_user)  # Inicializa FuncoesBotoes
        self.banco = DataBaseMarcacao(master, planilhas, file_path, app)
        self.sistema_contas = SistemaContas(file_path, current_user=self.current_user)
        self.gerenciador_planilhas = GerenciadorPlanilhas(master, self.sistema_contas)  # Instancia GerenciadorPlanilhas
        self.master = master
        self.file_path = file_path
        self.app = app
        self.create_widgets()  # Cria os widgets da interface

    def create_widgets(self):
        """Cria e organiza os widgets na interface principal."""
        # Título da tela
        title_label = Label(self, text="Gerenciamento de Pacientes", font=('Arial', 16, 'bold'), bg=self.cget('bg'), fg='#ECF0F1')
        title_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))  # Adiciona o título à grid

        # Botão para adicionar informações
        self.bt_adicionar_informacoes = Button(
            self, text='Adicionar Informações', command=self.adicionar_informacao)
        self.bt_adicionar_informacoes.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

        # Botão para exibir resultados de consulta
        self.bt_valores_atendimento = Button(
            self, text='Valores Atendimento', command=self.resultados_consulta)
        self.bt_valores_atendimento.grid(row=4, column=1, padx=10, pady=10, sticky='ew')

        self.bt_format_planilha = Button(
            self, text='Format Planilha', command=self.format_planilha)
        self.bt_format_planilha.grid(row=4, column=0, padx=10, pady=10, sticky='ew')

        self.bt_format_planilha = Button(
            self, text='Fechamento contas', command=self.fechamento_contas)
        self.bt_format_planilha.grid(row=4, column=2, padx=10, pady=10, sticky='ew')

        self.bt_format_planilha = Button(
            self, text='planilha ou sheet', command=self.planilha_sheet)
        self.bt_format_planilha.grid(row=5, column=2, padx=10, pady=10, sticky='ew')

        # Botão para excluir informações
        self.bt_excluir_informacao = Button(
            self, text='Excluir Informação', command=self.excluir_informacao)
        self.bt_excluir_informacao.grid(row=1, column=1, padx=10, pady=10, sticky='ew')

        # Botão para exibir informações
        self.bt_exibir_informacoes = Button(
            self, text='Exibir Informações', command=self.exibir)
        self.bt_exibir_informacoes.grid(row=1, column=2, padx=10, pady=10, sticky='ew')

        # Botão para exibir contas
        self.bt_exibir_contas = Button(
            self, text='Exibir Contas', command=self.exibir_contas)
        self.bt_exibir_contas.grid(row=2, column=0, padx=10, pady=10, sticky='ew')

        # Botão para enviar relatório via WhatsApp
        self.bt_enviar_relatorio_wpp = Button(
            self, text='Enviar Relatório Wpp', command=self.relatorio_wpp)
        self.bt_enviar_relatorio_wpp.grid(row=2, column=2, padx=10, pady=10, sticky='ew')

        # Botão para enviar relatório via e-mail
        self.bt_enviar_relatorio_email = Button(
            self, text='Enviar Relatório Email', command=self.relatorio_email)
        self.bt_enviar_relatorio_email.grid(row=2, column=1, padx=10, pady=10, sticky='ew')

        # Botão para emitir notas fiscais
        self.bt_emitir_ntfs = Button(
            self, text='Emitir NTFS-e', command=self.emitir_notas)
        self.bt_emitir_ntfs.grid(row=3, column=0, padx=10, pady=10, sticky='ew')

        self.bt_marcar_paciente = Button(
            self, text='Marcação', command=self.marcar_paciente)
        self.bt_marcar_paciente.grid(row=3, column=1, padx=10, pady=10, sticky='ew')

        self.bt_view_patients = Button(
            self, text='Visualizar Marcações', command=self.visu_marcacoes)
        self.bt_view_patients.grid(row=3, column=2, padx=10, pady=10, sticky='ew')

        # Configuração das colunas da grid para que elas se expandam igualmente
        for i in range(3):  # 3 colunas
            self.grid_columnconfigure(i, weight=1)

    # Métodos que chamam as funções correspondentes de FuncoesBotoes
    def adicionar_informacao(self):
        """Chama a função para adicionar informações."""
        self.funcoes_botoes.adicionar_informacao()

    def excluir_informacao(self):
        """Chama a função para excluir informações."""
        self.funcoes_botoes.excluir()

    def exibir(self):
        """Chama a função para exibir informações."""
        self.funcoes_botoes.exibir_informacao()
    
    def exibir_contas(self):
        """Chama a função para exibir contas."""
        self.funcoes_botoes.valores_totais()

    def emitir_notas(self):
        """Chama a função para processar notas fiscais."""
        self.funcoes_botoes.processar_notas_fiscais()

    def resultados_consulta(self):
        """Chama a função para exibir resultados de consulta."""
        self.funcoes_botoes.exibir_resultado()

    def relatorio_wpp(self):
        """Chama a função para enviar relatório via WhatsApp."""
        self.funcoes_botoes.enviar_whatsapp()

    def relatorio_email(self):
        """Chama a função para enviar relatório via e-mail."""
        self.funcoes_botoes.enviar_email()

    def marcar_paciente(self):
        """Chama a função para marcar paciente"""
        self.banco.add_user()

    def visu_marcacoes(self):
        self.banco.view_marcacoes()

    def format_planilha(self):
        self.funcoes_botoes.formatar_planilha()

    def fechamento_contas(self):
        self.sistema_contas.abrir_janela()

    def planilha_sheet(self):
        self.gerenciador_planilhas.abrir_gerenciador()





        