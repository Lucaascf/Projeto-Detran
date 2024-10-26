from tkinter import *


class App(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.tela()
        self.frames_da_tela()
        self.pagina_login()
        self.criando_botoes()
        self.pack()

    def pagina_login(self):
        # Rotulo usuario
        self.label_user = Label(self.master, text='Usuario:')
        self.label_user.pack(pady=5)
        # Campo entrada usuario
        self.entry_user = Entry(self.master)
        self.entry_user.pack(pady=5)

        # Rotulo senha
        self.label_password = Label(self.master, text="Senha:")
        self.label_password.pack(pady=5)
        # Campo de entrada de Senha
        self.entry_password = Entry(self.master, show='*')  # show='*' oculta a senha
        self.entry_password.pack(pady=5)

        # Botao Login
        self.login_button = Button(self.master, text='login', command=self.login)

    def login(self):
        user = self.entry_user.get()
        password = self.entry_password.get()
        if user == '' and password == '':
            print('Ola')
        else:
            print('Algo deu errado')


    def tela(self):
        self.master.title('Cadastro de Clientes')
        self.master.configure(background='#1e3743')
        self.master.geometry('700x500')
        self.master.resizable(True, True)
        self.master.maxsize(width=900, height=700)
        self.master.minsize(width=400, height=300)

    def frames_da_tela(self):
        self.frame_botoes = Frame(self.master, bd=4, bg='#dfe3ee',
                                  highlightbackground='#759fe6', highlightthickness=3)
        self.frame_botoes.pack(fill='both')  # Usar expand=True para ocupar o espaço

        self.frame_2 = Frame(self.master, bd=4, bg='#dfe3ee',
                             highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)

    def criando_botoes(self):
        # Criar os botões em uma grade
        self.bt_adicionar_informacoes = Button(self.frame_botoes, text='Adicionar Informações')
        self.bt_adicionar_informacoes.grid(row=0, column=0, padx=10, pady=10, sticky='ew')

        self.bt_excluir_informacao = Button(self.frame_botoes, text='Excluir Informação')
        self.bt_excluir_informacao.grid(row=0, column=1, padx=10, pady=10, sticky='ew')

        self.bt_exibir_informacoes = Button(self.frame_botoes, text='Exibir Informações')
        self.bt_exibir_informacoes.grid(row=0, column=2, padx=10, pady=10, sticky='ew')

        self.bt_exibir_contas = Button(self.frame_botoes, text='Exibir Contas')
        self.bt_exibir_contas.grid(row=1, column=0, padx=10, pady=10, sticky='ew')

        self.bt_enviar_relatorio = Button(self.frame_botoes, text='Enviar Relatório')
        self.bt_enviar_relatorio.grid(row=1, column=1, padx=10, pady=10, sticky='ew')

        self.bt_emitir_ntfs = Button(self.frame_botoes, text='Emitir NTFS-e')
        self.bt_emitir_ntfs.grid(row=1, column=2, padx=10, pady=10, sticky='ew')

        # Configurar as colunas para expandir igualmente
        for i in range(3):  # 3 colunas
            self.frame_botoes.grid_columnconfigure(i, weight=1)

if __name__ == '__main__':
    root = Tk()
    app = App(master=root)
    app.mainloop()
