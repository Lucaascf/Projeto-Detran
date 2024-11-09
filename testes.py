import tkinter as tk
from tkinter import ttk

class PacienteExibicao(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Exibição de Pacientes")
        self.geometry("700x400")

        # Títulos das colunas
        columns = ("nome", "renach", "pagamento", "medico", "psicologo")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        
        # Definindo cabeçalhos
        self.tree.heading("nome", text="Nome")
        self.tree.heading("renach", text="RENACH")
        self.tree.heading("pagamento", text="Forma de Pagamento")
        self.tree.heading("medico", text="Consultas Médico")
        self.tree.heading("psicologo", text="Consultas Psicólogo")
        
        # Ajustando largura das colunas
        self.tree.column("nome", width=150)
        self.tree.column("renach", width=100)
        self.tree.column("pagamento", width=120)
        self.tree.column("medico", width=120, anchor="center")
        self.tree.column("psicologo", width=120, anchor="center")

        # Exibindo a tabela
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Inserindo dados de exemplo
        self.inserir_dados_exemplo()

    def inserir_dados_exemplo(self):
        # Dados fictícios para exibição; substitua pelos dados do banco de dados ou planilha
        dados = [
            ("Ana Silva", "123456789", "Cartão de Crédito", 5, 3),
            ("Carlos Souza", "987654321", "Boleto", 2, 4),
            ("Mariana Costa", "564738291", "Dinheiro", 3, 1),
        ]
        
        # Inserindo dados na tabela
        for dado in dados:
            self.tree.insert("", tk.END, values=dado)

if __name__ == "__main__":
    app = PacienteExibicao()
    app.mainloop()
