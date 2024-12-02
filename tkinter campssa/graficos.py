import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
import sqlite3
from tkcalendar import DateEntry

class GraficoMarcacoes:
    def __init__(self, master, planilhas, file_path, app):
        self.master = master
        self.planilhas = planilhas
        self.file_path = file_path
        self.app = app
        self.db_name = "db_marcacao.db"
        
        # Paleta de cores melhorada
        self.colors = {
            'medico': '#2980B9',       # Azul para médico
            'psicologo': '#8E44AD',    # Roxo para psicólogo
            'compareceram': '#27AE60',  # Verde para presenças
            'faltaram': '#E74C3C',     # Vermelho para faltas
            'pendentes': '#F39C12',     # Laranja para pendentes
            'background': '#ECF0F1',    # Fundo claro
            'text': '#2C3E50'          # Texto escuro
        }

    def gerar_grafico(self):
        self.window = tk.Toplevel(self.master)
        self.window.title("Dashboard - Análise de Atendimentos")
        self.window.geometry("1400x800")
        self.window.configure(bg=self.colors['background'])

        # Frame principal
        main_frame = tk.Frame(self.window, bg=self.colors['background'])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Frame de controles
        control_frame = tk.Frame(main_frame, bg=self.colors['background'])
        control_frame.pack(fill="x", pady=(0, 20))

        # Seleção de período
        tk.Label(
            control_frame,
            text="Período de Análise:",
            bg=self.colors['background'],
            fg=self.colors['text'],
            font=("Arial", 12, "bold")
        ).pack(side="left", padx=10)

        self.start_date = DateEntry(
            control_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.start_date.pack(side="left", padx=5)

        tk.Label(
            control_frame,
            text="até",
            bg=self.colors['background'],
            fg=self.colors['text'],
            font=("Arial", 12)
        ).pack(side="left", padx=5)

        self.end_date = DateEntry(
            control_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.end_date.pack(side="left", padx=5)

        # Tipo de visualização
        tk.Label(
            control_frame,
            text="Tipo de Análise:",
            bg=self.colors['background'],
            fg=self.colors['text'],
            font=("Arial", 12, "bold")
        ).pack(side="left", padx=(20, 10))

        self.graph_type = ttk.Combobox(
            control_frame,
            values=[
                "Análise de Receitas",
                "Comparativo de Atendimentos",
                "Métricas por Profissional",
                "Análise de Frequência"
            ],
            width=25,
            state="readonly"
        )
        self.graph_type.pack(side="left", padx=5)
        self.graph_type.set("Análise de Receitas")

        ttk.Button(
            control_frame,
            text="Atualizar Visualização",
            command=self.update_graph
        ).pack(side="left", padx=20)

        # Frame para gráficos
        self.graph_frame = tk.Frame(main_frame, bg=self.colors['background'])
        self.graph_frame.pack(fill="both", expand=True)

        # Frame para sumário
        self.summary_frame = tk.Frame(main_frame, bg=self.colors['background'])
        self.summary_frame.pack(fill="x", pady=(20, 0))

    def get_data(self):
        try:
            start_date = self.start_date.get_date().strftime("%Y-%m-%d")
            end_date = self.end_date.get_date().strftime("%Y-%m-%d")
            graph_type = self.graph_type.get()
            
            queries = {
                "Análise de Receitas": """
                    SELECT 
                        date(data_agendamento) as data,
                        SUM(CASE WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' 
                            THEN 148.65 ELSE 0 END) as valor_medico,
                        SUM(CASE WHEN historico_comparecimento IS NOT NULL AND historico_comparecimento != '[]' 
                            THEN 192.61 ELSE 0 END) as valor_psicologo,
                        COUNT(DISTINCT CASE WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' 
                            THEN nome END) as pacientes_medico,
                        COUNT(DISTINCT CASE WHEN historico_comparecimento IS NOT NULL AND historico_comparecimento != '[]' 
                            THEN nome END) as pacientes_psicologo
                    FROM marcacoes 
                    WHERE data_agendamento BETWEEN ? AND ?
                        AND status_comparecimento = 'attended'
                    GROUP BY data
                    ORDER BY data
                """,

                "Comparativo de Atendimentos": """
                    SELECT 
                        CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 'medico'
                            ELSE 'psicologo'
                        END as tipo,
                        COUNT(*) as total_atendimentos,
                        COUNT(DISTINCT nome) as pacientes_unicos,
                        SUM(CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 148.65
                            ELSE 192.61 
                        END) as valor_total
                    FROM marcacoes
                    WHERE data_agendamento BETWEEN ? AND ?
                        AND status_comparecimento = 'attended'
                    GROUP BY 
                        CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 'medico'
                            ELSE 'psicologo'
                        END
                """,

                "Métricas por Profissional": """
                    SELECT 
                        CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 'medico'
                            ELSE 'psicologo'
                        END as tipo,
                        COUNT(*) as total_consultas,
                        COUNT(DISTINCT nome) as total_pacientes,
                        SUM(CASE WHEN status_comparecimento = 'attended' THEN 1 ELSE 0 END) as presencas,
                        SUM(CASE WHEN status_comparecimento = 'missed' THEN 1 ELSE 0 END) as faltas
                    FROM marcacoes
                    WHERE data_agendamento BETWEEN ? AND ?
                    GROUP BY 
                        CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 'medico'
                            ELSE 'psicologo'
                        END
                """,

                "Análise de Frequência": """
                    SELECT 
                        date(data_agendamento) as data,
                        CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 'medico'
                            ELSE 'psicologo'
                        END as tipo,
                        SUM(CASE WHEN status_comparecimento = 'attended' THEN 1 ELSE 0 END) as compareceram,
                        SUM(CASE WHEN status_comparecimento = 'missed' THEN 1 ELSE 0 END) as faltaram,
                        SUM(CASE WHEN status_comparecimento = 'pending' THEN 1 ELSE 0 END) as pendentes
                    FROM marcacoes
                    WHERE data_agendamento BETWEEN ? AND ?
                    GROUP BY data,
                        CASE 
                            WHEN historico_comparecimento IS NULL OR historico_comparecimento = '[]' THEN 'medico'
                            ELSE 'psicologo'
                        END
                    ORDER BY data, tipo
                """
            }

            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                query = queries[graph_type]
                
                # Todas as queries agora usam apenas dois parâmetros de data
                cursor.execute(query, (start_date, end_date))
                results = cursor.fetchall()
                
                if not results:
                    messagebox.showinfo("Aviso", "Nenhum dado encontrado para o período selecionado")
                    return None
                    
                return results

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao buscar dados: {str(e)}")
            return None

    def update_graph(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()
        for widget in self.summary_frame.winfo_children():
            widget.destroy()

        data = self.get_data()
        if not data:
            return

        graph_type = self.graph_type.get()
        
        fig, ax = plt.subplots(figsize=(12, 6))
        fig.patch.set_facecolor(self.colors['background'])
        ax.set_facecolor(self.colors['background'])

        if graph_type == "Análise de Receitas":
            self._plot_receitas(ax, data)
        elif graph_type == "Comparativo de Atendimentos":
            self._plot_comparativo(ax, data)
        elif graph_type == "Métricas por Profissional":
            self._plot_metricas(ax, data)
        else:
            self._plot_frequencia(ax, data)

        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Atualiza sumário
        self._update_summary(data, graph_type)

    def _plot_receitas(self, ax, data):
        dates = [datetime.strptime(row[0], '%Y-%m-%d').strftime('%d/%m') for row in data]
        medico = [float(row[1] or 0) for row in data]
        psicologo = [float(row[2] or 0) for row in data]

        width = 0.35
        ax.bar([i - width/2 for i in range(len(dates))], medico, width, 
               label='Médico', color=self.colors['medico'])
        ax.bar([i + width/2 for i in range(len(dates))], psicologo, width,
               label='Psicólogo', color=self.colors['psicologo'])

        ax.set_title('Receitas por Profissional')
        ax.set_xlabel('Data')
        ax.set_ylabel('Valor (R$)')
        ax.set_xticks(range(len(dates)))
        ax.set_xticklabels(dates, rotation=45)
        ax.legend()

    def _plot_comparativo(self, ax, data):
        tipos = ['Médico' if row[0] == 'medico' else 'Psicólogo' for row in data]
        atendimentos = [row[1] for row in data]
        pacientes = [row[2] for row in data]
        valores = [row[3] for row in data]

        x = range(len(tipos))
        width = 0.25

        ax.bar([i - width for i in x], atendimentos, width, label='Atendimentos')
        ax.bar([i for i in x], pacientes, width, label='Pacientes Únicos')
        ax.bar([i + width for i in x], valores, width, label='Valor Total (R$)')

        ax.set_title('Comparativo entre Profissionais')
        ax.set_xticks(x)
        ax.set_xticklabels(tipos)
        ax.legend()

    def _plot_metricas(self, ax, data):
        tipos = ['Médico' if row[0] == 'medico' else 'Psicólogo' for row in data]
        consultas = [row[1] for row in data]
        pacientes = [row[2] for row in data]
        presencas = [row[3] for row in data]
        faltas = [row[4] for row in data]

        x = range(len(tipos))
        width = 0.2

        ax.bar([i - 1.5*width for i in x], consultas, width, label='Total Consultas')
        ax.bar([i - 0.5*width for i in x], pacientes, width, label='Total Pacientes')
        ax.bar([i + 0.5*width for i in x], presencas, width, label='Presenças')
        ax.bar([i + 1.5*width for i in x], faltas, width, label='Faltas')

        ax.set_title('Métricas por Profissional')
        ax.set_xticks(x)
        ax.set_xticklabels(tipos)
        ax.legend()

    def _plot_frequencia(self, ax, data):
        dates = sorted(list(set(row[0] for row in data)))
        medico_data = {date: [0, 0, 0] for date in dates}
        psicologo_data = {date: [0, 0, 0] for date in dates}

        for row in data:
            data_dict = medico_data if row[1] == 'medico' else psicologo_data
            data_dict[row[0]] = [row[2], row[3], row[4]]

        width = 0.35
        x = range(len(dates))

        # Médico
        bottom_med = [0] * len(dates)
        for i, status in enumerate(['Compareceram', 'Faltaram', 'Pendentes']):
            values = [medico_data[d][i] for d in dates]
            ax.bar([xi - width/2 for xi in x], values, width, bottom=bottom_med,
                   label=f'Médico - {status}')
            bottom_med = [sum(x) for x in zip(bottom_med, values)]

        # Psicólogo
        bottom_psi = [0] * len(dates)
        for i, status in enumerate(['Compareceram', 'Faltaram', 'Pendentes']):
            values = [psicologo_data[d][i] for d in dates]
            ax.bar([xi + width/2 for xi in x], values, width, bottom=bottom_psi,
                   label=f'Psicólogo - {status}')
            bottom_psi = [sum(x) for x in zip(bottom_psi, values)]

        ax.set_title('Frequência por Profissional')
        ax.set_xlabel('Data')
        ax.set_ylabel('Quantidade')
        ax.set_xticks(x)
        ax.set_xticklabels([datetime.strptime(d, '%Y-%m-%d').strftime('%d/%m') for d in dates], 
                          rotation=45)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')

    def _update_summary(self, data, graph_type):
        if not data or not data[0]:  # Verifica se há dados válidos
            return
            
        summary_text = "Resumo da Análise:\n\n"
        
        try:
            if graph_type == "Análise de Receitas":
                # Calcula totais para médico e psicólogo
                total_medico = sum(float(row[1] or 0) for row in data)
                total_psicologo = sum(float(row[2] or 0) for row in data)
                total_pacientes_med = len(set(row[3] for row in data if row[3]))
                total_pacientes_psi = len(set(row[4] for row in data if row[4]))
                
                summary_text += f"""Médico:
• Receita Total: R$ {total_medico:,.2f}
• Total de Pacientes: {total_pacientes_med}
• Média por Paciente: R$ {(total_medico/total_pacientes_med if total_pacientes_med else 0):,.2f}

Psicólogo:
• Receita Total: R$ {total_psicologo:,.2f}
• Total de Pacientes: {total_pacientes_psi}
• Média por Paciente: R$ {(total_psicologo/total_pacientes_psi if total_pacientes_psi else 0):,.2f}

Total Geral: R$ {(total_medico + total_psicologo):,.2f}"""

            elif graph_type == "Comparativo de Atendimentos":
                for row in data:
                    prof_tipo = "Médico" if row[0] == "medico" else "Psicólogo"
                    atendimentos = int(row[1] or 0)
                    pacientes = int(row[2] or 0)
                    valor = float(row[3] or 0)
                    
                    summary_text += f"""{prof_tipo}:
• Total de Atendimentos: {atendimentos}
• Pacientes Únicos: {pacientes}
• Valor Total: R$ {valor:,.2f}
• Média por Atendimento: R$ {(valor/atendimentos if atendimentos else 0):,.2f}\n\n"""

            elif graph_type == "Métricas por Profissional":
                for row in data:
                    prof_tipo = "Médico" if row[0] == "medico" else "Psicólogo"
                    total_consultas = int(row[1] or 0)
                    total_pacientes = int(row[2] or 0)
                    presencas = int(row[3] or 0)
                    faltas = int(row[4] or 0)
                    
                    taxa_presenca = (presencas / total_consultas * 100) if total_consultas else 0
                    taxa_falta = (faltas / total_consultas * 100) if total_consultas else 0
                    
                    summary_text += f"""{prof_tipo}:
• Total de Consultas: {total_consultas}
• Total de Pacientes: {total_pacientes}
• Presenças: {presencas} ({taxa_presenca:.1f}%)
• Faltas: {faltas} ({taxa_falta:.1f}%)\n\n"""

            else:  # Análise de Frequência
                medico_data = [row for row in data if row[1] == 'medico']
                psi_data = [row for row in data if row[1] == 'psicologo']
                
                # Dados do Médico
                med_presencas = sum(int(row[2] or 0) for row in medico_data)
                med_faltas = sum(int(row[3] or 0) for row in medico_data)
                med_pendentes = sum(int(row[4] or 0) for row in medico_data)
                med_total = med_presencas + med_faltas
                
                # Dados do Psicólogo
                psi_presencas = sum(int(row[2] or 0) for row in psi_data)
                psi_faltas = sum(int(row[3] or 0) for row in psi_data)
                psi_pendentes = sum(int(row[4] or 0) for row in psi_data)
                psi_total = psi_presencas + psi_faltas
                
                summary_text += f"""Médico:
• Presenças: {med_presencas}
• Faltas: {med_faltas}
• Pendentes: {med_pendentes}
• Taxa de Comparecimento: {(med_presencas/med_total*100 if med_total else 0):.1f}%

Psicólogo:
• Presenças: {psi_presencas}
• Faltas: {psi_faltas}
• Pendentes: {psi_pendentes}
• Taxa de Comparecimento: {(psi_presencas/psi_total*100 if psi_total else 0):.1f}%"""

        except Exception as e:
            summary_text += f"\nErro ao processar dados: {str(e)}"

        finally:
            # Criar e configurar o widget de texto para o sumário
            text_widget = tk.Text(
                self.summary_frame,
                height=12,
                width=50,
                font=("Arial", 10),
                bg=self.colors['background'],
                fg=self.colors['text'],
                relief=tk.GROOVE,
                padx=10,
                pady=10
            )
            text_widget.insert("1.0", summary_text)
            text_widget.config(state="disabled")
            text_widget.pack(side="left", padx=20, pady=10)