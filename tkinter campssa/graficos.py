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
        
        self.colors = {
            'ganhos': '#2ecc71',
            'custos': '#e74c3c',
            'custos_medicos': '#3498db',
            'compareceram': '#2ecc71',
            'faltaram': '#e74c3c',
            'pendentes': '#f1c40f',
            'background': '#2C3E50',
            'text': '#ECF0F1'
        }

    def gerar_grafico(self):
        self.window = tk.Toplevel(self.master)
        self.window.title("Visualização de Dados")
        self.window.geometry("1200x800")
        self.window.configure(bg=self.colors['background'])

        # Frame de seleção
        selection_frame = tk.Frame(self.window, bg=self.colors['background'], padx=20, pady=20)
        selection_frame.pack(fill="x")

        # Frame para datas
        date_frame = tk.Frame(selection_frame, bg=self.colors['background'])
        date_frame.pack(side=tk.TOP, fill="x", pady=(0, 10))

        # Data inicial
        tk.Label(
            date_frame,
            text="Data Inicial:",
            bg=self.colors['background'],
            fg=self.colors['text'],
            font=("Arial", 12)
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.start_date = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.start_date.pack(side=tk.LEFT, padx=10)

        # Data final
        tk.Label(
            date_frame,
            text="Data Final:",
            bg=self.colors['background'],
            fg=self.colors['text'],
            font=("Arial", 12)
        ).pack(side=tk.LEFT, padx=(20, 10))

        self.end_date = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.end_date.pack(side=tk.LEFT, padx=10)

        # Frame para tipo de gráfico
        type_frame = tk.Frame(selection_frame, bg=self.colors['background'])
        type_frame.pack(side=tk.TOP, fill="x", pady=10)

        tk.Label(
            type_frame,
            text="Tipo de Visualização:",
            bg=self.colors['background'],
            fg=self.colors['text'],
            font=("Arial", 12)
        ).pack(side=tk.LEFT, padx=(0, 10))

        self.graph_type = ttk.Combobox(
            type_frame,
            values=[
                "Ganhos Totais",
                "Custos Totais", 
                "Custos Médicos",
                "Frequência de Pacientes"
            ],
            width=30,
            state="readonly"
        )
        self.graph_type.pack(side=tk.LEFT, padx=10)
        self.graph_type.set("Ganhos Totais")

        # Botão de atualização
        ttk.Button(
            selection_frame,
            text="Atualizar Gráfico",
            command=self.update_graph
        ).pack(side=tk.TOP, pady=10)

        # Frame para o gráfico
        self.graph_frame = tk.Frame(
            self.window, 
            bg=self.colors['background'],
            highlightbackground=self.colors['text'],
            highlightthickness=1
        )
        self.graph_frame.pack(fill="both", expand=True, padx=20, pady=20)

    def get_date_range(self):
        start_date = self.start_date.get_date()
        end_date = self.end_date.get_date()
        return start_date, end_date

    def get_data(self):
        try:
            graph_type = self.graph_type.get()
            start_date, end_date = self.get_date_range()
            
            queries = {
                "Ganhos Totais": """
                    SELECT date(data_agendamento) as data,
                        COUNT(CASE WHEN nome in (SELECT nome FROM marcacoes m2 
                            WHERE m2.nome = marcacoes.nome 
                            AND m2.data_agendamento BETWEEN date(?) AND date(?)
                            AND m2.status_comparecimento = 'attended'
                            GROUP BY m2.nome
                            HAVING COUNT(*) = 1) THEN 1 END) * 148.65 as valor_medico,
                        COUNT(CASE WHEN nome in (SELECT nome FROM marcacoes m2 
                            WHERE m2.nome = marcacoes.nome 
                            AND m2.data_agendamento BETWEEN date(?) AND date(?)
                            AND m2.status_comparecimento = 'attended'
                            GROUP BY m2.nome
                            HAVING COUNT(*) = 1) THEN 1 END) * 192.61 as valor_psicologo
                    FROM marcacoes 
                    WHERE date(data_agendamento) BETWEEN date(?) AND date(?)
                        AND status_comparecimento = 'attended'
                    GROUP BY data
                    ORDER BY data ASC
                """,
                "Custos Totais": """
                    SELECT 
                        nome,
                        SUM(CASE 
                            WHEN tipo = 'medico' THEN 49.00
                            WHEN tipo = 'psicologo' THEN 63.50
                            ELSE 0 
                        END) as valor_profissional
                    FROM (
                        SELECT 
                            nome,
                            CASE 
                                WHEN EXISTS (
                                    SELECT 1 FROM marcacoes m2 
                                    WHERE m2.nome = m1.nome 
                                    AND m2.data_agendamento BETWEEN date(?) AND date(?)
                                    AND m2.status_comparecimento = 'attended'
                                    GROUP BY m2.nome
                                    HAVING COUNT(*) = 1
                                ) THEN 'medico'
                                ELSE 'psicologo'
                            END as tipo
                        FROM marcacoes m1
                        WHERE date(data_agendamento) BETWEEN date(?) AND date(?)
                        AND status_comparecimento = 'attended'
                    ) subquery
                    GROUP BY nome
                    HAVING valor_profissional > 0
                    ORDER BY valor_profissional DESC
                """,
                "Custos Médicos": """
                    SELECT 
                        'Médico' as tipo,
                        COUNT(*) * 49.00 as valor_medico,
                        0 as valor_psicologo
                    FROM marcacoes
                    WHERE date(data_agendamento) BETWEEN date(?) AND date(?)
                        AND status_comparecimento = 'attended'
                        AND EXISTS (
                            SELECT 1 FROM marcacoes m2 
                            WHERE m2.nome = marcacoes.nome
                            AND m2.data_agendamento BETWEEN date(?) AND date(?)
                            AND m2.status_comparecimento = 'attended'
                            GROUP BY m2.nome
                            HAVING COUNT(*) = 1
                        )
                    UNION ALL
                    SELECT 
                        'Psicólogo' as tipo,
                        0 as valor_medico,
                        COUNT(*) * 63.50 as valor_psicologo
                    FROM marcacoes
                    WHERE date(data_agendamento) BETWEEN date(?) AND date(?)
                        AND status_comparecimento = 'attended'
                        AND EXISTS (
                            SELECT 1 FROM marcacoes m2 
                            WHERE m2.nome = marcacoes.nome
                            AND m2.data_agendamento BETWEEN date(?) AND date(?)
                            AND m2.status_comparecimento = 'attended'
                            GROUP BY m2.nome
                            HAVING COUNT(*) > 1
                        )
                    ORDER BY tipo
                """,
                "Frequência de Pacientes": """
                    SELECT date(data_agendamento) as data,
                        SUM(CASE WHEN status_comparecimento = 'attended' THEN 1 ELSE 0 END) as compareceram,
                        SUM(CASE WHEN status_comparecimento = 'missed' THEN 1 ELSE 0 END) as faltaram,
                        SUM(CASE WHEN status_comparecimento = 'pending' THEN 1 ELSE 0 END) as pendentes
                    FROM marcacoes
                    WHERE date(data_agendamento) BETWEEN date(?) AND date(?)
                    GROUP BY data
                    ORDER BY data ASC
                """
            }

            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                query = queries.get(graph_type)
                
                # Ajustar parâmetros baseado no tipo de consulta
                if graph_type == "Ganhos Totais":
                    cursor.execute(query, (start_date, end_date, start_date, end_date, start_date, end_date))
                elif graph_type == "Custos Totais":
                    cursor.execute(query, (start_date, end_date, start_date, end_date))
                elif graph_type == "Custos Médicos":
                    cursor.execute(query, (start_date, end_date, start_date, end_date, start_date, end_date, start_date, end_date))
                else:  # Frequência de Pacientes
                    cursor.execute(query, (start_date, end_date))
                
                results = cursor.fetchall()
                
                if not results:
                    messagebox.showinfo("Aviso", "Nenhum dado encontrado para o período selecionado")
                    return None
                    
                return results

        except sqlite3.Error as e:
            messagebox.showerror("Erro", f"Erro ao buscar dados: {e}")
            return None

    def update_graph(self):
        for widget in self.graph_frame.winfo_children():
            widget.destroy()

        data = self.get_data()
        if not data:
            return

        fig, ax = plt.subplots(figsize=(12, 7))
        fig.patch.set_facecolor(self.colors['background'])
        ax.set_facecolor(self.colors['background'])
        
        graph_type = self.graph_type.get()

        if graph_type in ["Ganhos Totais", "Custos Totais"]:
            self._plot_value_graph(ax, data, graph_type)
        elif graph_type == "Custos Médicos":
            self._plot_cost_graph(ax, data)
        else:  # Frequência de Pacientes
            self._plot_frequency_graph(ax, data)

        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _plot_value_graph(self, ax, data, graph_type):
        dates = [datetime.strptime(row[0], '%Y-%m-%d').strftime('%d/%m/%Y') for row in data]
        values = [float(row[1] or 0) for row in data]
        color = self.colors['ganhos'] if graph_type == "Ganhos Totais" else self.colors['custos']
        
        bars = ax.bar(dates, values, color=color, alpha=0.8)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width()/2.,
                height,
                f'R${height:,.2f}',
                ha='center', 
                va='bottom',
                color=self.colors['text']
            )
            
        ax.set_ylabel('Valor (R$)', color=self.colors['text'])
        ax.set_title(f"{graph_type}", color=self.colors['text'])
        plt.xticks(rotation=45, ha='right', color=self.colors['text'])
        plt.yticks(color=self.colors['text'])

    def _plot_cost_graph(self, ax, data):
        """Plot específico para custos médicos e psicológicos"""
        tipos = [row[0] for row in data]
        valores_medico = [float(row[1] or 0) for row in data]
        valores_psicologo = [float(row[2] or 0) for row in data]
        
        # Posições das barras
        x = range(len(tipos))
        width = 0.35
        
        # Criar barras
        bars1 = ax.bar(x, valores_medico, width, label='Médico', color=self.colors['custos'])
        bars2 = ax.bar([i + width for i in x], valores_psicologo, width, label='Psicólogo', color=self.colors['custos_medicos'])
        
        # Adicionar rótulos
        def autolabel(bars):
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height,
                            f'R${height:,.2f}',
                            ha='center', va='bottom', color=self.colors['text'])

        autolabel(bars1)
        autolabel(bars2)
        
        # Customizar eixos
        ax.set_ylabel('Valor Total (R$)', color=self.colors['text'])
        ax.set_title('Custos por Tipo de Profissional', color=self.colors['text'])
        ax.set_xticks([i + width/2 for i in x])
        ax.set_xticklabels(tipos)
        plt.xticks(color=self.colors['text'])
        plt.yticks(color=self.colors['text'])
        ax.legend(facecolor=self.colors['background'], labelcolor=self.colors['text'])

    def _plot_frequency_graph(self, ax, data):
        dates = [datetime.strptime(row[0], '%Y-%m-%d').strftime('%d/%m/%Y') for row in data]
        attended = [int(row[1] or 0) for row in data]
        missed = [int(row[2] or 0) for row in data]
        pending = [int(row[3] or 0) for row in data]
        
        width = 0.35
        
        p1 = ax.bar(dates, attended, width, label='Compareceram', color=self.colors['compareceram'])
        p2 = ax.bar(dates, missed, width, bottom=attended, label='Faltaram', color=self.colors['faltaram'])
        p3 = ax.bar(dates, pending, width, bottom=[i+j for i,j in zip(attended, missed)], 
                    label='Pendentes', color=self.colors['pendentes'])
        
        def auto_label(rects, heights, bottom=None):
            for i, rect in enumerate(rects):
                height = heights[i]
                if height > 0:
                    y = rect.get_y() + (rect.get_height()/2)
                    if bottom is not None:
                        y = bottom[i] + (rect.get_height()/2)
                    ax.text(rect.get_x() + rect.get_width()/2., y,
                            str(height),
                            ha='center', va='center', color=self.colors['text'])
        
        auto_label(p1, attended)
        auto_label(p2, missed, attended)
        auto_label(p3, pending, [i+j for i,j in zip(attended, missed)])
        
        ax.set_ylabel('Número de Pacientes', color=self.colors['text'])
        ax.set_title('Frequência de Pacientes', color=self.colors['text'])
        ax.legend(facecolor=self.colors['background'], labelcolor=self.colors['text'])
        plt.xticks(rotation=45, ha='right', color=self.colors['text'])
        plt.yticks(color=self.colors['text'])