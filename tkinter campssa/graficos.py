import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
import sqlite3
from tkcalendar import DateEntry
import json
import logging
from typing import List, Dict, Any, Optional, Tuple

class GraficoMarcacoes:
    def __init__(self, master: tk.Tk, planilhas: Any, caminho_arquivo: str, app: Any):
        self.master = master
        self.planilhas = planilhas
        self.caminho_arquivo = caminho_arquivo
        self.app = app
        self.nome_banco = "db_marcacao.db"
        
        # Configuração de logging
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        handler = logging.FileHandler('graficos.log')
        handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
        self.logger.addHandler(handler)
        
        # Paleta de cores moderna
        self.cores = {
            'medico': '#2563EB',       # Azul royal
            'psicologo': '#7C3AED',    # Roxo vibrante
            'compareceram': '#059669',  # Verde esmeralda
            'faltaram': '#DC2626',     # Vermelho intenso
            'pendentes': '#D97706',     # Laranja âmbar
            'fundo': '#1E293B',        # Azul escuro slate
            'texto': '#F1F5F9',        # Cinza claro slate
            'destaque': '#38BDF8'      # Azul celeste
        }

    def gerar_grafico(self) -> None:
        """Gera o dashboard principal de análises"""
        self.janela = tk.Toplevel(self.master)
        self.janela.title("Dashboard - Análise de Atendimentos")
        self.janela.geometry("1400x800")
        self.janela.configure(bg=self.cores['fundo'])

        # Frame principal
        frame_principal = tk.Frame(self.janela, bg=self.cores['fundo'])
        frame_principal.pack(fill="both", expand=True, padx=20, pady=20)

        # Frame de controles
        frame_controles = self._criar_frame_controles(frame_principal)
        frame_controles.pack(fill="x", pady=(0, 20))

        # Frame para gráficos
        self.frame_graficos = tk.Frame(frame_principal, bg=self.cores['fundo'])
        self.frame_graficos.pack(fill="both", expand=True)

        # Frame para resumo
        self.frame_resumo = tk.Frame(frame_principal, bg=self.cores['fundo'])
        self.frame_resumo.pack(fill="x", pady=(20, 0))

        # Inicializa com dados
        self.atualizar_visualizacao()

    def _criar_frame_controles(self, parent: tk.Frame) -> tk.Frame:
        """Cria o frame de controles do dashboard"""
        frame = tk.Frame(parent, bg=self.cores['fundo'])

        # Seleção de período
        self._criar_seletores_data(frame)

        # Tipo de visualização
        self._criar_seletor_visualizacao(frame)

        # Botão de atualização
        ttk.Button(
            frame,
            text="Atualizar Visualização",
            command=self.atualizar_visualizacao,
            style='Accent.TButton'
        ).pack(side="left", padx=20)

        return frame

    def _criar_seletores_data(self, frame: tk.Frame) -> None:
        """Cria os seletores de data inicial e final"""
        tk.Label(
            frame,
            text="Período de Análise:",
            bg=self.cores['fundo'],
            fg=self.cores['texto'],
            font=("Arial", 12, "bold")
        ).pack(side="left", padx=10)

        # Data inicial
        self.data_inicial = DateEntry(
            frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.data_inicial.pack(side="left", padx=5)

        tk.Label(
            frame,
            text="até",
            bg=self.cores['fundo'],
            fg=self.cores['texto'],
            font=("Arial", 12)
        ).pack(side="left", padx=5)

        # Data final
        self.data_final = DateEntry(
            frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy'
        )
        self.data_final.pack(side="left", padx=5)

    def _criar_seletor_visualizacao(self, frame: tk.Frame) -> None:
        """Cria o seletor de tipo de visualização"""
        tk.Label(
            frame,
            text="Tipo de Análise:",
            bg=self.cores['fundo'],
            fg=self.cores['texto'],
            font=("Arial", 12, "bold")
        ).pack(side="left", padx=(20, 10))

        self.tipo_grafico = ttk.Combobox(
            frame,
            values=[
                "Análise de Receitas",
                "Comparativo de Atendimentos",
                "Métricas por Profissional",
                "Análise de Frequência",
                "Tendências Mensais"
            ],
            width=25,
            state="readonly"
        )
        self.tipo_grafico.set("Análise de Receitas")
        self.tipo_grafico.pack(side="left", padx=5)

    def obter_dados(self) -> Optional[List[Tuple]]:
        """Obtém dados do banco de dados com base nos filtros selecionados"""
        try:
            data_inicial = self.data_inicial.get_date().strftime("%Y-%m-%d")
            data_final = self.data_final.get_date().strftime("%Y-%m-%d")
            tipo_grafico = self.tipo_grafico.get()
            
            with sqlite3.connect(self.nome_banco) as conn:
                cursor = conn.cursor()
                
                # Seleciona a query apropriada com base no tipo de gráfico
                query = self._obter_query_por_tipo(tipo_grafico)
                
                # Executa a query
                cursor.execute(query, (data_inicial, data_final))
                resultados = cursor.fetchall()
                
                if not resultados:
                    messagebox.showinfo("Aviso", "Nenhum dado encontrado para o período selecionado")
                    return None
                    
                return resultados

        except sqlite3.Error as e:
            self.logger.error(f"Erro ao buscar dados: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao buscar dados: {str(e)}")
            return None

    def _obter_query_por_tipo(self, tipo_grafico: str) -> str:
        """Retorna a query SQL apropriada para cada tipo de gráfico"""
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
                    SUM(CASE WHEN status_comparecimento = 'missed' THEN 1 ELSE 0 END) as faltas,
                    SUM(CASE WHEN status_comparecimento = 'pending' THEN 1 ELSE 0 END) as pendentes
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
            """,
            
            "Tendências Mensais": """
                SELECT 
                    strftime('%Y-%m', data_agendamento) as mes,
                    COUNT(*) as total_consultas,
                    SUM(CASE WHEN status_comparecimento = 'attended' THEN 1 ELSE 0 END) as presencas,
                    SUM(CASE WHEN status_comparecimento = 'missed' THEN 1 ELSE 0 END) as faltas,
                    SUM(CASE 
                        WHEN status_comparecimento = 'attended' AND 
                            (historico_comparecimento IS NULL OR historico_comparecimento = '[]')
                        THEN 148.65
                        WHEN status_comparecimento = 'attended' AND 
                            historico_comparecimento IS NOT NULL AND 
                            historico_comparecimento != '[]'
                        THEN 192.61
                        ELSE 0 
                    END) as receita_total
                FROM marcacoes
                WHERE data_agendamento BETWEEN ? AND ?
                GROUP BY mes
                ORDER BY mes
            """
        }
        
        return queries.get(tipo_grafico, queries["Análise de Receitas"])

    def atualizar_visualizacao(self) -> None:
        """Atualiza a visualização dos gráficos e resumos"""
        # Limpa frames existentes
        for widget in self.frame_graficos.winfo_children():
            widget.destroy()
        for widget in self.frame_resumo.winfo_children():
            widget.destroy()

        # Obtém novos dados
        dados = self.obter_dados()
        if not dados:
            return

        # Cria nova figura
        fig, ax = plt.subplots(figsize=(12, 6))
        fig.patch.set_facecolor(self.cores['fundo'])
        ax.set_facecolor(self.cores['fundo'])
        ax.tick_params(colors=self.cores['texto'])
        ax.xaxis.label.set_color(self.cores['texto'])
        ax.yaxis.label.set_color(self.cores['texto'])
        ax.title.set_color(self.cores['texto'])

        # Plota gráfico apropriado
        tipo_grafico = self.tipo_grafico.get()
        if tipo_grafico == "Análise de Receitas":
            self._plotar_receitas(ax, dados)
        elif tipo_grafico == "Comparativo de Atendimentos":
            self._plotar_comparativo(ax, dados)
        elif tipo_grafico == "Métricas por Profissional":
            self._plotar_metricas(ax, dados)
        elif tipo_grafico == "Análise de Frequência":
            self._plotar_frequencia(ax, dados)
        else:
            self._plotar_tendencias(ax, dados)

        plt.tight_layout()
        
        # Adiciona canvas
        canvas = FigureCanvasTkAgg(fig, self.frame_graficos)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        # Atualiza resumo
        self._atualizar_resumo(dados, tipo_grafico)

    def _plotar_receitas(self, ax: plt.Axes, dados: List[Tuple]) -> None:
        """Plota o gráfico de análise de receitas"""
        datas = [datetime.strptime(row[0], '%Y-%m-%d').strftime('%d/%m') for row in dados]
        medico = [float(row[1] or 0) for row in dados]
        psicologo = [float(row[2] or 0) for row in dados]

        largura = 0.35
        ax.bar([i - largura/2 for i in range(len(datas))], medico, largura, 
               label='Médico', color=self.cores['medico'])
        ax.bar([i + largura/2 for i in range(len(datas))], psicologo, largura,
               label='Psicólogo', color=self.cores['psicologo'])

        ax.set_title('Análise de Receitas por Profissional', color=self.cores['texto'])
        ax.set_xlabel('Data', color=self.cores['texto'])
        ax.set_ylabel('Valor (R$)', color=self.cores['texto'])
        ax.set_xticks(range(len(datas)))
        ax.set_xticklabels(datas, rotation=45)
        ax.legend(facecolor=self.cores['fundo'], labelcolor=self.cores['texto'])
        
        # Adiciona valores sobre as barras
        for i, (m, p) in enumerate(zip(medico, psicologo)):
            ax.text(i - largura/2, m, f'R${m:.2f}', ha='center', va='bottom', color=self.cores['texto'])
            ax.text(i + largura/2, p, f'R${p:.2f}', ha='center', va='bottom', color=self.cores['texto'])

    def _plotar_comparativo(self, ax: plt.Axes, dados: List[Tuple]) -> None:
        """Plota o gráfico comparativo de atendimentos"""
        tipos = ['Médico' if row[0] == 'medico' else 'Psicólogo' for row in dados]
        atendimentos = [row[1] for row in dados]
        pacientes = [row[2] for row in dados]
        valores = [float(row[3]) for row in dados]

        x = range(len(tipos))
        largura = 0.25

        # Plotagem das barras
        barras_atend = ax.bar([i - largura for i in x], atendimentos, largura, 
                             label='Atendimentos', color=self.cores['compareceram'])
        barras_pac = ax.bar([i for i in x], pacientes, largura, 
                           label='Pacientes Únicos', color=self.cores['destaque'])
        barras_val = ax.bar([i + largura for i in x], valores, largura, 
                           label='Valor Total (R$)', color=self.cores['medico'])

        ax.set_title('Comparativo entre Profissionais', color=self.cores['texto'])
        ax.set_xticks(x)
        ax.set_xticklabels(tipos)
        ax.legend(facecolor=self.cores['fundo'], labelcolor=self.cores['texto'])

        # Adiciona rótulos nas barras
        def autolabel(barras):
            for barra in barras:
                altura = barra.get_height()
                ax.text(barra.get_x() + barra.get_width()/2., altura,
                       f'{altura:.0f}' if altura < 1000 else f'R${altura:.2f}',
                       ha='center', va='bottom', color=self.cores['texto'])

        autolabel(barras_atend)
        autolabel(barras_pac)
        autolabel(barras_val)

    def _plotar_metricas(self, ax: plt.Axes, dados: List[Tuple]) -> None:
        """Plota o gráfico de métricas por profissional"""
        tipos = ['Médico' if row[0] == 'medico' else 'Psicólogo' for row in dados]
        consultas = [row[1] for row in dados]
        pacientes = [row[2] for row in dados]
        presencas = [row[3] for row in dados]
        faltas = [row[4] for row in dados]
        pendentes = [row[5] for row in dados]

        x = range(len(tipos))
        largura = 0.15

        # Plotagem das barras
        ax.bar([i - 2*largura for i in x], consultas, largura, 
               label='Total Consultas', color=self.cores['destaque'])
        ax.bar([i - largura for i in x], pacientes, largura, 
               label='Total Pacientes', color=self.cores['medico'])
        ax.bar([i for i in x], presencas, largura, 
               label='Presenças', color=self.cores['compareceram'])
        ax.bar([i + largura for i in x], faltas, largura, 
               label='Faltas', color=self.cores['faltaram'])
        ax.bar([i + 2*largura for i in x], pendentes, largura, 
               label='Pendentes', color=self.cores['pendentes'])

        ax.set_title('Métricas Detalhadas por Profissional', color=self.cores['texto'])
        ax.set_xticks(x)
        ax.set_xticklabels(tipos)
        ax.legend(facecolor=self.cores['fundo'], labelcolor=self.cores['texto'])

    def _plotar_frequencia(self, ax: plt.Axes, dados: List[Tuple]) -> None:
        """Plota o gráfico de análise de frequência"""
        datas = sorted(list(set(row[0] for row in dados)))
        dados_medico = {data: [0, 0, 0] for data in datas}
        dados_psicologo = {data: [0, 0, 0] for data in datas}

        for row in dados:
            dicionario = dados_medico if row[1] == 'medico' else dados_psicologo
            dicionario[row[0]] = [row[2], row[3], row[4]]

        largura = 0.35
        x = range(len(datas))

        # Plotagem para médico
        base_med = [0] * len(datas)
        for i, status in enumerate(['Compareceram', 'Faltaram', 'Pendentes']):
            valores = [dados_medico[d][i] for d in datas]
            ax.bar([xi - largura/2 for xi in x], valores, largura, bottom=base_med,
                   label=f'Médico - {status}', 
                   color=[self.cores['compareceram'], self.cores['faltaram'], 
                         self.cores['pendentes']][i])
            base_med = [sum(x) for x in zip(base_med, valores)]

        # Plotagem para psicólogo
        base_psi = [0] * len(datas)
        for i, status in enumerate(['Compareceram', 'Faltaram', 'Pendentes']):
            valores = [dados_psicologo[d][i] for d in datas]
            ax.bar([xi + largura/2 for xi in x], valores, largura, bottom=base_psi,
                   label=f'Psicólogo - {status}',
                   color=[self.cores['compareceram'], self.cores['faltaram'], 
                         self.cores['pendentes']][i],
                   alpha=0.7)
            base_psi = [sum(x) for x in zip(base_psi, valores)]

        ax.set_title('Análise de Frequência por Profissional', color=self.cores['texto'])
        ax.set_xlabel('Data', color=self.cores['texto'])
        ax.set_ylabel('Quantidade', color=self.cores['texto'])
        ax.set_xticks(x)
        ax.set_xticklabels([datetime.strptime(d, '%Y-%m-%d').strftime('%d/%m') 
                           for d in datas], rotation=45)
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', 
                 facecolor=self.cores['fundo'], labelcolor=self.cores['texto'])

    def _plotar_tendencias(self, ax: plt.Axes, dados: List[Tuple]) -> None:
        """Plota o gráfico de tendências mensais"""
        meses = [datetime.strptime(row[0], '%Y-%m').strftime('%b/%Y') for row in dados]
        total_consultas = [row[1] for row in dados]
        presencas = [row[2] for row in dados]
        faltas = [row[3] for row in dados]
        receitas = [float(row[4]) for row in dados]

        # Plotagem de linhas
        ax.plot(meses, total_consultas, 'o-', label='Total Consultas', 
                color=self.cores['destaque'])
        ax.plot(meses, presencas, 's-', label='Presenças', 
                color=self.cores['compareceram'])
        ax.plot(meses, faltas, '^-', label='Faltas', 
                color=self.cores['faltaram'])

        # Eixo secundário para receitas
        ax2 = ax.twinx()
        ax2.plot(meses, receitas, 'D-', label='Receita (R$)', 
                 color=self.cores['medico'])
        ax2.set_ylabel('Receita (R$)', color=self.cores['medico'])

        ax.set_title('Tendências Mensais', color=self.cores['texto'])
        ax.set_xlabel('Mês', color=self.cores['texto'])
        ax.set_ylabel('Quantidade', color=self.cores['texto'])
        ax.tick_params(axis='x', rotation=45)

        # Combina as legendas dos dois eixos
        linhas1, labels1 = ax.get_legend_handles_labels()
        linhas2, labels2 = ax2.get_legend_handles_labels()
        ax.legend(linhas1 + linhas2, labels1 + labels2, 
                 loc='upper left', bbox_to_anchor=(1.05, 1),
                 facecolor=self.cores['fundo'], labelcolor=self.cores['texto'])

    def _atualizar_resumo(self, dados: List[Tuple], tipo_grafico: str) -> None:
        """Atualiza o resumo estatístico com base no tipo de gráfico"""
        # Frame para o título do resumo
        frame_titulo = tk.Frame(self.frame_resumo, bg=self.cores['fundo'])
        frame_titulo.pack(fill="x", padx=20, pady=(10,5))
        
        tk.Label(
            frame_titulo,
            text="RESUMO DA ANÁLISE",
            font=("Arial", 12, "bold"),
            bg=self.cores['fundo'],
            fg=self.cores['texto']
        ).pack()

        # Frame para o conteúdo do resumo
        frame_conteudo = tk.Frame(self.frame_resumo, bg=self.cores['fundo'])
        frame_conteudo.pack(fill="both", expand=True, padx=20, pady=5)

        try:
            if tipo_grafico == "Análise de Receitas":
                total_medico = sum(float(row[1] or 0) for row in dados)
                total_psicologo = sum(float(row[2] or 0) for row in dados)
                total_pac_med = sum(int(row[3] or 0) for row in dados)
                total_pac_psi = sum(int(row[4] or 0) for row in dados)

                # Frame para dados do médico
                frame_medico = tk.LabelFrame(
                    frame_conteudo,
                    text="MÉDICO",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['medico'],
                    padx=10,
                    pady=5
                )
                frame_medico.pack(side="left", fill="both", expand=True, padx=5)

                metricas_medico = [
                    f"Receita Total: R$ {total_medico:,.2f}",
                    f"Total de Pacientes: {total_pac_med}",
                    f"Média por Paciente: R$ {(total_medico/total_pac_med if total_pac_med else 0):,.2f}"
                ]

                for metrica in metricas_medico:
                    tk.Label(
                        frame_medico,
                        text=metrica,
                        font=("Arial", 10),
                        bg=self.cores['fundo'],
                        fg=self.cores['texto']
                    ).pack(anchor="w")

                # Frame para dados do psicólogo
                frame_psicologo = tk.LabelFrame(
                    frame_conteudo,
                    text="PSICÓLOGO",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['psicologo'],
                    padx=10,
                    pady=5
                )
                frame_psicologo.pack(side="left", fill="both", expand=True, padx=5)

                metricas_psicologo = [
                    f"Receita Total: R$ {total_psicologo:,.2f}",
                    f"Total de Pacientes: {total_pac_psi}",
                    f"Média por Paciente: R$ {(total_psicologo/total_pac_psi if total_pac_psi else 0):,.2f}"
                ]

                for metrica in metricas_psicologo:
                    tk.Label(
                        frame_psicologo,
                        text=metrica,
                        font=("Arial", 10),
                        bg=self.cores['fundo'],
                        fg=self.cores['texto']
                    ).pack(anchor="w")

                # Frame para total geral
                frame_total = tk.LabelFrame(
                    frame_conteudo,
                    text="TOTAL GERAL",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['destaque'],
                    padx=10,
                    pady=5
                )
                frame_total.pack(side="left", fill="both", expand=True, padx=5)

                total_geral = total_medico + total_psicologo
                tk.Label(
                    frame_total,
                    text=f"R$ {total_geral:,.2f}",
                    font=("Arial", 12, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['texto']
                ).pack(expand=True)

            elif tipo_grafico == "Comparativo de Atendimentos":
                for row in dados:
                    tipo_prof = "MÉDICO" if row[0] == "medico" else "PSICÓLOGO"
                    cor_prof = self.cores['medico'] if row[0] == "medico" else self.cores['psicologo']
                    atendimentos = row[1]
                    pacientes = row[2]
                    valor = float(row[3])

                    frame_prof = tk.LabelFrame(
                        frame_conteudo,
                        text=tipo_prof,
                        font=("Arial", 10, "bold"),
                        bg=self.cores['fundo'],
                        fg=cor_prof,
                        padx=10,
                        pady=5
                    )
                    frame_prof.pack(side="left", fill="both", expand=True, padx=5)

                    metricas = [
                        f"Total de Atendimentos: {atendimentos}",
                        f"Pacientes Únicos: {pacientes}",
                        f"Valor Total: R$ {valor:,.2f}",
                        f"Média por Atendimento: R$ {(valor/atendimentos if atendimentos else 0):,.2f}"
                    ]

                    for metrica in metricas:
                        tk.Label(
                            frame_prof,
                            text=metrica,
                            font=("Arial", 10),
                            bg=self.cores['fundo'],
                            fg=self.cores['texto']
                        ).pack(anchor="w")
            elif tipo_grafico == "Métricas por Profissional":
                for row in dados:
                    tipo_prof = "MÉDICO" if row[0] == "medico" else "PSICÓLOGO"
                    cor_prof = self.cores['medico'] if row[0] == "medico" else self.cores['psicologo']
                    total_consultas = row[1]
                    total_pacientes = row[2]
                    presencas = row[3]
                    faltas = row[4]
                    pendentes = row[5]

                    taxa_presenca = (presencas / total_consultas * 100) if total_consultas else 0
                    taxa_falta = (faltas / total_consultas * 100) if total_consultas else 0

                    frame_prof = tk.LabelFrame(
                        frame_conteudo,
                        text=tipo_prof,
                        font=("Arial", 10, "bold"),
                        bg=self.cores['fundo'],
                        fg=cor_prof,
                        padx=10,
                        pady=5
                    )
                    frame_prof.pack(side="left", fill="both", expand=True, padx=5)

                    metricas = [
                        f"Total de Consultas: {total_consultas}",
                        f"Total de Pacientes: {total_pacientes}",
                        f"Presenças: {presencas} ({taxa_presenca:.1f}%)",
                        f"Faltas: {faltas} ({taxa_falta:.1f}%)",
                        f"Pendentes: {pendentes}"
                    ]

                    for metrica in metricas:
                        tk.Label(
                            frame_prof,
                            text=metrica,
                            font=("Arial", 10),
                            bg=self.cores['fundo'],
                            fg=self.cores['texto']
                        ).pack(anchor="w")

            elif tipo_grafico == "Análise de Frequência":
                # Separa dados por tipo de profissional
                dados_medico = [row for row in dados if row[1] == 'medico']
                dados_psicologo = [row for row in dados if row[1] == 'psicologo']
                
                # Frame para Médico
                frame_medico = tk.LabelFrame(
                    frame_conteudo,
                    text="MÉDICO",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['medico'],
                    padx=10,
                    pady=5
                )
                frame_medico.pack(side="left", fill="both", expand=True, padx=5)

                # Cálculos para médico
                total_compareceram_med = sum(row[2] for row in dados_medico)
                total_faltaram_med = sum(row[3] for row in dados_medico)
                total_pendentes_med = sum(row[4] for row in dados_medico)
                total_consultas_med = total_compareceram_med + total_faltaram_med + total_pendentes_med
                taxa_comparecimento_med = (total_compareceram_med / total_consultas_med * 100) if total_consultas_med else 0

                metricas_medico = [
                    f"Total de Consultas: {total_consultas_med}",
                    f"Compareceram: {total_compareceram_med}",
                    f"Faltaram: {total_faltaram_med}",
                    f"Pendentes: {total_pendentes_med}",
                    f"Taxa de Comparecimento: {taxa_comparecimento_med:.1f}%"
                ]

                for metrica in metricas_medico:
                    tk.Label(
                        frame_medico,
                        text=metrica,
                        font=("Arial", 10),
                        bg=self.cores['fundo'],
                        fg=self.cores['texto']
                    ).pack(anchor="w")

                # Frame para dados do psicólogo
                frame_psicologo = tk.LabelFrame(
                    frame_conteudo,
                    text="PSICÓLOGO",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['psicologo'],
                    padx=10,
                    pady=5
                )
                frame_psicologo.pack(side="left", fill="both", expand=True, padx=5)

                # Cálculos para psicólogo
                total_compareceram_psi = sum(row[2] for row in dados_psicologo)
                total_faltaram_psi = sum(row[3] for row in dados_psicologo)
                total_pendentes_psi = sum(row[4] for row in dados_psicologo)
                total_consultas_psi = total_compareceram_psi + total_faltaram_psi + total_pendentes_psi
                taxa_comparecimento_psi = (total_compareceram_psi / total_consultas_psi * 100) if total_consultas_psi else 0

                metricas_psicologo = [
                    f"Total de Consultas: {total_consultas_psi}",
                    f"Compareceram: {total_compareceram_psi}",
                    f"Faltaram: {total_faltaram_psi}",
                    f"Pendentes: {total_pendentes_psi}",
                    f"Taxa de Comparecimento: {taxa_comparecimento_psi:.1f}%"
                ]

                for metrica in metricas_psicologo:
                    tk.Label(
                        frame_psicologo,
                        text=metrica,
                        font=("Arial", 10),
                        bg=self.cores['fundo'],
                        fg=self.cores['texto']
                    ).pack(anchor="w")

                # Frame para total geral
                frame_total = tk.LabelFrame(
                    frame_conteudo,
                    text="TOTAL GERAL",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['destaque'],
                    padx=10,
                    pady=5
                )
                frame_total.pack(side="left", fill="both", expand=True, padx=5)

                # Cálculos totais
                total_consultas = total_consultas_med + total_consultas_psi
                total_compareceram = total_compareceram_med + total_compareceram_psi
                total_faltaram = total_faltaram_med + total_faltaram_psi
                total_pendentes = total_pendentes_med + total_pendentes_psi
                taxa_comparecimento_geral = (total_compareceram / total_consultas * 100) if total_consultas else 0

                metricas_total = [
                    f"Total Geral de Consultas: {total_consultas}",
                    f"Total de Presenças: {total_compareceram}",
                    f"Total de Faltas: {total_faltaram}",
                    f"Total Pendentes: {total_pendentes}",
                    f"Taxa Geral de Comparecimento: {taxa_comparecimento_geral:.1f}%"
                ]

                for metrica in metricas_total:
                    tk.Label(
                        frame_total,
                        text=metrica,
                        font=("Arial", 10),
                        bg=self.cores['fundo'],
                        fg=self.cores['texto']
                    ).pack(anchor="w")

            else:  # Tendências Mensais
                total_consultas = sum(row[1] for row in dados)
                total_presencas = sum(row[2] for row in dados)
                total_faltas = sum(row[3] for row in dados)
                total_receita = sum(float(row[4]) for row in dados)
                
                frame_totais = tk.LabelFrame(
                    frame_conteudo,
                    text="RESUMO DO PERÍODO",
                    font=("Arial", 10, "bold"),
                    bg=self.cores['fundo'],
                    fg=self.cores['destaque'],
                    padx=10,
                    pady=5
                )
                frame_totais.pack(fill="both", expand=True, padx=5)

                metricas = [
                    f"Total de Consultas: {total_consultas}",
                    f"Total de Presenças: {total_presencas}",
                    f"Total de Faltas: {total_faltas}",
                    f"Taxa de Comparecimento: {(total_presencas/total_consultas*100 if total_consultas else 0):.1f}%",
                    f"Receita Total: R$ {total_receita:,.2f}",
                    f"Média de Receita por Consulta: R$ {(total_receita/total_presencas if total_presencas else 0):,.2f}"
                ]

                for metrica in metricas:
                    tk.Label(
                        frame_totais,
                        text=metrica,
                        font=("Arial", 10),
                        bg=self.cores['fundo'],
                        fg=self.cores['texto']
                    ).pack(anchor="w")

        except Exception as e:
            self.logger.error(f"Erro ao processar dados do resumo: {str(e)}")
            tk.Label(
                frame_conteudo,
                text=f"Erro ao processar dados: {str(e)}",
                font=("Arial", 10),
                bg=self.cores['fundo'],
                fg='red'
            ).pack()

    def exportar_dados(self, formato: str = "csv") -> None:
        """Exporta os dados do gráfico atual para CSV ou Excel"""
        try:
            dados = self.obter_dados()
            if not dados:
                return

            tipo_grafico = self.tipo_grafico.get()
            data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"relatorio_{tipo_grafico.lower().replace(' ', '_')}_{data_atual}"

            if formato == "csv":
                import csv
                nome_arquivo += ".csv"
                with open(nome_arquivo, 'w', newline='', encoding='utf-8') as arquivo:
                    writer = csv.writer(arquivo)
                    # Escreve cabeçalho baseado no tipo de gráfico
                    headers = self._obter_headers_exportacao(tipo_grafico)
                    writer.writerow(headers)
                    writer.writerows(dados)
            
            elif formato == "excel":
                import pandas as pd
                nome_arquivo += ".xlsx"
                df = pd.DataFrame(dados, columns=self._obter_headers_exportacao(tipo_grafico))
                df.to_excel(nome_arquivo, index=False)

            messagebox.showinfo("Sucesso", f"Dados exportados para {nome_arquivo}")
            self.logger.info(f"Dados exportados com sucesso para {nome_arquivo}")

        except Exception as e:
            self.logger.error(f"Erro ao exportar dados: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao exportar dados: {str(e)}")

    def _obter_headers_exportacao(self, tipo_grafico: str) -> List[str]:
        """Retorna os cabeçalhos apropriados para cada tipo de gráfico na exportação"""
        headers = {
            "Análise de Receitas": [
                "Data", "Valor Médico", "Valor Psicólogo", 
                "Pacientes Médico", "Pacientes Psicólogo"
            ],
            "Comparativo de Atendimentos": [
                "Tipo Profissional", "Total Atendimentos", 
                "Pacientes Únicos", "Valor Total"
            ],
            "Métricas por Profissional": [
                "Tipo Profissional", "Total Consultas", "Total Pacientes",
                "Presenças", "Faltas", "Pendentes"
            ],
            "Análise de Frequência": [
                "Data", "Tipo Profissional", "Compareceram",
                "Faltaram", "Pendentes"
            ],
            "Tendências Mensais": [
                "Mês", "Total Consultas", "Presenças", 
                "Faltas", "Receita Total"
            ]
        }
        return headers.get(tipo_grafico, ["Data", "Valor"])

    def aplicar_tema_escuro(self, valor: bool = True) -> None:
        """Aplica ou remove o tema escuro da interface"""
        if valor:
            self.cores.update({
                'fundo': '#1E293B',
                'texto': '#F1F5F9',
                'destaque': '#38BDF8'
            })
        else:
            self.cores.update({
                'fundo': '#FFFFFF',
                'texto': '#1E293B',
                'destaque': '#2563EB'
            })
        
        # Atualiza a interface
        self.janela.configure(bg=self.cores['fundo'])
        for widget in self.janela.winfo_children():
            if isinstance(widget, (tk.Frame, tk.Label)):
                widget.configure(bg=self.cores['fundo'], fg=self.cores['texto'])
        
        # Atualiza visualização
        self.atualizar_visualizacao()

    def mostrar_ajuda(self) -> None:
        """Exibe uma janela de ajuda com informações sobre o uso do dashboard"""
        janela_ajuda = tk.Toplevel(self.master)
        janela_ajuda.title("Ajuda do Dashboard")
        janela_ajuda.geometry("600x400")
        janela_ajuda.configure(bg=self.cores['fundo'])

        texto_ajuda = """
        Dashboard de Análise de Atendimentos
        
        Como usar:
        1. Selecione o período de análise usando os calendários
        2. Escolha o tipo de visualização desejada
        3. Clique em "Atualizar Visualização" para ver os dados
        
        Tipos de Análise:
        • Análise de Receitas: Mostra o faturamento por profissional
        • Comparativo de Atendimentos: Compare métricas entre profissionais
        • Métricas por Profissional: Visão detalhada de cada profissional
        • Análise de Frequência: Acompanhe presenças e faltas
        • Tendências Mensais: Visualize tendências ao longo do tempo
        
        Dicas:
        • Passe o mouse sobre as barras para ver valores específicos
        • Use o botão direito do mouse para salvar os gráficos
        • Exporte os dados para análises mais detalhadas
        """

        text_widget = tk.Text(
            janela_ajuda,
            wrap=tk.WORD,
            bg=self.cores['fundo'],
            fg=self.cores['texto'],
            font=("Arial", 11),
            padx=20,
            pady=20
        )
        text_widget.pack(fill="both", expand=True)
        text_widget.insert("1.0", texto_ajuda)
        text_widget.config(state="disabled")