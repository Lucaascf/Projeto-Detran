import logging
import re
import openpyxl
import sqlite3
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
import tkinter as tk
from tkinter import *
from tkinter import (
    messagebox,
    filedialog,
    Frame,
    Label,
    Entry,
    Button,
    simpledialog,
    ttk,
)
import sys
from planilhas import Planilhas
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl
from openpyxl.styles import Font
import subprocess
from datetime import datetime
from tkcalendar import DateEntry


# Configurando logs
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def role_e_click(driver, xpath):
    """Rola a página até um elemento e clica nele.

    Args:
        driver: Instância do WebDriver do Selenium.
        xpath: XPath do elemento a ser clicado.
    """
    element = WebDriverWait(driver, 30).until(
        EC.visibility_of_element_located((By.XPATH, xpath))
    )
    # Rola para o elemento
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    element.click()


class FuncoesBotoes:
    """Classe que encapsula as funções relacionadas aos botões da interface."""

    def __init__(self, master: tk, planilhas: Planilhas, file_path: str, app, current_user=None):
        """Inicializa a classe FuncoesBotoes."""
        self.master = master
        self.planilhas = planilhas
        self.wb = self.planilhas.wb if self.planilhas else None
        self.file_path = file_path
        self.app = app
        self.current_user = current_user
        self.login_frame = None
        self.criar_conta_frame = None
        self.logger = logging.getLogger(__name__)

        # Variáveis para pagamento
        self._init_payment_vars()


    def _init_payment_vars(self):
        """Inicializa variáveis relacionadas a pagamento."""
        self.forma_pagamento_var = tk.StringVar(value="")
        self.radio_var = tk.StringVar(value="")
        self.payment_vars = {
            'D': tk.IntVar(),
            'C': tk.IntVar(),
            'E': tk.IntVar(),
            'P': tk.IntVar()
        }
        self.valor_entries = {}


    def set_current_user(self, user):
        """Define o usuário atual."""
        self.current_user = user


    def center(self, window):
        """Centraliza a janela na tela."""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
        window.deiconify()


    def get_active_workbook(self):
        """Obtém o workbook ativo atualizado."""
        if self.planilhas:
            self.planilhas.reload_workbook()
            self.wb = self.planilhas.wb
        return self.wb


    def _create_payment_frame(self, parent, cor_fundo, cor_texto, cor_selecionado):
        """Cria o frame de pagamento com todas as opções."""
        frame_pagamento = tk.LabelFrame(
            parent,
            text="Formas de Pagamento",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 12, "bold")
        )
        frame_pagamento.pack(padx=20, pady=10, fill="x")

        def on_payment_change():
            selected_count = sum(var.get() for var in self.payment_vars.values())
            for forma, entry in self.valor_entries.items():
                if selected_count > 1:
                    entry.config(state="normal")
                    if not entry.get():
                        entry.config(bg="#FFE5E5")
                else:
                    entry.delete(0, tk.END)
                    entry.config(state="disabled", bg="#F0F0F0")

        formas_pagamento = {
            'D': 'Débito',
            'C': 'Crédito',
            'E': 'Espécie',
            'P': 'PIX'
        }

        for codigo, nome in formas_pagamento.items():
            frame = tk.Frame(frame_pagamento, bg=cor_fundo)
            frame.pack(fill="x", padx=10, pady=2)

            cb = tk.Checkbutton(
                frame,
                text=nome,
                variable=self.payment_vars[codigo],
                bg=cor_fundo,
                fg=cor_texto,
                selectcolor=cor_selecionado,
                activebackground=cor_fundo,
                activeforeground=cor_texto,
                highlightthickness=0,
                command=on_payment_change
            )
            cb.pack(side=tk.LEFT, padx=(0, 10))

            valor_entry = tk.Entry(frame, width=15, state="disabled")
            valor_entry.pack(side=tk.LEFT)
            self.valor_entries[codigo] = valor_entry

            tk.Label(frame, text="R$", bg=cor_fundo, fg=cor_texto).pack(side=tk.LEFT, padx=(5, 0))

        return frame_pagamento


    def adicionar_informacao(self):
        """Cria uma nova janela para adicionar informações de pacientes."""
        self.adicionar_window = tk.Toplevel(self.master)
        self.adicionar_window.title("Adicionar Paciente")
        self.adicionar_window.geometry("500x450")
        self.adicionar_window.minsize(500, 450)
        self.adicionar_window.maxsize(500, 450)

        cor_fundo = self.master.cget("bg")
        cor_texto = "#ECF0F1"
        cor_selecionado = "#2C3E50"

        self.adicionar_window.configure(bg=cor_fundo)
        self.center(self.adicionar_window)

        # Configuração da interface
        self._setup_add_interface(cor_fundo, cor_texto, cor_selecionado)


    def _setup_add_interface(self, cor_fundo, cor_texto, cor_selecionado):
        """Configura a interface de adição de paciente."""
        # Título
        tk.Label(
            self.adicionar_window,
            text="Preencha as informações:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 16, "bold")
        ).pack(pady=(15, 5))

        # Frame para RadioButtons
        self._create_radio_frame(cor_fundo, cor_texto, cor_selecionado)

        # Label para mostrar valor da consulta
        self.valor_consulta_label = tk.Label(
            self.adicionar_window,
            text="Valor da consulta: R$ 0,00",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 10, "bold")
        )
        self.valor_consulta_label.pack(pady=5)

        # Função para atualizar o valor da consulta
        def atualizar_valor_consulta(*args):
            valores = {
                "medico": "148,65",
                "psicologo": "192,61",
                "ambos": "341,26"
            }
            valor = valores.get(self.radio_var.get(), "0,00")
            self.valor_consulta_label.config(text=f"Valor da consulta: R$ {valor}")

        # Associar a função ao radio_var
        self.radio_var.trace("w", atualizar_valor_consulta)

        # Entradas para nome e Renach (removido duplicatas)
        self.criar_entry("Nome:", "nome_entry", self.adicionar_window)
        self.criar_entry("Renach:", "renach_entry", self.adicionar_window)

        # Frame de pagamento
        self._create_payment_frame(self.adicionar_window, cor_fundo, cor_texto, cor_selecionado)

        # Botões
        self._create_button_frame(cor_fundo)

        # Texto de ajuda
        tk.Label(
            self.adicionar_window,
            text="Obs.: Para múltiplas formas de pagamento, informe o valor de cada uma.",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 9, "italic")
        ).pack(pady=(0, 10))


    def _create_radio_frame(self, cor_fundo, cor_texto, cor_selecionado):
        """Cria o frame com os radio buttons."""
        frame_radios = tk.Frame(self.adicionar_window, bg=cor_fundo)
        frame_radios.pack(pady=5)

        tipos = [("Médico", "medico"), ("Psicólogo", "psicologo"), ("Ambos", "ambos")]
        for tipo, valor in tipos:
            tk.Radiobutton(
                frame_radios,
                text=tipo,
                variable=self.radio_var,
                value=valor,
                bg=cor_fundo,
                fg=cor_texto,
                selectcolor=cor_selecionado,
                activebackground=cor_fundo,
                activeforeground=cor_texto,
                highlightthickness=0,
                font=("Arial", 12)
            ).pack(side=tk.LEFT, padx=2)


    def _create_button_frame(self, cor_fundo):
        """Cria o frame com os botões."""
        frame_botoes = tk.Frame(self.adicionar_window, bg=cor_fundo)
        frame_botoes.pack(pady=20)

        tk.Button(
            frame_botoes,
            text="Adicionar",
            command=self.salvar_informacao,
            width=15,
            highlightthickness=0,
            activebackground="#2C3E50",
            activeforeground="#ECF0F1"
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            frame_botoes,
            text="Voltar",
            command=self.adicionar_window.destroy,
            width=15,
            activebackground="#2C3E50",
            activeforeground="#ECF0F1"
        ).pack(side=tk.LEFT, padx=5)


    def _adicionar_paciente_ao_banco(self, nome, renach, pagamentos, escolha):
        """Adiciona ou atualiza um paciente no banco de dados de marcação."""
        try:
            # Conecta ao banco de dados de marcação
            conn = sqlite3.connect('db_marcacao.db')
            cursor = conn.cursor()
            
            # Verifica se o paciente já existe
            cursor.execute("SELECT appointment_date FROM patients WHERE renach = ?", (renach,))
            existing_patient = cursor.fetchone()
            
            # Data atual para o registro
            data_atual = datetime.now().strftime("%Y-%m-%d")
            
            if isinstance(pagamentos, list):
                pagamento_info = ' | '.join(pagamentos)
            else:
                pagamento_info = str(pagamentos)
                
            observation = f"Tipo: {escolha}\nPagamento: {pagamento_info}\nRegistrado em: {data_atual}"
            
            if existing_patient:
                # Atualiza a data e observações do paciente existente
                cursor.execute("""
                    UPDATE patients 
                    SET appointment_date = ?,
                        observation = ?
                    WHERE renach = ?
                """, (data_atual, observation, renach))
                
            else:
                # Insere novo paciente
                cursor.execute("""
                    INSERT INTO patients (name, renach, phone, appointment_date, observation)
                    VALUES (?, ?, ?, ?, ?)
                """, (nome, renach, '', data_atual, observation))
            
            conn.commit()
            return True
            
        except sqlite3.Error as e:
            self.logger.error(f"Erro SQLite ao adicionar/atualizar paciente: {str(e)}")
            raise Exception(f"Erro ao salvar no banco de dados: {str(e)}")
            
        except Exception as e:
            self.logger.error(f"Erro ao adicionar/atualizar paciente: {str(e)}")
            raise
            
        finally:
            if 'conn' in locals():
                conn.close()


    def contar_pagamento(self, col_inicial, col_final):
        """Conta o número de pessoas e a quantidade de pagamentos."""
        n_pessoa = 0
        cont_pag = {"D": 0, "C": 0, "E": 0, "P": 0}

        # Usa a sheet ativa correta
        wb = self.get_active_workbook()
        ws = wb.active

        # Verifica se há conteúdo nas células antes de contar
        for row in ws.iter_rows(
            min_row=3, max_row=ws.max_row, min_col=col_inicial, max_col=col_final
        ):
            nome = row[0].value
            if isinstance(nome, str) and nome.strip():
                n_pessoa += 1

                # Verifica a forma de pagamento
                pag = row[4].value
                if isinstance(pag, str):
                    # Extrai apenas o código do pagamento (D, C, E ou P)
                    # considerando que pode ter valor após o código
                    codigo_pag = pag.split(":")[0].strip()
                    if codigo_pag in cont_pag:
                        cont_pag[codigo_pag] += 1

        return n_pessoa, cont_pag


    def criar_entry(self, frame_nome, var_name, parent):
        """Cria um frame com label e entry para entradas de texto."""
        frame = tk.Frame(parent, bg=parent.cget("bg"))
        frame.pack(pady=2)

        tk.Label(
            frame,
            text=frame_nome,
            bg=parent.cget("bg"),
            fg="#ECF0F1",
            font=("Arial", 12)
        ).pack(side=tk.LEFT, anchor="w", padx=5)

        entry = tk.Entry(frame)
        entry.pack(side=tk.LEFT, padx=2)
        setattr(self, var_name, entry)


    def salvar_informacao(self):
        """Salva os dados no banco de dados e na planilha."""
        try:
            # Obter dados dos campos
            nome = self.nome_entry.get().strip().upper()
            renach = self.renach_entry.get().strip()
            escolha = self.radio_var.get()

            # Valores máximos por tipo de atendimento
            VALOR_MEDICO = 148.65
            VALOR_PSICOLOGO = 192.61
            VALOR_AMBOS = 341.26

            # Definir valor máximo baseado na escolha
            valor_maximo = {
                "medico": VALOR_MEDICO,
                "psicologo": VALOR_PSICOLOGO,
                "ambos": VALOR_AMBOS
            }.get(escolha)

            # Validar dados básicos
            if not nome or not renach:
                messagebox.showerror("Erro", "Por favor, preencha os campos de nome e RENACH.")
                return

            if not renach.isdigit():
                messagebox.showerror("Erro", "O RENACH deve ser um número inteiro.")
                return

            if not escolha:
                messagebox.showerror("Erro", "Selecione o tipo de atendimento.")
                return

            # Verificar formas de pagamento selecionadas
            formas_selecionadas = {
                forma: var.get() for forma, var in self.payment_vars.items()
            }

            if not any(formas_selecionadas.values()):
                messagebox.showerror("Erro", "Selecione pelo menos uma forma de pagamento.")
                return

            # Contar quantas formas de pagamento foram selecionadas
            num_formas_selecionadas = sum(formas_selecionadas.values())

            # Processar pagamentos com as novas regras de validação
            pagamentos = []
            soma_valores = 0
            
            # Se apenas uma forma de pagamento está selecionada
            if num_formas_selecionadas == 1:
                forma_selecionada = next(forma for forma, sel in formas_selecionadas.items() if sel)
                valor_entrada = self.valor_entries[forma_selecionada].get().strip()
                
                # Se um valor foi especificado, use-o (desde que seja igual ao valor máximo)
                if valor_entrada:
                    try:
                        valor_float = float(valor_entrada.replace(',', '.'))
                        if abs(valor_float - valor_maximo) > 0.01:
                            messagebox.showerror(
                                "Erro",
                                f"O valor deve ser igual ao valor da consulta ({valor_maximo:.2f})"
                            )
                            return
                    except ValueError:
                        messagebox.showerror(
                            "Erro",
                            f"O valor informado não é um número válido"
                        )
                        return
                else:
                    # Se nenhum valor foi especificado, use o valor máximo
                    valor_float = valor_maximo
                
                valor_formatado = f"{valor_float:.2f}".replace('.', ',')
                pagamentos.append(f"{forma_selecionada}:{valor_formatado}")
                soma_valores = valor_float

            # Se múltiplas formas de pagamento estão selecionadas
            else:
                for codigo, selecionado in formas_selecionadas.items():
                    if selecionado:
                        valor = self.valor_entries[codigo].get().strip()
                        if not valor:
                            messagebox.showerror(
                                "Erro",
                                f"É obrigatório informar o valor para {codigo}"
                            )
                            return
                        
                        try:
                            valor_float = float(valor.replace(',', '.'))
                            soma_valores += valor_float
                        except ValueError:
                            messagebox.showerror(
                                "Erro",
                                f"O valor informado para {codigo} não é um número válido"
                            )
                            return
                        
                        valor_formatado = f"{valor_float:.2f}".replace('.', ',')
                        pagamentos.append(f"{codigo}:{valor_formatado}")

                # Validar soma dos valores
                if abs(soma_valores - valor_maximo) > 0.01:
                    messagebox.showerror(
                        "Erro",
                        f"A soma dos valores ({soma_valores:.2f}) deve ser igual ao "
                        f"valor da consulta ({valor_maximo:.2f})"
                    )
                    return

            # Continua com o salvamento após validações
            if self._adicionar_paciente_ao_banco(nome, renach, pagamentos, escolha):
                try:
                    wb = self.get_active_workbook()
                    ws = wb.active

                    # Formatar string de pagamento (com espaço após a barra vertical)
                    info_pagamento = " | ".join(pagamentos)

                    # Função para encontrar próxima linha vazia ou criar nova
                    def encontrar_proxima_linha(coluna_inicial):
                        ultima_linha = 3
                        for row in range(3, ws.max_row + 2):
                            if ws[f"{coluna_inicial}{row}"].value is None:
                                return row
                            ultima_linha = row
                        return ultima_linha + 1

                    # Salvar dados conforme o tipo de atendimento
                    if escolha in ["medico", "ambos"]:
                        nova_linha_medico = encontrar_proxima_linha("B")
                        ws[f"B{nova_linha_medico}"] = nome
                        ws[f"C{nova_linha_medico}"] = renach
                        ws[f"F{nova_linha_medico}"] = info_pagamento

                    if escolha in ["psicologo", "ambos"]:
                        nova_linha_psicologo = encontrar_proxima_linha("H")
                        ws[f"H{nova_linha_psicologo}"] = nome
                        ws[f"I{nova_linha_psicologo}"] = renach
                        ws[f"L{nova_linha_psicologo}"] = info_pagamento

                    # Salvar e formatar
                    wb.save(self.file_path)
                    self.formatar_planilha()  # Formata a planilha após salvar
                    messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")
                    self.adicionar_window.destroy()

                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar na planilha: {str(e)}")
                    
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar informações: {str(e)}")


    def salvar_na_planilha(self, nome, renach, pagamentos, escolha):
        """Salva os dados na planilha."""
        try:
            wb = self.get_active_workbook()
            ws = wb.active

            # Formatar string de pagamento
            if len(pagamentos) == 1 and ":" not in pagamentos[0]:
                # Uma única forma de pagamento sem valor
                info_pagamento = pagamentos[0]
            else:
                # Múltiplas formas de pagamento com valores
                info_pagamento = " | ".join(pagamentos)

            # Encontrar próximas linhas vazias
            nova_linha_medico = next(
                (row for row in range(3, ws.max_row + 2) if not ws[f"B{row}"].value),
                None,
            )
            nova_linha_psicologo = next(
                (row for row in range(3, ws.max_row + 2) if not ws[f"H{row}"].value),
                None,
            )

            if not nova_linha_medico or not nova_linha_psicologo:
                raise Exception("Não há linhas vazias disponíveis na planilha")

            if escolha in ["medico", "ambos"]:
                ws[f"B{nova_linha_medico}"] = nome
                ws[f"C{nova_linha_medico}"] = renach
                ws[f"F{nova_linha_medico}"] = info_pagamento

            if escolha in ["psicologo", "ambos"]:
                ws[f"H{nova_linha_psicologo}"] = nome
                ws[f"I{nova_linha_psicologo}"] = renach
                ws[f"L{nova_linha_psicologo}"] = info_pagamento

            wb.save(self.file_path)
            messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")

        except Exception as e:
            raise Exception(f"Erro ao salvar na planilha: {str(e)}")


    def excluir(self):
        """Remove informações de pacientes da planilha com base no RENACH fornecido pelo usuário."""
        try:
            wb = self.get_active_workbook()
            ws = wb.active

            def realizar_exclusao():
                try:
                    renach = int(renach_entry.get().strip())

                    def limpar_linha(row_num, start_col, end_col):
                        """Limpa os valores de uma linha específica"""
                        for col in range(start_col, end_col + 1):
                            cell = ws.cell(row=row_num, column=col)
                            # Verifica se não é uma célula mesclada
                            if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                                cell.value = None

                    def mover_conteudo(start_row, start_col, end_col):
                        """Move o conteúdo das células para cima"""
                        max_row = ws.max_row
                        # Move de baixo para cima para evitar sobrescrever dados
                        for row in range(start_row, max_row):
                            for col in range(start_col, end_col + 1):
                                current_cell = ws.cell(row=row, column=col)
                                next_cell = ws.cell(row=row + 1, column=col)
                                
                                # Só copia se a célula atual não for mesclada
                                if not isinstance(current_cell, openpyxl.cell.cell.MergedCell):
                                    if isinstance(next_cell, openpyxl.cell.cell.MergedCell):
                                        current_cell.value = None
                                    else:
                                        current_cell.value = next_cell.value

                        # Limpa a última linha
                        limpar_linha(max_row, start_col, end_col)

                    def encontrar_paciente(col_renach):
                        """Encontra a linha do paciente pelo RENACH"""
                        for row in range(3, ws.max_row + 1):
                            cell = ws.cell(row=row, column=col_renach)
                            if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                                if cell.value and str(cell.value).strip() == str(renach):
                                    return row
                        return None

                    # Procura nas seções de médico e psicólogo
                    linha_medico = encontrar_paciente(3)  # Coluna C
                    linha_psi = encontrar_paciente(9)     # Coluna I

                    alteracoes = False

                    if linha_medico:
                        mover_conteudo(linha_medico, 2, 6)  # Colunas B-F
                        alteracoes = True
                        messagebox.showinfo("Sucesso", "Removido da seção de médicos")

                    if linha_psi:
                        mover_conteudo(linha_psi, 8, 12)    # Colunas H-L
                        alteracoes = True
                        messagebox.showinfo("Sucesso", "Removido da seção de psicólogos")

                    if alteracoes:
                        wb.save(self.file_path)
                        excluir_window.destroy()
                    else:
                        messagebox.showwarning("Aviso", "RENACH não encontrado")

                except ValueError:
                    messagebox.showerror("Erro", "Por favor, insira um RENACH válido (apenas números)")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao excluir paciente: {str(e)}")

            # Interface da janela de exclusão
            excluir_window = tk.Toplevel(self.master)
            excluir_window.title("Excluir Paciente")
            excluir_window.geometry("400x150")
            excluir_window.resizable(False, False)
            excluir_window.configure(bg=self.master.cget("bg"))
            excluir_window.transient(self.master)
            excluir_window.grab_set()

            # Frame principal
            main_frame = tk.Frame(excluir_window, bg=self.master.cget("bg"))
            main_frame.pack(expand=True, fill='both', padx=20, pady=20)

            # Label
            tk.Label(
                main_frame,
                text="Informe o RENACH:",
                bg=self.master.cget("bg"),
                fg="#ECF0F1",
                font=("Arial", 14, "bold")
            ).pack(pady=10)

            # Entry frame
            entry_frame = tk.Frame(main_frame, bg=self.master.cget("bg"))
            entry_frame.pack(fill='x', pady=5)

            renach_entry = tk.Entry(entry_frame, justify='center')
            renach_entry.pack(pady=5)
            renach_entry.focus()
            renach_entry.bind('<Return>', lambda e: realizar_exclusao())

            # Botões
            button_frame = tk.Frame(main_frame, bg=self.master.cget("bg"))
            button_frame.pack(pady=10)

            tk.Button(
                button_frame,
                text="Excluir",
                command=realizar_exclusao,
                bg="#ff4444",
                fg="white",
                font=("Arial", 10, "bold"),
                width=15
            ).pack(side=tk.LEFT, padx=5)

            tk.Button(
                button_frame,
                text="Cancelar",
                command=excluir_window.destroy,
                width=15
            ).pack(side=tk.LEFT, padx=5)

            self.center(excluir_window)

            def on_closing():
                wb.close()
                excluir_window.destroy()

            excluir_window.protocol("WM_DELETE_WINDOW", on_closing)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao iniciar exclusão: {str(e)}")


    def exibir_resultado(self):
        """Exibe os resultados detalhados de contagem e valores por forma de pagamento."""
        try:
            # Criar janela
            janela_exibir_resultado = tk.Toplevel(self.master)
            janela_exibir_resultado.title("Valores de Atendimento")
            janela_exibir_resultado.geometry("500x600")
            janela_exibir_resultado.maxsize(width=500, height=600)
            janela_exibir_resultado.minsize(width=500, height=600)

            # Usando a cor de fundo da janela principal
            cor_fundo = self.master.cget("bg")
            janela_exibir_resultado.configure(bg=cor_fundo)

            def processar_pagamentos(pagamento_str):
                """Processa uma string de pagamento e retorna um dicionário com os valores."""
                resultado = {'D': [], 'C': [], 'E': [], 'P': []}
                if not pagamento_str or not isinstance(pagamento_str, str):
                    return resultado

                # Divide os diferentes pagamentos por pipe
                pagamentos = [p.strip() for p in pagamento_str.split('|')]
                
                for pag in pagamentos:
                    try:
                        # Remove espaços extras
                        pag = pag.strip()
                        
                        # Encontra a forma de pagamento (primeira letra)
                        forma = None
                        for letra in pag:
                            if letra.upper() in ['D', 'C', 'E', 'P']:
                                forma = letra.upper()
                                break
                        
                        if not forma:
                            continue
                            
                        # Procura pelo valor após o ':' se existir
                        if ':' in pag:
                            # Pega tudo após o primeiro ':'
                            valor_texto = pag[pag.index(':')+1:].strip()
                            # Limpa o texto do valor
                            valor_texto = valor_texto.replace('R$', '').replace(' ', '').replace(',', '.')
                            try:
                                valor = float(valor_texto)
                                resultado[forma].append(valor)
                            except ValueError:
                                print(f"Valor inválido ignorado: {valor_texto}")
                        else:
                            # Se não tem valor especificado e é o único pagamento
                            if len(pagamentos) == 1:
                                resultado[forma].append(148.65)  # Valor padrão médico
                    except Exception as e:
                        print(f"Erro ao processar pagamento '{pag}': {str(e)}")
                        continue

                return resultado

            def calcular_totais(lista_pagamentos):
                """Calcula totais por forma de pagamento."""
                totais = {'D': [], 'C': [], 'E': [], 'P': []}
                
                for pagamento in lista_pagamentos:
                    valores_pagamento = processar_pagamentos(pagamento)
                    
                    for forma, valores in valores_pagamento.items():
                        totais[forma].extend(valores)
                        
                # Processa os resultados finais
                resultados = {}
                for forma, valores in totais.items():
                    if valores:  # Só inclui formas que têm valores
                        resultados[forma] = {
                            'quantidade': len(valores),
                            'valor_total': sum(valores),
                            'valores': valores,
                            'media': sum(valores) / len(valores) if valores else 0
                        }
                        
                return resultados

            # Criar frame com scrollbar
            main_frame = tk.Frame(janela_exibir_resultado, bg=cor_fundo)
            main_frame.pack(fill="both", expand=True)

            canvas = tk.Canvas(main_frame, bg=cor_fundo)
            scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=cor_fundo)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            def criar_secao(titulo, n_pacientes, dados, row_start):
                """Cria uma seção de informações na janela."""
                tk.Label(
                    scrollable_frame,
                    text=titulo,
                    font=("Arial", 16, "bold"),
                    bg=cor_fundo,
                    fg="#ECF0F1",
                ).pack(pady=(20 if row_start > 0 else 15, 5))

                tk.Label(
                    scrollable_frame,
                    text=f"Total de Pacientes: {n_pacientes}",
                    font=("Arial", 12),
                    bg=cor_fundo,
                    fg="#ECF0F1",
                ).pack(pady=(0, 10))

                total_secao = 0
                formas_nome = {'D': 'Débito', 'C': 'Crédito', 'E': 'Espécie', 'P': 'PIX'}
                
                for forma, info in dados.items():
                    if info['quantidade'] > 0:
                        nome_forma = formas_nome.get(forma, forma)
                        
                        # Sumário da forma de pagamento
                        tk.Label(
                            scrollable_frame,
                            text=f"{nome_forma}: {info['quantidade']} pagamento(s) - Total: R$ {info['valor_total']:.2f}",
                            font=("Arial", 12),
                            bg=cor_fundo,
                            fg="#ECF0F1",
                        ).pack(pady=2)
                        
                        # Média por pagamento
                        tk.Label(
                            scrollable_frame,
                            text=f"Média por pagamento: R$ {info['media']:.2f}",
                            font=("Arial", 10),
                            bg=cor_fundo,
                            fg="#ECF0F1",
                        ).pack(pady=(0, 2))
                        
                        # Lista de valores individuais
                        valores_txt = ", ".join([f"R$ {v:.2f}" for v in info['valores']])
                        tk.Label(
                            scrollable_frame,
                            text=f"Valores: {valores_txt}",
                            font=("Arial", 10),
                            bg=cor_fundo,
                            fg="#ECF0F1",
                            wraplength=450
                        ).pack(pady=(0, 5))
                        
                        total_secao += info['valor_total']

                tk.Label(
                    scrollable_frame,
                    text=f"Total da Seção: R$ {total_secao:.2f}",
                    font=("Arial", 14, "bold"),
                    bg=cor_fundo,
                    fg="#ECF0F1",
                ).pack(pady=(5, 10))
                
                return total_secao

            # Obter dados do médico e psicólogo
            wb = self.get_active_workbook()
            ws = wb.active

            def contar_pacientes_e_coletar_pagamentos(col_inicial, col_final):
                """Conta pacientes e coleta pagamentos de uma seção."""
                pagamentos = []
                n_pacientes = 0
                for row in range(3, ws.max_row + 1):
                    nome = ws.cell(row=row, column=col_inicial).value
                    if nome and isinstance(nome, str) and nome.strip():
                        n_pacientes += 1
                        pagamento = ws.cell(row=row, column=col_final).value
                        if pagamento and isinstance(pagamento, str):
                            pagamentos.append(pagamento)
                return n_pacientes, pagamentos

            # Coletar dados (colunas B-F para médico, H-L para psicólogo)
            n_medico, pagamentos_medico = contar_pacientes_e_coletar_pagamentos(2, 6)    # B a F
            n_psicologo, pagamentos_psi = contar_pacientes_e_coletar_pagamentos(8, 12)   # H a L

            # Calcular totais
            totais_medico = calcular_totais(pagamentos_medico)
            totais_psi = calcular_totais(pagamentos_psi)

            # Exibir resultados
            total_med = criar_secao("MÉDICO", n_medico, totais_medico, 0)
            total_psi = criar_secao("PSICÓLOGO", n_psicologo, totais_psi, 1)

            # Separador
            tk.Frame(
                scrollable_frame,
                height=2,
                bg="#34495E",
                width=450
            ).pack(pady=10, fill='x', padx=20)

            # Total geral
            tk.Label(
                scrollable_frame,
                text="TOTAL GERAL",
                font=("Arial", 16, "bold"),
                bg=cor_fundo,
                fg="#ECF0F1",
            ).pack(pady=(10, 5))

            tk.Label(
                scrollable_frame,
                text=f"R$ {(total_med + total_psi):.2f}",
                font=("Arial", 14, "bold"),
                bg=cor_fundo,
                fg="#ECF0F1",
            ).pack(pady=(5, 20))

            # Configurar o layout do scroll
            canvas.pack(side="left", fill="both", expand=True, padx=5)
            scrollbar.pack(side="right", fill="y")

            # Configurar o scroll com o mouse
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
            def _on_closing():
                canvas.unbind_all("<MouseWheel>")
                janela_exibir_resultado.destroy()

            janela_exibir_resultado.protocol("WM_DELETE_WINDOW", _on_closing)

            # Centralizar a janela
            self.center(janela_exibir_resultado)
            wb.close()

        except Exception as e:
            self.logger.error(f"Erro ao exibir resultados: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro ao exibir os resultados: {str(e)}")
            if 'wb' in locals():
                wb.close()


    def valores_totais(self):
        """Exibe os valores totais em uma janela Tkinter."""
        # Recarrega o workbook para garantir dados atualizados
        self.planilhas.reload_workbook()

        # Obtém as contagens
        n_medico, pag_medico = self.planilhas.contar_medico()
        n_psicologo, pag_psicologo = self.planilhas.contar_psi()

        # Valores fixos
        VALOR_CONSULTA_MEDICO = 148.65
        VALOR_PAGAR_MEDICO = 49.00
        VALOR_CONSULTA_PSI = 192.61
        VALOR_PAGAR_PSI = 63.50

        # Cálculos
        total_medico = n_medico * VALOR_CONSULTA_MEDICO
        total_psicologo = n_psicologo * VALOR_CONSULTA_PSI
        valor_medico = n_medico * VALOR_PAGAR_MEDICO
        valor_psicologo = n_psicologo * VALOR_PAGAR_PSI

        # Criação da janela
        janela_contas = tk.Toplevel(self.master)
        janela_contas.title("Contas")
        janela_contas.geometry("300x400")
        janela_contas.configure(bg="#2C3E50")

        # Médico
        tk.Label(
            janela_contas,
            text="Contas Médico:",
            font=("Arial", 16, "bold"),
            bg="#2C3E50",
            fg="#ECF0F1",
        ).pack(pady=(15, 5))

        tk.Label(
            janela_contas,
            text=f"Número de pacientes: {n_medico}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(
            janela_contas,
            text=f"Valor Total: R$ {total_medico:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(
            janela_contas,
            text=f"Valor a Pagar: R$ {valor_medico:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(janela_contas, text="", bg="#2C3E50").pack(pady=10)

        # Psicólogo
        tk.Label(
            janela_contas,
            text="Contas Psicólogo:",
            font=("Arial", 16, "bold"),
            bg="#2C3E50",
            fg="#ECF0F1",
        ).pack(pady=5)

        tk.Label(
            janela_contas,
            text=f"Número de pacientes: {n_psicologo}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(
            janela_contas,
            text=f"Valor Total: R$ {total_psicologo:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(
            janela_contas,
            text=f"Valor a Pagar: R$ {valor_psicologo:.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(janela_contas, text="", bg="#2C3E50").pack(pady=10)

        # Resumo Geral
        tk.Label(
            janela_contas,
            text="Resumo Geral:",
            font=("Arial", 16, "bold"),
            bg="#2C3E50",
            fg="#ECF0F1",
        ).pack(pady=5)

        tk.Label(
            janela_contas,
            text=f"Total Geral: R$ {(total_medico + total_psicologo):.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        tk.Label(
            janela_contas,
            text=f"Total a Pagar: R$ {(valor_medico + valor_psicologo):.2f}",
            bg="#2C3E50",
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(pady=2)

        self.center(janela_contas)


    def processar_notas_fiscais(self):
        driver = webdriver.Chrome()
        cpfs = {"medico": [], "psicologo": [], "ambos": []}

        try:
            # Ler o arquivo Excel
            logging.info("Lendo o arquivo Excel")
            df = pd.read_excel(
                self.file_path, skiprows=1, sheet_name="17.10", dtype={"Renach": str}
            )
        except Exception as e:
            logging.error(f"Erro ao ler o arquivo Excel: {e}")
            messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
            return

        # logging.info(f'DataFrame lido: {df.head()}')
        logging.info("DataFrame lido!")

        try:
            renach_c = df.iloc[:, 2].dropna().tolist()
            renach_i = df.iloc[:, 8].dropna().tolist()

            # Acessando o site e fazendo login
            logging.info("Acessando o site do DETRAN e fazendo login")
            driver.get("https://clinicas.detran.ba.gov.br/")
            usuario = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="documento"]'))
            )
            doc = "11599160000115"
            for numero in doc:
                usuario.send_keys(numero)

            actions = ActionChains(driver)
            actions.send_keys(Keys.TAB).perform()
            time.sleep(1)

            senha = "475869"
            campo_senha = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="senha"]'))
            )
            for numero in senha:
                campo_senha.send_keys(numero)

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="acessar"]'))
            ).click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/aside/section/ul/li[2]/a/span[1]")
                )
            ).click()
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "/html/body/aside/section/ul/li[2]/ul/li/a/span")
                )
            ).click()

            barra_pesquisa = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="list_items_filter"]/label/input')
                )
            )

            # coletando informações do cliente
            logging.info("Coletando informações de CPFs")

            def coletar_cpf(dados, tipo):
                for dado in dados:
                    dado = str(dado).strip()
                    barra_pesquisa.clear()
                    barra_pesquisa.send_keys(dado)
                    time.sleep(2)
                    try:
                        paciente = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located(
                                (By.XPATH, '//*[@id="list_items"]/tbody/tr/td[3]')
                            )
                        )
                        cpf = paciente.text

                        print(f"coletado dado: {dado}, tipo: {tipo}")

                        if tipo == "medico" and dado in renach_i:
                            cpfs["ambos"].append(cpf)
                        elif tipo == "medico":
                            cpfs["medico"].append(cpf)
                        elif tipo == "psicologo" and cpf not in cpfs["ambos"]:
                            cpfs["psicologo"].append(cpf)
                    except Exception as e:
                        logging.error(f"Error ao coletar CPF: {e}")

            coletar_cpf(renach_c, "medico")
            coletar_cpf(renach_i, "psicologo")

            cpfs["medico"] = [cpf for cpf in cpfs["medico"] if cpf not in cpfs["ambos"]]
            cpfs["psicologo"] = [
                cpf for cpf in cpfs["psicologo"] if cpf not in cpfs["ambos"]
            ]

            logging.info("Acessando site para emissão de NTFS-e")
            driver.get("https://nfse.salvador.ba.gov.br/")

            # usuario
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="txtLogin"]'))
            ).send_keys("11599160000115")
            # senha
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="txtSenha"]'))
            ).send_keys("486258camp")
            # esperar resolver o captcha
            WebDriverWait(driver, 30).until(
                EC.invisibility_of_element_located((By.XPATH, '//*[@id="img1"]'))
            )
            # emissao NFS-e
            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//*[@id="menu-lateral"]/li[1]/a')
                )
            ).click()

            def emitir_nota(cpf, tipo):
                try:
                    barra_pesquisa = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, '//*[@id="tbCPFCNPJTomador"]')
                        )
                    )
                    barra_pesquisa.clear()
                    barra_pesquisa.send_keys(cpf)
                    WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btAvancar"]'))
                    ).click()

                    # Complete as informações da nota fiscal
                    role_e_click(driver, '//*[@id="ddlCNAE_chosen"]/a')
                    print("cnae clicada")
                    # opcao cane
                    WebDriverWait(driver, 30).until(
                        EC.visibility_of_element_located(
                            (By.XPATH, '//*[@id="ddlCNAE_chosen"]/div/ul/li[2]')
                        )
                    ).click()
                    print("opcao cnae visivel")
                    # aliq %
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//*[@id="tbAliquota"]')
                        )
                    ).send_keys("2,5")

                    servicos = {
                        "ambos": "Exame de sanidade física e mental",
                        "psicologo": "Exame de sanidade mental",
                        "medico": "Exame de sanidade física",
                    }
                    # preenchendo o tipo de serviço
                    tipo_servico = servicos.get(
                        tipo, "Exame de sanidade física"
                    )  # valor padrao, caso 'tipo' nao esteja no dicionario

                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.XPATH, '//*[@id="tbDiscriminacao"]')
                        )
                    ).send_keys(tipo_servico)

                    valor_nota = (
                        "148,65"
                        if tipo == "medico"
                        else "192,61" if tipo == "psicolgo" else "341,26"
                    )
                    # valor pago na consulta
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="tbValor"]'))
                    ).send_keys(valor_nota)
                    # emitindo nota
                    WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btEmitir"]'))
                    ).click()
                    # aceitando o alerta
                    WebDriverWait(driver, 20).until(EC.alert_is_present())
                    alert = Alert(driver)
                    alert.accept()
                    # botao voltar - voltando para emissao de nota fiscal por cpf
                    WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="btVoltar"]'))
                    ).click()
                    logging.info(f"Nota emitida para o CPF: {cpf}, Valor: {valor_nota}")

                except Exception as e:
                    logging.error("Erro ao emitir nota: {e}")

            # emitir notas
            try:
                for cpf in cpfs["medico"]:
                    emitir_nota(cpf, "medico")
                for cpf in cpfs["psicologo"]:
                    emitir_nota(cpf, "psicologo")
                for cpf in cpfs["ambos"]:
                    emitir_nota(cpf, "ambos")
            except Exception as e:
                logging.error(f"Erro na emissao das notas: {e}")
        finally:
            driver.quit()
            logging.info("Processo finalizado")
            return cpfs
            
    def exibir_informacao(self):
        """Exibe informações dos pacientes em uma interface organizada com opções de filtragem e detalhes de pagamento."""
        try:
            # Carrega o workbook e seleciona a sheet correta de forma segura
            wb = self.get_active_workbook()
            
            try:
                if hasattr(self.planilhas, 'sheet_name') and self.planilhas.sheet_name:
                    ws = wb[self.planilhas.sheet_name]
                else:
                    ws = wb.active
            except KeyError:
                ws = wb.active
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao acessar a planilha: {str(e)}")
                if wb:
                    wb.close()
                return
            
            if not ws:
                messagebox.showerror("Erro", "Não foi possível encontrar uma planilha válida.")
                if wb:
                    wb.close()
                return
            
            # Coleta e estrutura os dados
            medico, psi = [], []
            
            def processar_pagamento(valor):
                """Processa e formata informações de pagamento."""
                if not valor:
                    return ""
                valor_str = str(valor).strip()
                if not valor_str:
                    return ""
                
                # Se for um valor numérico, formata como moeda
                try:
                    valor_float = float(valor_str.replace("R$", "").replace(",", ".").strip())
                    return f"R$ {valor_float:.2f}"
                except ValueError:
                    return valor_str

            # Coleta dados do médico (colunas B-F)
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=2, max_col=6):
                nome = row[0].value
                if nome and isinstance(nome, (str, int)) and str(nome).strip():
                    medico.append({
                        'nome': str(nome).strip(),
                        'renach': str(row[1].value).strip() if row[1].value else '',
                        'forma_pagamento': processar_pagamento(row[4].value)
                    })

            # Coleta dados do psicólogo (colunas H-L)
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=8, max_col=12):
                nome = row[0].value
                if nome and isinstance(nome, (str, int)) and str(nome).strip():
                    psi.append({
                        'nome': str(nome).strip(),
                        'renach': str(row[1].value).strip() if row[1].value else '',
                        'forma_pagamento': processar_pagamento(row[4].value)
                    })

            wb.close()

            if not medico and not psi:
                messagebox.showinfo("Aviso", "Nenhuma informação encontrada!")
                return

            # Criando a janela principal
            janela_informacao = tk.Toplevel(self.master)
            janela_informacao.title("Informações dos Pacientes")
            janela_informacao.geometry("1200x800")
            cor_fundo = self.master.cget("bg")
            cor_texto = "#ECF0F1"
            cor_header = "#2C3E50"
            cor_destaque = "#34495E"
            janela_informacao.configure(bg=cor_fundo)

            # Frame principal
            main_frame = tk.Frame(janela_informacao, bg=cor_fundo)
            main_frame.pack(fill="both", expand=True, padx=20, pady=10)

            # Frame superior para controles
            control_frame = tk.Frame(main_frame, bg=cor_fundo)
            control_frame.pack(fill="x", pady=(0, 10))

            # Frame para a tabela com scrollbar
            table_container = tk.Frame(main_frame)
            table_container.pack(fill="both", expand=True)

            # Canvas e scrollbar
            canvas = tk.Canvas(table_container, bg=cor_fundo)
            scrollbar = tk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
            
            # Frame para a tabela
            table_frame = tk.Frame(canvas, bg=cor_fundo)
            
            def aplicar_filtros(*args):
                """Aplica os filtros de tipo e busca aos dados."""
                # Limpa a tabela atual
                for widget in table_frame.winfo_children():
                    if int(widget.grid_info()['row']) > 0:  # Preserva o cabeçalho
                        widget.destroy()

                dados_filtrados = []
                filtro = filtro_var.get()
                termo_busca = busca_var.get().lower()
                forma_pagamento_filtro = forma_pagamento_var.get().lower()

                def check_filtros(pac, tipo):
                    """Verifica se o paciente atende aos critérios de filtro."""
                    busca_match = (termo_busca in str(pac['nome']).lower() or 
                                termo_busca in str(pac['renach']).lower())
                    pagamento_match = (not forma_pagamento_filtro or 
                                    forma_pagamento_filtro in str(pac['forma_pagamento']).lower())
                    return busca_match and pagamento_match

                # Aplica os filtros
                if filtro in ["todos", "medico"]:
                    for i, pac in enumerate(medico):
                        if check_filtros(pac, "Médico"):
                            dados_filtrados.append((i+1, pac, "Médico"))

                if filtro in ["todos", "psi"]:
                    offset = len(medico) if filtro == "todos" else 0
                    for i, pac in enumerate(psi):
                        if check_filtros(pac, "Psicólogo"):
                            dados_filtrados.append((i+1+offset, pac, "Psicólogo"))

                # Atualiza estatísticas
                total_filtrado = len(dados_filtrados)
                stats_text = (f"Exibindo: {total_filtrado} paciente(s) | "
                            f"Total Geral: {len(medico) + len(psi)} | "
                            f"Médico: {len(medico)} | Psicólogo: {len(psi)}")
                stats_label.config(text=stats_text)

                # Preenche a tabela
                for row, (num, pac, tipo) in enumerate(dados_filtrados, start=1):
                    # Define cor de fundo alternada para melhor legibilidade
                    bg_color = cor_destaque if row % 2 == 0 else cor_fundo

                    items = [
                        (str(num), "center", 5),
                        (pac['nome'], "w", 30),
                        (pac['renach'], "center", 10),
                        (pac['forma_pagamento'], "w", 20),
                        (tipo, "center", 10)
                    ]

                    for col, (text, anchor, width) in enumerate(items):
                        label = tk.Label(
                            table_frame,
                            text=text,
                            bg=bg_color,
                            fg=cor_texto,
                            font=("Arial", 10),
                            padx=10,
                            pady=5,
                            anchor=anchor,
                            width=width
                        )
                        label.grid(row=row, column=col, sticky="nsew", padx=1, pady=1)

                # Atualiza a região de rolagem
                table_frame.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))

            # Variáveis de controle
            filtro_var = tk.StringVar(value="todos")
            busca_var = tk.StringVar()
            forma_pagamento_var = tk.StringVar()

            # Trace para as variáveis
            for var in [busca_var, forma_pagamento_var]:
                var.trace("w", aplicar_filtros)

            # Frame para os filtros
            filtros_frame = tk.Frame(control_frame, bg=cor_fundo)
            filtros_frame.pack(fill="x", padx=5)

            # Tipo de atendimento
            tipo_frame = tk.LabelFrame(filtros_frame, text="Tipo de Atendimento", bg=cor_fundo, fg=cor_texto)
            tipo_frame.pack(side="left", padx=5, pady=5)

            for valor, texto in [("todos", "Todos"), ("medico", "Médico"), ("psi", "Psicólogo")]:
                tk.Radiobutton(
                    tipo_frame,
                    text=texto,
                    variable=filtro_var,
                    value=valor,
                    command=aplicar_filtros,
                    bg=cor_fundo,
                    fg=cor_texto,
                    selectcolor=cor_header,
                    activebackground=cor_fundo,
                    activeforeground=cor_texto
                ).pack(side="left", padx=5)

            # Busca
            busca_frame = tk.LabelFrame(filtros_frame, text="Buscar por Nome/RENACH", bg=cor_fundo, fg=cor_texto)
            busca_frame.pack(side="left", padx=5, pady=5)
            tk.Entry(busca_frame, textvariable=busca_var, width=30).pack(padx=5, pady=2)

            # Forma de pagamento
            pagamento_frame = tk.LabelFrame(filtros_frame, text="Filtrar por Forma de Pagamento", bg=cor_fundo, fg=cor_texto)
            pagamento_frame.pack(side="left", padx=5, pady=5)
            tk.Entry(pagamento_frame, textvariable=forma_pagamento_var, width=20).pack(padx=5, pady=2)

            # Configuração do canvas
            canvas.create_window((0, 0), window=table_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Cabeçalho da tabela
            headers = [
                ("Nº", 5),
                ("Nome", 30),
                ("RENACH", 10),
                ("Forma de Pagamento", 20),
                ("Tipo", 10)
            ]

            for col, (header, width) in enumerate(headers):
                tk.Label(
                    table_frame,
                    text=header,
                    bg=cor_header,
                    fg=cor_texto,
                    font=("Arial", 11, "bold"),
                    padx=10,
                    pady=8,
                    relief="raised",
                    width=width
                ).grid(row=0, column=col, sticky="nsew", padx=1, pady=1)

            # Frame de estatísticas
            stats_frame = tk.Frame(main_frame, bg=cor_fundo)
            stats_frame.pack(fill="x", pady=10)
            
            stats_label = tk.Label(
                stats_frame,
                text="",
                bg=cor_fundo,
                fg=cor_texto,
                font=("Arial", 10, "bold")
            )
            stats_label.pack(pady=5)

            # Configuração final do layout
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # Configuração de eventos de rolagem
            def on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            def on_closing():
                canvas.unbind_all("<MouseWheel>")
                canvas.unbind_all("<Button-4>")
                canvas.unbind_all("<Button-5>")
                janela_informacao.destroy()

            # Configuração do scroll do mouse
            if sys.platform.startswith("win") or sys.platform == "darwin":
                canvas.bind_all("<MouseWheel>", on_mousewheel)
            else:
                canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
                canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

            janela_informacao.protocol("WM_DELETE_WINDOW", on_closing)

            # Aplica os filtros iniciais
            aplicar_filtros()

            # Centraliza a janela
            self.center(janela_informacao)

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao exibir as informações: {str(e)}")
            if 'wb' in locals():
                wb.close()

    def exibir_resultado(self):
        """Exibe os resultados detalhados de contagem e valores por forma de pagamento."""
        janela_exibir_resultado = tk.Toplevel(self.master)
        janela_exibir_resultado.geometry("500x600")
        janela_exibir_resultado.maxsize(width=500, height=600)
        janela_exibir_resultado.minsize(width=500, height=600)

        # Usando a cor de fundo da janela principal
        cor_fundo = self.master.cget("bg")
        janela_exibir_resultado.configure(bg=cor_fundo)
        
        def processar_pagamentos(pagamento_str):
            """Processa uma string de pagamento e retorna um dicionário com os valores."""
            resultado = {'D': [], 'C': [], 'E': [], 'P': []}
            if not pagamento_str or not isinstance(pagamento_str, str):
                return resultado

            try:
                # Dividir por pipes para separar múltiplos pagamentos
                for pagamento in [p.strip() for p in pagamento_str.split('|')]:
                    # Encontrar primeira ocorrência de cada forma de pagamento
                    for forma in ['D', 'C', 'E', 'P']:
                        if forma in pagamento:
                            # Se tem valor associado (formato FORMA:VALOR)
                            if ':' in pagamento:
                                try:
                                    # Pegar o valor após a forma de pagamento
                                    valor_str = pagamento[pagamento.find(':') + 1:].strip()
                                    # Limpar e converter o valor
                                    valor_str = valor_str.replace('R$', '').replace(' ', '').replace(',', '.')
                                    valor = float(valor_str)
                                    resultado[forma].append(valor)
                                except (ValueError, IndexError):
                                    print(f"Erro ao processar valor em: {pagamento}")
                            else:
                                # Se não tem valor especificado, usa valor padrão
                                resultado[forma].append(148.65)  # Valor padrão
            except Exception as e:
                print(f"Erro ao processar pagamento: {str(e)}")

            return resultado

        def calcular_totais(lista_pagamentos):
            """Calcula os totais para cada forma de pagamento."""
            totais = {'D': [], 'C': [], 'E': [], 'P': []}
            
            try:
                for pagamento in lista_pagamentos:
                    valores_pagamento = processar_pagamentos(pagamento)
                    for forma, valores in valores_pagamento.items():
                        totais[forma].extend(valores)
                
                return {
                    forma: {
                        'quantidade': len(valores),
                        'valor_total': sum(valores),
                        'valores': valores,
                        'media': sum(valores) / len(valores) if valores else 0
                    }
                    for forma, valores in totais.items()
                    if valores  # Só inclui formas que têm valores
                }
            except Exception as e:
                print(f"Erro ao calcular totais: {str(e)}")
                return {}

        # Obter dados do médico e psicólogo
        wb = self.get_active_workbook()
        ws = wb.active

        def contar_pacientes_e_coletar_pagamentos(col_inicial, col_final):
            """Conta pacientes e coleta pagamentos de uma seção."""
            pagamentos = []
            n_pacientes = 0
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
                nome = row[col_inicial-1].value
                if nome and isinstance(nome, str) and nome.strip():
                    n_pacientes += 1
                    pagamento = row[col_final-1].value
                    if pagamento and isinstance(pagamento, str):
                        pagamentos.append(pagamento)
            return n_pacientes, pagamentos

        # Coletar dados (colunas B-F para médico, H-L para psicólogo)
        n_medico, pagamentos_medico = contar_pacientes_e_coletar_pagamentos(2, 6)
        n_psicologo, pagamentos_psi = contar_pacientes_e_coletar_pagamentos(8, 12)

        # Calcular totais
        totais_medico = calcular_totais(pagamentos_medico)
        totais_psi = calcular_totais(pagamentos_psi)

        # Criar labels para exibição
        def criar_secao(titulo, n_pacientes, dados, row_start):
            """Cria uma seção de informações com observações para valores diferentes do padrão."""
            tk.Label(
                janela_exibir_resultado,
                text=titulo,
                font=("Arial", 16, "bold"),
                bg=cor_fundo,
                fg="#ECF0F1",
            ).pack(pady=(20 if row_start > 0 else 15, 5))

            tk.Label(
                janela_exibir_resultado,
                text=f"Total de Pacientes: {n_pacientes}",
                font=("Arial", 12),
                bg=cor_fundo,
                fg="#ECF0F1",
            ).pack(pady=(0, 10))

            total_secao = 0
            formas_nome = {'D': 'Débito', 'C': 'Crédito', 'E': 'Espécie', 'P': 'PIX'}
            valor_padrao = 148.65 if titulo == "MÉDICO" else 192.61
            
            for forma, info in dados.items():
                if info['quantidade'] > 0:
                    nome_forma = formas_nome.get(forma, forma)
                    
                    # Sumário da forma de pagamento
                    tk.Label(
                        janela_exibir_resultado,
                        text=f"{nome_forma}: {info['quantidade']} pagamento(s) - R$ {info['valor_total']:.2f}",
                        font=("Arial", 12),
                        bg=cor_fundo,
                        fg="#ECF0F1",
                    ).pack(pady=2)
                    
                    # Encontrar valores diferentes do padrão
                    valores_diferentes = [v for v in info['valores'] if abs(v - valor_padrao) > 0.01]
                    valores_unicos = set(valores_diferentes)  # Remove duplicatas
                    
                    # Criar observações para cada valor diferente
                    for valor in valores_unicos:
                        contagem = valores_diferentes.count(valor)
                        obs_texto = (f"Obs: {'um' if contagem == 1 else str(contagem)} " 
                                f"pagamento{'s' if contagem > 1 else ''} no valor de R$ {valor:.2f}")
                        
                        tk.Label(
                            janela_exibir_resultado,
                            text=obs_texto,
                            font=("Arial", 10, "italic"),
                            bg=cor_fundo,
                            fg="#ECF0F1",
                            wraplength=450
                        ).pack(pady=(0, 5))
                    
                    total_secao += info['valor_total']

            tk.Label(
                janela_exibir_resultado,
                text=f"Total da Seção: R$ {total_secao:.2f}",
                font=("Arial", 14, "bold"),
                bg=cor_fundo,
                fg="#ECF0F1",
            ).pack(pady=(5, 10))
            
            return total_secao

        # Exibir resultados
        total_med = criar_secao("MÉDICO", n_medico, totais_medico, 0)
        total_psi = criar_secao("PSICÓLOGO", n_psicologo, totais_psi, 1)

        # Total geral
        tk.Label(
            janela_exibir_resultado,
            text="TOTAL GERAL",
            font=("Arial", 16, "bold"),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack(pady=(20, 5))

        tk.Label(
            janela_exibir_resultado,
            text=f"R$ {(total_med + total_psi):.2f}",
            font=("Arial", 14, "bold"),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack(pady=(5, 20))

        self.center(janela_exibir_resultado)
        wb.close()


    def enviar_whatsapp(self):
        # Janela número ou nome do grupo
        janela_wpp = tk.Toplevel(self.master)
        janela_wpp.geometry("300x210")
        cor_fundo = self.master.cget("bg")
        janela_wpp.configure(bg=cor_fundo)
        self.center(janela_wpp)

        tk.Label(
            janela_wpp,
            text="Enviar para:",
            font=("Arial", 16, "bold"),
            bg=cor_fundo,
            fg="#ECF0F1",
        ).pack(anchor="center", padx=5, pady=5)

        self.wpp_entry = tk.Entry(janela_wpp)
        self.wpp_entry.pack(padx=5, pady=5)

        # Checkbutton para salvar as informações
        tk.Button(
            janela_wpp, text="Enviar", command=self.processar_envio_whatsapp
        ).pack(pady=10)


    def processar_envio_whatsapp(self):
        # Captura o valor do campo de entrada
        group_name = self.wpp_entry.get().strip()

        if not group_name:
            messagebox.showerror("Erro", "Insira um número, grupo ou nome")
            return

        # Preparar as informações para enviar a mensagem
        n_medico, pag_medico = self.planilhas.contar_medico()
        n_psicologo, pag_psicologo = self.planilhas.contar_psi()

        valor_medico = n_medico * 49
        valor_psicologo = n_psicologo * 63.50

        message_medico = f"Valor medico: {valor_medico}"
        message_psicologo = f"Valor psicologo: {valor_psicologo}"

        dir_path = os.getcwd()
        profile = os.path.join(dir_path, "profile", "wpp")

        # Configurar opções do Chrome
        logging.info("Configurando Chrome...")
        options = Options()
        options.add_argument(r"user-data-dir={}".format(profile))
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        # Inicializar o WebDriver
        logging.info("Inicializando o WebDriver...")
        service = Service(executable_path="/usr/local/bin/chromedriver")
        driver = webdriver.Chrome(service=service, options=options)

        # Acessar o WhatsApp Web
        logging.info("Abrindo WhatsApp...")
        driver.get("https://web.whatsapp.com/")

        # Aguardar até que a página esteja totalmente carregada
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"]'))
            )
        except Exception as e:
            messagebox.showerror("Erro" f"Erro ao esnaear QR Code: {e}")
            logging.error(f"Erro ao escanear QR Code: {e}")
            driver.quit()
            return

        # Selecionar o grupo
        try:
            logging.info("Enviando mensagem...")
            barra_pesquisa = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located(
                    (By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div/p')
                )
            )
            barra_pesquisa.send_keys(group_name)
            time.sleep(1)
            barra_pesquisa.send_keys(Keys.ENTER)
        except Exception as e:
            messagebox.showerror(f"Erro ao selecionar grupo: {e}")
            logging.error(f"Erro ao seleconar o grupo: {e}")
            driver.quit()
            return

        # Enviar a mensagem
        try:
            logging.info("Enviando mensagem...")
            enviar_mensagem = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p',
                    )
                )
            )
            enviar_mensagem.send_keys(message_medico)
            time.sleep(1)
            enviar_mensagem.send_keys(Keys.ENTER)

            enviar_mensagem = WebDriverWait(driver, 30).until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div[1]/div/div[1]/p',
                    )
                )
            )
            enviar_mensagem.send_keys(message_psicologo)
            time.sleep(1)
            enviar_mensagem.send_keys(Keys.ENTER)
            messagebox.showinfo("Mensagens enviadas")

            time.sleep(7)
        except Exception as e:
            messagebox.showerror(f"Erro ao enviar as mensagens: {e}")
        finally:
            driver.quit()


    def enviar_email(self):
        janela_email = tk.Toplevel(self.master)
        janela_email.geometry("300x400")
        cor_fundo = self.master.cget("bg")
        janela_email.configure(bg=cor_fundo)
        self.center(janela_email)

        tk.Label(
            janela_email,
            text="Email:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_email = tk.Entry(janela_email)
        entry_email.pack(pady=5)

        tk.Label(
            janela_email,
            text="Senha:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_senha = tk.Entry(janela_email, show="*")  # Ocultar senha
        entry_senha.pack(pady=5)

        tk.Label(
            janela_email,
            text="Destinatário:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_destinatario = tk.Entry(janela_email)
        entry_destinatario.pack(pady=5)

        tk.Label(
            janela_email,
            text="Assunto:",
            bg=cor_fundo,
            fg="#ECF0F1",
            font=("Arial", 14, "bold"),
        ).pack(pady=5)
        entry_assunto = tk.Entry(janela_email)
        entry_assunto.pack(pady=5)

        tk.Button(
            janela_email,
            text="Selecionar XLSX",
            command=lambda: self.selecionar_xlsx(
                entry_email.get(),
                entry_senha.get(),
                entry_destinatario.get(),
                entry_assunto.get(),
            ),
        ).pack(pady=20)


    def selecionar_xlsx(self, email, senha, destinatario, assunto):
        """
        Abre diálogo para selecionar arquivo XLSX
        """
        if not all([email, senha, destinatario, assunto]):
            messagebox.showerror("Erro", "Preencha todos os campos!")
            return

        arquivo_xlsx = filedialog.askopenfilename(
            title="Selecione o arquivo XLSX",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
        )

        if arquivo_xlsx:
            self.enviar(email, senha, destinatario, assunto, arquivo_xlsx)


    def enviar(self, email, senha, destinatario, assunto, caminho_xlsx):
        """
        Envia e-mail com arquivo XLSX anexado
        """
        smtp_server = "smtp.gmail.com"  # Para Gmail
        smtp_port = 587

        try:
            # Criando a mensagem
            msg = MIMEMultipart()
            msg["From"] = email
            msg["To"] = destinatario
            msg["Subject"] = assunto

            # Corpo do e-mail padrão
            corpo = "Segue em anexo o arquivo XLSX conforme solicitado."
            msg.attach(MIMEText(corpo, "plain"))

            # Anexar arquivo XLSX
            with open(caminho_xlsx, "rb") as arquivo:
                parte_xlsx = MIMEApplication(arquivo.read(), _subtype="xlsx")
                parte_xlsx.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=os.path.basename(caminho_xlsx),
                )
                msg.attach(parte_xlsx)

            # Contexto SSL para conexão segura
            context = ssl.create_default_context()

            # Enviando o e-mail
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls(context=context)  # Inicia a segurança TLS
                server.login(email, senha)  # Faz login no servidor
                server.send_message(msg)  # Envia a mensagem

            # Mensagem de sucesso
            messagebox.showinfo(
                "Sucesso",
                f"E-mail enviado com sucesso para {destinatario}!\nAnexo: {os.path.basename(caminho_xlsx)}",
            )

        except smtplib.SMTPAuthenticationError:
            messagebox.showerror(
                "Erro de Autenticação",
                "Verifique seu email e senha. Use uma senha de aplicativo para o Gmail.",
            )
        except Exception as e:
            messagebox.showerror("Erro ao Enviar", f"Ocorreu um erro: {str(e)}")


    def configurar_frames(self, login_frame, criar_conta_frame):
        self.login_frame = login_frame
        self.criar_conta_frame = criar_conta_frame


    def mostrar_criar_conta(self):
        """Alterna para o frame de criar conta."""
        self.login_frame.hide()
        self.criar_conta_frame.show()


    def voltar_para_login(self):
        """Alterna de volta para o frame de login."""
        self.criar_conta_frame.hide()
        self.login_frame.show()


    def formatar_planilha(self):
        """
        Formata a planilha com os dados do usuário e data atual.
        """
        try:
            wb = self.get_active_workbook()
            if not wb:
                return False
                
            ws = wb.active
            if not ws:
                return False

            # Remover todas as mesclagens existentes (conversão para lista)
            for range_str in list(ws.merged_cells.ranges):
                ws.unmerge_cells(str(range_str))

            # Define estilos base de borda
            borda = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )

            font_bold = Font(name="Arial", bold=True, size=11, color="000000")
            font_regular = Font(name="Arial", size=11, color="000000")
            alignment_center = Alignment(horizontal="center", vertical="center")
            alignment_left = Alignment(horizontal="left", vertical="center")

            # Limpar a formatação existente
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=12):
                for cell in row:
                    cell.border = None
                    cell.font = font_regular
                    cell.alignment = alignment_left

            # Configurar larguras das colunas
            ws.column_dimensions['B'].width = 55
            ws.column_dimensions['C'].width = 13
            ws.column_dimensions['H'].width = 55
            ws.column_dimensions['I'].width = 13

            # Configurar cabeçalhos
            data_atual = datetime.now().strftime("%d/%m/%Y")
            usuario = self.current_user if hasattr(self, 'current_user') and self.current_user else "Usuário"

            # Cabeçalho da seção médico
            ws['A1'] = f"({usuario}) Atendimento Médico {data_atual}"
            ws.merge_cells('A1:E1')
            ws['A1'].font = font_bold
            ws['A1'].alignment = alignment_center

            # Cabeçalho da seção psicólogo
            ws['G1'] = f"({usuario}) Atendimento Psicológico {data_atual}"
            ws.merge_cells('G1:K1')
            ws['G1'].font = font_bold
            ws['G1'].alignment = alignment_center

            # Cabeçalhos das colunas
            headers = ["Ordem", "Nome", "Renach", "Reexames", "Valor"]
            for start_col in [1, 7]:  # A e G
                for idx, header in enumerate(headers):
                    cell = ws.cell(row=2, column=start_col + idx)
                    cell.value = header
                    cell.font = font_bold
                    cell.alignment = alignment_center
                    cell.border = borda

            # Processar linhas de dados
            def processar_dados(start_col):
                num_pacientes = 0
                ultima_linha = 2
                
                for row in range(3, ws.max_row + 1):
                    nome = ws.cell(row=row, column=start_col + 1).value
                    if nome and str(nome).strip():
                        num_pacientes += 1
                        ultima_linha = row
                        
                        # Valor fixo
                        valor = 148.65 if start_col == 1 else 192.61
                        valor_cell = ws.cell(row=row, column=start_col + 4)
                        valor_cell.value = valor
                        valor_cell.number_format = '"R$"#,##0.00'
                        valor_cell.alignment = alignment_center
                        
                        # Formatar linha
                        for col in range(start_col, start_col + 5):
                            cell = ws.cell(row=row, column=col)
                            cell.border = borda
                            if col == start_col + 1:  # Coluna nome
                                cell.alignment = alignment_left
                            else:
                                cell.alignment = alignment_center
                
                return num_pacientes, ultima_linha

            # Processar dados de médicos e psicólogos
            num_medico, ultima_linha_med = processar_dados(1)
            num_psi, ultima_linha_psi = processar_dados(7)

            # Limpar área após os dados
            ultima_linha = max(ultima_linha_med, ultima_linha_psi)
            for row in range(ultima_linha + 1, ws.max_row + 1):
                for col in range(1, 12):
                    cell = ws.cell(row=row, column=col)
                    cell.value = None
                    cell.border = None

            # Adicionar resumo se houver dados
            if num_medico > 0 or num_psi > 0:
                linha_atual = ultima_linha + 2
                
                # Cálculos
                total_medico = num_medico * 148.65
                total_psi = num_psi * 192.61
                pagamento_medico = num_medico * 49.00
                pagamento_psi = num_psi * 63.50
                
                resumo = [
                    ("Atendimento Médico", total_medico),
                    ("Atendimento Psicológico", total_psi),
                    ("Total", total_medico + total_psi),
                    None,  # Linha em branco
                    ("Pagamento Médico", pagamento_medico),
                    ("Pagamento Psicológico", pagamento_psi),
                    ("Total Clínica", (total_medico + total_psi) - (pagamento_medico + pagamento_psi))
                ]

                for item in resumo:
                    if item is None:
                        linha_atual += 1
                        continue
                        
                    texto, valor = item
                    
                    # Limpar área do resumo primeiro
                    for col in range(8, 11):
                        cell = ws.cell(row=linha_atual, column=col)
                        cell.value = None
                        cell.border = None
                    
                    # Adicionar texto
                    texto_cell = ws.cell(row=linha_atual, column=8)
                    texto_cell.value = texto
                    texto_cell.font = font_bold
                    texto_cell.alignment = alignment_center
                    
                    ws.merge_cells(f'H{linha_atual}:J{linha_atual}')
                    
                    # Adicionar valor
                    valor_cell = ws.cell(row=linha_atual, column=10)
                    valor_cell.value = valor
                    valor_cell.number_format = '"R$"#,##0.00'
                    valor_cell.font = font_bold
                    valor_cell.border = borda
                    valor_cell.alignment = alignment_center
                    
                    linha_atual += 1
                    
            wb.save(self.file_path)
            return True

        except Exception as e:
            logging.error(f"Erro ao formatar planilha: {str(e)}")
            if 'wb' in locals():
                wb.close()
            return False


    def _adicionar_totais(self, ws, linha_inicio, col_inicio, valor_consulta, valor_profissional, num_pacientes, borda, font_bold, alignment_center):
        """Adiciona os totais para uma seção (médico ou psicólogo)"""
        # Soma
        ws.cell(row=linha_inicio, column=col_inicio + 2, value="Soma")
        soma_cell = ws.cell(row=linha_inicio, column=col_inicio + 4, value=valor_consulta * num_pacientes)
        soma_cell.number_format = '"R$"#,##0.00'
        
        # Valor profissional
        ws.cell(row=linha_inicio + 1, column=col_inicio + 2, value="Profissional")
        prof_cell = ws.cell(row=linha_inicio + 1, column=col_inicio + 4, value=valor_profissional * num_pacientes)
        prof_cell.number_format = '"R$"#,##0.00'
        
        # Total
        ws.cell(row=linha_inicio + 2, column=col_inicio + 2, value="Total")
        total_cell = ws.cell(row=linha_inicio + 2, column=col_inicio + 4, 
                            value=(valor_consulta - valor_profissional) * num_pacientes)
        total_cell.number_format = '"R$"#,##0.00'
        
        # Aplica formatação
        for row in range(linha_inicio, linha_inicio + 3):
            for col in range(col_inicio + 2, col_inicio + 5):
                cell = ws.cell(row=row, column=col)
                cell.border = borda
                cell.font = font_bold
                cell.alignment = alignment_center


    def _adicionar_resumo_geral(self, ws, linha_inicio, num_medico, num_psi, borda, font_bold, alignment_center):
        """Adiciona o resumo geral na planilha"""
        # Configurações
        valor_medico = 148.65
        valor_psi = 192.61
        pagamento_medico = 49.00
        pagamento_psi = 63.50
        
        # Cálculos
        total_medico = num_medico * valor_medico
        total_psi = num_psi * valor_psi
        total_geral = total_medico + total_psi
        
        pagamento_total_medico = num_medico * pagamento_medico
        pagamento_total_psi = num_psi * pagamento_psi
        total_clinica = total_geral - (pagamento_total_medico + pagamento_total_psi)
        
        # Lista de valores a serem adicionados
        resumo = [
            ("Atendimento Médico", total_medico),
            ("Atendimento Psicológico", total_psi),
            ("Total", total_geral),
            ("", None),  # Linha em branco
            ("Pagamento Médico", pagamento_total_medico),
            ("Pagamento Psicológico", pagamento_total_psi),
            ("Total Clínica", total_clinica)
        ]
        
        # Adiciona os valores
        for idx, (texto, valor) in enumerate(resumo):
            if texto:  # Pula linha em branco
                ws.cell(row=linha_inicio + idx, column=8, value=texto)
                if valor is not None:
                    valor_cell = ws.cell(row=linha_inicio + idx, column=10, value=valor)
                    valor_cell.number_format = '"R$"#,##0.00'
                
                # Aplica formatação
                for col in [8, 9, 10]:
                    cell = ws.cell(row=linha_inicio + idx, column=col)
                    cell.border = borda
                    cell.font = font_bold
                    cell.alignment = alignment_center
                
                # Mescla células para o texto
                ws.merge_cells(f"H{linha_inicio + idx}:I{linha_inicio + idx}")


class SistemaContas:
    def __init__(self, file_path: str, current_user=None):
        self.file_path = file_path
        self.current_user = current_user
        self.sheet_name = "Contas Fechamento"
        self.criar_sheet_se_nao_existir()


    def abrir_janela(self):
        """Cria uma nova janela para o sistema de contas"""
        self.window = tk.Toplevel()
        self.window.title("Sistema de Gerenciamento de Contas")
        self.window.geometry("500x400")
        self.criar_interface()

        # Configurar a janela como modal
        self.window.transient(self.window.master)
        self.window.grab_set()
        self.window.focus_set()


    def criar_sheet_se_nao_existir(self):
        """Cria a planilha e a aba (sheet) se não existirem."""
        if os.path.exists(self.file_path):
            wb = load_workbook(self.file_path)
            if self.sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(title=self.sheet_name)
                ws.append(["DATA", "CONTAS", "VALOR"])
                wb.save(self.file_path)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = self.sheet_name
            ws.append(["DATA", "CONTAS", "VALOR"])
            wb.save(self.file_path)


    def criar_interface(self):
        """Cria a interface gráfica usando grid layout"""
        # Configurando o frame principal
        main_frame = tk.Frame(self.window, padx=20, pady=20)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Configurando expansão do grid
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # Título
        title_label = tk.Label(
            main_frame, text="Gerenciamento de Contas", font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Data
        tk.Label(main_frame, text="Data:", font=("Arial", 10, "bold")).grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.date_entry = DateEntry(
            main_frame,
            width=20,
            date_pattern="dd/mm/yyyy",
            background="darkblue",
            foreground="white",
            borderwidth=2,
        )
        self.date_entry.grid(row=1, column=1, sticky="we", padx=(5, 0), pady=5)

        # Descrição
        tk.Label(main_frame, text="Descrição:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky="w", pady=5
        )
        self.info_entry = tk.Entry(main_frame)
        self.info_entry.grid(row=2, column=1, sticky="we", padx=(5, 0), pady=5)

        # Valor
        tk.Label(main_frame, text="Valor (R$):", font=("Arial", 10, "bold")).grid(
            row=3, column=0, sticky="w", pady=5
        )
        self.valor_entry = tk.Entry(main_frame)
        self.valor_entry.grid(row=3, column=1, sticky="we", padx=(5, 0), pady=5)

        # Frame para botões
        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        button_frame.grid_columnconfigure((0, 1), weight=1)

        # Botões
        save_button = tk.Button(
            button_frame,
            text="Salvar",
            command=self.capturar_dados,
            width=20,
            bg="#4CAF50",
            fg="white",
        )
        save_button.grid(row=0, column=0, padx=5)

        clear_button = tk.Button(
            button_frame, text="Limpar", command=self.limpar_campos, width=20
        )
        clear_button.grid(row=0, column=1, padx=5)

        # Botão Fechar
        close_button = tk.Button(
            button_frame, text="Fechar", command=self.window.destroy, width=20
        )
        close_button.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        # Frame para mensagens de status
        self.status_frame = tk.Frame(main_frame)
        self.status_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=(10, 0))

        self.status_label = tk.Label(self.status_frame, text="", foreground="green")
        self.status_label.grid(row=0, column=0, sticky="we")

        # Configurar foco inicial
        self.info_entry.focus()


    def salvar_informacoes(self, data_escolhida, info, valor):
        """Salva as informações na planilha, agrupando por data e colocando informações na mesma célula."""
        try:
            wb = load_workbook(self.file_path)
            ws = wb[self.sheet_name]

            try:
                data_formatada = datetime.strptime(data_escolhida, "%d/%m/%Y").date()
            except ValueError:
                messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA")
                return False

            dados = []
            for row in ws.iter_rows(min_row=2):
                if row[0].value:
                    dados.append(
                        {
                            "data": row[0].value.date(),
                            "info": row[1].value,
                            "valor": row[2].value,
                            "linha": row[0].row,
                        }
                    )

            data_existe = False
            for i, dado in enumerate(dados):
                if dado["data"] == data_formatada:
                    data_existe = True
                    dados[i]["info"] = (
                        f"{dado['info']}\n{info}" if dado["info"] else info
                    )
                    dados[i]["valor"] = (
                        f"{dado['valor']}\n{valor}" if dado["valor"] else valor
                    )
                    break

            if not data_existe:
                dados.append(
                    {
                        "data": data_formatada,
                        "info": info,
                        "valor": valor,
                        "linha": None,
                    }
                )

            dados_ordenados = sorted(dados, key=lambda x: x["data"])

            # Limpa os dados existentes
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.value = None

            for i, dado in enumerate(dados_ordenados, start=2):
                ws.cell(row=i, column=1).value = dado["data"]
                ws.cell(row=i, column=2).value = dado["info"]

                # Atribui o valor à célula e formata como moeda
                cell_valor = ws.cell(row=i, column=3)
                if dado["valor"] is not None:
                    cell_valor.value = dado["valor"]  # Aqui você armazena o valor
                    cell_valor.number_format = '"R$"#,##0.00'  # Formato de moeda
                else:
                    cell_valor.value = valor  # Caso não tenha dado anterior
                    cell_valor.number_format = '"R$"#,##0.00'  # Formato de moeda

                # Centraliza a data
                ws.cell(row=i, column=1).alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            # Ajusta a formatação de texto e alinhamento
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="center")

            # Ajusta a largura das colunas
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = max_length + 2
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(self.file_path)
            messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")
            return True

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar informações: {str(e)}")
            return False


    def validar_campos(self):
        """Valida os campos antes de salvar"""
        info = self.info_entry.get().strip()
        valor = self.valor_entry.get().strip()
        data = self.date_entry.get().strip()

        if not all([data, info, valor]):
            messagebox.showerror("Erro", "Todos os campos são obrigatórios!")
            return False

        try:
            float(valor.replace(",", "."))
            return True
        except ValueError:
            messagebox.showerror("Erro", "O valor deve ser um número válido!")
            return False


    def limpar_campos(self):
        """Limpa os campos após salvar"""
        self.info_entry.delete(0, tk.END)
        self.valor_entry.delete(0, tk.END)


    def capturar_dados(self):
        """Captura e processa os dados do formulário"""
        if self.validar_campos():
            data = self.date_entry.get()
            info = self.info_entry.get()
            valor = self.valor_entry.get().replace(",", ".")

            try:
                # Converte o valor para float e formata como moeda
                valor_float = float(valor)
                valor_formatado = f"R$ {valor_float:,.2f}"  # Formatação para moeda

                # Chama a função de salvar com o valor formatado
                if self.salvar_informacoes(data, info, valor_formatado):
                    self.limpar_campos()
            except ValueError:
                messagebox.showerror(
                    "Erro", "Por favor, insira um valor numérico válido."
                )


class GerenciadorPlanilhas:
    def __init__(self, master, sistema_contas):
        self.master = master
        self.sistema_contas = sistema_contas
        self.file_path = None
        self.sheet_name = None
        self.active_window = None


    def abrir_gerenciador(self):
        """Abre a janela de gerenciamento de planilhas"""
        if self.active_window:
            self.active_window.lift()
            return

        self.active_window = Toplevel(self.master)
        self.active_window.title("Gerenciador de Planilhas")
        self.active_window.geometry("600x700")
        self.active_window.resizable(False, False)

        # Centralizar a janela
        window_width = 600
        window_height = 700
        screen_width = self.active_window.winfo_screenwidth()
        screen_height = self.active_window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.active_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Configurar grid da janela
        self.active_window.grid_columnconfigure(0, weight=1)
        self.active_window.grid_rowconfigure(0, weight=1)

        self._setup_interface()

        # Cleanup quando a janela for fechada
        self.active_window.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Tornar a janela modal
        self.active_window.transient(self.master)
        self.active_window.grab_set()


    def _setup_interface(self):
        """Configura a interface do gerenciador"""
        # Frame principal com padding
        main_frame = ttk.Frame(self.active_window, padding="20 20 20 20")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        # Título
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        title_frame.grid_columnconfigure(0, weight=1)

        title_label = ttk.Label(
            title_frame,
            text="Gerenciador de Planilhas Excel",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0)

        # Frame para arquivo atual
        file_frame = ttk.LabelFrame(main_frame, text="Arquivo Atual", padding="10")
        file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        file_frame.grid_columnconfigure(0, weight=1)

        self.lbl_arquivo = ttk.Label(
            file_frame,
            text=(
                self.sistema_contas.file_path
                if hasattr(self.sistema_contas, "file_path")
                else "Nenhum arquivo selecionado"
            ),
            wraplength=500,
        )
        self.lbl_arquivo.grid(row=0, column=0, sticky="ew", padx=5)

        # Frame para lista de sheets
        list_frame = ttk.LabelFrame(
            main_frame, text="Planilhas Disponíveis", padding="10"
        )
        list_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        # Container para lista e scrollbars
        list_container = ttk.Frame(list_frame)
        list_container.grid(row=0, column=0, sticky="nsew")
        list_container.grid_columnconfigure(0, weight=1)
        list_container.grid_rowconfigure(0, weight=1)

        self.listbox = Listbox(
            list_container,
            font=("Arial", 10),
            selectmode=SINGLE,
            height=10,
            borderwidth=1,
            relief="solid",
        )
        self.listbox.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(
            list_container, orient=VERTICAL, command=self.listbox.yview
        )
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(
            list_container, orient=HORIZONTAL, command=self.listbox.xview
        )
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        self.listbox.configure(
            yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set
        )

        # Frame para criar nova sheet
        create_frame = ttk.LabelFrame(
            main_frame, text="Criar Nova Planilha", padding="10"
        )
        create_frame.grid(row=3, column=0, sticky="ew", pady=(0, 20))
        create_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(create_frame, text="Nome:", font=("Arial", 10)).grid(
            row=0, column=0, padx=(0, 10), sticky="w"
        )

        self.nova_sheet_entry = ttk.Entry(create_frame)
        self.nova_sheet_entry.grid(row=0, column=1, sticky="ew")

        # Frame para botões
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, sticky="ew")
        for i in range(2):
            button_frame.grid_columnconfigure(i, weight=1)

        # Primeira linha de botões
        ttk.Button(
            button_frame, text="Nova Planilha Excel", command=self.criar_nova_planilha
        ).grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        ttk.Button(
            button_frame, text="Abrir Planilha Existente", command=self.abrir_planilha
        ).grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Segunda linha de botões
        ttk.Button(
            button_frame, text="Selecionar Sheet", command=self.selecionar_sheet
        ).grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        ttk.Button(
            button_frame, text="Criar Nova Sheet", command=self.criar_nova_sheet
        ).grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        self.atualizar_lista_sheets()


    def criar_nova_planilha(self):
        """Cria um novo arquivo Excel"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )

        if file_path:
            try:
                wb = Workbook()
                wb.save(file_path)
                self.sistema_contas.file_path = file_path
                self.lbl_arquivo.config(text=file_path)
                self.atualizar_lista_sheets()
                messagebox.showinfo(
                    "Sucesso", "Nova planilha Excel criada com sucesso!"
                )
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao criar planilha: {str(e)}")


    def abrir_planilha(self):
        """Abre uma planilha Excel existente"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )

        if file_path:
            try:
                wb = load_workbook(file_path)
                self.sistema_contas.file_path = file_path

                # Pega a sheet ativa atual
                active_sheet = wb.active
                self.sistema_contas.sheet_name = active_sheet.title

                wb.close()
                self.lbl_arquivo.config(text=file_path)
                self.atualizar_lista_sheets()
                messagebox.showinfo(
                    "Sucesso",
                    f"Planilha aberta com sucesso! Sheet ativa: {active_sheet.title}",
                )
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir planilha: {str(e)}")


    def atualizar_lista_sheets(self):
        """Atualiza a lista de sheets disponíveis"""
        self.listbox.delete(0, END)
        if (
            hasattr(self.sistema_contas, "file_path")
            and self.sistema_contas.file_path
            and os.path.exists(self.sistema_contas.file_path)
        ):
            try:
                wb = load_workbook(self.sistema_contas.file_path)
                for sheet in wb.sheetnames:
                    self.listbox.insert(END, sheet)
                wb.close()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao listar planilhas: {str(e)}")


    def selecionar_sheet(self):
        """Seleciona uma sheet existente e a torna ativa"""
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showerror("Erro", "Selecione uma planilha!")
            return

        nome_sheet = self.listbox.get(selection[0])
        try:
            wb = load_workbook(self.sistema_contas.file_path)
            if nome_sheet in wb.sheetnames:
                # Define a sheet selecionada como ativa
                wb.active = wb[nome_sheet]
                wb.save(self.sistema_contas.file_path)

                # Atualiza o nome da sheet no sistema_contas
                self.sistema_contas.sheet_name = nome_sheet

                wb.close()
                messagebox.showinfo(
                    "Sucesso", f"Planilha '{nome_sheet}' selecionada e ativada!"
                )
                self.active_window.destroy()
                self.active_window = None
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar planilha: {str(e)}")


    def criar_nova_sheet(self):
        """Cria uma nova sheet e a torna ativa"""
        nome_sheet = self.nova_sheet_entry.get().strip()
        if not nome_sheet:
            messagebox.showerror("Erro", "Digite um nome para a nova planilha!")
            return

        if (
            not hasattr(self.sistema_contas, "file_path")
            or not self.sistema_contas.file_path
        ):
            messagebox.showerror("Erro", "Primeiro abra ou crie uma planilha Excel!")
            return

        try:
            wb = load_workbook(self.sistema_contas.file_path)
            if nome_sheet in wb.sheetnames:
                messagebox.showerror("Erro", "Já existe uma planilha com este nome!")
                wb.close()
                return

            # Cria nova sheet e a torna ativa
            new_sheet = wb.create_sheet(title=nome_sheet)
            wb.active = new_sheet
            wb.save(self.sistema_contas.file_path)
            wb.close()

            # Atualiza o nome da sheet no sistema_contas
            self.sistema_contas.sheet_name = nome_sheet

            self.atualizar_lista_sheets()
            messagebox.showinfo(
                "Sucesso", f"Planilha '{nome_sheet}' criada e ativada com sucesso!"
            )
            self.active_window.destroy()
            self.active_window = None
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao criar planilha: {str(e)}")


    def _on_closing(self):
        """Handler para quando a janela for fechada"""
        self.active_window.destroy()
        self.active_window = None
