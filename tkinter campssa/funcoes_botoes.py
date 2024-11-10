from dataclasses import dataclass
import sys
from typing import Dict, List, Optional, Tuple
import tkinter as tk
from tkinter import messagebox, filedialog, Frame, Label, Entry, Button, Toplevel, SINGLE, VERTICAL, HORIZONTAL, END, Listbox
from tkinter import ttk
from functools import lru_cache
import logging
import json
import os
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, Border, Side
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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import ssl
from datetime import datetime
from tkcalendar import DateEntry
import sqlite3
from database_connection import DatabaseConnection


# Configurando logs
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

"""Rola a página até um elemento e clica nele"""


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
    """
    SEÇÃO 1: INICIALIZAÇÃO E CONFIGURAÇÃO
    """

    # Código de inicialização...
    def __init__(self, master, planilhas, file_path, app, current_user=None):
        self.master = master
        self.planilhas = planilhas
        self.file_path = file_path
        self.app = app
        self.current_user = current_user
        self.login_frame = None
        self.criar_conta_frame = None
        self.logger = logging.getLogger(__name__)

        # Variáveis para pagamento
        self._init_payment_vars()

        # Initialize entry attributes
        self.nome_entry = None
        self.renach_entry = None
        self.valor_entries = {}
        self.dinheiro_entry = None
        self.cartao_entry = None
        self.pix_entry = None

    # Código de inicialização de variáveis de pagamento...
    def _init_payment_vars(self):
        """Inicializa variáveis relacionadas a pagamento."""
        self.forma_pagamento_var = tk.StringVar(value="")
        self.radio_var = tk.StringVar(value="")
        self.payment_vars = {
            "D": tk.IntVar(),
            "C": tk.IntVar(),
            "E": tk.IntVar(),
            "P": tk.IntVar(),
        }

    # Código de configuração de usuário...
    def set_current_user(self, user):
        """Define o usuário atual."""
        self.current_user = user

    # Código de configuração de frames...
    def configurar_frames(self, login_frame, criar_conta_frame):
        """Configura os frames de login e criar conta."""
        self.login_frame = login_frame
        self.criar_conta_frame = criar_conta_frame

    """
    SEÇÃO 2: INTERFACE GRÁFICA E WIDGETS
    """

    # Código de centralização...
    def center(self, window):
        """Centraliza a janela na tela."""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")
        window.deiconify()

    # Código de criação do frame de pagamento...
    def _create_payment_frame(self, parent, cor_fundo, cor_texto, cor_selecionado):
        """Cria o frame de pagamento com todas as opções."""
        frame_pagamento = tk.LabelFrame(
            parent,
            text="Formas de Pagamento",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 12, "bold"),
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

        formas_pagamento = {"D": "Débito", "C": "Crédito", "E": "Espécie", "P": "PIX"}

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
                command=on_payment_change,
            )
            cb.pack(side=tk.LEFT, padx=(0, 10))

            valor_entry = tk.Entry(frame, width=15, state="disabled")
            valor_entry.pack(side=tk.LEFT)
            self.valor_entries[codigo] = valor_entry

            # Atribuir as entradas aos atributos correspondentes
            if codigo == "D":
                self.dinheiro_entry = valor_entry
            elif codigo == "C":
                self.cartao_entry = valor_entry
            elif codigo == "P":
                self.pix_entry = valor_entry

            tk.Label(frame, text="R$", bg=cor_fundo, fg=cor_texto).pack(
                side=tk.LEFT, padx=(5, 0)
            )

        return frame_pagamento

    # Código do frame de radio buttons...
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
                font=("Arial", 12),
            ).pack(side=tk.LEFT, padx=2)

    # Código do frame de botões...
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
            activeforeground="#ECF0F1",
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            frame_botoes,
            text="Voltar",
            command=self.adicionar_window.destroy,
            width=15,
            activebackground="#2C3E50",
            activeforeground="#ECF0F1",
        ).pack(side=tk.LEFT, padx=5)

    # Código de criação de entry...
    def criar_entry(self, frame_nome, var_name, parent):
        """Cria um frame com label e entry para entradas de texto."""
        frame = tk.Frame(parent, bg=parent.cget("bg"))
        frame.pack(pady=2)

        tk.Label(
            frame,
            text=frame_nome,
            bg=parent.cget("bg"),
            fg="#ECF0F1",
            font=("Arial", 12),
        ).pack(side=tk.LEFT, anchor="w", padx=5)

        entry = tk.Entry(frame)
        entry.pack(side=tk.LEFT, padx=2)
        setattr(self, var_name, entry)
        return entry  # Return the entry widget

    # Configura a interface para adição de novo paciente
    def _setup_add_interface(self, cor_fundo, cor_texto, cor_selecionado):
        """Configura a interface de adição de paciente."""

        # Título
        tk.Label(
            self.adicionar_window,
            text="Preencha as informações:",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 16, "bold"),
        ).pack(pady=(15, 5))

        # Frame para RadioButtons
        self._create_radio_frame(cor_fundo, cor_texto, cor_selecionado)

        # Label para mostrar valor da consulta
        self.valor_consulta_label = tk.Label(
            self.adicionar_window,
            text="Valor da consulta: R$ 0,00",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 10, "bold"),
        )
        self.valor_consulta_label.pack(pady=5)

        # Função para atualizar o valor da consulta
        def atualizar_valor_consulta(*args):
            valores = {"medico": "148,65", "psicologo": "192,61", "ambos": "341,26"}
            valor = valores.get(self.radio_var.get(), "0,00")
            self.valor_consulta_label.config(text=f"Valor da consulta: R$ {valor}")

        # Associar a função ao radio_var
        self.radio_var.trace("w", atualizar_valor_consulta)

        # Entradas para nome e Renach
        self.nome_entry = self.criar_entry("Nome:", "nome_entry", self.adicionar_window)
        self.renach_entry = self.criar_entry(
            "Renach:", "renach_entry", self.adicionar_window
        )

        # Frame de pagamento
        self._create_payment_frame(
            self.adicionar_window, cor_fundo, cor_texto, cor_selecionado
        )

        def limpar_campos():
            self.nome_entry.delete(0, tk.END)
            self.renach_entry.delete(0, tk.END)

            # Limpar campos de forma de pagamento
            for entry in self.valor_entries.values():
                entry.delete(0, tk.END)
                entry.config(state="disabled", bg="#F0F0F0")

            # Desmarcar checkbuttons de forma de pagamento
            for var in self.payment_vars.values():
                var.set(0)

            # Limpar seleção dos RadioButtons
            self.radio_var.set("")

        def adicionar_paciente():
            if self.verificar_soma_pagamentos():  # Chamando o método da classe
                if self.salvar_informacao():
                    if self.adicionar_window.winfo_exists():
                        limpar_campos()
                        self.adicionar_window.destroy()

        # Botões
        button_frame = tk.Frame(self.adicionar_window, bg=cor_fundo)
        button_frame.pack(pady=10)

        tk.Button(
            button_frame,
            text="Adicionar",
            command=adicionar_paciente,
            bg="#2980B9",
            fg="white",
            font=("Arial", 12),
            width=10,
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            button_frame,
            text="Cancelar",
            command=self.adicionar_window.destroy,
            bg="#95A5A6",
            fg="white",
            font=("Arial", 12),
            width=10,
        ).pack(side=tk.LEFT, padx=5)

        # Texto de ajuda
        tk.Label(
            self.adicionar_window,
            text="Obs.: Para múltiplas formas de pagamento, informe o valor de cada uma.",
            bg=cor_fundo,
            fg=cor_texto,
            font=("Arial", 9, "italic"),
        ).pack(pady=(0, 10))

        # Limpar campos após criar todos os widgets
        limpar_campos()

    """
    SEÇÃO 3: MANIPULAÇÃO DE DADOS
    """

    # Código de obtenção do workbook...
    def get_active_workbook(self):
        """Obtém o workbook ativo atualizado."""
        if self.planilhas:
            self.planilhas.reload_workbook()
            return self.planilhas.wb
        return None

    # Código de verificação de pagamentos...
    def verificar_soma_pagamentos(self):
        """Verifica se a soma dos valores de pagamento está correta."""

        def convert_to_float(value):
            """Converte string de valor monetário para float."""
            if not value:
                return 0.0
            # Remove R$ e espaços, substitui vírgula por ponto
            clean_value = value.replace("R$", "").replace(" ", "").replace(",", ".")
            try:
                return float(clean_value)
            except ValueError:
                messagebox.showerror("Erro", f"Valor inválido: {value}")
                return None

        try:
            # Obtém o valor da consulta
            valor_consulta_str = self.valor_consulta_label.cget("text").split("R$ ")[1]
            valor_consulta = convert_to_float(valor_consulta_str)
            if valor_consulta is None:
                return False

            # Obtém valores dos campos de pagamento
            valor_dinheiro = convert_to_float(self.dinheiro_entry.get())
            valor_cartao = convert_to_float(self.cartao_entry.get())
            valor_pix = convert_to_float(self.pix_entry.get())

            if any(v is None for v in [valor_dinheiro, valor_cartao, valor_pix]):
                return False

            # Verifica quantas formas de pagamento foram selecionadas
            formas_selecionadas = sum(var.get() for var in self.payment_vars.values())

            if formas_selecionadas > 1:
                # Múltiplas formas de pagamento selecionadas
                soma_pagamentos = valor_dinheiro + valor_cartao + valor_pix

                # Usa uma pequena margem de erro para comparações de ponto flutuante
                if abs(soma_pagamentos - valor_consulta) > 0.01:
                    messagebox.showerror(
                        "Erro",
                        f"A soma dos valores de pagamento (R$ {soma_pagamentos:.2f}) "
                        f"deve ser igual ao valor da consulta (R$ {valor_consulta:.2f})",
                    )
                    return False
            else:
                # Apenas uma forma de pagamento selecionada
                valor_pagamento = valor_dinheiro + valor_cartao + valor_pix

                # Verifica se há algum valor e se está correto
                if valor_pagamento > 0 and abs(valor_pagamento - valor_consulta) > 0.01:
                    messagebox.showerror(
                        "Erro",
                        f"O valor do pagamento (R$ {valor_pagamento:.2f}) "
                        f"deve ser igual ao valor da consulta (R$ {valor_consulta:.2f})",
                    )
                    return False

            return True

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao verificar valores: {str(e)}")
            return False

    # Código de contagem de pagamentos...
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

    # Adiciona ou atualiza informações do paciente no banco de dados
    def _adicionar_paciente_ao_banco(self, nome, renach, pagamentos, escolha):
        """
        Adiciona ou atualiza paciente no banco de dados de marcação com status 'Compareceu'.
        """
        try:
            with DatabaseConnection("db_marcacao.db") as conn:
                cursor = conn.cursor()

                data_atual = datetime.now().strftime("%Y-%m-%d")
                STATUS_COMPARECEU = (
                    "attended"  # Usando o mesmo status que é verificado na listagem
                )

                # Prepara informações de pagamento
                pagamento_info = (
                    " | ".join(pagamentos)
                    if isinstance(pagamentos, list)
                    else str(pagamentos)
                )
                observacao = f"Tipo: {escolha}\nPagamento: {pagamento_info}\nRegistrado em: {data_atual}"

                # Verifica se o paciente existe
                cursor.execute(
                    """
                    SELECT data_agendamento, status_comparecimento 
                    FROM marcacoes 
                    WHERE renach = ?
                    ORDER BY data_agendamento DESC 
                    LIMIT 1
                """,
                    (renach,),
                )

                resultado = cursor.fetchone()

                if resultado:
                    # Se o paciente existe, atualiza os dados
                    data_anterior = resultado[0]
                    status_anterior = resultado[1]

                    cursor.execute(
                        """
                        UPDATE marcacoes 
                        SET nome = ?,
                            data_agendamento = ?,
                            status_comparecimento = ?,
                            observacao = ?
                        WHERE renach = ?
                    """,
                        (nome, data_atual, STATUS_COMPARECEU, observacao, renach),
                    )

                    # Atualiza histórico
                    cursor.execute(
                        """
                        SELECT historico_comparecimento 
                        FROM marcacoes 
                        WHERE renach = ?
                    """,
                        (renach,),
                    )

                    historico_atual = cursor.fetchone()[0]
                    try:
                        historico = (
                            json.loads(historico_atual) if historico_atual else []
                        )
                    except:
                        historico = []

                    historico.append(
                        {
                            "data_anterior": data_anterior,
                            "data_nova": data_atual,
                            "status_anterior": status_anterior,
                            "status_novo": STATUS_COMPARECEU,
                            "atualizado_em": datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                        }
                    )

                    cursor.execute(
                        """
                        UPDATE marcacoes 
                        SET historico_comparecimento = ?
                        WHERE renach = ?
                    """,
                        (json.dumps(historico), renach),
                    )

                    self.logger.info(
                        f"Atualizado status do paciente {renach} para Compareceu"
                    )
                    messagebox.showinfo(
                        "Sucesso", "Paciente atualizado com status Compareceu!"
                    )

                else:
                    # Se o paciente não existe, insere novo registro
                    cursor.execute(
                        """
                        INSERT INTO marcacoes (
                            nome,
                            renach,
                            telefone,
                            data_agendamento,
                            status_comparecimento,
                            observacao,
                            historico_comparecimento
                        ) VALUES (?, ?, ?, ?, ?, ?, '[]')
                    """,
                        (nome, renach, "", data_atual, STATUS_COMPARECEU, observacao),
                    )

                    self.logger.info(
                        f"Novo paciente {renach} adicionado com status Compareceu"
                    )
                    messagebox.showinfo(
                        "Sucesso", "Novo paciente adicionado com status Compareceu!"
                    )

                conn.commit()
                return True

        except sqlite3.Error as e:
            self.logger.error(f"Erro no banco de dados: {e}")
            messagebox.showerror(
                "Erro", f"Erro ao processar operação no banco de dados: {e}"
            )
            return False
        except Exception as e:
            self.logger.error(f"Erro inesperado: {e}")
            messagebox.showerror("Erro", f"Erro inesperado: {e}")
            return False

    """
    SEÇÃO 4: OPERAÇÕES COM PACIENTES
    """

    # Código de adição de informações...
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

    # Código de salvamento de informações...
    def salvar_informacao(self):
        """Salva os dados no banco de dados e na planilha."""
        try:
            # Obter dados dos campos
            nome = self.nome_entry.get().strip().upper()
            renach = self.renach_entry.get().strip()
            escolha = self.radio_var.get()

            if not nome or not renach or not escolha:
                messagebox.showerror(
                    "Erro",
                    "Por favor, preencha todos os campos obrigatórios (nome, RENACH e tipo de atendimento).",
                )
                return False

            # Mapear escolha do radio button
            tipo_escolha = {
                "medico": "medico",
                "psicologo": "psicologo",
                "ambos": "ambos",
            }.get(escolha)

            if not tipo_escolha:
                messagebox.showerror("Erro", "Selecione o tipo de atendimento.")
                return

            # Validar dados básicos
            if not nome or not renach:
                messagebox.showerror(
                    "Erro", "Por favor, preencha os campos de nome e RENACH."
                )
                return

            if not renach.isdigit():
                messagebox.showerror("Erro", "O RENACH deve ser um número inteiro.")
                return

            # Verificar formas de pagamento selecionadas
            formas_selecionadas = {
                forma: var.get() for forma, var in self.payment_vars.items()
            }

            if not any(formas_selecionadas.values()):
                messagebox.showerror(
                    "Erro", "Selecione pelo menos uma forma de pagamento."
                )
                return

            # Processar pagamentos
            pagamentos = []
            soma_valores = 0
            num_formas_selecionadas = sum(formas_selecionadas.values())

            # Se apenas uma forma de pagamento está selecionada
            if num_formas_selecionadas == 1:
                forma_selecionada = next(
                    forma for forma, sel in formas_selecionadas.items() if sel
                )
                valor_entrada = self.valor_entries[forma_selecionada].get().strip()

                if valor_entrada:  # Se um valor foi especificado
                    try:
                        valor_float = float(valor_entrada.replace(",", "."))
                        valor_formatado = f"{valor_float:.2f}".replace(".", ",")
                        pagamentos.append(f"{forma_selecionada}:{valor_formatado}")
                    except ValueError:
                        messagebox.showerror(
                            "Erro", "O valor informado não é um número válido"
                        )
                        return
                else:  # Se não houver valor, adiciona apenas a forma de pagamento
                    pagamentos.append(forma_selecionada)

            else:  # Múltiplas formas de pagamento
                for codigo, selecionado in formas_selecionadas.items():
                    if selecionado:
                        valor = self.valor_entries[codigo].get().strip()

                        if valor:  # Se um valor foi especificado
                            try:
                                valor_float = float(valor.replace(",", "."))
                                valor_formatado = f"{valor_float:.2f}".replace(".", ",")
                                pagamentos.append(f"{codigo}:{valor_formatado}")
                                soma_valores += valor_float
                            except ValueError:
                                messagebox.showerror(
                                    "Erro",
                                    f"O valor informado para {codigo} não é um número válido",
                                )
                                return
                        else:  # Se não houver valor, adiciona apenas a forma de pagamento
                            pagamentos.append(codigo)

                # Verifica a soma apenas se todos os pagamentos têm valores
                if all(":" in pag for pag in pagamentos):
                    valores_maximos = {
                        "medico": 148.65,
                        "psicologo": 192.61,
                        "ambos": 341.26,
                    }
                    valor_maximo = valores_maximos[tipo_escolha]

                    if abs(soma_valores - valor_maximo) > 0.01:
                        messagebox.showerror(
                            "Erro",
                            f"A soma dos valores ({soma_valores:.2f}) deve ser igual ao valor da consulta ({valor_maximo:.2f})",
                        )
                        return

            # Tenta salvar na planilha
            self.logger.info(
                f"Tentando salvar na planilha: nome={nome}, renach={renach}, tipo={tipo_escolha}"
            )
            if not self.planilhas.wb:
                self.planilhas.reload_workbook()

            ws = self.planilhas.get_active_sheet()

            # Encontrar próxima linha vazia
            def encontrar_proxima_linha(coluna_letra):
                for row in range(3, ws.max_row + 2):
                    if not ws[f"{coluna_letra}{row}"].value:
                        return row
                return ws.max_row + 1

            alteracoes_feitas = False

            # String de pagamento formatada
            info_pagamento = " | ".join(pagamentos)

            # Salvar dados conforme o tipo de atendimento
            if tipo_escolha in ["medico", "ambos"]:
                nova_linha = encontrar_proxima_linha("B")
                ws[f"B{nova_linha}"] = nome
                ws[f"C{nova_linha}"] = renach
                ws[f"F{nova_linha}"] = info_pagamento
                alteracoes_feitas = True
                self.logger.info(f"Dados médicos salvos na linha {nova_linha}")

            if tipo_escolha in ["psicologo", "ambos"]:
                nova_linha = encontrar_proxima_linha("H")
                ws[f"H{nova_linha}"] = nome
                ws[f"I{nova_linha}"] = renach
                ws[f"L{nova_linha}"] = info_pagamento
                alteracoes_feitas = True
                self.logger.info(f"Dados psicológicos salvos na linha {nova_linha}")

            if alteracoes_feitas:
                self.planilhas.wb.save(self.planilhas.file_path)
                self.logger.info("Planilha salva com sucesso")

                # Após salvar na planilha, salva no banco
                if self._adicionar_paciente_ao_banco(
                    nome, renach, pagamentos, tipo_escolha
                ):
                    messagebox.showinfo("Sucesso", "Informações salvas com sucesso!")
                    self.adicionar_window.destroy()
                    return True
            else:
                messagebox.showerror(
                    "Erro", "Não foi possível salvar as informações na planilha"
                )
                return False

        except Exception as e:
            self.logger.error(f"Erro ao salvar informações: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao salvar informações: {str(e)}")
            return False

    # Código de exclusão...
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
                                if not isinstance(
                                    current_cell, openpyxl.cell.cell.MergedCell
                                ):
                                    if isinstance(
                                        next_cell, openpyxl.cell.cell.MergedCell
                                    ):
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
                                if cell.value and str(cell.value).strip() == str(
                                    renach
                                ):
                                    return row
                        return None

                    # Procura nas seções de médico e psicólogo
                    linha_medico = encontrar_paciente(3)  # Coluna C
                    linha_psi = encontrar_paciente(9)  # Coluna I

                    alteracoes = False

                    if linha_medico:
                        mover_conteudo(linha_medico, 2, 6)  # Colunas B-F
                        alteracoes = True
                        messagebox.showinfo("Sucesso", "Removido da seção de médicos")

                    if linha_psi:
                        mover_conteudo(linha_psi, 8, 12)  # Colunas H-L
                        alteracoes = True
                        messagebox.showinfo(
                            "Sucesso", "Removido da seção de psicólogos"
                        )

                    if alteracoes:
                        wb.save(self.file_path)
                        excluir_window.destroy()
                    else:
                        messagebox.showwarning("Aviso", "RENACH não encontrado")

                except ValueError:
                    messagebox.showerror(
                        "Erro", "Por favor, insira um RENACH válido (apenas números)"
                    )
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
            main_frame.pack(expand=True, fill="both", padx=20, pady=20)

            # Label
            tk.Label(
                main_frame,
                text="Informe o RENACH:",
                bg=self.master.cget("bg"),
                fg="#ECF0F1",
                font=("Arial", 14, "bold"),
            ).pack(pady=10)

            # Entry frame
            entry_frame = tk.Frame(main_frame, bg=self.master.cget("bg"))
            entry_frame.pack(fill="x", pady=5)

            renach_entry = tk.Entry(entry_frame, justify="center")
            renach_entry.pack(pady=5)
            renach_entry.focus()
            renach_entry.bind("<Return>", lambda e: realizar_exclusao())

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
                width=15,
            ).pack(side=tk.LEFT, padx=5)

            tk.Button(
                button_frame, text="Cancelar", command=excluir_window.destroy, width=15
            ).pack(side=tk.LEFT, padx=5)

            self.center(excluir_window)

            def on_closing():
                wb.close()
                excluir_window.destroy()

            excluir_window.protocol("WM_DELETE_WINDOW", on_closing)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao iniciar exclusão: {str(e)}")

    # Código de exibição de informações...
    def exibir_informacao(self):
        """Exibe informações dos pacientes usando o PatientInfoDisplay."""
        try:
            # Cria uma instância do PatientInfoDisplay passando os parâmetros necessários
            display = PatientInfoDisplay(
                master=self.master,          # Janela principal
                planilhas=self.planilhas,    # Objeto que gerencia as planilhas
                logger=self.logger           # Logger para registro de eventos
            )
            
            # Chama o método display para mostrar as informações
            display.display()
            
            return True
            
        except Exception as e:
            self.logger.error(f"Erro ao exibir informações: {str(e)}")
            messagebox.showerror(
                "Erro", 
                f"Ocorreu um erro ao exibir as informações: {str(e)}"
            )
            return False
    """
    SEÇÃO 5: MANIPULAÇÃO DE PLANILHAS
    """

    # Código de formatação...
    def formatar_planilha(self):
        """
        Formata a planilha preservando as informações necessárias.
        """
        try:
            if not self.planilhas:
                return False

            self.planilhas.reload_workbook()
            ws = self.planilhas.get_active_sheet()

            if not ws:
                return False

            # Definir estilos
            thin_side = Side(style="thin")
            borda = Border(
                left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
            )
            font_bold = Font(name="Arial", bold=True, size=11, color="000000")
            font_regular = Font(name="Arial", size=11, color="000000")
            alignment_center = Alignment(horizontal="center", vertical="center")
            alignment_left = Alignment(horizontal="left", vertical="center")

            # Coletar dados existentes
            dados_medicos = []
            dados_psicologos = []
            max_row = ws.max_row + 1

            # Encontrar última linha de dados válidos
            ultima_linha_dados = 3
            for row in range(3, max_row):
                nome_med = ws.cell(row=row, column=2).value
                nome_psi = ws.cell(row=row, column=8).value

                if (isinstance(nome_med, str) and nome_med.strip()) or (
                    isinstance(nome_psi, str) and nome_psi.strip()
                ):
                    ultima_linha_dados = row

            # Coletar dados até a última linha válida
            for row in range(3, ultima_linha_dados + 1):
                # Dados médicos
                nome_med = ws.cell(row=row, column=2).value
                if (
                    isinstance(nome_med, str)
                    and nome_med.strip()
                    and not any(
                        palavra in str(nome_med).lower()
                        for palavra in ["soma", "médico", "total"]
                    )
                ):
                    dados_medicos.append(
                        {
                            "nome": nome_med.strip(),
                            "renach": str(
                                ws.cell(row=row, column=3).value or ""
                            ).strip(),
                            "reexames": str(
                                ws.cell(row=row, column=4).value or ""
                            ).strip(),
                            "pagamento": str(
                                ws.cell(row=row, column=6).value or ""
                            ).strip(),
                        }
                    )

                # Dados psicólogos
                nome_psi = ws.cell(row=row, column=8).value
                if (
                    isinstance(nome_psi, str)
                    and nome_psi.strip()
                    and not any(
                        palavra in str(nome_psi).lower()
                        for palavra in ["soma", "psicólogo", "total"]
                    )
                ):
                    dados_psicologos.append(
                        {
                            "nome": nome_psi.strip(),
                            "renach": str(
                                ws.cell(row=row, column=9).value or ""
                            ).strip(),
                            "reexames": str(
                                ws.cell(row=row, column=10).value or ""
                            ).strip(),
                            "pagamento": str(
                                ws.cell(row=row, column=12).value or ""
                            ).strip(),
                        }
                    )

            # Limpar planilha
            for row in range(1, max_row):
                for col in range(1, 13):
                    cell = ws.cell(row=row, column=col)
                    if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                        cell.value = None
                        cell.border = borda
                        cell.font = font_regular
                        cell.alignment = alignment_center

            # Configurar cabeçalhos
            from datetime import datetime

            data_atual = datetime.now().strftime("%d/%m/%Y")
            usuario = (
                self.current_user
                if hasattr(self, "current_user") and self.current_user
                else "Usuário"
            )

            # Cabeçalhos principais
            ws["A1"] = f"({usuario}) Atendimento Médico {data_atual}"
            ws.merge_cells("A1:F1")
            ws["G1"] = f"({usuario}) Atendimento Psicológico {data_atual}"
            ws.merge_cells("G1:L1")

            for col in range(1, 7):
                cell = ws.cell(row=1, column=col)
                cell.font = font_bold
                cell.alignment = alignment_center

            for col in range(7, 13):
                cell = ws.cell(row=1, column=col)
                cell.font = font_bold
                cell.alignment = alignment_center

            # Cabeçalhos das colunas
            headers = ["Ordem", "Nome", "Renach", "Reexames", "Valor", "Pagamento"]
            for start_col in [1, 7]:
                for idx, header in enumerate(headers):
                    cell = ws.cell(row=2, column=start_col + idx)
                    cell.value = header
                    cell.font = font_bold
                    cell.alignment = alignment_center
                    cell.border = borda

            # Restaurar dados médicos
            for idx, dados in enumerate(dados_medicos, start=3):
                # Ordem
                ws.cell(row=idx, column=1).value = idx - 2

                # Dados do paciente
                ws.cell(row=idx, column=2).value = dados["nome"]
                ws.cell(row=idx, column=2).alignment = alignment_left

                ws.cell(row=idx, column=3).value = dados["renach"]
                ws.cell(row=idx, column=4).value = dados["reexames"] or "D"

                # Valor fixo
                valor_cell = ws.cell(row=idx, column=5)
                valor_cell.value = 148.65
                valor_cell.number_format = '"R$"#,##0.00'

                # Forma de pagamento
                ws.cell(row=idx, column=6).value = dados["pagamento"]

            # Restaurar dados psicólogos
            for idx, dados in enumerate(dados_psicologos, start=3):
                # Ordem
                ws.cell(row=idx, column=7).value = idx - 2

                # Dados do paciente
                ws.cell(row=idx, column=8).value = dados["nome"]
                ws.cell(row=idx, column=8).alignment = alignment_left

                ws.cell(row=idx, column=9).value = dados["renach"]
                ws.cell(row=idx, column=10).value = dados["reexames"] or "D"

                # Valor fixo
                valor_cell = ws.cell(row=idx, column=11)
                valor_cell.value = 192.61
                valor_cell.number_format = '"R$"#,##0.00'

                # Forma de pagamento
                ws.cell(row=idx, column=12).value = dados["pagamento"]

            # Adicionar totais médicos (uma linha abaixo do último paciente)
            if dados_medicos:
                linha_med = len(dados_medicos) + 3

                # Soma
                ws.cell(row=linha_med, column=4).value = "Soma"
                ws.cell(row=linha_med, column=5).value = len(dados_medicos) * 148.65
                ws.cell(row=linha_med, column=5).number_format = '"R$"#,##0.00'

                # Médico
                ws.cell(row=linha_med + 1, column=4).value = "Médico"
                ws.cell(row=linha_med + 1, column=5).value = len(dados_medicos) * 49.00
                ws.cell(row=linha_med + 1, column=5).number_format = '"R$"#,##0.00'

                # Total
                ws.cell(row=linha_med + 2, column=4).value = "Total"
                ws.cell(row=linha_med + 2, column=5).value = (
                    len(dados_medicos) * 148.65
                ) - (len(dados_medicos) * 49.00)
                ws.cell(row=linha_med + 2, column=5).number_format = '"R$"#,##0.00'

            # Adicionar totais psicólogos (uma linha abaixo do último paciente)
            if dados_psicologos:
                linha_psi = len(dados_psicologos) + 3

                # Soma
                ws.cell(row=linha_psi, column=10).value = "Soma"
                ws.cell(row=linha_psi, column=11).value = len(dados_psicologos) * 192.61
                ws.cell(row=linha_psi, column=11).number_format = '"R$"#,##0.00'

                # Psicólogo
                ws.cell(row=linha_psi + 1, column=10).value = "Psicólogo"
                ws.cell(row=linha_psi + 1, column=11).value = (
                    len(dados_psicologos) * 63.50
                )
                ws.cell(row=linha_psi + 1, column=11).number_format = '"R$"#,##0.00'

                # Total
                ws.cell(row=linha_psi + 2, column=10).value = "Total"
                ws.cell(row=linha_psi + 2, column=11).value = (
                    len(dados_psicologos) * 192.61
                ) - (len(dados_psicologos) * 63.50)
                ws.cell(row=linha_psi + 2, column=11).number_format = '"R$"#,##0.00'

            # Ajustar largura das colunas
            larguras = {
                "A": 8,
                "B": 40,
                "C": 12,
                "D": 12,
                "E": 12,
                "F": 15,
                "G": 8,
                "H": 40,
                "I": 12,
                "J": 12,
                "K": 12,
                "L": 15,
            }
            for coluna, largura in larguras.items():
                ws.column_dimensions[coluna].width = largura

            self.planilhas.wb.save(self.file_path)
            return True

        except Exception as e:
            self.logger.error(f"Erro ao formatar planilha: {str(e)}")
            return False

    # Código de salvamento na planilha...
    def salvar_na_planilha(self, nome, renach, pagamentos, escolha):
        """Salva os dados na planilha."""
        try:
            # Garantir que temos o objeto planilhas
            if not self.planilhas:
                raise Exception("Objeto planilhas não inicializado")

            # Recarregar workbook e obter sheet ativa
            self.planilhas.reload_workbook()
            ws = self.planilhas.get_active_sheet()

            if not ws:
                raise Exception("Não foi possível acessar a planilha ativa")

            # Formatar string de pagamento
            info_pagamento = " | ".join(pagamentos)

            # Encontrar próximas linhas vazias para médico e psicólogo
            def encontrar_proxima_linha(coluna_letra):
                for row in range(3, ws.max_row + 2):
                    if not ws[f"{coluna_letra}{row}"].value:
                        return row
                return ws.max_row + 1

            alteracoes_feitas = False

            # Salvar dados do médico
            if escolha in ["medico", "ambos"]:
                nova_linha = encontrar_proxima_linha("B")
                ws[f"B{nova_linha}"] = nome
                ws[f"C{nova_linha}"] = renach
                ws[f"F{nova_linha}"] = info_pagamento
                alteracoes_feitas = True
                self.logger.info(f"Dados médicos salvos na linha {nova_linha}")

            # Salvar dados do psicólogo
            if escolha in ["psicologo", "ambos"]:
                nova_linha = encontrar_proxima_linha("H")
                ws[f"H{nova_linha}"] = nome
                ws[f"I{nova_linha}"] = renach
                ws[f"L{nova_linha}"] = info_pagamento
                alteracoes_feitas = True
                self.logger.info(f"Dados psicológicos salvos na linha {nova_linha}")

            if alteracoes_feitas:
                self.planilhas.wb.save(self.file_path)
                self.logger.info("Planilha salva com sucesso")
                return True
            else:
                raise Exception("Nenhuma alteração foi realizada na planilha")

        except Exception as e:
            self.logger.error(f"Erro ao salvar na planilha: {str(e)}")
            return False

    # Código de adição de totais...
    def _adicionar_totais(
        self,
        ws,
        linha_inicio,
        col_inicio,
        valor_consulta,
        valor_profissional,
        num_pacientes,
        borda,
        font_bold,
        alignment_center,
    ):
        """Adiciona os totais para uma seção (médico ou psicólogo)"""
        # Soma
        ws.cell(row=linha_inicio, column=col_inicio + 2, value="Soma")
        soma_cell = ws.cell(
            row=linha_inicio,
            column=col_inicio + 4,
            value=valor_consulta * num_pacientes,
        )
        soma_cell.number_format = '"R$"#,##0.00'

        # Valor profissional
        ws.cell(row=linha_inicio + 1, column=col_inicio + 2, value="Profissional")
        prof_cell = ws.cell(
            row=linha_inicio + 1,
            column=col_inicio + 4,
            value=valor_profissional * num_pacientes,
        )
        prof_cell.number_format = '"R$"#,##0.00'

        # Total
        ws.cell(row=linha_inicio + 2, column=col_inicio + 2, value="Total")
        total_cell = ws.cell(
            row=linha_inicio + 2,
            column=col_inicio + 4,
            value=(valor_consulta - valor_profissional) * num_pacientes,
        )
        total_cell.number_format = '"R$"#,##0.00'

        # Aplica formatação
        for row in range(linha_inicio, linha_inicio + 3):
            for col in range(col_inicio + 2, col_inicio + 5):
                cell = ws.cell(row=row, column=col)
                cell.border = borda
                cell.font = font_bold
                cell.alignment = alignment_center

    # Código de adição de resumo...
    def _adicionar_resumo_geral(
        self, ws, linha_inicio, num_medico, num_psi, borda, font_bold, alignment_center
    ):
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
            ("Total Clínica", total_clinica),
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

    """
    SEÇÃO 6: COMUNICAÇÃO E INTEGRAÇÃO
    """

    # Código de envio de WhatsApp...
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

    # Código de processamento de envio...
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

    # Código de envio de email...
    def enviar_email(self):
        """Configura e exibe interface para envio de email."""
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

    # Código de seleção de arquivo...
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

    # Código de envio...
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

    """
    SEÇÃO 7: RELATÓRIOS E VISUALIZAÇÃO
    """

    # Código de exibição de totais...
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

    # Código de processamento de notas...
    def processar_notas_fiscais(self):
        """Processa e emite notas fiscais."""
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

    def mostrar_valores_atendimentos(self):
        # Carrega os dados da planilha
        wb = self.get_active_workbook()
        ws = wb.active

        # Inicializa o total de pagamentos acumulados para cada método
        total_medico = {"Débito": 0, "Crédito": 0, "Espécie": 0, "PIX": 0}
        total_psicologo = {"Débito": 0, "Crédito": 0, "Espécie": 0, "PIX": 0}

        # Valores fixos para consulta
        VALOR_MEDICO = 148.65
        VALOR_PSICOLOGO = 192.61

        # Itera sobre as linhas da planilha para calcular os valores
        for row in range(3, ws.max_row + 1):
            pagamento_medico = ws[f"F{row}"].value
            pagamento_psicologo = ws[f"L{row}"].value

            # Processa os valores de pagamento para médico
            if pagamento_medico:
                if pagamento_medico in ["D", "C", "E", "P"]:
                    metodo = self._traduzir_metodo(pagamento_medico)
                    total_medico[metodo] += VALOR_MEDICO
                else:
                    # Para formatos como "D:100,65|C:48,00"
                    for parte in pagamento_medico.split("|"):
                        metodo, valor = parte.split(":")
                        total_medico[self._traduzir_metodo(metodo.strip())] += float(valor.replace(",", "."))

            # Processa os valores de pagamento para psicólogo
            if pagamento_psicologo:
                if pagamento_psicologo in ["D", "C", "E", "P"]:
                    metodo = self._traduzir_metodo(pagamento_psicologo)
                    total_psicologo[metodo] += VALOR_PSICOLOGO
                else:
                    # Para formatos como "D:100,65|C:48,00"
                    for parte in pagamento_psicologo.split("|"):
                        metodo, valor = parte.split(":")
                        total_psicologo[self._traduzir_metodo(metodo.strip())] += float(valor.replace(",", "."))

        # Criação da janela para exibir os valores
        janela_valores = tk.Toplevel(self.master)
        janela_valores.title("Valores dos Atendimentos")
        janela_valores.geometry("400x400")
        janela_valores.configure(bg="#2C3E50")

        # Exibindo valores acumulados para médico
        tk.Label(janela_valores, text="Valores - Médico:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 20, "bold")).pack(pady=5)
        for metodo, valor in total_medico.items():
            tk.Label(janela_valores, text=f"{metodo}: R$ {valor:.2f}", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 12, "bold")).pack()

        # Espaço entre seções
        tk.Label(janela_valores, text="", bg="#2C3E50").pack()

        # Exibindo valores acumulados para psicólogo
        tk.Label(janela_valores, text="Valores - Psicólogo:", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 20, "bold")).pack(pady=5)
        for metodo, valor in total_psicologo.items():
            tk.Label(janela_valores, text=f"{metodo}: R$ {valor:.2f}", bg="#2C3E50", fg="#ECF0F1", font=("Arial", 12, "bold")).pack()

        # Centraliza a janela
        self.center(janela_valores)

    def _traduzir_metodo(self, codigo):
        """Converte os códigos de pagamento em textos legíveis."""
        return {
            "D": "Débito",
            "C": "Crédito",
            "E": "Espécie",
            "P": "PIX"
        }.get(codigo, "Desconhecido")

    """
    SEÇÃO 8: NAVEGAÇÃO DE INTERFACE
    """

    # Código de navegação...
    def mostrar_criar_conta(self):
        """Alterna para o frame de criar conta."""
        self.login_frame.hide()
        self.criar_conta_frame.show()

    # Código de navegação...
    def voltar_para_login(self):
        """Alterna de volta para o frame de login."""
        self.criar_conta_frame.hide()
        self.login_frame.show()


class SistemaContas:
    """
    SEÇÃO 1: INICIALIZAÇÃO E CONFIGURAÇÃO
    """

    # Inicializa o sistema com arquivo e usuário
    def __init__(self, file_path: str, current_user=None):
        """Inicializa o sistema de contas."""
        self.file_path = file_path
        self.current_user = current_user
        self.sheet_name = "Contas Fechamento"
        self.criar_sheet_se_nao_existir()

    # Cria planilha base se não existir
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

    """
    SEÇÃO 2: INTERFACE GRÁFICA
    """

    # Abre a janela principal do sistema
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

    # Cria interface com campos e botões
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

    """
    SEÇÃO 3: MANIPULAÇÃO DE DADOS
    """

    # Salva dados na planilha agrupados por data
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

    # Processa os dados do formulário para salvar
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

    """
    SEÇÃO 4: VALIDAÇÃO E UTILITÁRIOS
    """

    # Verifica se os campos estão preenchidos corretamente
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

    # Limpa todos os campos do formulário
    def limpar_campos(self):
        """Limpa os campos após salvar"""
        self.info_entry.delete(0, tk.END)
        self.valor_entry.delete(0, tk.END)


class GerenciadorPlanilhas:
    """
    SEÇÃO 1: INICIALIZAÇÃO E CONFIGURAÇÃO
    """

    # Inicializa o gerenciador com janela principal e sistema de contas
    def __init__(self, master, sistema_contas):
        """Inicializa o gerenciador de planilhas."""
        self.master = master
        self.sistema_contas = sistema_contas
        self.file_path = None
        self.sheet_name = None
        self.active_window = None

    """
    SEÇÃO 2: INTERFACE GRÁFICA
    """

    # Abre janela principal do gerenciador
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

    # Configura os elementos da interface
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

    # Fecha a janela do gerenciador
    def _on_closing(self):
        """Handler para quando a janela for fechada"""
        self.active_window.destroy()
        self.active_window = None

    """
    SEÇÃO 3: MANIPULAÇÃO DE ARQUIVOS
    """

    # Cria nova planilha Excel
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

    # Abre planilha Excel existente
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

    """
    SEÇÃO 4: GERENCIAMENTO DE SHEETS
    """

    # Atualiza lista de sheets disponíveis
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

    # Seleciona e ativa uma sheet
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

    # Cria nova sheet e a torna ativa
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


@dataclass
class PatientData:
    """Estrutura de dados imutável para informações do paciente."""
    nome: str
    renach: str
    pagamento: str
    tipo: str
    search_text: str

class PatientInfoDisplay:
    """Classe otimizada para exibição de informações de pacientes."""
    
    def __init__(self, master: tk.Tk, planilhas, logger=None):
        self.master = master
        self.planilhas = planilhas
        self.logger = logger or logging.getLogger(__name__)
        
        # Sistema de cache melhorado com TTL
        self.data_cache = {
            'medico': [],
            'psi': [],
            'last_update': None,
            'last_filters': {},
            'timer': None,
            'ttl': 300  # 5 minutos de TTL para o cache
        }
        
        # Temas e estilos modernos
        self.theme = {
            'background': '#1a1a1a',
            'secondary_bg': '#2d2d2d',
            'text': '#ffffff',
            'accent': '#3498db',
            'header': '#2c3e50',
            'highlight': '#34495e',
            'separator': '#7f8c8d',
            'hover': '#3e4d5c',
            'error': '#e74c3c'
        }
        
        # Configurações de UI responsiva
        self.ui_config = {
            'min_width': 800,
            'min_height': 600,
            'padding': 10,
            'animation_duration': 200
        }
        
        # Referências de UI com estado
        self.ui_refs = {}
        
        # Estado de ordenação
        self.sort_state = {
            'column': None,
            'reverse': False
        }


    @lru_cache(maxsize=1000)
    def _process_payment(self, value: str) -> str:
        """Processa e formata valores de pagamento com cache."""
        if not value:
            return ""
        try:
            if isinstance(value, (int, float)):
                return f"R$ {float(value):.2f}"
            value_str = str(value).replace("R$", "").replace(",", ".").strip()
            return f"R$ {float(value_str):.2f}"
        except:
            return str(value).strip()

    def _load_data(self) -> bool:
            """Carrega e processa os dados da planilha com cache inteligente."""
            try:
                current_time = time.time()
                
                # Verifica se o cache ainda é válido
                if (self.data_cache['last_update'] and 
                    current_time - self.data_cache['last_update'] < self.data_cache['ttl']):
                    return True
                    
                self.planilhas.reload_workbook()
                wb = self.planilhas.wb
                
                if not wb:
                    return False

                ws = wb.active
                if not ws:
                    messagebox.showerror("Erro", "Planilha inválida")
                    return False

                # Processamento em lote otimizado
                med_data = []
                psi_data = []
                
                # Pré-aloca as listas para melhor performance
                max_rows = ws.max_row
                med_data = []
                psi_data = []
                
                # Processa em chunks para melhor performance
                chunk_size = 100
                for start_row in range(3, max_rows + 1, chunk_size):
                    end_row = min(start_row + chunk_size, max_rows + 1)
                    
                    rows = list(ws.iter_rows(min_row=start_row, max_row=end_row))
                    
                    for row in rows:
                        # Processamento de médicos
                        if row[1].value:
                            nome = str(row[1].value).strip().upper()
                            if not any(x in nome.lower() for x in ["soma", "médico", "psicólogo", "total"]):
                                med_data.append(PatientData(
                                    nome=nome,
                                    renach=str(row[2].value or "").strip(),
                                    pagamento=self._process_payment(row[5].value),
                                    tipo="Médico",
                                    search_text=f"{nome.lower()} {str(row[2].value or '').lower()}"
                                ))
                        
                        # Processamento de psicólogos
                        if len(row) > 7 and row[7].value:
                            nome = str(row[7].value).strip().upper()
                            if not any(x in nome.lower() for x in ["soma", "médico", "psicólogo", "total"]):
                                psi_data.append(PatientData(
                                    nome=nome,
                                    renach=str(row[8].value or "").strip(),
                                    pagamento=self._process_payment(row[11].value),
                                    tipo="Psicólogo",
                                    search_text=f"{nome.lower()} {str(row[8].value or '').lower()}"
                                ))

                # Atualiza o cache
                self.data_cache.update({
                    'medico': med_data,
                    'psi': psi_data,
                    'last_update': current_time
                })

                return bool(med_data or psi_data)

            except Exception as e:
                self.logger.error(f"Erro ao carregar dados: {e}")
                return False
            
    def _create_ui(self) -> Tuple[tk.Toplevel, Dict]:
        """Cria e retorna a interface do usuário."""
        window = tk.Toplevel(self.master)
        window.title("Informações dos Pacientes")
        window.geometry("1200x800")
        window.configure(bg=self.theme['background'])
        
        # Frames principais
        frames = self._create_frames(window)
        
        # Controles de filtro
        filters = self._create_filters(frames['control'])
        
        # Tabela
        table = self._create_table(frames['table'])
        
        # Barra de status
        stats_label = tk.Label(
            frames['stats'],
            bg=self.theme['background'],
            fg=self.theme['text'],
            font=("Arial", 10, "bold")
        )
        stats_label.pack(pady=5)
        
        self.ui_refs = {
            'window': window,
            'frames': frames,
            'filters': filters,
            'table': table,
            'stats': stats_label
        }
        
        return window, self.ui_refs
    
    def _create_frames(self, window: tk.Toplevel) -> Dict[str, tk.Frame]:
        """Cria e retorna os frames principais."""
        frames = {
            'main': tk.Frame(window, bg=self.theme['background']),
            'control': tk.Frame(window, bg=self.theme['background']),
            'table': tk.Frame(window, bg=self.theme['background']),
            'stats': tk.Frame(window, bg=self.theme['background'])
        }
        
        frames['main'].pack(fill="both", expand=True, padx=20, pady=10)
        frames['control'].pack(in_=frames['main'], fill="x", pady=(0, 10))
        frames['table'].pack(in_=frames['main'], fill="both", expand=True)
        frames['stats'].pack(in_=frames['main'], fill="x", pady=10)
        
        return frames
    
    def _create_filters(self, parent: tk.Frame) -> Dict[str, tk.Variable]:
        """Cria e retorna os controles de filtro."""
        filters = {
            'type': tk.StringVar(value="todos"),
            'search': tk.StringVar(),
            'payment': tk.StringVar()
        }
        
        filter_frame = tk.Frame(parent, bg=self.theme['background'])
        filter_frame.pack(fill="x", padx=5)
        
        # Tipo de atendimento
        type_frame = self._create_filter_section(filter_frame, "Tipo de Atendimento")
        options = [("todos", "Todos"), ("medico", "Médico"), ("psi", "Psicólogo")]
        for value, text in options:
            tk.Radiobutton(
                type_frame,
                text=text,
                variable=filters['type'],
                value=value,
                bg=self.theme['background'],
                fg=self.theme['text'],
                selectcolor=self.theme['header'],
                command=lambda: self._delayed_filter()
            ).pack(side="left", padx=5)
        
        # Busca
        search_frame = self._create_filter_section(filter_frame, "Buscar por Nome/RENACH")
        tk.Entry(
            search_frame,
            textvariable=filters['search'],
            width=30
        ).pack(padx=5, pady=2)
        
        # Pagamento
        payment_frame = self._create_filter_section(filter_frame, "Filtrar por Forma de Pagamento")
        tk.Entry(
            payment_frame,
            textvariable=filters['payment'],
            width=20
        ).pack(padx=5, pady=2)
        
        for var in filters.values():
            var.trace("w", lambda *args: self._delayed_filter())
        
        return filters
    
    def _create_filter_section(self, parent: tk.Frame, title: str) -> tk.LabelFrame:
        """Cria uma seção de filtro com título."""
        frame = tk.LabelFrame(
            parent,
            text=title,
            bg=self.theme['background'],
            fg=self.theme['text'],
            font=("Arial", 10)
        )
        frame.pack(side="left", padx=5, pady=5)
        return frame
    
    def _create_table(self, parent: tk.Frame) -> Dict:
        """Cria uma tabela moderna e responsiva."""
        # Frame principal com bordas arredondadas
        table_container = tk.Frame(
            parent,
            bg=self.theme['background'],
            highlightbackground=self.theme['accent'],
            highlightthickness=1
        )
        table_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Canvas com scrollbar suave
        canvas = tk.Canvas(
            table_container,
            bg=self.theme['background'],
            highlightthickness=0
        )
        
        scrollbar = ttk.Scrollbar(
            table_container,
            orient="vertical",
            command=canvas.yview
        )
        
        table_frame = tk.Frame(
            canvas,
            bg=self.theme['background']
        )
        
        # Configuração do scroll
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            
        def _bound_to_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)
            
        def _unbound_to_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")
            
        def _on_mousewheel(event):
            # Windows
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
            else:
                # Tratamento do evento para Windows
                if event.delta > 0:
                    canvas.yview_scroll(-1, "units")
                else:
                    canvas.yview_scroll(1, "units")
                    
        table_frame.bind('<Configure>', _on_frame_configure)
        
        # Vincula os eventos de scroll
        canvas.bind('<Enter>', _bound_to_mousewheel)
        canvas.bind('<Leave>', _unbound_to_mousewheel)
        
        # Cabeçalhos clicáveis para ordenação
        headers = [
            ("Nº", 5),
            ("Nome", 30),
            ("RENACH", 10),
            ("Forma de Pagamento", 20),
            ("Tipo", 10)
        ]
        
        for col, (header, width) in enumerate(headers):
            header_frame = tk.Frame(
                table_frame,
                bg=self.theme['header']
            )
            header_frame.grid(row=0, column=col, sticky="nsew", padx=1, pady=1)
            
            label = tk.Label(
                header_frame,
                text=header,
                bg=self.theme['header'],
                fg=self.theme['text'],
                font=("Arial", 11, "bold"),
                padx=10,
                pady=8
            )
            label.pack(fill="both", expand=True)
            
            # Adiciona funcionalidade de ordenação
            label.bind('<Button-1>', lambda e, col=col: self._sort_table(col))
            label.bind('<Enter>', lambda e, widget=label: self._on_header_hover(widget, True))
            label.bind('<Leave>', lambda e, widget=label: self._on_header_hover(widget, False))
        
        # Configuração do canvas e scrollbar
        canvas.create_window((0, 0), window=table_frame, anchor="nw", tags=("table",))
        
        def _on_canvas_configure(event):
            canvas.itemconfig("table", width=event.width)
        
        canvas.bind('<Configure>', _on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Layout responsivo - alterada a ordem de empacotamento
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configurações de grid
        for i in range(5):
            table_frame.grid_columnconfigure(i, weight=1)
        
        return {
            'container': table_container,
            'frame': table_frame,
            'canvas': canvas,
            'scrollbar': scrollbar
        }
    
    def _smooth_scroll(self, scrollbar):
            """Implementa scrolling suave."""
            def scroll(*args):
                scrollbar.set(*args)
                self.ui_refs['table']['canvas'].yview_moveto(args[0])
            return scroll

    def _sort_table(self, column: int):
        """Ordena a tabela pela coluna clicada."""
        if self.sort_state['column'] == column:
            self.sort_state['reverse'] = not self.sort_state['reverse']
        else:
            self.sort_state['column'] = column
            self.sort_state['reverse'] = False
        
        self._apply_filters()

    def _on_header_hover(self, widget: tk.Label, entering: bool):
        """Efeito hover nos cabeçalhos."""
        widget.configure(
            bg=self.theme['hover'] if entering else self.theme['header']
        )

    def _update_table(self, data: List[PatientData]) -> None:
        """Atualiza a tabela com os dados filtrados."""
        table = self.ui_refs['table']
        
        # Limpa tabela preservando cabeçalho
        for widget in table['frame'].winfo_children():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()

        # Separa os dados por tipo
        medicos = [p for p in data if p.tipo == "Médico"]
        psicologos = [p for p in data if p.tipo == "Psicólogo"]
        
        row = 1
        # Processa médicos
        for idx, patient in enumerate(medicos, 1):
            bg_color = self.theme['highlight'] if idx % 2 == 0 else self.theme['background']
            
            cells = [
                (str(idx), "center", 5),
                (patient.nome, "w", 30),
                (patient.renach, "center", 10),
                (patient.pagamento, "w", 20),
                (patient.tipo, "center", 10)
            ]
            
            for col, (text, anchor, width) in enumerate(cells):
                tk.Label(
                    table['frame'],
                    text=text,
                    bg=bg_color,
                    fg=self.theme['text'],
                    font=("Arial", 10),
                    anchor=anchor,
                    width=width,
                    padx=10,
                    pady=5
                ).grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
            
            row += 1

        # Adiciona separador se houver médicos e psicólogos
        if medicos and psicologos:
            separator = tk.Frame(
                table['frame'],
                height=2,
                bg=self.theme['separator']
            )
            separator.grid(
                row=row,
                column=0,
                columnspan=5,
                sticky="ew",
                pady=5
            )
            row += 1

        # Processa psicólogos
        for idx, patient in enumerate(psicologos, 1):
            bg_color = self.theme['highlight'] if row % 2 == 0 else self.theme['background']
            
            cells = [
                (str(idx), "center", 5),
                (patient.nome, "w", 30),
                (patient.renach, "center", 10),
                (patient.pagamento, "w", 20),
                (patient.tipo, "center", 10)
            ]
            
            for col, (text, anchor, width) in enumerate(cells):
                tk.Label(
                    table['frame'],
                    text=text,
                    bg=bg_color,
                    fg=self.theme['text'],
                    font=("Arial", 10),
                    anchor=anchor,
                    width=width,
                    padx=10,
                    pady=5
                ).grid(row=row, column=col, sticky="nsew", padx=1, pady=1)
            
            row += 1

        # Atualiza scroll region
        table['frame'].update_idletasks()
        table['canvas'].configure(scrollregion=table['canvas'].bbox("all"))
        
    def _update_stats(self, filtered_data: List[PatientData]) -> None:
        """Atualiza as estatísticas."""
        med_count = sum(1 for p in filtered_data if p.tipo == "Médico")
        psi_count = sum(1 for p in filtered_data if p.tipo == "Psicólogo")
        total = len(filtered_data)
        
        stats = f"Total: {total} | Médico: {med_count} | Psicólogo: {psi_count}"
        self.ui_refs['stats'].config(text=stats)

    def _filter_data(self) -> List[PatientData]:
            """Sistema de filtragem otimizado com cache de resultados."""
            filters = self.ui_refs['filters']
            current_filters = {
                'type': filters['type'].get(),
                'search': filters['search'].get().lower(),
                'payment': filters['payment'].get().lower()
            }
            
            # Verifica se os filtros mudaram
            if current_filters == self.data_cache['last_filters']:
                return self.data_cache.get('last_result', [])
            
            def matches_criteria(patient: PatientData) -> bool:
                if current_filters['search'] and current_filters['search'] not in patient.search_text:
                    return False
                if current_filters['payment'] and current_filters['payment'] not in patient.pagamento.lower():
                    return False
                return True
            
            filtered = []
            if current_filters['type'] in ["todos", "medico"]:
                filtered.extend([p for p in self.data_cache['medico'] if matches_criteria(p)])
            
            if current_filters['type'] in ["todos", "psi"]:
                filtered.extend([p for p in self.data_cache['psi'] if matches_criteria(p)])
            
            # Aplica ordenação
            if self.sort_state['column'] is not None:
                key_funcs = [
                    lambda x: int(x.nome.split()[0]) if x.nome.split()[0].isdigit() else 0,
                    lambda x: x.nome.lower(),
                    lambda x: x.renach,
                    lambda x: x.pagamento,
                    lambda x: x.tipo
                ]
                
                filtered.sort(
                    key=key_funcs[self.sort_state['column']],
                    reverse=self.sort_state['reverse']
                )
            
            # Atualiza cache
            self.data_cache['last_filters'] = current_filters.copy()
            self.data_cache['last_result'] = filtered
            
            return filtered
    
    def _delayed_filter(self) -> None:
            """Implementa filtragem com delay para melhor performance."""
            if self.data_cache['timer']:
                self.master.after_cancel(self.data_cache['timer'])
            self.data_cache['timer'] = self.master.after(300, self._apply_filters)

    def _apply_filters(self) -> None:
        """Aplica filtros com animação suave."""
        if self.data_cache['timer']:
            self.master.after_cancel(self.data_cache['timer'])
        
        def animate():
            table = self.ui_refs['table']['frame']
            table.tk_setPalette(background=self.theme['background'])
            
            filtered_data = self._filter_data()
            self._update_table(filtered_data)
            self._update_stats(filtered_data)
            
            table.tk_setPalette(background=self.theme['background'])
        
        self.data_cache['timer'] = self.master.after(150, animate)

    def display(self) -> None:
        """Exibe a interface principal com tratamento de erros aprimorado."""
        try:
            if not self._load_data():
                messagebox.showerror("Erro", "Não foi possível carregar os dados")
                return
            
            window, _ = self._create_ui()
            self._apply_filters()
            
            # Centralização e dimensionamento responsivo
            window.update_idletasks()
            width = max(window.winfo_width(), self.ui_config['min_width'])
            height = max(window.winfo_height(), self.ui_config['min_height'])
            x = (window.winfo_screenwidth() // 2) - (width // 2)
            y = (window.winfo_screenheight() // 2) - (height // 2)
            window.geometry(f"{width}x{height}+{x}+{y}")
            
            # Configuração de redimensionamento
            window.minsize(self.ui_config['min_width'], self.ui_config['min_height'])
            
            def on_closing():
                if self.data_cache['timer']:
                    self.master.after_cancel(self.data_cache['timer'])
                window.destroy()
            
            window.protocol("WM_DELETE_WINDOW", on_closing)
            
        except Exception as e:
            self.logger.error(f"Erro ao exibir interface: {e}")
            messagebox.showerror(
                "Erro",
                f"Ocorreu um erro ao exibir a interface: {str(e)}\nPor favor, contate o suporte."
            )