from tkinter import Tk, Frame, Label, Toplevel, BooleanVar
import tkinter as tk
from tkinter import messagebox
from config import ConfigManager
import json
import base64
from cryptography.fernet import Fernet
import os
import re


class EmitirNota:
    def __init__(self, master):
        self.master = master
        self.config_manager = ConfigManager()
        self.ui_config = self.config_manager.get_config("UI_CONFIG")
        self.window = None
        self.remember_var = None
        self.key = b"YOUR_SECRET_KEY_HERE"  # Substitua por uma chave segura em produção
        self.cipher_suite = Fernet(base64.urlsafe_b64encode(self.key.ljust(32)[:32]))
        self.credentials_file = "saved_credentials.enc"
        self.second_window = None
        self.entry_user = None
        self.entry_password = None

    def format_cnpj(self, cnpj):
        """Formate o CNPJ com máscaras: XX.XXX.XXX/XXXX-XX"""
        cnpj = "".join(filter(str.isdigit, cnpj))
        if len(cnpj) != 14:
            return cnpj
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    def encrypt_data(self, data):
        return self.cipher_suite.encrypt(json.dumps(data).encode())

    def decrypt_data(self, encrypted_data):
        try:
            return json.loads(self.cipher_suite.decrypt(encrypted_data).decode())
        except:
            return None

    def save_credentials(self, username, password):
        data = {"username": username, "password": password}
        encrypted_data = self.encrypt_data(data)
        with open(self.credentials_file, "wb") as f:
            f.write(encrypted_data)

    def load_credentials(self):
        try:
            if os.path.exists(self.credentials_file):
                with open(self.credentials_file, "rb") as f:
                    encrypted_data = f.read()
                return self.decrypt_data(encrypted_data)
        except:
            pass
        return None

    def clear_saved_credentials(self):
        if os.path.exists(self.credentials_file):
            os.remove(self.credentials_file)

    def show(self):
        """Mostra a primeira janela de login"""
        if self.window is not None:
            try:
                self.window.state()
                self.window.lift()
                return
            except tk.TclError:
                self.window = None

        self.window = Toplevel(self.master)
        self.window.title("Emitir Nota")
        self._configure_window()
        self.create_widgets()

        # Carrega credenciais salvas
        saved_creds = self.load_credentials()
        if saved_creds:
            self.entry_user.insert(0, saved_creds["username"])
            self.entry_password.insert(0, saved_creds["password"])
            self.remember_var.set(True)

        self.window.transient(self.master)
        self.window.grab_set()

    def _configure_window(self):
        width, height = 400, 350
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.minsize(width, height)
        self.window.maxsize(width, height)
        self.window.resizable(False, False)
        self.window.configure(bg=self.ui_config["colors"]["background"])

    def on_username_change(self, *args):
        current_password = self.entry_password.get()
        username = self.entry_user.get()
        formatted_username = self.format_cnpj(username)

        cursor_position = self.entry_user.index(tk.INSERT)
        self.entry_user.delete(0, tk.END)
        self.entry_user.insert(0, formatted_username)

        try:
            self.entry_user.icursor(cursor_position)
        except tk.TclError:
            pass

        if self.remember_var.get():
            saved_creds = self.load_credentials()
            if saved_creds and saved_creds["username"] == formatted_username:
                self.entry_password.delete(0, tk.END)
                self.entry_password.insert(0, saved_creds["password"])
            else:
                self.entry_password.delete(0, tk.END)
                self.entry_password.insert(0, current_password)

    def on_remember_change(self):
        if self.remember_var.get():
            saved_creds = self.load_credentials()
            if saved_creds and saved_creds["username"] == self.entry_user.get():
                self.entry_password.delete(0, tk.END)
                self.entry_password.insert(0, saved_creds["password"])
        else:
            self.clear_saved_credentials()

    def create_widgets(self):
        colors = self.ui_config["colors"]
        fonts = self.ui_config["fonts"]

        container = tk.Frame(self.window, bg=colors["background"])
        container.pack(expand=True)

        # Title
        tk.Label(
            container,
            text="Conta Clinica",
            font=fonts["title"],
            bg=colors["background"],
            fg=colors["title"],
        ).grid(row=0, column=0, columnspan=2, pady=20)

        # CNPJ field
        tk.Label(
            container,
            text="CNPJ:",
            font=fonts["normal"],
            bg=colors["background"],
            fg=colors["text"],
        ).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="e")

        self.entry_user = tk.Entry(
            container,
            font=fonts["normal"],
            width=20,
            bg=colors["frame"],
            fg=colors["text"],
            insertbackground=colors["text"],
        )
        self.entry_user.grid(row=1, column=1, pady=5, sticky="w")

        # Password field
        tk.Label(
            container,
            text="Senha:",
            font=fonts["normal"],
            bg=colors["background"],
            fg=colors["text"],
        ).grid(row=2, column=0, padx=(0, 10), pady=5, sticky="e")

        self.entry_password = tk.Entry(
            container,
            font=fonts["normal"],
            show="*",
            width=20,
            bg=colors["frame"],
            fg=colors["text"],
            insertbackground=colors["text"],
        )
        self.entry_password.grid(row=2, column=1, pady=5, sticky="w")

        # Remember me checkbox
        self.remember_var = BooleanVar()
        tk.Checkbutton(
            container,
            text="Lembrar credenciais",
            variable=self.remember_var,
            font=fonts["normal"],
            bg=colors["background"],
            fg=colors["text"],
            selectcolor=colors["frame"],
            command=self.on_remember_change,
        ).grid(row=3, column=0, columnspan=2, pady=10)

        # Track username changes
        self.entry_user.bind("<KeyRelease>", self.on_username_change)

        # Buttons
        button_frame = tk.Frame(container, bg=colors["background"])
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)

        tk.Button(
            button_frame,
            text="Entrar",
            font=fonts["button"],
            bg=colors["button"],
            fg=colors["text"],
            width=15,
            command=self.login,
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame,
            text="Cancelar",
            font=fonts["button"],
            bg=colors["button"],
            fg=colors["text"],
            width=15,
            command=self.close_window,
        ).pack(side="left", padx=5)

    def close_window(self):
        if self.window:
            self.window.destroy()
            self.window = None

    def login(self):
        usuario = self.entry_user.get()
        senha = self.entry_password.get()

        if usuario and senha:
            if self.remember_var.get():
                self.save_credentials(usuario, senha)
            else:
                self.clear_saved_credentials()
            
            if self.login_callback:
                self.login_callback(usuario, senha)  # Passando os argumentos corretamente
                self.close_window()
            return True
        else:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos!")
            return False

    def show_second_window(self):
        """Mostra a segunda janela de login"""
        if self.second_window is None:
            self.second_window = SecondLoginWindow(self.master, self.config_manager)
        self.second_window.show()
        return self.second_window




class SecondLoginWindow:
    def __init__(self, master, config_manager):
        self.master = master
        self.config_manager = config_manager
        self.ui_config = config_manager.get_config("UI_CONFIG")
        self.window = None
        self.remember_var = None
        self.key = b"YOUR_SECRET_KEY_HERE"  # Use the same key for consistency
        self.cipher_suite = Fernet(base64.urlsafe_b64encode(self.key.ljust(32)[:32]))
        self.credentials_file = (
            "second_credentials.enc"  # Different file for second window
        )
        self.login_callback = None

    def format_cnpj(self, cnpj):
        """Formate o CNPJ com máscaras: XX.XXX.XXX/XXXX-XX"""
        cnpj = "".join(filter(str.isdigit, cnpj))
        if len(cnpj) != 14:
            return cnpj
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    def unformat_cnpj(self, cnpj):
        """Remove CNPJ formatting"""
        return "".join(filter(str.isdigit, cnpj))

    def encrypt_data(self, data):
        return self.cipher_suite.encrypt(json.dumps(data).encode())

    def decrypt_data(self, encrypted_data):
        try:
            return json.loads(self.cipher_suite.decrypt(encrypted_data).decode())
        except:
            return None

    def save_credentials(self, username, password):
        data = {"username": username, "password": password}
        encrypted_data = self.encrypt_data(data)
        with open(self.credentials_file, "wb") as f:
            f.write(encrypted_data)

    def load_credentials(self):
        try:
            if os.path.exists(self.credentials_file):
                with open(self.credentials_file, "rb") as f:
                    encrypted_data = f.read()
                return self.decrypt_data(encrypted_data)
        except:
            pass
        return None

    def clear_saved_credentials(self):
        if os.path.exists(self.credentials_file):
            os.remove(self.credentials_file)

    def show(self):
        if self.window is not None:
            try:
                self.window.state()
                self.window.lift()
                return
            except tk.TclError:
                self.window = None

        self.window = Toplevel(self.master)
        self.window.title("Segunda Autenticação")
        self._configure_window()
        self.create_widgets()

        # Carrega credenciais salvas
        saved_creds = self.load_credentials()
        if saved_creds:
            self.entry_user.insert(0, saved_creds["username"])
            self.entry_password.insert(0, saved_creds["password"])
            self.remember_var.set(True)

        self.window.transient(self.master)
        self.window.grab_set()

        def on_closing():
            self.window.destroy()
            self.window = None

        self.window.protocol("WM_DELETE_WINDOW", on_closing)

    def _configure_window(self):
        width, height = 400, 350
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        self.window.minsize(width, height)
        self.window.maxsize(width, height)
        self.window.resizable(False, False)
        self.window.configure(bg=self.ui_config["colors"]["background"])

    def on_username_change(self, *args):
        current_password = self.entry_password.get()
        username = self.entry_user.get()
        formatted_username = self.format_cnpj(username)

        cursor_position = self.entry_user.index(tk.INSERT)
        self.entry_user.delete(0, tk.END)
        self.entry_user.insert(0, formatted_username)

        try:
            self.entry_user.icursor(cursor_position)
        except tk.TclError:
            pass

        if self.remember_var.get():
            saved_creds = self.load_credentials()
            if saved_creds and saved_creds["username"] == formatted_username:
                self.entry_password.delete(0, tk.END)
                self.entry_password.insert(0, saved_creds["password"])
            else:
                self.entry_password.delete(0, tk.END)
                self.entry_password.insert(0, current_password)

    def on_remember_change(self):
        if self.remember_var.get():
            saved_creds = self.load_credentials()
            if saved_creds and saved_creds["username"] == self.entry_user.get():
                self.entry_password.delete(0, tk.END)
                self.entry_password.insert(0, saved_creds["password"])
        else:
            self.clear_saved_credentials()

    def create_widgets(self):
        colors = self.ui_config["colors"]
        fonts = self.ui_config["fonts"]

        container = tk.Frame(self.window, bg=colors["background"])
        container.pack(expand=True)

        # Title
        tk.Label(
            container,
            text="Segunda Autenticação",
            font=fonts["title"],
            bg=colors["background"],
            fg=colors["title"],
        ).grid(row=0, column=0, columnspan=2, pady=20)

        # CNPJ field
        tk.Label(
            container,
            text="CNPJ:",
            font=fonts["normal"],
            bg=colors["background"],
            fg=colors["text"],
        ).grid(row=1, column=0, padx=(0, 10), pady=5, sticky="e")

        self.entry_user = tk.Entry(
            container,
            font=fonts["normal"],
            width=20,
            bg=colors["frame"],
            fg=colors["text"],
            insertbackground=colors["text"],
        )
        self.entry_user.grid(row=1, column=1, pady=5, sticky="w")

        # Password field
        tk.Label(
            container,
            text="Senha:",
            font=fonts["normal"],
            bg=colors["background"],
            fg=colors["text"],
        ).grid(row=2, column=0, padx=(0, 10), pady=5, sticky="e")

        self.entry_password = tk.Entry(
            container,
            font=fonts["normal"],
            show="*",
            width=20,
            bg=colors["frame"],
            fg=colors["text"],
            insertbackground=colors["text"],
        )
        self.entry_password.grid(row=2, column=1, pady=5, sticky="w")

        # Remember me checkbox
        self.remember_var = BooleanVar()
        tk.Checkbutton(
            container,
            text="Lembrar credenciais",
            variable=self.remember_var,
            font=fonts["normal"],
            bg=colors["background"],
            fg=colors["text"],
            selectcolor=colors["frame"],
            command=self.on_remember_change,
        ).grid(row=3, column=0, columnspan=2, pady=10)

        # Track username changes
        self.entry_user.bind("<KeyRelease>", self.on_username_change)

        # Buttons
        button_frame = tk.Frame(container, bg=colors["background"])
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)

        tk.Button(
            button_frame,
            text="Confirmar",
            font=fonts["button"],
            bg=colors["button"],
            fg=colors["text"],
            width=15,
            command=self.login,
        ).pack(side="left", padx=5)

        tk.Button(
            button_frame,
            text="Cancelar",
            font=fonts["button"],
            bg=colors["button"],
            fg=colors["text"],
            width=15,
            command=lambda: self.close_window(),
        ).pack(side="left", padx=5)

    def close_window(self):
        if self.window is not None:
            self.window.destroy()
            self.window = None

    def login(self):
        usuario = self.entry_user.get()
        senha = self.entry_password.get()

        if usuario and senha:
            if self.remember_var.get():
                self.save_credentials(usuario, senha)
            else:
                self.clear_saved_credentials()

            messagebox.showinfo("Sucesso", "Segunda autenticação realizada com sucesso!")
            
            # Armazenar os valores em variáveis antes de fechar a janela
            self.usuario = usuario
            self.senha = senha
            
            self.close_window()

            if self.login_callback:
                self.login_callback(True)  # Chama o callback com o resultado do login

            return True
        else:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos!")
            return False