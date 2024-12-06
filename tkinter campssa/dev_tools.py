import tkinter as tk
from tkinter import ttk, messagebox
import logging
import sqlite3
import json
import hashlib
from datetime import datetime
from database_connection import DatabaseConnection


class StandaloneDevTools:
    PERMISSIONS = {
        'manage_users': 'Gerenciar Usuários',
        'view_reports': 'Visualizar Relatórios',
        'edit_data': 'Editar Dados',
        'export_data': 'Exportar Dados',
        'admin_access': 'Acesso Administrativo'
    }
    
    DEFAULT_PERMISSIONS = {
        'admin': list(PERMISSIONS.keys()),
        'manager': ['view_reports', 'edit_data', 'export_data'],
        'employee': ['view_reports', 'edit_data']
    }

    def __init__(self):
        self.root = tk.Tk()
        self.setup_logging()
        self.setup_database()
        self.setup_main_window()
        
    def setup_logging(self):
        """Configura o sistema de logging"""
        logging.basicConfig(
            filename='devtools.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
        
    def setup_database(self):
        """Inicializa a estrutura do banco de dados"""
        try:
            with DatabaseConnection('login.db') as conn:
                cursor = conn.cursor()
                
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user TEXT NOT NULL UNIQUE,
                        password TEXT NOT NULL,
                        role TEXT NOT NULL,
                        permissions TEXT,
                        is_active INTEGER DEFAULT 1,
                        created_by TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        last_login TIMESTAMP
                    )
                """)
                
                # Verifica se existe um usuário admin
                cursor.execute("SELECT COUNT(*) FROM users WHERE user = 'admin'")
                if cursor.fetchone()[0] == 0:
                    # Cria usuário admin padrão se não existir
                    hashed_password = hashlib.sha256('admin123'.encode()).hexdigest()
                    cursor.execute("""
                        INSERT INTO users (user, password, role, permissions, is_active)
                        VALUES (?, ?, ?, ?, ?)
                    """, ('admin', hashed_password, 'admin', json.dumps(['all']), 1))
                
                conn.commit()
                self.logger.info("Database initialized successfully")
                
        except sqlite3.Error as e:
            self.logger.error(f"Database setup error: {e}")
            raise
        
    def setup_main_window(self):
        """Configura a janela principal"""
        self.root.title("Developer Tools - User Management")
        self.root.geometry("800x600")
        self.root.configure(bg='#2c3e50')
        
        style = ttk.Style()
        style.configure('User.TFrame', background='#2c3e50')
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Create users tab
        self.setup_users_tab()
        
    def setup_users_tab(self):
        """Configura a aba de gerenciamento de usuários"""
        users_frame = ttk.Frame(self.notebook, padding="20", style='User.TFrame')
        self.notebook.add(users_frame, text="Gerenciamento de Usuários")
        
        # User list section
        list_frame = ttk.LabelFrame(users_frame, text="Lista de Usuários", padding="10")
        list_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Create Treeview
        columns = ('Username', 'Role', 'Status', 'Created By', 'Last Login')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings')
        
        # Configure columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Button frame
        btn_frame = ttk.Frame(users_frame)
        btn_frame.pack(fill='x', pady=5)
        
        ttk.Button(
            btn_frame,
            text="Atualizar Lista",
            command=self.load_users
        ).pack(side='left', padx=5)
        
        ttk.Button(
            btn_frame,
            text="Adicionar Usuário",
            command=self.show_add_user
        ).pack(side='left', padx=5)
        
        ttk.Button(
            btn_frame,
            text="Detalhes do Usuário",
            command=self.show_user_details
        ).pack(side='left', padx=5)
        
        # Carrega usuários inicialmente
        self.load_users()
        
    def check_user_exists(self, username):
        """Verifica se um usuário já existe"""
        try:
            with DatabaseConnection('login.db') as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM users WHERE user = ?", (username,))
                return cursor.fetchone()[0] > 0
        except sqlite3.Error as e:
            self.logger.error(f"Erro ao verificar usuário: {e}")
            return False

    def create_user_in_db(self, username, password, role, permissions):
        """Cria um novo usuário no banco de dados"""
        try:
            with DatabaseConnection('login.db') as conn:
                cursor = conn.cursor()
                
                hashed_password = hashlib.sha256(password.encode()).hexdigest()
                permissions_json = json.dumps(permissions)
                
                cursor.execute("""
                    INSERT INTO users (
                        user, password, role, permissions, is_active
                    ) VALUES (?, ?, ?, ?, ?)
                """, (username, hashed_password, role, permissions_json, 1))
                
                conn.commit()
                return True
                
        except sqlite3.Error as e:
            self.logger.error(f"Erro ao criar usuário: {e}")
            return False

    def load_users(self):
        """Carrega a lista de usuários"""
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            with DatabaseConnection('login.db') as conn:
                cursor = conn.cursor()
                
                cursor.execute("""
                    SELECT user, role, is_active, created_by, last_login
                    FROM users
                    ORDER BY user
                """)
                
                for user in cursor.fetchall():
                    username, role, is_active, created_by, last_login = user
                    status = "Ativo" if is_active else "Inativo"
                    created_by = created_by if created_by else "Sistema"
                    
                    self.tree.insert('', 'end', values=(
                        username,
                        role,
                        status,
                        created_by,
                        last_login or "Nunca"
                    ))
                
        except sqlite3.Error as e:
            self.logger.error(f"Erro ao carregar usuários: {e}")
            messagebox.showerror("Erro", "Falha ao carregar usuários do banco")

    def show_add_user(self):
        """Mostra a janela de adição de usuário"""
        window = tk.Toplevel(self.root)
        window.title("Criar Usuário")
        window.geometry("500x400")
        window.configure(bg='#2c3e50')
        
        frame = ttk.Frame(window, padding="20")
        frame.pack(fill='both', expand=True)
        
        # Username field
        ttk.Label(frame, text="Username:").pack(anchor='w', pady=2)
        username_entry = ttk.Entry(frame, width=40)
        username_entry.pack(fill='x', pady=2)
        
        # Password field
        ttk.Label(frame, text="Password:").pack(anchor='w', pady=2)
        password_entry = ttk.Entry(frame, show="*", width=40)
        password_entry.pack(fill='x', pady=2)
        
        # Role selection
        ttk.Label(frame, text="Role:").pack(anchor='w', pady=2)
        role_var = tk.StringVar(value='employee')
        role_frame = ttk.Frame(frame)
        role_frame.pack(fill='x', pady=2)
        
        ttk.Radiobutton(role_frame, text="Employee", variable=role_var, value='employee').pack(side='left', padx=10)
        ttk.Radiobutton(role_frame, text="Manager", variable=role_var, value='manager').pack(side='left', padx=10)
        ttk.Radiobutton(role_frame, text="Admin", variable=role_var, value='admin').pack(side='left', padx=10)
        
        # Permissions Frame
        perm_frame = ttk.LabelFrame(frame, text="Permissions", padding=10)
        perm_frame.pack(fill='x', pady=10)
        
        permission_vars = {}
        for perm_key, perm_name in self.PERMISSIONS.items():
            var = tk.BooleanVar()
            permission_vars[perm_key] = var
            ttk.Checkbutton(
                perm_frame, 
                text=perm_name, 
                variable=var
            ).pack(anchor='w')
        
        def update_permissions(*args):
            """Atualiza permissões baseado no papel selecionado"""
            role = role_var.get()
            default_perms = self.DEFAULT_PERMISSIONS.get(role, [])
            for perm_key, var in permission_vars.items():
                var.set(perm_key in default_perms)
        
        role_var.trace('w', update_permissions)
        update_permissions()
        
        def create_user():
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            role = role_var.get()
            
            if not username or not password:
                messagebox.showerror("Erro", "Username e password são obrigatórios")
                return
            
            if self.check_user_exists(username):
                messagebox.showerror("Erro", "Username já existe")
                return
            
            selected_permissions = [
                perm for perm, var in permission_vars.items()
                if var.get()
            ]
            
            if self.create_user_in_db(username, password, role, selected_permissions):
                messagebox.showinfo("Sucesso", "Usuário criado com sucesso!")
                window.destroy()
                self.load_users()
            else:
                messagebox.showerror("Erro", "Falha ao criar usuário")
        
        ttk.Button(
            frame,
            text="Criar Usuário",
            command=create_user
        ).pack(pady=20)
        
        self.center(window)

    def show_user_details(self):
        """Mostra os detalhes do usuário selecionado"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Por favor, selecione um usuário")
            return
            
        user_data = self.tree.item(selection[0])['values']
        username = user_data[0]
        
        window = tk.Toplevel(self.root)
        window.title(f"Detalhes do Usuário - {username}")
        window.geometry("600x400")
        window.configure(bg='#2c3e50')
        
        frame = ttk.Frame(window, padding="20")
        frame.pack(fill='both', expand=True)
        
        try:
            with DatabaseConnection('login.db') as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT user, role, created_by, last_login, is_active, permissions
                    FROM users WHERE user = ?
                """, (username,))
                user = cursor.fetchone()
                
                if user:
                    details = ttk.LabelFrame(frame, text="Informações do Usuário", padding="10")
                    details.pack(fill='x', pady=5)
                    
                    ttk.Label(details, text=f"Username: {user[0]}").pack(anchor='w')
                    ttk.Label(details, text=f"Role: {user[1]}").pack(anchor='w')
                    ttk.Label(details, text=f"Status: {'Ativo' if user[4] else 'Inativo'}").pack(anchor='w')
                    ttk.Label(details, text=f"Criado por: {user[2] or 'Sistema'}").pack(anchor='w')
                    ttk.Label(details, text=f"Último login: {user[3] or 'Nunca'}").pack(anchor='w')
                    
                    # Permissions
                    perm_frame = ttk.LabelFrame(frame, text="Permissões", padding="10")
                    perm_frame.pack(fill='x', pady=5)
                    
                    try:
                        permissions = json.loads(user[5]) if user[5] else []
                        for perm in permissions:
                            perm_name = self.PERMISSIONS.get(perm, perm)
                            ttk.Label(perm_frame, text=f"• {perm_name}").pack(anchor='w')
                    except json.JSONDecodeError:
                        ttk.Label(perm_frame, text="Sem permissões definidas").pack(anchor='w')
                    
                    # Action buttons frame
                    action_frame = ttk.Frame(frame)
                    action_frame.pack(fill='x', pady=10)
                    
                    def toggle_status():
                        """Altera o status do usuário entre ativo/inativo"""
                        if username == 'admin':
                            messagebox.showerror("Erro", "Não é possível modificar o usuário admin")
                            return
                        
                        new_status = not user[4]
                        try:
                            with DatabaseConnection('login.db') as conn:
                                cursor = conn.cursor()
                                cursor.execute("""
                                    UPDATE users 
                                    SET is_active = ? 
                                    WHERE user = ?
                                """, (new_status, username))
                                conn.commit()
                            
                            messagebox.showinfo("Sucesso", 
                                            f"Usuário {'ativado' if new_status else 'desativado'}")
                            self.load_users()
                            window.destroy()
                        except sqlite3.Error as e:
                            messagebox.showerror("Erro", f"Falha ao atualizar status: {str(e)}")
                            self.logger.error(f"Erro ao atualizar status do usuário: {e}")
                    
                    def delete_user():
                        """Exclui o usuário do sistema"""
                        if username == 'admin':
                            messagebox.showerror("Erro", "Não é possível excluir o usuário admin")
                            return
                            
                        if messagebox.askyesno("Confirmar Exclusão", 
                                            f"Tem certeza que deseja excluir o usuário {username}?\n"
                                            "Esta ação não pode ser desfeita."):
                            try:
                                with DatabaseConnection('login.db') as conn:
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM users WHERE user = ?", (username,))
                                    conn.commit()
                                
                                messagebox.showinfo("Sucesso", "Usuário excluído com sucesso")
                                self.load_users()
                                window.destroy()
                            except sqlite3.Error as e:
                                messagebox.showerror("Erro", f"Falha ao excluir usuário: {str(e)}")
                                self.logger.error(f"Erro ao excluir usuário: {e}")

                    def reset_password():
                        """Redefine a senha do usuário"""
                        if username == 'admin':
                            messagebox.showerror("Erro", "Não é possível redefinir a senha do admin")
                            return
                            
                        reset_window = tk.Toplevel(window)
                        reset_window.title("Redefinir Senha")
                        reset_window.geometry("300x150")
                        reset_window.configure(bg='#2c3e50')
                        
                        reset_frame = ttk.Frame(reset_window, padding="20")
                        reset_frame.pack(fill='both', expand=True)
                        
                        ttk.Label(reset_frame, text="Nova Senha:").pack(anchor='w', pady=2)
                        password_entry = ttk.Entry(reset_frame, show="*")
                        password_entry.pack(fill='x', pady=2)
                        
                        def do_reset():
                            """Executa a redefinição da senha"""
                            new_password = password_entry.get().strip()
                            if not new_password:
                                messagebox.showerror("Erro", "A senha não pode estar vazia")
                                return
                                
                            try:
                                with DatabaseConnection('login.db') as conn:
                                    cursor = conn.cursor()
                                    hashed_password = hashlib.sha256(new_password.encode()).hexdigest()
                                    cursor.execute("""
                                        UPDATE users 
                                        SET password = ? 
                                        WHERE user = ?
                                    """, (hashed_password, username))
                                    conn.commit()
                                
                                messagebox.showinfo("Sucesso", "Senha redefinida com sucesso")
                                reset_window.destroy()
                            except sqlite3.Error as e:
                                messagebox.showerror("Erro", f"Falha ao redefinir senha: {str(e)}")
                                self.logger.error(f"Erro ao redefinir senha: {e}")
                        
                        ttk.Button(
                            reset_frame,
                            text="Redefinir",
                            command=do_reset
                        ).pack(pady=10)
                        
                        self.center(reset_window)
                    
                    # Status toggle button
                    ttk.Button(
                        action_frame,
                        text=f"{'Desativar' if user[4] else 'Ativar'} Usuário",
                        command=toggle_status
                    ).pack(side='left', padx=5)
                    
                    # Reset password button
                    ttk.Button(
                        action_frame,
                        text="Redefinir Senha",
                        command=reset_password
                    ).pack(side='left', padx=5)
                    
                    # Delete user button
                    ttk.Button(
                        action_frame,
                        text="Excluir Usuário",
                        command=delete_user
                    ).pack(side='left', padx=5)
        
        except sqlite3.Error as e:
            self.logger.error(f"Erro ao carregar detalhes do usuário: {e}")
            messagebox.showerror("Erro", "Falha ao carregar detalhes do usuário")

        self.center(window)

    def center(self, window):
        """Centraliza uma janela na tela"""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')
        
    def run(self):
        """Inicia a aplicação"""
        self.center(self.root)
        self.root.mainloop()

if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('devtools.log'),
            logging.StreamHandler()
        ]
    )
    
    try:
        app = StandaloneDevTools()
        app.run()
    except Exception as e:
        logging.error(f"Fatal error: {e}")
        if hasattr(app, 'root'):
            messagebox.showerror("Error", f"Fatal error occurred: {str(e)}")