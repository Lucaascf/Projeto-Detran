# /home/lusca/py_excel/tkinter campssa/auth/user_manager.py
import sqlite3
import hashlib
import json
from dataclasses import dataclass
from typing import List, Dict, Optional
from datetime import datetime
import logging



@dataclass
class User:
    """Representa um usuário do sistema"""

    id: int
    username: str
    role: str
    permissions: List[str]
    created_by: Optional[int] = None
    created_at: str = ""
    last_login: str = ""
    is_active: bool = True


class UserManager:
    """Gerencia todas as operações relacionadas a usuários e permissões"""

    ROLES = {"admin": "Administrador", "manager": "Gerente", "employee": "Funcionário"}

    PERMISSIONS = {
        "manage_users": "Gerenciar Usuários",
        "view_reports": "Visualizar Relatórios",
        "edit_patients": "Editar Pacientes",
        "add_patients": "Adicionar Pacientes",
        "delete_patients": "Excluir Pacientes",
        "manage_appointments": "Gerenciar Agendamentos",
        "financial_access": "Acesso Financeiro",
        "export_data": "Exportar Dados",
    }

    DEFAULT_PERMISSIONS = {
        "admin": list(PERMISSIONS.keys()),
        "manager": [
            "view_reports",
            "edit_patients",
            "add_patients",
            "manage_appointments",
        ],
        "employee": ["add_patients", "view_reports"],
    }

    def __init__(self, db_path: str = "users.db"):
        self.db_path = db_path
        self.current_user: Optional[User] = None
        self.setup_database()

    def setup_database(self):
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()

                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        user TEXT NOT NULL UNIQUE,         -- Column for username
                        password TEXT NOT NULL,
                        role TEXT NOT NULL,
                        permissions TEXT NOT NULL,
                        created_by INTEGER,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        last_login TIMESTAMP,
                        is_active BOOLEAN DEFAULT 1,
                        FOREIGN KEY (created_by) REFERENCES users(id)
                    )
                """
                )

                cursor.execute("SELECT id FROM users WHERE user = 'admin'")
                if not cursor.fetchone():
                    admin_password = self._hash_password("admin123")
                    admin_permissions = json.dumps(self.DEFAULT_PERMISSIONS["admin"])
                    cursor.execute(
                        'INSERT INTO users (user, password, role, permissions) VALUES (?, ?, "admin", ?)',
                        ("admin", admin_password, admin_permissions),
                    )

                conn.commit()

        except sqlite3.Error as e:
            logging.error(f"Database setup error: {e}")
            raise

    def _hash_password(self, password: str) -> str:
        """Cria hash seguro da senha"""
        return hashlib.sha256(password.encode()).hexdigest()

    def authenticate(self, username: str, password: str) -> Optional[User]:
        """Autentica um usuário"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT id, role, permissions, created_by, created_at, last_login, is_active 
                    FROM users 
                    WHERE user = ? AND password = ? AND is_active = 1
                """,
                    (username, self._hash_password(password)),
                )

                result = cursor.fetchone()
                if not result:
                    return None

                # Atualiza último login
                cursor.execute(
                    """
                    UPDATE users 
                    SET last_login = CURRENT_TIMESTAMP 
                    WHERE id = ?
                """,
                    (result[0],),
                )

                self.current_user = User(
                    id=result[0],
                    username=username,
                    role=result[1],
                    permissions=json.loads(result[2]),
                    created_by=result[3],
                    created_at=result[4],
                    last_login=result[5],
                    is_active=result[6],
                )

                conn.commit()
                logging.info(f"User authenticated: {username}")
                return self.current_user

        except sqlite3.Error as e:
            logging.error(f"Authentication error: {e}")
            return None


    def create_user(
        self, username: str, password: str, role: str, permissions: List[str]
    ) -> bool:
        """Cria um novo usuário"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()

                # Verifica se usuário já existe
                cursor.execute("SELECT id FROM users WHERE user = ?", (username,))
                if cursor.fetchone():
                    return False

                # Insere novo usuário
                cursor.execute(
                    """
                    INSERT INTO users (user, password, role, permissions, created_by)
                    VALUES (?, ?, ?, ?, ?)
                """,
                    (
                        username,
                        self._hash_password(password),
                        role,
                        json.dumps(permissions),
                        self.current_user.id if self.current_user else None,
                    ),
                )

                conn.commit()
                logging.info(f"New user created: {username}")
                return True

        except sqlite3.Error as e:
            logging.error(f"Error creating user: {e}")
            return False


    def create_first_user(self, activation_key: str, username: str, password: str) -> bool:
        """Cria o primeiro usuário com base na chave de ativação."""
        if not self.token_manager.validate_admin_token(activation_key):  # Usando o método correto
            return False
        
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Verifica se já existe algum usuário
                cursor.execute("SELECT COUNT(*) FROM users")
                if cursor.fetchone()[0] > 0:
                    return False
                
                # Cria o primeiro usuário como admin
                hashed_password = self._hash_password(password)
                admin_permissions = json.dumps(self.DEFAULT_PERMISSIONS["admin"])
                
                cursor.execute(
                    """
                    INSERT INTO users (user, password, role, permissions)
                    VALUES (?, ?, 'admin', ?)
                    """,
                    (username, hashed_password, admin_permissions)
                )
                
                conn.commit()
                return True
        
        except sqlite3.Error as e:
            logging.error(f"Error creating first user: {e}")
            return False


    def update_user(self, user_id: int, updates: Dict) -> bool:
        """Atualiza informações do usuário"""
        if not self.current_user or "manage_users" not in self.current_user.permissions:
            return False

        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()

                update_fields = []
                update_values = []

                # Prepara campos para atualização
                if "password" in updates:
                    update_fields.append("password = ?")
                    update_values.append(self._hash_password(updates["password"]))

                if "role" in updates:
                    update_fields.append("role = ?")
                    update_values.append(updates["role"])

                if "permissions" in updates:
                    update_fields.append("permissions = ?")
                    update_values.append(json.dumps(updates["permissions"]))

                if "is_active" in updates:
                    update_fields.append("is_active = ?")
                    update_values.append(updates["is_active"])

                if not update_fields:
                    return False

                # Executa atualização
                update_values.append(user_id)
                cursor.execute(
                    f"""
                    UPDATE users 
                    SET {", ".join(update_fields)}
                    WHERE id = ?
                """,
                    update_values,
                )

                # Registra no histórico
                cursor.execute(
                    """
                    INSERT INTO user_history (user_id, action, details, performed_by)
                    VALUES (?, 'update', ?, ?)
                """,
                    (
                        user_id,
                        f"User updated: {', '.join(updates.keys())}",
                        self.current_user.id,
                    ),
                )

                conn.commit()
                logging.info(f"User {user_id} updated by {self.current_user.username}")
                return True

        except sqlite3.Error as e:
            logging.error(f"Error updating user: {e}")
            return False

    def get_users(self) -> List[User]:
        """Retorna lista de todos os usuários"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT id, user, role, permissions, created_by, created_at, last_login, is_active 
                    FROM users
                    ORDER BY user
                """
                )

                users = []
                for row in cursor.fetchall():
                    users.append(
                        User(
                            id=row[0],
                            username=row[1],
                            role=row[2],
                            permissions=json.loads(row[3]),
                            created_by=row[4],
                            created_at=row[5],
                            last_login=row[6],
                            is_active=row[7],
                        )
                    )

                return users

        except sqlite3.Error as e:
            logging.error(f"Error fetching users: {e}")
            return []
        


    def verificar_permissao(self, permissao: str) -> bool:
        """Verifica se o usuário atual tem a permissão especificada"""
        if not self.current_user:
            return False
        
        if permissao in self.current_user.permissions:
            return True
        
        if self.current_user.role == "admin":
            return True
        
        if self.current_user.role == "manager" and permissao in self.DEFAULT_PERMISSIONS["manager"]:
            return True
        
        if self.current_user.role == "employee" and permissao in self.DEFAULT_PERMISSIONS["employee"]:
            return True
        
        return False