import sqlite3


def create_db():
    conn = sqlite3.connect("login.db")
    cursor = conn.cursor()

    # Criação da tabela de usuários
    cursor.execute(
        """CREATE TABLE IF NOT EXISTS users (
                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                   user TEXT NOT NULL,
                   password TEXT NOT NULL)"""
    )

    conn.commit()
    conn.close()


# Função CRUD
class DataBase:
    def __init__(self, db_name='login.db'):
        self.db_name = db_name
        self.create_db()

    # Função para criar o banco de dados e a tabela de usuários
    def create_db(self):
        conn = sqlite3.connect("login.db")
        cursor = conn.cursor()

        # Criação da tabela de usuários
        cursor.execute(
            """CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user TEXT NOT NULL,
                    password TEXT NOT NULL)"""
        )

        conn.commit()
        conn.close()

    # Função para criar novo usuário
    def create_user(self, user, password):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('INSERT INTO users (user, password) VALUES (?, ?)', (user, password))
        conn.commit()
        conn.close()
        print(f'usuario {user} criado com sucesso')

    # Função para ser um usuário com base no user
    def read_user(self, user):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE user = ?', (user,))
        usuario = cursor.fetchone()
        conn.close()
        return usuario
    
    # Função para atualizar a senha de um usuário
    def update_user(self, user, new_password):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('UPDATE users SET password = ? WHERE user =?', (new_password, user))

     # Função para deletar um usuário com base no user
    def delete_user(self, user):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM users WHERE user = ?', (user,))
        conn.comit()
        conn.close()