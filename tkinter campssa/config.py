# config.py
import json
import os
from pathlib import Path
from typing import Dict, Any
import logging

class ConfigManager:
    """Gerenciador de configurações dinâmicas da aplicação."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.current_user = None
        self.config_dir = Path("configs")
        self.config_dir.mkdir(exist_ok=True)
        self._load_default_config()

    def _load_default_config(self):
        """
        Carrega as configurações padrão do sistema.
        Esta função define o estado inicial e os valores padrão para todas as configurações.
        
        A estrutura de configuração inclui:
        - APP_CONFIG: Configurações gerais da aplicação
        - UI_CONFIG: Configurações de interface do usuário
        - USER_PREFERENCES: Preferências específicas do usuário
        - SYSTEM_SETTINGS: Configurações do sistema
        """
        self.default_config = {
            'APP_CONFIG': {
                'title': '',
                'version': '1.0.0',
                'background_color': '#2C3E50',
                'initial_geometry': '350x250',
                'main_geometry': '700x500',
                'file_types': [
                    ("Arquivos Excel", "*.xlsx"),
                    ("Todos os Arquivos", "*.*")
                ],
                'max_recent_files': 5,
                'auto_save': True,
                'auto_save_interval': 300  # segundos
            },
            'UI_CONFIG': {
                'colors': {
                    'background': '#2C3E50',
                    'text': '#ECF0F1',
                    'success_button': '#2ecc71',
                    'success_button_active': '#27ae60',
                    'danger_button': '#e74c3c',
                    'danger_button_active': '#c0392b',
                    'warning': '#f1c40f',
                    'info': '#3498db',
                    'disabled': '#95a5a6'
                },
                'fonts': {
                    'title': ('Arial', 14, 'bold'),
                    'header': ('Arial', 12, 'bold'),
                    'normal': ('Arial', 10),
                    'small': ('Arial', 9),
                    'monospace': ('Courier', 10)
                },
                'padding': {
                    'default': 10,
                    'small': 5,
                    'large': 20,
                    'extra_large': 30
                },
                'margins': {
                    'default': 10,
                    'small': 5,
                    'large': 20
                },
                'animations': {
                    'enabled': True,
                    'duration': 200
                }
            },
            'USER_PREFERENCES': {
                'theme': 'default',
                'language': 'pt_BR',
                'notifications': {
                    'enabled': True,
                    'sound': True,
                    'desktop': True
                },
                'display': {
                    'show_toolbar': True,
                    'show_statusbar': True,
                    'compact_mode': False
                },
                'table_settings': {
                    'rows_per_page': 10,
                    'sort_column': 'name',
                    'sort_direction': 'asc'
                },
                'recent_files': [],
                'default_save_location': ''
            },
            'SYSTEM_SETTINGS': {
                'debug_mode': False,
                'log_level': 'INFO',
                'backup': {
                    'enabled': True,
                    'interval': 24,  # horas
                    'keep_last': 7,  # dias
                    'location': './backups'
                },
                'performance': {
                    'cache_enabled': True,
                    'cache_size': 100,  # MB
                    'max_threads': 4
                },
                'security': {
                    'session_timeout': 30,  # minutos
                    'max_login_attempts': 3,
                    'password_expiry': 90,  # dias
                    'require_strong_password': True
                },
                'paths': {
                    'temp': './temp',
                    'logs': './logs',
                    'data': './data'
                }
            }
        }
        
        # Cria diretórios necessários
        self._create_required_directories()

    def _create_required_directories(self):
        """Cria os diretórios necessários para o funcionamento do sistema."""
        try:
            paths = self.default_config['SYSTEM_SETTINGS']['paths']
            for path in paths.values():
                Path(path).mkdir(parents=True, exist_ok=True)
                
            # Cria diretório de backup
            backup_path = self.default_config['SYSTEM_SETTINGS']['backup']['location']
            Path(backup_path).mkdir(parents=True, exist_ok=True)
            
            self.logger.info("Diretórios do sistema criados com sucesso")
        except Exception as e:
            self.logger.error(f"Erro ao criar diretórios do sistema: {str(e)}")

    def reset_to_default(self, config_section=None):
        """
        Reseta as configurações para o padrão.
        
        Args:
            config_section (str, optional): Seção específica para resetar.
                                          Se None, reseta todas as configurações.
        """
        try:
            if config_section:
                if config_section in self.default_config:
                    current_config = self._get_user_config_path()
                    if current_config.exists():
                        with open(current_config, 'r+', encoding='utf-8') as f:
                            user_config = json.load(f)
                            user_config[config_section] = self.default_config[config_section].copy()
                            f.seek(0)
                            json.dump(user_config, f, indent=4)
                            f.truncate()
                    self.logger.info(f"Configuração {config_section} resetada com sucesso")
            else:
                self._create_default_user_config()
                self.logger.info("Todas as configurações foram resetadas para o padrão")
        except Exception as e:
            self.logger.error(f"Erro ao resetar configurações: {str(e)}")
            raise

    def set_current_user(self, username: str):
        """Define o usuário atual e carrega suas configurações."""
        self.current_user = username
        self._load_user_config()
        self.logger.info(f"Configurações carregadas para o usuário: {username}")

    def _get_user_config_path(self) -> Path:
        """Retorna o caminho do arquivo de configuração do usuário."""
        return self.config_dir / f"{self.current_user}_config.json"

    def _load_user_config(self):
        """Carrega as configurações específicas do usuário."""
        if not self.current_user:
            return

        config_path = self._get_user_config_path()
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                    self._update_config(user_config)
            except Exception as e:
                self.logger.error(f"Erro ao carregar configurações do usuário: {str(e)}")
                self._create_default_user_config()
        else:
            self._create_default_user_config()

    def _create_default_user_config(self):
        """Cria um arquivo de configuração padrão para o usuário."""
        if not self.current_user:
            return

        config_path = self._get_user_config_path()
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.default_config, f, indent=4)
            self.logger.info(f"Criado arquivo de configuração padrão para {self.current_user}")
        except Exception as e:
            self.logger.error(f"Erro ao criar configuração padrão: {str(e)}")

    def _update_config(self, user_config: Dict[str, Any]):
        """Atualiza as configurações com as preferências do usuário."""
        for key in self.default_config:
            if key in user_config:
                self.default_config[key].update(user_config[key])

    def get_config(self, config_type: str) -> Dict[str, Any]:
        """Retorna as configurações específicas do tipo solicitado."""
        return self.default_config.get(config_type, {})

    def update_user_config(self, config_type: str, new_config: Dict[str, Any]):
        """Atualiza as configurações do usuário."""
        if not self.current_user:
            return

        try:
            self.default_config[config_type].update(new_config)
            config_path = self._get_user_config_path()
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.default_config, f, indent=4)
            self.logger.info(f"Configurações atualizadas para {self.current_user}")
        except Exception as e:
            self.logger.error(f"Erro ao atualizar configurações: {str(e)}")

# Configurações do banco de dados (estas permanecem fixas por questões de segurança)
DB_CONFIG = {
    'login_db': 'login.db',
    'marcacao_db': 'db_marcacao.db',
    'queries': {
        'create_users_table': '''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user TEXT NOT NULL,
                password TEXT NOT NULL
            )
        ''',
        'create_patients_table': '''
            CREATE TABLE IF NOT EXISTS patients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                renach TEXT NOT NULL,
                phone TEXT,
                appointment_date TEXT NOT NULL,
                observation TEXT,
                created_by TEXT NOT NULL
            )
        ''',
        # ... outras queries SQL ...
    }
}

# Configurações de logging (estas também permanecem fixas)
LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'standard': {
            'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        },
    },
    'handlers': {
        'file': {
            'class': 'logging.FileHandler',
            'filename': 'app.log',
            'formatter': 'standard'
        },
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'standard'
        }
    },
    'loggers': {
        '': {
            'handlers': ['file', 'console'],
            'level': 'INFO',
            'propagate': True
        }
    }
}

# Instância global do gerenciador de configurações
config_manager = ConfigManager()