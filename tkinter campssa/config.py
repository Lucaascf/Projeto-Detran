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

    def get_config(self, config_type: str) -> Dict[str, Any]:
        """
        Retorna as configurações específicas do tipo solicitado.
        
        Args:
            config_type (str): Tipo de configuração ('APP_CONFIG', 'UI_CONFIG', etc.)
            
        Returns:
            Dict[str, Any]: Dicionário com as configurações solicitadas
        """
        if not hasattr(self, 'default_config'):
            self._load_default_config()
        
        return self.default_config.get(config_type, {})

    def set_current_user(self, username: str):
        """Define o usuário atual e carrega suas configurações."""
        self.current_user = username
        self._load_user_config()
        self.logger.info(f"Configurações carregadas para o usuário: {username}")

    def _load_default_config(self):
        """Carrega as configurações padrão do sistema."""
        self.default_config = {
            "APP_CONFIG": {
                "title": "Gerenciamento de Pacientes - A",
                "version": "1.0.0",
                "background_color": "#2C3E50",
                "initial_geometry": "1200x600",
                "main_geometry": "1200x600",
                "window": {
                    "min_width": 1000,
                    "min_height": 600,
                    "max_width": 1600,
                    "max_height": 1000,
                },
                "file_types": [
                    ("Arquivos Excel", "*.xlsx"),
                    ("Todos os Arquivos", "*.*"),
                ],
            },
            "UI_CONFIG": {
                "colors": {
                "background": "#0D1117",      # Azul muito escuro (quase preto)
                "frame": "#161B22",           # Azul escuro para frames
                "button": "#2D4B6D",          # Azul médio-escuro para botões
                "button_hover": "#1D3557",    # Azul mais escuro para hover
                "text": "#E6EDF3",           # Azul muito claro para texto
                "title": "#FFFFFF",          # Branco puro para títulos
                "border": "#1B2129",         # Borda sutil
            },
                "fonts": {
                    "title": ("Segoe UI", 18, "bold"),
                    "header": ("Segoe UI", 11, "bold"),
                    "button": ("Segoe UI", 10),
                    "normal": ("Segoe UI", 10),
                    "small": ("Segoe UI", 9),
                },
                "styles": {
                    "button": {
                        "relief": "flat",
                        "borderwidth": 0,
                        "highlightthickness": 0,
                        "cursor": "hand2",
                        "width": 20,
                    },
                    "frame": {
                        "relief": "flat",
                        "borderwidth": 1,
                        "padding": 8,
                    },
                },
                "padding": {
                    "default": 8,
                    "small": 4,
                    "large": 15,
                    "button": 6,
                    "section": 8,
                    "title": 20,
                },
                "margins": {
                    "default": 8,
                    "small": 4,
                    "large": 15,
                },
            },
        }

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
                with open(config_path, "r", encoding="utf-8") as f:
                    user_config = json.load(f)
                    self._update_config(user_config)
            except Exception as e:
                self.logger.error(f"Erro ao carregar configurações do usuário: {str(e)}")
                self._create_default_user_config()
        else:
            self._create_default_user_config()

    def _update_config(self, user_config: Dict[str, Any]):
        """Atualiza as configurações com as preferências do usuário."""
        for key in self.default_config:
            if key in user_config:
                self.default_config[key].update(user_config[key])

    def _create_default_user_config(self):
        """Cria um arquivo de configuração padrão para o usuário."""
        if not self.current_user:
            return

        config_path = self._get_user_config_path()
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.default_config, f, indent=4)
            self.logger.info(f"Criado arquivo de configuração padrão para {self.current_user}")
        except Exception as e:
            self.logger.error(f"Erro ao criar configuração padrão: {str(e)}")

# Instância global do gerenciador de configurações
config_manager = ConfigManager()