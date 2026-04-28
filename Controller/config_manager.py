import json
import os
import shutil
import sys
from pathlib import Path

# Глобальное название программы
NAME = "DocGenerator"

class ConfigManager:
    def __init__(self):
        self.app_dir = self._get_app_dir(NAME)
        self.templates_dir = self.app_dir / "templates"
        self.config_file = self.app_dir / "config.json"
        self.belly_dir = self.resource_path("exe_belly")

        # Проверка наличия ВCEЙ папки в AppData
        if not self.app_dir.exists():
            self._initial_setup()
        else:
            # На всякий случай гарантируем наличие подпапки шаблонов
            self.templates_dir.mkdir(parents=True, exist_ok=True)

        self.defaults = self._load_belly_defaults()
        self.config = self.load()

    def resource_path(self, relative_path):
        """ Получает абсолютный путь к ресурсу, работает для обычного запуска и для PyInstaller """
        try:
            # PyInstaller создает временную папку и хранит путь в _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return Path(os.path.join(base_path, relative_path))
    def _get_app_dir(self, app_name):
        if sys.platform == 'win32':
            base = Path(os.environ.get('APPDATA', Path.home() / 'AppData' / 'Roaming'))
        elif sys.platform == 'darwin':
            base = Path.home() / 'Library' / 'Application Support'
        else:
            base = Path.home() / '.config'
        return base / app_name

    def _initial_setup(self):
        """Первый запуск: создание папок и копирование всего из belly"""
        self.app_dir.mkdir(parents=True, exist_ok=True)
        self.templates_dir.mkdir(parents=True, exist_ok=True)

        # Копируем конфиг
        belly_config = self.belly_dir / "config.json"
        if belly_config.exists():
            shutil.copy(belly_config, self.config_file)

        # Копируем шаблоны
        if self.belly_dir.exists():
            for item in self.belly_dir.glob("*.docx"):
                shutil.copy(item, self.templates_dir)

    def _load_belly_defaults(self):
        belly_config = self.belly_dir / "config.json"
        base_defaults = {
            "appearance_mode": "dark", 
            "color_theme": "blue", 
            "last_data_folder": str(Path.home()),
            "last_template_folder": str(Path.home()),
            "last_output_folder": str(Path.home() / "Downloads"),
            "street_types": {
                "ул.": ["улица", "ул"],
                "микр-н.": ["микрорайон", "микр-н", "микр", "мкр"],
                "тер.": ["территория", "тер"],
                "пр-кт.": ["проспект", "пр-кт", "пр"],
                "пер.": ["переулок", "пер"],
                "ш.": ["шоссе", "ш"]
            }
        }
        
        if belly_config.exists():
            try:
                with open(belly_config, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return {**base_defaults, **data}
            except Exception:
                pass
        return base_defaults

    def load(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    config = {**self.defaults, **data}
                    
                    # Миграция со старых ключей если новых нет
                    if "last_folder" in config:
                        if "last_data_folder" not in data:
                            config["last_data_folder"] = config["last_folder"]
                        if "last_template_folder" not in data:
                            config["last_template_folder"] = config["last_folder"]
                    if "save_folder" in config and "last_output_folder" not in data:
                        config["last_output_folder"] = config["save_folder"]
                        
                    return config
            except Exception:
                return self.defaults
        return self.defaults

    def save(self):
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def get_templates(self):
        return [f.name for f in self.templates_dir.glob("*.docx")]

    def add_template(self, source_path):
        path = Path(source_path)
        if path.exists() and path.suffix.lower() == '.docx':
            dest = self.templates_dir / path.name
            shutil.copy(path, dest)
            return True, path.name
        return False, None

    def remove_template(self, filename):
        file_path = self.templates_dir / filename
        if file_path.exists():
            file_path.unlink()
            return True
        return False
