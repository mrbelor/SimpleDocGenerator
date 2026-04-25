import os
import sys
from pathlib import Path
from Model import load_data, transform_time, shablon
from .config_manager import ConfigManager, NAME

class MainController:
    def __init__(self):
        self.config = ConfigManager()
        self.name = NAME
        self.source_data = None
        self.source_path = None

    def load_source(self, file_path):
        try:
            raw_data = load_data(file_path)
            self.source_data = transform_time(raw_data)
            self.source_path = file_path
            
            # Сохраняем папку данных
            self.config.config["last_data_folder"] = os.path.dirname(file_path)
            self.config.save()
            
            return True, f"Загружено: {os.path.basename(file_path)}"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def get_templates(self):
        return self.config.get_templates()

    def add_template(self, path):
        success, name = self.config.add_template(path)
        if success:
            # Сохраняем папку шаблонов
            self.config.config["last_template_folder"] = os.path.dirname(path)
            self.config.save()
        return success, name

    def remove_template(self, name):
        return self.config.remove_template(name)

    def generate_document(self, template_name, custom_path=None):
        if not self.source_data:
            return False, "Нет данных!"
        if not template_name or template_name == "Шаблон не выбран":
            return False, "Шаблон не выбран!"
        
        template_path = self.config.templates_dir / template_name
        if not template_path.exists():
            return False, "Файл шаблона не найден!"

        try:
            if custom_path:
                final_output_path = custom_path
                # Сохраняем папку сохранения
                self.config.config["last_output_folder"] = os.path.dirname(custom_path)
                self.config.save()
            else:
                save_dir = Path(self.config.config["last_output_folder"])
                if not save_dir.exists():
                    save_dir.mkdir(parents=True, exist_ok=True)
                final_output_path = str(save_dir / f"Результат_{template_name}")

            # Вызов генерации из Model
            final_path = shablon(self.source_data, str(template_path), final_output_path)
            return True, f"Успешно сохранено"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"

    def open_templates_folder(self):
        path = str(self.config.templates_dir)
        if sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
            
    def set_appearance_mode(self, mode):
        self.config.config["appearance_mode"] = mode.lower()
        self.config.save()

    # Вспомогательные методы для получения путей
    def get_last_data_folder(self):
        return self.config.config.get("last_data_folder", str(Path.home()))

    def get_last_template_folder(self):
        return self.config.config.get("last_template_folder", str(Path.home()))

    def get_last_output_folder(self):
        return self.config.config.get("last_output_folder", str(Path.home() / "Downloads"))

    def get_save_folder_name(self):
        folder = self.get_last_output_folder()
        if not folder:
            return "Не выбрана"
        
        # Если это ".", превращаем в абсолютный путь, чтобы вытянуть имя папки
        abs_path = os.path.abspath(folder)
        name = os.path.basename(abs_path)
        
        # На случай если мы в корне диска и basename пустой
        return name if name else abs_path
