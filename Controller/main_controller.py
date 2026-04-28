import os
import sys
import time
from pathlib import Path
from Model import load_data, transform_time, transform_address, shablon
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
            street_types_config = self.config.config.get("street_types")
            self.source_data = transform_address(self.source_data, street_types=street_types_config)
            self.source_path = file_path
            
            # Сохраняем папку данных
            self.config.config["last_data_folder"] = os.path.dirname(file_path)
            self.config.save()
            
            return True, f"Выбрано: {os.path.basename(file_path)}"
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

    def get_suggested_filename(self, template_name):
        template_names = self.config.config.get("template_names", {})
        if template_name in template_names:
            return template_names[template_name]
        
        # Убираем расширение у оригинального шаблона
        base_tmpl = os.path.splitext(template_name)[0]
        return f"Результат_{base_tmpl}"

    def generate_document(self, template_name, custom_path=None):
        if custom_path:
            # Сохраняем папку сохранения в любом случае
            self.config.config["last_output_folder"] = os.path.dirname(custom_path)
            
            # Ассоциативное запоминание имени (без расширения)
            base_name = os.path.splitext(os.path.basename(custom_path))[0]
            if "template_names" not in self.config.config:
                self.config.config["template_names"] = {}
            self.config.config["template_names"][template_name] = base_name
            self.config.save()
            
            final_output_path = custom_path
        else:
            save_dir = Path(self.config.config["last_output_folder"])
            if not save_dir.exists():
                save_dir.mkdir(parents=True, exist_ok=True)
                
            base_name = self.get_suggested_filename(template_name)
            final_output_path = str(save_dir / f"{base_name}.docx")
            
            # Если файл существует, приписываем unix timestamp
            if os.path.exists(final_output_path):
                unix_time = int(time.time())
                final_output_path = str(save_dir / f"{base_name}_{unix_time}.docx")

        if not self.source_data:
            return False, "Нет данных!"
        if not template_name or template_name == "Шаблон не выбран":
            return False, "Шаблон не выбран!"
        
        template_path = self.config.templates_dir / template_name
        if not template_path.exists():
            return False, "Файл шаблона не найден!"

        try:
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
