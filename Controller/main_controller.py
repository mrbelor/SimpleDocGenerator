"""Контроллер: связывает View с Model и ConfigManager."""

import os
import sys
import traceback
from pathlib import Path

from Model import load_data, get_sheet_names, transform_time, generate_documents
from .config_manager import ConfigManager
from constants import APP_NAME, ERR_NO_DATA, ERR_NO_TEMPLATE, ERR_TEMPLATE_FILE_MISSING, TEMPLATE_PLACEHOLDER


class MainController:
    """Основной контроллер приложения."""

    def __init__(self) -> None:
        self.config = ConfigManager()
        self.name = APP_NAME
        self.source_data: list | None = None
        self.source_path: str | None = None
        self.source_sheet: str | None = None

    # ------------------------------------------------------------------
    # Данные
    # ------------------------------------------------------------------

    def get_sheets(self, file_path: str) -> list:
        """Листы файла данных — для выбора во вьюхе."""
        return get_sheet_names(file_path)

    def load_source(self, file_path: str, sheet: str | None = None) -> tuple:
        """Загружает Excel/CSV, трансформирует время.

        sheet=None — берутся все листы, иначе только указанный.
        Адреса тут не трогаем: нормализация только по явному тегу {{Поле:addr}}.
        """
        try:
            raw_data = load_data(file_path, sheet=sheet)
            data = transform_time(raw_data)

            self.source_data = data
            self.source_path = file_path
            self.source_sheet = sheet
            self.config.config["last_data_folder"] = os.path.dirname(file_path)
            self.config.save()

            name = os.path.basename(file_path)
            return True, f"Выбрано: {name}" if sheet is None else f"Выбрано: {name} → {sheet}"
        except Exception as e:
            traceback.print_exc()
            return False, f"Ошибка: {e}"

    # ------------------------------------------------------------------
    # Шаблоны
    # ------------------------------------------------------------------

    def get_templates(self) -> list:
        return self.config.get_templates()

    def add_template(self, path: str) -> tuple:
        success, name = self.config.add_template(path)
        if success:
            self.config.config["last_template_folder"] = os.path.dirname(path)
            self.config.save()
        return success, name

    def remove_template(self, name: str) -> bool:
        return self.config.remove_template(name)

    def get_suggested_filename(self, template_name: str) -> str:
        saved = self.config.config.get("template_names", {})
        if template_name in saved:
            return saved[template_name]
        base = os.path.splitext(template_name)[0]
        return f"Результат_{base}"

    def get_output_name(self) -> str | None:
        """Имя файла для кнопки «Сформировать» - от источника, без памяти по шаблону."""
        if not self.source_path:
            return None
        return f"{Path(self.source_path).stem}_результат"

    # ------------------------------------------------------------------
    # Генерация
    # ------------------------------------------------------------------

    def generate_document(self, template_name: str, custom_path: str | None = None) -> tuple:
        """Генерирует документ из шаблона и загруженных данных."""
        if not self.source_data:
            return False, ERR_NO_DATA
        if not template_name or template_name == TEMPLATE_PLACEHOLDER:
            return False, ERR_NO_TEMPLATE

        template_path = self.config.templates_dir / template_name
        if not template_path.exists():
            return False, ERR_TEMPLATE_FILE_MISSING

        output_path = self._resolve_output_path(template_name, custom_path)
        self._persist_output_meta(template_name, output_path, custom_path)

        try:
            addr_cfg = str(self.config.address_config_file)
            generate_documents(self.source_data, str(template_path), output_path, config_path=addr_cfg)
            return True, "Успешно сохранено"
        except Exception as e:
            traceback.print_exc()
            return False, f"Ошибка: {e}"

    def _resolve_output_path(self, template_name: str, custom_path: str | None) -> str:
        if custom_path:
            return custom_path

        save_dir = Path(self.config.config["last_output_folder"])
        save_dir.mkdir(parents=True, exist_ok=True)

        base_name = self.get_output_name()
        tmpl_suffix = Path(template_name).suffix.lower()
        out_suffix = tmpl_suffix if tmpl_suffix in (".xls", ".xlsx") else ".docx"

        candidate = save_dir / f"{base_name}{out_suffix}"

        if candidate.exists():
            import time
            candidate = save_dir / f"{base_name}_{int(time.time())}{out_suffix}"

        return str(candidate)

    def _persist_output_meta(self, template_name: str, output_path: str, custom_path: str | None) -> None:
        """Запоминает папку сохранения и ассоциирует имя файла с шаблоном."""
        self.config.config["last_output_folder"] = os.path.dirname(output_path)

        if custom_path:
            base_name = os.path.splitext(os.path.basename(custom_path))[0]
            self.config.config.setdefault("template_names", {})[template_name] = base_name

        self.config.save()

    # ------------------------------------------------------------------
    # UI-хелперы
    # ------------------------------------------------------------------

    def open_templates_folder(self) -> None:
        path = str(self.config.templates_dir)
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')

    def set_appearance_mode(self, mode: str) -> None:
        self.config.config["appearance_mode"] = mode.lower()
        self.config.save()

    def get_last_data_folder(self) -> str:
        return self.config.config.get("last_data_folder", str(Path.home()))

    def get_last_template_folder(self) -> str:
        return self.config.config.get("last_template_folder", str(Path.home()))

    def get_last_output_folder(self) -> str:
        return self.config.config.get("last_output_folder", str(Path.home() / "Downloads"))

    def get_save_folder_name(self) -> str:
        """Возвращает читаемое имя папки сохранения."""
        folder = self.get_last_output_folder()
        if not folder:
            return "Не выбрана"
        abs_path = os.path.abspath(folder)
        name = os.path.basename(abs_path)
        return name if name else abs_path
