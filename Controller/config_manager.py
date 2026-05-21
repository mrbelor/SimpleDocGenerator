"""Управление конфигурацией приложения и шаблонами."""

import json
import os
import shutil
import sys
from pathlib import Path

from constants import APP_NAME


class ConfigManager:
    """Отвечает за хранение настроек и управление файлами шаблонов."""

    _DEFAULTS: dict = {
        "appearance_mode": "dark",
        "color_theme": "blue",
        "last_data_folder": str(Path.home()),
        "last_template_folder": str(Path.home()),
        "last_output_folder": str(Path.home() / "Downloads"),
    }

    def __init__(self) -> None:
        self.app_dir = self._get_app_dir(APP_NAME)
        self.templates_dir = self.app_dir / "templates"
        self.config_file = self.app_dir / "config.json"
        self.address_config_file = self.app_dir / "address_config.json"
        self.belly_dir = self.resource_path("exe_belly")

        if not self.app_dir.exists():
            self._initial_setup()
        else:
            if not self.address_config_file.exists():
                self._copy_address_config()
            self.templates_dir.mkdir(parents=True, exist_ok=True)

        self._sync_templates()

        self.defaults = self._load_belly_defaults()
        self.config = self.load()

    @staticmethod
    def resource_path(relative_path: str) -> Path:
        """Абсолютный путь к ресурсу — работает и в dev-режиме, и после PyInstaller."""
        try:
            base_path = sys._MEIPASS  # type: ignore[attr-defined]
        except AttributeError:
            base_path = os.path.abspath(".")
        return Path(base_path) / relative_path

    @staticmethod
    def _get_app_dir(app_name: str) -> Path:
        if sys.platform == "win32":
            base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
        elif sys.platform == "darwin":
            base = Path.home() / "Library" / "Application Support"
        else:
            base = Path.home() / ".config"
        return base / app_name

    def _copy_address_config(self) -> None:
        source = self.belly_dir / "address_config.json"
        if source.exists():
            shutil.copy(source, self.address_config_file)

    def _initial_setup(self) -> None:
        """Первый запуск: создаём папки и копируем дефолтные файлы из belly."""
        self.app_dir.mkdir(parents=True, exist_ok=True)
        self.templates_dir.mkdir(parents=True, exist_ok=True)

        belly_config = self.belly_dir / "config.json"
        if belly_config.exists():
            shutil.copy(belly_config, self.config_file)

        self._copy_address_config()

    def _sync_templates(self) -> None:
        """Добавляет новые шаблоны из belly и обновляет устаревшие."""
        if not self.belly_dir.exists():
            return

        for source in self.belly_dir.glob("**/*.docx"):
            dest = self.templates_dir / source.name
            if not dest.exists() or source.stat().st_mtime > dest.stat().st_mtime:
                shutil.copy2(source, dest)

    def _load_belly_defaults(self) -> dict:
        belly_config = self.belly_dir / "config.json"
        if belly_config.exists():
            try:
                with open(belly_config, encoding="utf-8") as f:
                    return {**self._DEFAULTS, **json.load(f)}
            except (json.JSONDecodeError, OSError):
                pass
        return dict(self._DEFAULTS)

    def load(self) -> dict:
        """Загружает конфиг из AppData, мигрирует старые ключи."""
        if not self.config_file.exists():
            return dict(self.defaults)

        try:
            with open(self.config_file, encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            return dict(self.defaults)

        config = {**self.defaults, **data}

        # Миграция старых ключей
        if "last_folder" in config:
            config.setdefault("last_data_folder", config["last_folder"])
            config.setdefault("last_template_folder", config["last_folder"])
        if "save_folder" in config:
            config.setdefault("last_output_folder", config["save_folder"])

        return config

    def save(self) -> None:
        with open(self.config_file, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def get_templates(self) -> list:
        return [f.name for f in self.templates_dir.glob("*.docx")]

    def add_template(self, source_path: str) -> tuple:
        path = Path(source_path)
        if path.exists() and path.suffix.lower() == ".docx":
            shutil.copy(path, self.templates_dir / path.name)
            return True, path.name
        return False, None

    def remove_template(self, filename: str) -> bool:
        target = self.templates_dir / filename
        if target.exists():
            target.unlink()
            return True
        return False
