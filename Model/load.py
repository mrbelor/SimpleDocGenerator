"""Загрузка и трансформация данных из Excel/CSV."""

import datetime
import math
import os
import re
from functools import lru_cache
from pathlib import Path

import pandas as pd

from .address_module import AddressNormaliser

# Регулярки для парсинга диапазонов/одиночных значений времени
_RE_TIME_RANGE = re.compile(r"^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$")
_RE_TIME_SINGLE = re.compile(r"^(\d{1,2}:\d{2})$")


# Кэшированный фабричный метод для AddressNormaliser.
# Без кэша при обработке 1000 строк парсится JSON 1000 раз.
@lru_cache(maxsize=4)
def _get_normaliser(config_path):
    return AddressNormaliser(source=config_path)


def _normalize_date_value(val):
    """Приводит pandas-даты к строке 'дд.мм.гггг', остальное не трогает."""
    if val is pd.NaT:
        return ""
    if isinstance(val, float) and math.isnan(val):
        return ""
    if isinstance(val, pd.Timestamp):
        if pd.isna(val):
            return ""
        dt = val.to_pydatetime()
        if dt.hour == dt.minute == dt.second == dt.microsecond == 0:
            return dt.strftime("%d.%m.%Y")
        return val
    if isinstance(val, datetime.datetime):
        if val.hour == val.minute == val.second == val.microsecond == 0:
            return val.strftime("%d.%m.%Y")
        return val
    if isinstance(val, datetime.time):
        return val
    if isinstance(val, datetime.date):
        return val.strftime("%d.%m.%Y")
    if isinstance(val, str):
        try:
            dt = datetime.datetime.fromisoformat(val)
            if dt.hour == dt.minute == dt.second == dt.microsecond == 0:
                return dt.strftime("%d.%m.%Y")
        except ValueError:
            pass
    return val


def load_data(file_path):
    """Загружает Excel или CSV-файл, возвращает список словарей."""
    path = Path(file_path.strip(" '\""))
    if not path.exists():
        raise FileNotFoundError(f"Файл не найден: {path}")

    ext = path.suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(path, sep=None, engine="python")
    elif ext in (".xls", ".xlsx"):
        try:
            all_sheets = pd.read_excel(path, sheet_name=None, dtype=str)
            df = pd.concat(all_sheets.values(), ignore_index=True)
        except Exception:
            df = pd.read_csv(path, sep=None, engine="python")
    else:
        raise ValueError(f"Формат не поддерживается: {ext}. Ожидается .csv, .xls или .xlsx")

    records = df.fillna("").to_dict(orient="records")
    return [{k: _normalize_date_value(v) for k, v in row.items()} for row in records]


def _transform_time_value(val):
    """Преобразует одно значение времени: 'HH:MM-HH:MM' -> 'с HH:MM до HH:MM'."""
    if isinstance(val, str):
        if m := _RE_TIME_RANGE.match(val):
            return f"с {m.group(1)} до {m.group(2)}"
        if m := _RE_TIME_SINGLE.match(val):
            return f"в {m.group(1)}"
    elif isinstance(val, datetime.time):
        return f"в {val.strftime('%H:%M')}"
    return val


def transform_time(data, time_key="Время"):
    """Применяет трансформацию времени ко всем записям."""
    result = []
    for item in data:
        item = item.copy()
        item[time_key] = _transform_time_value(item.get(time_key))
        result.append(item)
    return result


def _transform_address_value(val, config_path=None, **_):
    """Нормализует одно значение адреса через AddressNormaliser."""
    if not isinstance(val, str) or not val:
        return val

    if config_path is None:
        raise RuntimeError("Путь к address_config.json не задан (config_path=None)")
    if not os.path.exists(config_path):
        raise RuntimeError(f"Файл конфига адресов не найден: {config_path}")

    normaliser = _get_normaliser(config_path)

    # Добавляем пробелы после точек, чтобы "г.Новосибирск" не склеивался
    val_spaced = val.replace(".", ". ")
    parsed = normaliser.parse(val_spaced)

    if not parsed:
        return val

    parts = []
    for label, value in parsed:
        if label in ("?", "неизвестно", ""):
            parts.append(value)
        else:
            parts.append(f"{label} {value}")
    return ", ".join(parts)


def transform_address(data, address_key="Адрес", strict=False, config_path=None):
    """Нормализует адреса во всех записях. Поиск колонки нечувствителен к регистру."""
    result = []
    for item in data:
        item = item.copy()
        actual_key = _find_address_key(item, address_key, strict)
        if actual_key:
            item[actual_key] = _transform_address_value(item[actual_key], config_path=config_path)
        result.append(item)
    return result


def _find_address_key(item, address_key, strict):
    if strict:
        return address_key if address_key in item else None
    key_lower = address_key.lower()
    return next((k for k in item if key_lower in str(k).lower()), None)


# Светофор — диспетчер форматтеров для шаблонов ({{Ключ:addr}} и т.п.)
def _format_time(val=None, **_):
    return _transform_time_value(_normalize_date_value(val))


def _get_table_generator(val=None, **_):
    from .table import TableGenerator
    return TableGenerator()


FORMATTERS = {
    "addr": _transform_address_value,
    "time": _format_time,
    "table": _get_table_generator,
}


def get_available_specifiers():
    return list(FORMATTERS.keys())


def traffic_light(value, specifier, config_path=None):
    """Диспетчер: выбирает нужный форматтер по спецификатору из шаблона."""
    formatter = FORMATTERS.get(specifier)
    if formatter:
        return formatter(value, config_path=config_path)
    return _normalize_date_value(value)
