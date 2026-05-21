"""Общие константы проекта."""

APP_NAME = "DocGenerator"

# Плейсхолдеры для дропдауна шаблонов (не реальные файлы)
TEMPLATE_PLACEHOLDER = "Шаблон не выбран"
TEMPLATE_NOT_FOUND = "Файл не найден"
INVALID_TEMPLATES = frozenset({TEMPLATE_PLACEHOLDER, TEMPLATE_NOT_FOUND})

# Сообщения об ошибках генерации (возвращаются контроллером, проверяются вьюхой)
ERR_NO_DATA = "Нет данных!"
ERR_NO_TEMPLATE = "Шаблон не выбран!"
ERR_TEMPLATE_FILE_MISSING = "Файл шаблона не найден!"
POPUP_ERRORS = frozenset({ERR_NO_DATA, ERR_NO_TEMPLATE, ERR_TEMPLATE_FILE_MISSING})
