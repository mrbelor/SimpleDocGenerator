"""Генерация Word-документов из шаблона и данных."""

import io
import re
import time
from pathlib import Path

from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
from jinja2 import DebugUndefined, UndefinedError, Environment
import jinja2.exceptions

from .load import traffic_light, get_available_specifiers


class MyDocxTemplate(DocxTemplate):
    """DocxTemplate с поддержкой спецификаторов {{Ключ:addr}}, {{Ключ:time}}, {{:table}}."""

    def render_xml_part(self, src_xml, part, context, jinja_env):
        # {{Ключ:addr}} -> {{Ключ__addr}}, {{:table}} -> {{__table}}
        src_xml = re.sub(r"(\{\{[^}]*?):([^}]*?\}\})", r"\1__\2", src_xml)
        return super().render_xml_part(src_xml, part, context, jinja_env)


def _safe_render(tpl: MyDocxTemplate, context: dict, config_path: str | None = None) -> None:
    """Рендерит шаблон с расширенным контекстом (спецификаторы + raw-значения)."""
    jinja_env = Environment(undefined=DebugUndefined)
    specifiers = get_available_specifiers()

    expanded = dict(context)

    # Добавляем безаргументные спецификаторы: __table, __addr и т.п.
    for spec in specifiers:
        expanded[f"__{spec}"] = traffic_light(None, spec, config_path=config_path)

    # Для каждого ключа добавляем форматированные варианты и raw
    for key, val in context.items():
        if not isinstance(key, str) or key.startswith("__"):
            continue
        for spec in specifiers:
            expanded[f"{key}__{spec}"] = traffic_light(val, spec, config_path=config_path)
        expanded[f"{key}__raw"] = val

    try:
        tpl.render(expanded, jinja_env=jinja_env)
    except UndefinedError as e:
        match = re.search(r"'(.+)' is undefined", str(e))
        name = match.group(1) if match else str(e)
        raise ValueError(
            f'Недопустимые символы в названии столбца: "{name}"\n'
            f"Переименуйте столбец в Excel и в шаблоне (разрешены буквы, цифры, _)."
        ) from e
    except jinja2.exceptions.TemplateSyntaxError as e:
        raise ValueError(
            "Ошибка в шаблоне {{ ... }}: нельзя использовать пробелы.\n"
            "Замените пробелы на '_'. Например: {{Адрес_УК}}"
        ) from e


def generate_documents(
    data: list,
    path_to_template: str,
    output_path: str = "./result.docx",
    config_path: str | None = None,
) -> Path:
    """генерирует документ из шаблона и данных; если шаблон excel — идёт в ExcelGenerator, иначе Word."""
    if not data:
        return None

    if Path(path_to_template).suffix.lower() in (".xls", ".xlsx"):
        from .excel_generator import ExcelGenerator
        return ExcelGenerator().generate(data, path_to_template, output_path)

    output_path = Path(output_path)

    from .table import TableGenerator
    tg = TableGenerator()

    import sys
    table_result = tg.process_table_logic(data, path_to_template, output_path, sys.modules[__name__], config_path=config_path)
    if table_result:
        return table_result

    # Штатная логика: первый документ инициализирует Composer
    tpl = MyDocxTemplate(path_to_template)
    _safe_render(tpl, data[0], config_path=config_path)
    stream = io.BytesIO()
    tpl.save(stream)
    composer = Composer(Document(stream))

    for item in data[1:]:
        tpl = MyDocxTemplate(path_to_template)
        _safe_render(tpl, item, config_path=config_path)
        stream = io.BytesIO()
        tpl.save(stream)
        composer.doc.add_page_break()
        composer.append(Document(stream))

    final_path = _unique_path(output_path)
    composer.save(final_path)
    return final_path


def _unique_path(path: Path) -> Path:
    """Если файл уже существует — добавляет unix timestamp к имени."""
    if path.exists():
        return path.with_stem(f"{path.stem}_{int(time.time())}")
    return path


# Обратная совместимость: старое имя функции
shablon = generate_documents
