"""Генерация Word-документов из шаблона и данных."""

import io
import os
import re
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
    progress=None,
) -> Path:
    """генерирует документ из шаблона и данных; если шаблон excel - идёт в ExcelGenerator, иначе Word.

    progress(готово, всего) зовётся после каждой записи (в :table - после каждой группы)
    из того потока, где идёт генерация. готово == всего: осталось только записать файл.
    """
    if not data:
        return None

    output_path = Path(output_path)
    suffix = Path(path_to_template).suffix.lower()
    if suffix in (".xls", ".xlsx"):
        from .excel_generator import ExcelGenerator
        doc = ExcelGenerator().generate(data, path_to_template, config_path, progress)
        output_path = output_path.with_suffix(suffix)  # тихой конвертации форматов нет
    else:
        doc = _compose_word(data, path_to_template, config_path, progress)

    # свободное имя подбирает контроллер; здесь пишем ровно туда, куда сказали.
    # Пишем в <имя>.part и переименовываем одним ходом: прерванная запись
    # (закрыли окно, упала программа) не оставит недописанный файл с настоящим именем
    part = output_path.with_name(f"{output_path.name}.part")
    doc.save(str(part))
    os.replace(part, output_path)
    return output_path


def _compose_word(data: list, path_to_template: str, config_path: str | None, progress) -> Composer:
    """Все записи в один Word-документ: ветка {{:table}} или по документу на запись."""
    from .table import TableGenerator
    import sys
    table_doc = TableGenerator().process_table_logic(data, path_to_template, sys.modules[__name__], config_path=config_path, progress=progress)
    if table_doc:
        return table_doc

    # Штатная логика: первый документ инициализирует Composer
    tpl = MyDocxTemplate(path_to_template)
    _safe_render(tpl, data[0], config_path=config_path)
    stream = io.BytesIO()
    tpl.save(stream)
    composer = Composer(Document(stream))

    for done, item in enumerate(data[1:], 2):
        tpl = MyDocxTemplate(path_to_template)
        _safe_render(tpl, item, config_path=config_path)
        stream = io.BytesIO()
        tpl.save(stream)
        composer.doc.add_page_break()
        composer.append(Document(stream))
        if progress:
            progress(done, len(data))
    return composer
