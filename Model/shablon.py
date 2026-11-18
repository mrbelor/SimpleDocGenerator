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
from .sign import Signer, mark_placeholders, mark_text_names, swap_placeholders, text_signs


class MyDocxTemplate(DocxTemplate):
    """DocxTemplate с поддержкой спецификаторов {{Ключ:addr}}, {{Ключ:time}}, {{Ключ:sign}}, {{:table}}."""

    # Чистка XML (patch_xml) зависит только от текста, а шаблон один на все записи:
    # чистим один раз, дальше берём готовое. Результат тот же.
    _patched: dict = {}

    def patch_xml(self, src_xml):
        if src_xml not in self._patched:
            if len(self._patched) > 32:
                self._patched.clear()
            self._patched[src_xml] = super().patch_xml(src_xml)
        return self._patched[src_xml]

    def render_xml_part(self, src_xml, part, context, jinja_env):
        src_xml = mark_placeholders(src_xml)  # заглушка подписи: её {{Ключ:sign}} не для jinja
        src_xml = mark_text_names(src_xml, self, context)  # {{"Имя":sign}} в тексте
        # {{Ключ:addr}} -> {{Ключ__addr}}, {{:table}} -> {{__table}}
        src_xml = re.sub(r"(\{\{[^}]*?):([^}]*?\}\})", r"\1__\2", src_xml)
        return super().render_xml_part(src_xml, part, context, jinja_env)


class _Composer(Composer):
    """Composer без лишних проходов по всему документу на каждой записи.

    Обычный Composer после каждой записи заново нумерует их по всему документу,
    поэтому время растёт квадратично. Каждая такая нумерация полностью затирает
    прошлую, и больше эти номера никто не читает: одна нумерация после последней
    записи даёт тот же файл байт в байт.
    """

    _deferred = True

    def renumber_bookmarks(self):
        if not self._deferred:
            super().renumber_bookmarks()

    def renumber_docpr_ids(self):
        if not self._deferred:
            super().renumber_docpr_ids()

    def renumber_nvpicpr_ids(self):
        if not self._deferred:
            super().renumber_nvpicpr_ids()

    # fix_section_types и fix_header_and_footers начинают с проверки
    # «len(секций всего документа) == 1 or len(секций записи) == 1» и при ней ничего не делают.
    # Секции всего документа считаются проходом по всему документу на каждой записи.
    # Когда в записи одна секция, ответ известен без этого прохода.

    def fix_section_types(self, doc):
        if len(doc.sections) != 1:
            super().fix_section_types(doc)

    def fix_header_and_footers(self, doc):
        if len(doc.sections) != 1:
            super().fix_header_and_footers(doc)

    def renumber_all(self):
        self._deferred = False
        self.renumber_bookmarks()
        self.renumber_docpr_ids()
        self.renumber_nvpicpr_ids()


def _safe_render(tpl: MyDocxTemplate, context: dict, config_path: str | None = None,
                 signer: Signer | None = None, row: int = 1) -> None:
    """Рендерит шаблон с расширенным контекстом (спецификаторы + raw-значения).

    signer - подписи {{Ключ:sign}} (row - номер записи для сообщений об ошибке); без него подписей нет.
    """
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
    tpl.signer, tpl.sign_row = signer, row  # для {{"Имя":sign}} в тексте (render_xml_part)
    if signer:
        expanded.update(text_signs(tpl, context, signer, row))

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
    if signer:
        swap_placeholders(tpl, context, row, signer)


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
    signer = Signer(path_to_template)
    tpl = MyDocxTemplate(path_to_template)
    _safe_render(tpl, data[0], config_path=config_path, signer=signer, row=1)
    stream = io.BytesIO()
    tpl.save(stream)
    composer = _Composer(Document(stream))

    for done, item in enumerate(data[1:], 2):
        tpl = MyDocxTemplate(path_to_template)
        _safe_render(tpl, item, config_path=config_path, signer=signer, row=done)
        composer.doc.add_page_break()
        composer.append(tpl.docx)
        if progress:
            progress(done, len(data))
    if len(data) > 1:  # одна запись: обычный Composer тоже ничего не нумерует
        composer.renumber_all()
    return composer
