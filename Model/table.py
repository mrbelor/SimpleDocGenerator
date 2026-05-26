"""Генерация таблиц Word при наличии тега {{...:table}} в шаблоне.

Формат тега:
  {{ук:счет:адрес+дом+кв+нп:дата:время:table}}

Пример:
  {{УК:Лицевой_счет:Адрес объекта+Дом+Квартира+НП:Дата:Время:table}}

Без аргументов {{:table}} — использует поля по умолчанию.
"""

import datetime
import io
import re
from collections import defaultdict, namedtuple

from docx import Document
from docx.shared import Cm, Pt


_TABLE_MARKER_RE = re.compile(r'\{\{([^{}]*):table\}\}')
_TABLE_COLS = 3
_COL_WIDTHS = [Cm(1.0), Cm(3.0), Cm(14.5)]
_HEADER_LABELS = ["№", "Лицевой счет абонента", "Полный адрес (город, улица, дом, квартира)"]
_TABLE_FONT = "Times New Roman"

_ACCOUNT_KEYS = ("Лицевой_счет", "Лицевой счет", "Лицевой счет абонента")

_TableSpec = namedtuple("_TableSpec", ["uk", "account", "addr", "house", "apt", "np_field", "date", "time"])
_TableSpec.__new__.__defaults__ = ("УК", None, "Адрес объекта", "Дом", "Квартира", "НП", "Дата", "Время")


class TableGenerator:
    """Обрабатывает шаблоны с тегом {{...:table}}: группирует данные и строит таблицу."""

    def process_table_logic(self, data_list: list, path_to_template, output_path, template_module, config_path=None):
        """
        Точка входа из shablon.py.
        Возвращает Path к файлу если шаблон содержит {{...:table}}, иначе False.
        """
        spec = self._load_marker_spec(path_to_template)
        if spec is None:
            return False

        grouped = self._make_grouped_data(data_list, spec)

        from docxcompose.composer import Composer
        composer = None

        for group_item in grouped:
            out_stream = self._render_group(group_item, path_to_template, template_module, config_path, spec)

            if composer is None:
                composer = Composer(Document(out_stream))
            else:
                composer.doc.add_page_break()
                composer.append(Document(out_stream))

        final_path = self._resolve_final_path(output_path)
        composer.save(final_path)
        return final_path

    # ------------------------------------------------------------------
    # Группировка данных по УК
    # ------------------------------------------------------------------

    def _make_grouped_data(self, data_list: list, spec) -> list:
        """Группирует строки по полю УК, вычисляет диапазон дат для каждой группы."""
        groups: dict = defaultdict(list)
        for item in data_list:
            groups[str(item.get(spec.uk, "Неизвестная УК"))].append(item)

        result = []
        for uk, items in groups.items():
            date_label = self._build_date_label(items, spec.date)
            base_item = items[0].copy()
            base_item[spec.date] = date_label
            base_item["__table_records__"] = items
            result.append(base_item)
        return result

    def _build_date_label(self, items: list, date_field: str) -> str:
        dates = [self._parse_date(i.get(date_field, "")) for i in items]
        dates = [d for d in dates if d]
        if not dates:
            return str(items[0].get(date_field, ""))
        min_d, max_d = min(dates), max(dates)
        if min_d == max_d:
            return min_d.strftime("%d.%m.%Y")
        return f"с {min_d.strftime('%d.%m.%Y')} по {max_d.strftime('%d.%m.%Y')}"

    @staticmethod
    def _parse_date(date_str: str) -> datetime.datetime | None:
        try:
            return datetime.datetime.strptime(str(date_str).strip(), "%d.%m.%Y")
        except (ValueError, TypeError):
            return None

    # ------------------------------------------------------------------
    # Рендер одной группы
    # ------------------------------------------------------------------

    def _render_group(self, group_item: dict, path_to_template, template_module, config_path, spec) -> io.BytesIO:
        doc = Document(path_to_template)
        marker_paragraph = self._find_marker_paragraph(doc)

        if marker_paragraph:
            marker_paragraph.text = _TABLE_MARKER_RE.sub("", marker_paragraph.text)
            table = self._build_table(doc, group_item["__table_records__"], config_path, spec)
            marker_paragraph._p.addnext(table._tbl)

        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)

        tpl = template_module.MyDocxTemplate(stream)
        template_module._safe_render(tpl, group_item, config_path=config_path)

        out = io.BytesIO()
        tpl.save(out)
        out.seek(0)
        return out

    # ------------------------------------------------------------------
    # Построение таблицы
    # ------------------------------------------------------------------

    def _build_table(self, doc: Document, records: list, config_path, spec) -> object:
        table = doc.add_table(rows=1, cols=_TABLE_COLS)
        table.style = "Table Grid"
        table.autofit = False
        table.allow_autofit = False

        for j, col in enumerate(table.columns):
            col.width = _COL_WIDTHS[j]

        self._fill_header(table, _COL_WIDTHS)

        for i, item in enumerate(records):
            self._add_data_row(table, i, item, config_path, _COL_WIDTHS, spec)

        self._apply_cell_style(table)
        return table

    @staticmethod
    def _fill_header(table, col_widths: list) -> None:
        header_cells = table.rows[0].cells
        for j, (cell, label) in enumerate(zip(header_cells, _HEADER_LABELS)):
            cell.text = label
            cell.width = col_widths[j]

    def _add_data_row(self, table, row_index: int, item: dict, config_path, col_widths: list, spec) -> None:
        row_cells = table.add_row().cells

        for j, cell in enumerate(row_cells):
            cell.width = col_widths[j]

        row_cells[0].text = str(row_index + 1)
        row_cells[1].text = str(self._get_account_number(item, spec.account))
        row_cells[2].text = self._build_address_cell_text(item, config_path, spec)

    @staticmethod
    def _get_account_number(item: dict, account_field=None) -> str:
        if account_field is not None:
            return str(item.get(account_field, ""))
        for key in _ACCOUNT_KEYS:
            if key in item:
                return str(item[key])
        return ""

    def _build_address_cell_text(self, item: dict, config_path, spec) -> str:
        address = self._normalise_address(item, config_path, spec)
        date_val = str(item.get(spec.date, ""))
        time_val = self._extract_time(item, spec.time)

        cell_text = address
        if time_val:
            cell_text += f"\n{date_val} в период {time_val} часов"
        else:
            cell_text += f"\n{date_val}"
        return cell_text

    def _normalise_address(self, item: dict, config_path, spec) -> str:
        """Собирает и нормализует адрес из полей записи."""
        from .load import _transform_address_value

        addr_raw = str(item.get(spec.addr, "")).strip()
        addr_formatted = self._capitalise_address(addr_raw)

        is_full = "," in addr_formatted or " д." in addr_formatted.lower() or " кв." in addr_formatted.lower()

        if is_full:
            addr_str = self._append_house_and_apt(addr_formatted, item, spec)
        else:
            addr_str = self._build_address_from_parts(addr_formatted, item, spec)

        normalised = _transform_address_value(addr_str, config_path=config_path)
        return normalised if normalised else addr_str

    @staticmethod
    def _capitalise_address(addr_raw: str) -> str:
        """Делает первую букву каждого слова заглавной, сохраняет маркеры типа 'д.', 'ул.'."""
        low_markers = {"д.", "кв.", "ул.", "г.", "с.", "п.", "рп."}
        words = []
        for word in addr_raw.split():
            if word.lower() in low_markers:
                words.append(word.lower())
            elif "-" in word:
                words.append("-".join(part.capitalize() for part in word.split("-")))
            else:
                words.append(word.capitalize())
        return " ".join(words)

    @staticmethod
    def _is_valid(val) -> bool:
        if val is None:
            return False
        return bool(str(val).strip() and str(val).strip().lower() not in ("nan", "none"))

    def _append_house_and_apt(self, addr_str: str, item: dict, spec) -> str:
        house = item.get(spec.house)
        apt = item.get(spec.apt)
        if self._is_valid(house) and f"д. {house}" not in addr_str:
            addr_str += f", д. {self._clean_float(str(house))}"
        if self._is_valid(apt) and f"кв. {apt}" not in addr_str:
            addr_str += f", кв. {self._clean_float(str(apt))}"
        return addr_str

    def _build_address_from_parts(self, addr_formatted: str, item: dict, spec) -> str:
        clean_street = re.sub(r"(?i)\bул(?:ица)?\.?", "", addr_formatted).strip()
        addr_str = f"ул. {clean_street}" if clean_street else ""

        house_num = item.get(spec.house)
        apt_num = item.get(spec.apt)

        if self._is_valid(house_num):
            addr_str += f", д. {self._clean_float(str(house_num))}"
        if self._is_valid(apt_num):
            addr_str += f", кв. {self._clean_float(str(apt_num))}"

        settlement = self._build_settlement_prefix(item, spec.np_field)
        if settlement:
            addr_str = f"{settlement}, {addr_str}" if addr_str else settlement

        return addr_str

    @staticmethod
    def _build_settlement_prefix(item: dict, np_field: str) -> str:
        """Строит префикс населённого пункта: 'г. Саратов', 'с. Ивановка' и т.п."""
        np_val = str(item.get(np_field, "")).strip()
        if not np_val or np_val.lower() in ("nan", "none", ""):
            return ""

        prefix_map = {
            frozenset({"с", "село"}): "с.",
            frozenset({"п", "пос", "поселок", "посёлок"}): "п.",
            frozenset({"рп", "рабпос"}): "рп.",
            frozenset({"г", "город"}): "г.",
            frozenset({"нп", "нп."}): "нп.",
        }
        prefix = "г."
        clean_words = []

        for word in np_val.split():
            word_lower = word.lower().strip(".")
            matched = next((p for keys, p in prefix_map.items() if word_lower in keys), None)
            if matched:
                prefix = matched
            else:
                clean_words.append(word.capitalize())

        return f"{prefix} {' '.join(clean_words)}" if clean_words else ""

    @staticmethod
    def _clean_float(val: str) -> str:
        """Убирает '.0' у числел, которые pandas читает как float: '42.0' -> '42'."""
        return val[:-2] if val.endswith(".0") else val

    @staticmethod
    def _extract_time(item: dict, time_field: str) -> str:
        """Вытаскивает значение времени из записи, включая колонки Unnamed."""
        time_val = str(item.get(time_field, ""))
        if not time_val:
            for k, v in item.items():
                if isinstance(k, str) and "Unnamed" in k and re.match(r"^\d{2}:\d{2}-\d{2}:\d{2}$", str(v)):
                    return str(v)
        return time_val

    # ------------------------------------------------------------------
    # Вспомогательные
    # ------------------------------------------------------------------

    @staticmethod
    def _load_marker_spec(path_to_template):
        """Открывает шаблон, ищет маркер и возвращает _TableSpec. None если маркера нет."""
        doc = Document(path_to_template)
        para = next((p for p in doc.paragraphs if _TABLE_MARKER_RE.search(p.text)), None)
        if para is None:
            return None
        return TableGenerator._parse_marker_spec(para.text)

    @staticmethod
    def _find_marker_paragraph(doc: Document):
        return next((p for p in doc.paragraphs if _TABLE_MARKER_RE.search(p.text)), None)

    @staticmethod
    def _parse_marker_spec(text: str):
        """Разбирает тег → _TableSpec.

        Формат: {{ук:счет:адрес+дом+кв+нп:дата:время:table}}
        """
        m = _TABLE_MARKER_RE.search(text)
        if not m or not m.group(1):
            return _TableSpec()
        parts = [p.strip() for p in m.group(1).split(":") if p.strip()]
        addr_parts = [f.strip() for f in parts[2].split("+")] if len(parts) > 2 else []
        return _TableSpec(
            uk=parts[0] if len(parts) > 0 else "УК",
            account=parts[1] if len(parts) > 1 else None,
            addr=addr_parts[0] if len(addr_parts) > 0 else "Адрес объекта",
            house=addr_parts[1] if len(addr_parts) > 1 else "Дом",
            apt=addr_parts[2] if len(addr_parts) > 2 else "Квартира",
            np_field=addr_parts[3] if len(addr_parts) > 3 else "НП",
            date=parts[3] if len(parts) > 3 else "Дата",
            time=parts[4] if len(parts) > 4 else "Время",
        )

    @staticmethod
    def _apply_cell_style(table) -> None:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.space_before = Pt(0)
                    for run in paragraph.runs:
                        run.font.name = _TABLE_FONT

    @staticmethod
    def _resolve_final_path(output_path) -> object:
        import time
        from pathlib import Path
        output_path = Path(output_path)
        if output_path.exists():
            return output_path.with_stem(f"{output_path.stem}_{int(time.time())}")
        return output_path
