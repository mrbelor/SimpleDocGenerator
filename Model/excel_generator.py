# excel → excel: два генератора под два формата.
#
# xlsx шаблон → xlsx выход: клонируем строки внутри одного openpyxl-воркбука —
#   dst._style = src._style гарантирует точное совпадение стилей.
#
# xls шаблон → xls выход: xlrd читает с formatting_info=True, xlwt пишет.
#   оба используют BIFF-константы (цвета, границы, выравнивание — те же числа),
#   поэтому перенос стилей бесшовный без каких-либо маппингов.
#
# синтаксис ячеек:
#   новый:  {{НП}}, {{Адрес объекта}}, д. {{Дом}}
#   старый: {{НП}} + ", " + {{Адрес объекта}} + ", д. " + {{Дом}}

import copy as _copy
import re
import time
from pathlib import Path

import xlrd
import openpyxl

_TAG_RE = re.compile(r"\{\{([^{}]+)\}\}")
_OLD_CONCAT_RE = re.compile(r"\}\}\s*\+|\"[^\"]*\"\s*\+\s*\{\{")
_TAG_NUM = ":num"  # {{:num}} → порядковый номер записи


# ---------------------------------------------------------------------------
# роутер
# ---------------------------------------------------------------------------

class ExcelGenerator:
    """выбирает правильный генератор по расширению шаблона"""

    def generate(self, data_list: list, template_path, output_path) -> Path:
        template_path = Path(template_path)
        if template_path.suffix.lower() == ".xls":
            return XlsGenerator().generate(data_list, template_path, output_path)
        return XlsxGenerator().generate(data_list, template_path, output_path)


# ---------------------------------------------------------------------------
# xlsx → xlsx
# ---------------------------------------------------------------------------

class XlsxGenerator:

    def generate(self, data_list: list, template_path, output_path) -> Path:
        return self._from_xlsx(data_list, Path(template_path), output_path)

    def _from_xlsx(self, data_list: list, template_path: Path, output_path) -> Path:
        wb = openpyxl.load_workbook(str(template_path))

        for sheet_name in list(wb.sheetnames):
            ws_src = wb[sheet_name]
            info = _find_template_rows_xlsx(ws_src)
            if info is None:
                del wb[sheet_name]  # листы без тегов в результат не попадают
                continue

            header_count, tmpl_rows = info
            ws_out = wb.create_sheet(f"__tmp__{sheet_name}")

            for rx in range(1, header_count + 1):
                _clone_row_xlsx(ws_src, rx, ws_out, rx)

            out_row = header_count + 1
            for rec_idx, record in enumerate(data_list):
                _eval = lambda s, r, n=rec_idx + 1: _evaluate(s, r, n)
                for tmpl_rx in tmpl_rows:
                    _clone_row_xlsx(ws_src, tmpl_rx, ws_out, out_row, record, _eval)
                    out_row += 1

            for col, dim in ws_src.column_dimensions.items():
                new_dim = _copy.copy(dim)
                new_dim.sheet = ws_out
                ws_out.column_dimensions[col] = new_dim

            position = wb.sheetnames.index(sheet_name)
            del wb[sheet_name]
            ws_out.title = sheet_name
            wb.move_sheet(sheet_name, offset=position - len(wb.sheetnames))

        out_path = _unique_path(Path(output_path).with_suffix(".xlsx"))
        wb.save(str(out_path))
        return out_path


def _find_template_rows_xlsx(ws):
    tmpl_rows = [
        row[0].row for row in ws.iter_rows()
        if any(cell.value and _TAG_RE.search(str(cell.value)) for cell in row)
    ]
    if not tmpl_rows:
        return None
    return tmpl_rows[0] - 1, tmpl_rows


def _clone_row_xlsx(ws_src, src_rx, ws_dst, dst_rx, record=None, evaluate_fn=None):
    for cell in ws_src[src_rx]:
        dst = ws_dst.cell(row=dst_rx, column=cell.column)
        if (record is not None and evaluate_fn is not None
                and isinstance(cell.value, str) and "{{" in cell.value):
            val = evaluate_fn(cell.value, record)
            dst.value = val if val else None
        else:
            dst.value = cell.value
        dst._style = cell._style  # style id валиден — один воркбук

    if src_rx in ws_src.row_dimensions:
        ws_dst.row_dimensions[dst_rx].height = ws_src.row_dimensions[src_rx].height


# ---------------------------------------------------------------------------
# xls → xls
# ---------------------------------------------------------------------------

class XlsGenerator:
    """xls → xls через xlutils.copy: копируем воркбук целиком (XF таблица, палитра,
    все стили), добавляем выходной лист В ТОМ ЖЕ воркбуке и пишем в него ячейки
    через put_cell с теми же xf_idx — никакой ручной трансляции стилей."""

    def generate(self, data_list: list, template_path, output_path) -> Path:
        from xlutils.copy import copy as xl_copy

        template_path = Path(template_path)
        rb = xlrd.open_workbook(str(template_path), formatting_info=True)
        wb = xl_copy(rb)  # xlwt-воркбук со всеми стилями из шаблона

        # патчим add_style: когда приходит _RawXf — возвращаем готовый idx без
        # пересчёта через весь XFStyle machinery (он для нашего объекта всё равно упадёт)
        _orig_add_style = wb.add_style
        def _patched_add_style(style):
            if isinstance(style, _RawXf):
                return style._idx
            return _orig_add_style(style)
        wb.add_style = _patched_add_style

        output_sheets = []

        for si in range(rb.nsheets):
            rs = rb.sheet_by_index(si)
            info = _find_template_rows_xls(rs)
            if info is None:
                continue

            ws_src = wb.get_sheet(si)  # xlwt-копия шаблонного листа
            header_count, tmpl_rows = info

            ws_out = wb.add_sheet(f"__out__{rs.name}")
            output_sheets.append((ws_out, rs.name))

            out_row = 0
            for rx in range(header_count):
                _dup_row(rs, ws_src, rx, ws_out, out_row)
                out_row += 1

            for rec_idx, record in enumerate(data_list):
                for tmpl_rx in tmpl_rows:
                    _dup_row(rs, ws_src, tmpl_rx, ws_out, out_row, record, rec_idx + 1)
                    out_row += 1

            for cx, colinfo in rs.colinfo_map.items():
                ws_out.col(cx).width = colinfo.width

            for (rlo, rhi, clo, chi) in rs.merged_cells:
                if rlo < header_count:
                    ws_out.merge(rlo, rhi - 1, clo, chi - 1)

            tmpl_row_count = len(tmpl_rows)
            tmpl_row_map = {rx: i for i, rx in enumerate(tmpl_rows)}
            for (rlo, rhi, clo, chi) in rs.merged_cells:
                if rlo in tmpl_row_map:
                    local_idx = tmpl_row_map[rlo]
                    row_span = rhi - rlo
                    for rec_idx in range(len(data_list)):
                        offset = header_count + rec_idx * tmpl_row_count + local_idx
                        ws_out.merge(offset, offset + row_span - 1, clo, chi - 1)


        # убираем шаблонные листы, оставляем только выходные
        keep = {id(ws) for ws, _ in output_sheets}
        wb._Workbook__worksheets[:] = [
            ws for ws in wb._Workbook__worksheets if id(ws) in keep
        ]
        for ws_out, name in output_sheets:
            ws_out.name = name

        out_path = _unique_path(Path(output_path).with_suffix(".xls"))
        wb.save(str(out_path))
        return out_path


class _RawXf:
    """минимальный стиль-обёртка вокруг готового xf_idx — чтобы не гонять
    xlwt через весь XFStyle machinery когда индекс уже известен."""
    class _Font:
        height = 200  # нужно только для __adjust_height; высоты ставим сами
    font = _Font()

    def __init__(self, idx: int):
        self._idx = idx

    def xf_index(self, wb):
        return self._idx


def _dup_row(rs_xlrd, ws_xlwt_src, src_rx, ws_out, dst_rx, record=None, row_num=None):
    """дублирует строку: xf_idx берём из xlwt-источника (тот же воркбук),
    значения — из xlrd; подстановку тегов делаем только для строковых ячеек."""
    if src_rx in rs_xlrd.rowinfo_map:
        ri = rs_xlrd.rowinfo_map[src_rx]
        ws_out.row(dst_rx).height = ri.height
        ws_out.row(dst_rx).height_mismatch = True

    src_xlwt_row = ws_xlwt_src._Worksheet__rows.get(src_rx)
    out_row_obj = ws_out.row(dst_rx)

    for cx in range(rs_xlrd.ncols):
        cell = rs_xlrd.cell(src_rx, cx)

        xf_idx = 0
        if src_xlwt_row:
            xlwt_cell = src_xlwt_row._Row__cells.get(cx)
            if xlwt_cell:
                xf_idx = xlwt_cell.xf_idx

        style = _RawXf(xf_idx)

        val = cell.value
        if record is not None and isinstance(val, str) and "{{" in val:
            val = _evaluate(val, record, row_num)

        ctype = cell.ctype
        if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            out_row_obj.set_cell_blank(cx, style)
        elif ctype == xlrd.XL_CELL_TEXT:
            out_row_obj.set_cell_text(cx, val if val is not None else "", style)
        elif ctype in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_DATE):
            out_row_obj.set_cell_number(cx, val, style)
        elif ctype == xlrd.XL_CELL_BOOLEAN:
            out_row_obj.set_cell_boolean(cx, val, style)
        else:
            out_row_obj.set_cell_blank(cx, style)


def _find_template_rows_xls(rs):
    tmpl_rows = [
        rx for rx in range(rs.nrows)
        if any(
            isinstance(rs.cell_value(rx, cx), str) and "{{" in rs.cell_value(rx, cx)
            for cx in range(rs.ncols)
        )
    ]
    if not tmpl_rows:
        return None
    return tmpl_rows[0], tmpl_rows


# ---------------------------------------------------------------------------
# общее
# ---------------------------------------------------------------------------

def _evaluate(template_str: str, record: dict, row_num: int | None = None):
    if not isinstance(template_str, str) or not template_str or "{{" not in template_str:
        return template_str

    def _replace(tag: str) -> str:
        tag = tag.strip()
        if row_num is not None and tag == _TAG_NUM:
            return str(row_num)
        return str(record.get(tag, ""))

    if _OLD_CONCAT_RE.search(template_str):
        parts = re.findall(r'\{\{([^{}]+)\}\}|"([^"]*)"', template_str)
        return "".join(_replace(f) if f else lit for f, lit in parts)
    return _TAG_RE.sub(lambda m: _replace(m.group(1)), template_str)



def _unique_path(path: Path) -> Path:
    if path.exists():
        return path.with_stem(f"{path.stem}_{int(time.time())}")
    return path


