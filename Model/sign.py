"""Подписи монтёров в Word: {{Поле:sign}}.

Подпись берётся из папки <папка шаблонов>/Подписи/<значение поля>/: любая картинка, случайно,
у одного и того же человека без повтора подряд. Два способа отметить место:
- картинка-заглушка с замещающим текстом {{Поле:sign}}: подпись встаёт в её рамку
  (размер, место и обтекание из шаблона, пропорции подписи сохраняются);
- тег {{Поле:sign}} прямо в тексте: подпись встаёт в строку, высотой SIGN_HEIGHT.

Имя в данных и имя папки сравниваются только по буквам и цифрам, без регистра, ё = е,
латинские буквы-двойники = русские: «Александров А.Т.» = папка «Александров А Т».
Инициалы вместо имени, другой порядок слов, опечатка - разные люди: ошибка, а не догадка.
"""

import io
import random
import re
from pathlib import Path

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.shared import Mm
from docxtpl import InlineImage
from PIL import Image

SIGNS_DIR = "Подписи"
SIGN_PLACEHOLDER = "заглушка_подписи.png"  # из exe_belly в папку Подписи: чем отметить место в шаблоне
SIGN_HEIGHT = Mm(12)  # подпись тегом в тексте
_EXT = (".png", ".jpg", ".jpeg")
_MARK = "@@sign@@"  # метка в замещающем тексте заглушки между render и подменой
_ALT_RE = re.compile(r'descr="\{\{\s*([^{}:"]+?)\s*:\s*sign\s*\}\}"')


def mark_placeholders(xml: str) -> str:
    """Замещающий текст заглушки {{Поле:sign}} -> метка, чтобы jinja его не трогал."""
    return _ALT_RE.sub(lambda m: f'descr="{_MARK}{m.group(1)}"', xml)


_TWINS = str.maketrans("abcehkmoptxy", "авсенкмортху")  # латиница, похожая на кириллицу


def _key(name) -> str:
    s = str(name).lower().replace("ё", "е").translate(_TWINS)
    return "".join(ch for ch in s if ch.isalnum())


class Signer:
    """Выбирает подписи на одну генерацию: помнит предыдущую у каждого человека."""

    def __init__(self, template_path):
        self.root = Path(template_path).parent / SIGNS_DIR
        self._dirs = None
        self._last = {}

    def _folders(self) -> dict:
        if self._dirs is None:
            self._dirs = {}
            if self.root.is_dir():
                for d in self.root.iterdir():
                    if d.is_dir():
                        self._dirs.setdefault(_key(d.name), []).append(d)
        return self._dirs

    def pick(self, name, row) -> Path:
        key = _key(name)
        if not key:
            raise ValueError(f"Запись {row}: пустое имя, не знаю, чью подпись ставить")
        found = self._folders().get(key, [])
        if len(found) > 1:
            names = ", ".join(f'"{d.name}"' for d in found)
            raise ValueError(f'Для "{name}" подходит несколько папок с подписями: {names}')
        files = sorted(
            p for p in found[0].iterdir()
            if p.is_file() and not p.name.startswith(".") and p.suffix.lower() in _EXT
        ) if found else []
        if not files:
            raise ValueError(
                f'Нет подписей для "{name}" (запись {row}).\n'
                f"Положите картинки в папку «{SIGNS_DIR}/{str(name).strip()}» в папке шаблонов."
            )
        f = random.choice([p for p in files if p != self._last.get(key)] or files)
        self._last[key] = f
        return f


class _TextSign(InlineImage):
    """{{Поле:sign}} в тексте: подпись выбирается, только если тег реально есть в шаблоне."""

    def __init__(self, tpl, signer, name, row):
        super().__init__(tpl, None, height=SIGN_HEIGHT)
        self.signer, self.name, self.row = signer, name, row

    def _insert_image(self):
        self.image_descriptor = str(self.signer.pick(self.name, self.row))
        return super()._insert_image()


def text_signs(tpl, context: dict, signer: Signer, row: int) -> dict:
    """Ключи Поле__sign для тегов в тексте."""
    return {f"{k}__sign": _TextSign(tpl, signer, v, row)
            for k, v in context.items() if isinstance(k, str) and not k.startswith("__")}


def _story_parts(docx):
    yield docx.part
    for rel in docx.part.rels.values():
        if rel.reltype in (RT.HEADER, RT.FOOTER):
            yield rel.target_part


def _fit(drawing, path: Path) -> None:
    """Вписать подпись в рамку заглушки без искажения, по центру рамки."""
    ext = drawing.find(qn("wp:extent"))
    cx, cy = int(ext.get("cx")), int(ext.get("cy"))
    w, h = Image.open(path).size
    if w * cy > h * cx:  # подпись шире рамки
        ncx, ncy = cx, max(1, cx * h // w)
    else:
        ncx, ncy = max(1, cy * w // h), cy
    for e in (ext, drawing.find(".//" + qn("a:xfrm") + "/" + qn("a:ext"))):
        if e is not None:
            e.set("cx", str(ncx))
            e.set("cy", str(ncy))
    for tag, shift in (("wp:positionH", (cx - ncx) // 2), ("wp:positionV", (cy - ncy) // 2)):
        off = drawing.find(qn(tag) + "/" + qn("wp:posOffset"))
        if off is not None:
            off.text = str(int(off.text) + shift)


def swap_placeholders(tpl, item: dict, row: int, signer: Signer) -> None:
    """После render: каждая заглушка с меткой получает свою подпись."""
    for part in _story_parts(tpl.docx):
        for docpr in list(part.element.iter(qn("wp:docPr"))):
            descr = docpr.get("descr", "")
            if not descr.startswith(_MARK):
                continue
            col = descr[len(_MARK):]
            if col not in item:
                raise ValueError(f'В шаблоне подпись по полю "{col}", а в данных такого столбца нет')
            name = item[col]
            path = signer.pick(name, row)
            drawing = docpr.getparent()
            image = part.package.get_or_add_image_part(io.BytesIO(path.read_bytes()))
            drawing.find(".//" + qn("a:blip")).set(qn("r:embed"), part.relate_to(image, RT.IMAGE))
            _fit(drawing, path)
            for e in (docpr, drawing.find(".//" + qn("pic:cNvPr"))):
                if e is not None and e.get("descr", "").startswith(_MARK):
                    e.set("descr", str(name))
