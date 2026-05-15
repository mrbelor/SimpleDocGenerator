import io
import re
import time
from pathlib import Path
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
from jinja2 import DebugUndefined, UndefinedError, Environment
import jinja2.exceptions

class MyDocxTemplate(DocxTemplate):
    def render_xml_part(self, src_xml, part, context, jinja_env):
        # Поддержка {{Ключ:addr}}, {{Ключ:time}} и {{:table}}
        # Заменяем двоеточия: {{Ключ:addr}} -> {{Ключ__addr}}, но {{:table}} -> {{__table}}
        src_xml = re.sub(r'(\{\{[^}]*?):([^}]*?\}\})', r'\1__\2', src_xml)
        return super().render_xml_part(src_xml, part, context, jinja_env)

def _safe_render(tpl, context, config_path=None):
    jinja_env = Environment(undefined=DebugUndefined)
    
    try:
        from .load import traffic_light, get_available_specifiers
    except ImportError:
        from load import traffic_light, get_available_specifiers
        
    expanded_context = dict(context)
    specifiers = get_available_specifiers()
    
    for spec in specifiers:
        expanded_context[f"__{spec}"] = traffic_light(None, spec, config_path=config_path)

    for k, v in context.items():
        if isinstance(k, str) and not k.startswith("__"):
            for spec in specifiers:
                expanded_context[f"{k}__{spec}"] = traffic_light(v, spec, config_path=config_path)
            expanded_context[f"{k}__raw"] = v
            
    try:
        tpl.render(expanded_context, jinja_env=jinja_env)
    except UndefinedError as e:
        match = re.search(r"'(.+)' is undefined", str(e))
        name = match.group(1) if match else str(e)
        raise ValueError(
            f"Недопустимые символы в названии столбца: \"{name}\"\n"
            f"Переименуйте столбец в Excel и в шаблоне (разрешены буквы, цифры, _)."
        ) from e
    except jinja2.exceptions.TemplateSyntaxError as e:
        raise ValueError(
            f"Ошибка в шаблоне {{{{ ... }}}}: нельзя использовать пробелы.\n"
            f"Замените пробелы на '_'. Например: {{{{Адрес_УК}}}}"
        ) from e


def shablon(data, path_to_shablon, output_path="./result.docx", config_path=None):
	if not data: return
	output_path = Path(output_path)

	# Пытаемся обработать как таблицу (группировка по УК)
	import sys
	try:
		from table import tableGen
	except ImportError:
		from .table import tableGen
		
	tg = tableGen()
	# Передаем текущий модуль (sys.modules[__name__]), чтобы tableGen мог вызвать _safe_render
	table_result = tg.process_table_logic(data, path_to_shablon, output_path, sys.modules[__name__], config_path=config_path)
	if table_result:
		return table_result

	# Если {{:table}} нет в шаблоне, работает штатная логика: первый док для инициализации
	tpl = MyDocxTemplate(path_to_shablon)
	_safe_render(tpl, data[0], config_path=config_path)
	stream = io.BytesIO()
	tpl.save(stream)
	composer = Composer(Document(stream))

	# остальные в цикле
	for item in data[1:]:
		tpl = MyDocxTemplate(path_to_shablon)
		_safe_render(tpl, item, config_path=config_path)
		stream = io.BytesIO()
		tpl.save(stream)

		composer.doc.add_page_break()
		composer.append(Document(stream))

	# предотвращение перезаписи файлов
	if output_path.exists(): # если файл уже существует - создать с временной меткой
		new_stem = f"{output_path.stem}_{int(time.time())}"
		final_path = output_path.with_stem(new_stem)
	else:
		final_path = output_path

	composer.save(final_path)
	print("Done!")
	return final_path


def test():
	from load import load_data, transform_time
	from pprint import pprint as pp

	data = load_data("./test/data_example.xls")
	# переводит из "13:00-17:00" в "с 13:00 до 17:00"
	data = transform_time(data, "Время")

	path_to_shablon = "./test/notif_shablon.docx"
	
	shablon(data, path_to_shablon, config_path="./address_module/address_config.json")


if __name__ == '__main__':
	test()
