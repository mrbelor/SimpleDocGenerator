import io
import time
from pathlib import Path
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer


def shablon(data, path_to_shablon, output_path="./result.docx"):
	if not data: return
	output_path = Path(output_path)

	# первый док для инициализации
	tpl = DocxTemplate(path_to_shablon)
	tpl.render(data[0])
	stream = io.BytesIO()
	tpl.save(stream)
	composer = Composer(Document(stream))

	# остальные в цикле
	for item in data[1:]:
		tpl = DocxTemplate(path_to_shablon)
		tpl.render(item)
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

	data = load_data("./test/EXPORT.xls")
	# переводит из "13:00-17:00" в "с 13:00 до 17:00"
	data = transform_time(data, "Время")

	path_to_shablon = "./test/notif_shablon.docx"
	
	shablon(data, path_to_shablon)


if __name__ == '__main__':
	test()
