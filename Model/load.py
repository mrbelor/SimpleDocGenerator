import pandas as pd
from pathlib import Path
import re
import datetime
from .address_module import AddressNormaliser

# компиляция регулярок для времени
range_regex = re.compile(r'^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$')
single_regex = re.compile(r'^(\d{1,2}:\d{2})$')

def load_data(file_path):
	path = Path(file_path.strip(" '\"")) # НА ВСЯКИЙ убираем по краям пути мусор

	if not path.exists():
		raise FileNotFoundError(f"Файл по указанному пути не найден: {path}")

	ext = path.suffix.lower() # расширение файла
	# загрузка файлов (без доверия к расширениям файлов)
	if ext == '.csv':
		df = pd.read_csv(path, sep=None, engine='python')
	elif ext in ('.xls', '.xlsx'):
		try:
			# Читаем все листы в словарь DataFrames
			all_sheets = pd.read_excel(path, sheet_name=None)
			# Конкатенируем все листы в один DataFrame
			df = pd.concat(all_sheets.values(), ignore_index=True)
		except Exception as e:
			# Если не удалось прочитать как Excel, пробуем как CSV
			# (некоторые системы выгружают CSV с расширением .xls)
			print(f"Не удалось прочитать {path.name} как Excel: {e}. Попытка в csv...")
			df = pd.read_csv(path, sep=None, engine='python')
	else:
		raise ValueError("Формат не поддерживается. Ожидается .csv, .xls или .xlsx")

	# NaN заменяем на пустые строки и конвертируем в список словарей
	return df.fillna("").to_dict(orient='records')

def transform_time(data, time_key = 'Время'):
	res = []
	for item in data:
		item = item.copy() # работаем с копией, чтобы не мутировать исходные данные
		val = item.get(time_key) # значение по ключу типа 'Время': '09:00-13:00'

		if isinstance(val, str):
			# диапазон (например, "13:00-17:00" или "13:00 - 17:00")
			if match := range_regex.match(val):
				item[time_key] = f"с {match.group(1)} до {match.group(2)}"
			# одиночное время (например, "17:00")
			elif match := single_regex.match(val):
				item[time_key] = f"в {match.group(1)}"

		# обработка datetime (если в excel это было время, а pandas автоматом преобразовал его в datetime)
		elif isinstance(val, datetime.time):
			item[time_key] = f"в {val.strftime('%H:%M')}"

		res.append(item)

	return res

def transform_address(data, address_key='Адрес', strict=False, config_path=None):
	# Инициализируем нормализатор. Если путь не передан, он сам попробует найти дефолтный.
	try:
		normalizer = AddressNormaliser(source=config_path)
	except Exception as e:
		print(f"Предупреждение: не удалось инициализировать AddressNormaliser: {e}")
		return data

	res = []
	for item in data:
		item = item.copy()
		
		actual_key = None
		if strict:
			if address_key in item:
				actual_key = address_key
		else:
			for k in item.keys():
				if address_key.lower() in str(k).lower():
					actual_key = k
					break
					
		if actual_key and isinstance(item.get(actual_key), str):
			val = item[actual_key]
			parsed = normalizer.parse(val)
			
			if parsed:
				# Формируем строку: "лейбл значение". Если лейбл "неизвестно", выводим только значение.
				formatted_parts = []
				for label, value in parsed:
					if label == "неизвестно":
						formatted_parts.append(value)
					else:
						formatted_parts.append(f"{label} {value}")
				
				item[actual_key] = ", ".join(formatted_parts)
				
		res.append(item)

	return res

def test():
	from pprint import pprint as pp
	
	data = load_data("./test/data_example.xls")
	data = transform_time(data, "Время")
	data = transform_address(data, address_key="Адрес")
	
	pp(data)

if __name__ == "__main__":
	test()
