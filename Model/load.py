import pandas as pd
from pathlib import Path
import re
import datetime

# компиляция регулярок для времени
range_regex = re.compile(r'^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$')
single_regex = re.compile(r'^(\d{1,2}:\d{2})$')

import json
import sys
import os

# компиляция регулярок для адреса
def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

street_types_path = get_resource_path(os.path.join("exe_belly", "street_types.json"))
with open(street_types_path, 'r', encoding='utf-8') as f:
    STREET_TYPES = json.load(f)

STREET_LOOKUP = {
    synonym: std_name 
    for std_name, synonyms in STREET_TYPES.items() 
    for synonym in synonyms
}

all_street_synonyms = sorted(STREET_LOOKUP.keys(), key=len, reverse=True)
street_pattern = '|'.join(map(re.escape, all_street_synonyms))

address_city_re = re.compile(r'(?:^|\s|,)(?:г\.|гор\.|г\s+|гор\s+)([А-Яа-яЁёA-Za-z\-]+)', re.IGNORECASE)
address_street_re = re.compile(rf'(?:^|\s|,)({street_pattern})\.?\s+([^,]+)', re.IGNORECASE)
address_house_re = re.compile(r'(?:^|\s|,)(?:д|дом)\.?\s*([0-9A-Za-zА-Яа-яЁё\/\-]+)', re.IGNORECASE)
address_flat_re = re.compile(r'(?:^|\s|,)(?:кв|квартира)\.?\s*([0-9A-Za-zА-Яа-яЁё]+)', re.IGNORECASE)


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

def transform_address(data, street_types=None, address_key='Адрес', strict=False):
	if not street_types:
		street_types = {
			"ул.": ["улица", "ул"],
			"микр-н.": ["микрорайон", "микр-н", "микр", "мкр"],
			"тер.": ["территория", "тер"],
			"пр-кт.": ["проспект", "пр-кт", "пр"],
			"пер.": ["переулок", "пер"],
			"ш.": ["шоссе", "ш"]
		}
		
	STREET_LOOKUP = {
		synonym: std_name 
		for std_name, synonyms in street_types.items() 
		for synonym in synonyms
	}
	all_street_synonyms = sorted(STREET_LOOKUP.keys(), key=len, reverse=True)
	street_pattern = '|'.join(map(re.escape, all_street_synonyms))
	address_street_re = re.compile(rf'(?:^|\s|,)({street_pattern})\.?\s+([^,]+)', re.IGNORECASE)

	res = []
	for item in data:
		item = item.copy() # работаем с копией
		
		# Ищем подходящий ключ
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
			
			city_match = address_city_re.search(val)
			street_match = address_street_re.search(val)
			house_match = address_house_re.search(val)
			flat_match = address_flat_re.search(val)
			
			city = city_match.group(1).strip() if city_match else None
			street_type = ""
			street_name = ""
			if street_match:
				st = street_match.group(1).strip().lower()
				street_type = STREET_LOOKUP.get(st, f"{st}.")
				street_name = street_match.group(2).strip()
			
			house = house_match.group(1).strip() if house_match else None
			flat = flat_match.group(1).strip() if flat_match else None
			
			parts = []
			if city:
				city = city.title() if city.islower() else city
				parts.append(f"г. {city}")
			if street_name:
				parts.append(f"{street_type} {street_name}")
			if house:
				parts.append(f"д. {house}")
			if flat:
				parts.append(f"кв. {flat}")
			
			if parts:
				item[actual_key] = ", ".join(parts)
				
		res.append(item)

	return res

def test():
	from pprint import pprint as pp
	
	data = load_data("./test/data_example.xls")
	data = transform_time(data, "Время") # переводит из "13:00-17:00" в "с 13:00 до 17:00"
	data = transform_address(data, "Адрес") # форматирует адреса
	
	pp(data)
	pp(data[4]['Адрес'] if len(data) > 4 and 'Адрес' in data[4] else "Нет адреса")

if __name__ == "__main__":
	test()
