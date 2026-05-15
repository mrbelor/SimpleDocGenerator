import pandas as pd
from pathlib import Path
import re
import datetime
import os

try:
	from .address_module import AddressNormaliser
except ImportError:
	from address_module import AddressNormaliser

# компиляция регулярок для времени
range_regex = re.compile(r'^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$')
single_regex = re.compile(r'^(\d{1,2}:\d{2})$')

def _normalize_date_value(val):
	# pd.NaT / NaN -> ""
	try:
		if val is pd.NaT:
			return ""
	except Exception:
		pass
	if isinstance(val, float):
		try:
			import math
			if math.isnan(val):
				return ""
		except Exception:
			pass

	if isinstance(val, pd.Timestamp):
		if pd.isna(val):
			return ""
		py = val.to_pydatetime()
		if py.hour == 0 and py.minute == 0 and py.second == 0 and py.microsecond == 0:
			return py.strftime('%d.%m.%Y')
		return val

	if isinstance(val, datetime.datetime):
		if val.hour == 0 and val.minute == 0 and val.second == 0 and val.microsecond == 0:
			return val.strftime('%d.%m.%Y')
		return val

	if isinstance(val, datetime.time):
		return val

	if isinstance(val, datetime.date):
		return val.strftime('%d.%m.%Y')

	return val

def load_data(file_path):
	path = Path(file_path.strip(" '\"")) 
	if not path.exists():
		raise FileNotFoundError(f"Файл по указанному пути не найден: {path}")

	ext = path.suffix.lower()
	if ext == '.csv':
		df = pd.read_csv(path, sep=None, engine='python')
	elif ext in ('.xls', '.xlsx'):
		try:
			all_sheets = pd.read_excel(path, sheet_name=None)
			df = pd.concat(all_sheets.values(), ignore_index=True)
		except Exception as e:
			df = pd.read_csv(path, sep=None, engine='python')
	else:
		raise ValueError("Формат не поддерживается. Ожидается .csv, .xls или .xlsx")

	records = df.fillna("").to_dict(orient='records')
	return [{k: _normalize_date_value(v) for k, v in row.items()} for row in records]

def _transform_time(val):
	"""Трансформирует одиночное значение времени"""
	if isinstance(val, str):
		if match := range_regex.match(val):
			return f"с {match.group(1)} до {match.group(2)}"
		elif match := single_regex.match(val):
			return f"в {match.group(1)}"
	elif isinstance(val, datetime.time):
		return f"в {val.strftime('%H:%M')}"
	return val

def _transform_address(val, config_path=None, **kwargs):
	"""Трансформирует одиночное значение адреса"""
	if not isinstance(val, str) or not val:
		return val

	if config_path is None:
		raise RuntimeError("КРИТИЧЕСКАЯ ОШИБКА: Путь к address_config.json не задан (config_path=None)!")

	if not os.path.exists(config_path):
		raise RuntimeError(f"КРИТИЧЕСКАЯ ОШИБКА: Файл конфига не найден по пути: {config_path}")

	try:
		normalizer = AddressNormaliser(source=config_path)
	except Exception as e:
		raise RuntimeError(f"Не удалось инициализировать AddressNormaliser: {e}")

	# Добавляем пробелы после точек, чтобы г.Новосибирск не склеивался
	val_spaced = val.replace(".", ". ")
	
	parsed = normalizer.parse(val_spaced)
	if parsed:
		formatted_parts = []
		for label, value in parsed:
			if label in ["?", "неизвестно", ""]:
				formatted_parts.append(value)
			else:
				formatted_parts.append(f"{label} {value}")
		return ", ".join(formatted_parts)
	return val

def transform_time(data, time_key = 'Время'):
	res = []
	for item in data:
		item = item.copy()
		val = item.get(time_key)
		item[time_key] = _transform_time(val)
		res.append(item)
	return res

def transform_address(data, address_key='Адрес', strict=False, config_path=None):
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
					
		if actual_key:
			item[actual_key] = _transform_address(item[actual_key], config_path=config_path)
				
		res.append(item)
	return res

def _format_time(val, **kwargs):
	return _transform_time(_normalize_date_value(val))

def _get_table_gen(val=None, **kwargs):
	try:
		from .table import tableGen
	except (ImportError, ValueError):
		import table
		tableGen = table.tableGen
	return tableGen()

FORMATTERS = {
	'addr': _transform_address,
	'time': _format_time,
	'table': _get_table_gen
}

def get_available_specifiers():
	return list(FORMATTERS.keys())

def traffic_light(value, specifier, config_path=None):
	"""
	Светофор для выбора метода обработки данных на основе спецификатора.
	"""
	formatter = FORMATTERS.get(specifier)
	if formatter:
		print(f"сработал спецификатор {specifier}")
		return formatter(value, config_path=config_path)
	return _normalize_date_value(value)

def test_date_normalization():
	# pd.Timestamp с временем 00:00 -> только дата
	assert _normalize_date_value(pd.Timestamp('2024-05-01')) == '01.05.2024'
	# pd.Timestamp с временем -> без изменений
	ts = pd.Timestamp('2024-05-01 14:30')
	assert _normalize_date_value(ts) is ts
	# datetime.date -> только дата
	assert _normalize_date_value(datetime.date(2024, 5, 1)) == '01.05.2024'
	# pd.NaT -> ""
	assert _normalize_date_value(pd.NaT) == ""
	print("test_date_normalization: OK")

def test_traffic_light():
	# Тест времени
	assert traffic_light("13:00-17:00", "time") == "с 13:00 до 17:00"
	# Тест адреса (хотя бы что возвращает строку)
	addr_res = traffic_light("Москва, Ленина 1", "addr")
	assert isinstance(addr_res, str)
	# Тест таблицы
	from table import tableGen
	assert isinstance(traffic_light(None, "table"), tableGen)
	print("test_traffic_light: OK")

def test_full_load():
	from pprint import pprint as pp
	try:
		data = load_data("./test/data_example.xls")
		if data:
			print(f"Loaded {len(data)} rows")
			# Проверка точечной трансформации через светофор
			sample_val = data[0].get("Время", "")
			print(f"Original time: {sample_val} -> Transformed: {traffic_light(sample_val, 'time')}")
	except Exception as e:
		print(f"Full load test skipped or failed: {e}")

if __name__ == "__main__":
	test_date_normalization()
	test_traffic_light()
