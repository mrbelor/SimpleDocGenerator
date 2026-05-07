import re
import json

class AddressNormaliser:

	def __init__(self, source="address_config.json"):
		try:
			with open(source, encoding="utf-8") as f:
				cfg = json.load(f)
			self.aliases = cfg["aliases"]
			self.types   = cfg["types"]
		except FileNotFoundError:
			raise FileNotFoundError(f"конфиг не найден: {source}")
		except KeyError as e:
			raise KeyError(f"в конфиге отсутствует ключ: {e}")

	def _resolve(self, tok):
		return self.aliases.get(tok.lower())

	def _clean(self, s):
		# убираем всё кроме кириллицы, цифр, дефиса и пробелов
		return re.sub(r"[^\u0400-\u04ffа-яёА-ЯЁ0-9\-\s]", "", s).strip()

	def _tokenize(self, s):
		# режем по пробелам и запятым, пустые строки не берём (чтоб нули не удалить)
		return [word for word in re.split(r"[\s,]+", s) if word]

	def parse(self, raw):
		tokens = self._tokenize(self._clean(raw))
		parts, i = [], 0

		while i < len(tokens):
			tok = tokens[i].lower()
			nxt = tokens[i + 1].lower() if i + 1 < len(tokens) else None

			# двухсловный маркер: "ж/д ст", "р п"
			two = f"{tok} {nxt}" if nxt else None
			if two and self._resolve(two):
				canon = self._resolve(two)
				val = tokens[i + 2] if i + 2 < len(tokens) else ""
				parts.append((self.types[canon]["label"], val.lower(), self.types[canon]["level"]))
				i += 3
				continue

			# обычный маркер + значение (возможно многословное)
			if self._resolve(tok):
				canon = self._resolve(tok)
				val = tokens[i + 1] if i + 1 < len(tokens) else ""
				j = i + 2
				while j < len(tokens) and not self._resolve(tokens[j].lower()):
					val += " " + tokens[j]
					j += 1
				parts.append((self.types[canon]["label"], val.strip().lower(), self.types[canon]["level"]))
				i = j
				continue

			# слитный токен без пробела: "д102", "кв3"
			m = re.match(r"^([а-яё\-/]+)\.?(\d.*)$", tok)
			if m and self._resolve(m.group(1)):
				canon = self._resolve(m.group(1))
				parts.append((self.types[canon]["label"], m.group(2), self.types[canon]["level"]))
			else:
				# нераспознанный токен — сохраняем как неизвестно
				parts.append((self.types["?"]["label"], tokens[i].lower(), self.types["?"]["level"]))
			
			i += 1

		parts.sort(key=lambda x: x[2]) # ранжирование по уровню
		return [(label, value) for label, value, level in parts]


def main():
	n = AddressNormaliser()
	with open("test_addresses.txt") as f:
		for line in map(str.strip, f):
			print(n.parse(line))


if __name__ == '__main__':
	main()
