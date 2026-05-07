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

	def _parse_chunk(self, chunk: str):
		tokens = [w for w in re.split(r"\s+", chunk.strip()) if w]
		if not tokens:
			return None
		# 1) Двухсловный маркер в начале: "жд ст Энгельс"
		if len(tokens) >= 2:
			two = f"{tokens[0].lower()} {tokens[1].lower()}"
			if self._resolve(two):
				canon = self._resolve(two)
				val = " ".join(tokens[2:]).lower()
				return (canon, val)
		# 2) Однословный маркер в начале: "ул Архангела Михаила"
		if self._resolve(tokens[0].lower()):
			canon = self._resolve(tokens[0].lower())
			val = " ".join(tokens[1:]).lower()
			return (canon, val)
		# 3) Двухсловный маркер в конце
		if len(tokens) >= 3:
			two = f"{tokens[-2].lower()} {tokens[-1].lower()}"
			if self._resolve(two):
				canon = self._resolve(two)
				val = " ".join(tokens[:-2]).lower()
				return (canon, val)
		# 4) Однословный маркер в конце: "Саратовская область"
		if len(tokens) >= 2 and self._resolve(tokens[-1].lower()):
			canon = self._resolve(tokens[-1].lower())
			val = " ".join(tokens[:-1]).lower()
			return (canon, val)
		# 5) Слитный токен: "д102", "кв3"
		if len(tokens) == 1:
			m = re.match(r"^([а-яё\-/]+)\.?(\d.*)$", tokens[0].lower())
			if m and self._resolve(m.group(1)):
				canon = self._resolve(m.group(1))
				return (canon, m.group(2))
			# 6) Нераспознанный чанк
		return ("?", " ".join(tokens).lower())

	def parse(self, raw):
		chunks = [self._clean(c) for c in raw.split(",") if c.strip()]

		parts = []
		for chunk in chunks:
			result = self._parse_chunk(chunk)
			if result:
				canon, val = result
				info = self.types.get(canon)
				if info:
					parts.append((info["label"], val.strip(), info["level"]))
		parts.sort(key=lambda x: x[2])
		return [(label, value.title()) for label, value, _ in parts]


def main():
	n = AddressNormaliser()
	with open("test_addresses.txt") as f:
		for line in map(str.strip, f):
			print(n.parse(line))


if __name__ == '__main__':
	main()
