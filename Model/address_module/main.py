"""Нормализация адресов по конфигу (aliases + types)."""

import json
import os
import re


class AddressNormaliser:
    """Парсит сырую строку адреса в структурированный список (label, value)."""

    def __init__(self, source=None):
        source = source or "address_config.json"

        if not os.path.exists(source):
            raise FileNotFoundError(f"Конфиг адресов не найден: {source}")

        try:
            with open(source, encoding="utf-8") as f:
                cfg = json.load(f)
            self.aliases = cfg["aliases"]
            self.types = cfg["types"]
        except KeyError as e:
            raise KeyError(f"В конфиге отсутствует ключ: {e}") from e

    def parse(self, raw):
        """Разбивает строку адреса на отсортированные по уровню части."""
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

    def _resolve(self, token):
        return self.aliases.get(token.lower())

    @staticmethod
    def _clean(s):
        return re.sub(r"[^Ѐ-ӿа-яёА-ЯЁa-zA-Z0-9\-/\s]", "", s).strip()

    @staticmethod
    def _tokenize(s):
        return [w for w in re.split(r"[\s,]+", s) if w]

    def _parse_chunk(self, chunk):
        tokens = [w for w in re.split(r"\s+", chunk.strip()) if w]
        if not tokens:
            return None

        # Двухсловный маркер в начале: "жд ст Энгельс"
        if len(tokens) >= 2:
            two = f"{tokens[0].lower()} {tokens[1].lower()}"
            if self._resolve(two):
                return self._resolve(two), " ".join(tokens[2:]).lower()

        # Однословный маркер в начале: "ул Архангела Михаила"
        if self._resolve(tokens[0].lower()):
            return self._resolve(tokens[0].lower()), " ".join(tokens[1:]).lower()

        # Двухсловный маркер в конце
        if len(tokens) >= 3:
            two = f"{tokens[-2].lower()} {tokens[-1].lower()}"
            if self._resolve(two):
                return self._resolve(two), " ".join(tokens[:-2]).lower()

        # Однословный маркер в конце: "Саратовская область"
        if len(tokens) >= 2 and self._resolve(tokens[-1].lower()):
            return self._resolve(tokens[-1].lower()), " ".join(tokens[:-1]).lower()

        # Слитный токен: "д102", "кв3"
        if len(tokens) == 1:
            m = re.match(r"^([а-яё\-/]+)\.?(\d.*)$", tokens[0].lower())
            if m and self._resolve(m.group(1)):
                return self._resolve(m.group(1)), m.group(2)

        # Почтовый индекс без маркера
        joined = "".join(tokens)
        if 4 <= len(joined) <= 10:
            has_enough_digits = len(re.findall(r"\d", joined)) >= 3
            is_index_chars = re.match(r"^[a-zA-Z0-9\-]+$", joined)
            if is_index_chars and has_enough_digits:
                if joined.isdigit() and 5 <= len(joined) <= 6:
                    return "индекс", " ".join(tokens)
                elif not joined.isdigit():
                    return "индекс", " ".join(tokens).upper()

        return "?", " ".join(tokens).lower()
