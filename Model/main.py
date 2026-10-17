"""Прогон модели без UI - для отдельного тестирования.

python Model/main.py                              - примеры из Model/test, результат в deleteme/
python Model/main.py данные.xlsx шаблон.docx [выход]
"""

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))

from Model import load_data, transform_time, generate_documents

CFG_PATH = HERE / "address_module" / "address_config.json"
OUT_DIR = ROOT / "deleteme"  # в .gitignore

EXAMPLES = [
    (HERE / "test/excel2word/data_example.xls", HERE / "test/excel2word/notif_shablon.docx"),
    (HERE / "test/excel2excel/1/source.xlsx", HERE / "test/excel2excel/1/shblon.xls"),
]


def run(data_path, template_path, output_path=None):
    template_path = Path(template_path)
    if output_path is None:
        OUT_DIR.mkdir(exist_ok=True)
        output_path = OUT_DIR / f"res_{template_path.stem}{template_path.suffix}"

    # то же, что делает контроллер: загрузка + время, адреса только по тегу :addr
    data = transform_time(load_data(data_path))
    result = generate_documents(data, str(template_path), str(output_path), config_path=str(CFG_PATH))
    print(f"{template_path.name}: {len(data)} записей -> {result}")


def main():
    if len(sys.argv) > 1:
        run(*sys.argv[1:4])
        return
    for data_path, template_path in EXAMPLES:
        run(data_path, template_path)


if __name__ == "__main__":
    main()
