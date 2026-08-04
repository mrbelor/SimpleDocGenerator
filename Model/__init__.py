from .load import load_data, get_sheet_names, traffic_light, transform_time, transform_address
from .shablon import generate_documents, shablon  # shablon — обратная совместимость

__all__ = [
    "load_data",
    "get_sheet_names",
    "traffic_light",
    "transform_time",
    "transform_address",
    "generate_documents",
    "shablon",
]
