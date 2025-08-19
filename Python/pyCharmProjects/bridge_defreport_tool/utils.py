# utils.py
import uuid

def generate_uid():
    """Создаёт уникальный идентификатор для дефекта"""
    return str(uuid.uuid4())

def sort_placements(placements):
    """Сортирует список разделов по числовому префиксу"""
    return sorted(placements, key=lambda x: int(x.split('.', 1)[0]))