#project_storage.py
import json
from datetime import datetime

def save_json(path, project: dict):
    data = dict(project)
    data["saved_at"] = datetime.now().isoformat()

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def load_json(path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ValueError("Некорректный формат проекта: ожидался JSON-объект")

    for key in ("bridge", "spans", "piers", "defects"):
        if key not in data:
            raise ValueError(f"Некорректный формат проекта: нет ключа '{key}'")
    return data