import json
from datetime import datetime

def save_json(path, report_data):
    data = {
        "saved_at": datetime.now().isoformat(),
        "defects": report_data
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("defects", [])