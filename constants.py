# constants.py
import os
import sys

def resource_path(relative_path):
    """Определяет путь к файлу как в dev, так и в сборке exe/app"""
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

DB_PATH = resource_path("bridge_defects.db")
TEMPLATE_PATH = resource_path("report_template.docx")