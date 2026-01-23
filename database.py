# database.py
import sqlite3
from constants import DB_PATH
from utils import sort_placements

class Database:
    def __init__(self):
        self.conn = sqlite3.connect(DB_PATH)
        self.cursor = self.conn.cursor()

    def get_placements(self):
        """Возвращает список разделов, отсортированных по номеру"""
        self.cursor.execute("SELECT DISTINCT name FROM placements")
        placements = [row[0] for row in self.cursor.fetchall()]
        return sort_placements(placements)

    def get_defects_by_placement(self, placement):
        """Возвращает дефекты для конкретного раздела"""
        self.cursor.execute(
            "SELECT num_ODM, name, option, safetyClass, durabilityClass, "
            "repairabilityClass, loadCapacity, localizationODM, units, qty_rule "
            "FROM defect_types WHERE placement = ?",
            (placement,)
        )
        return self.cursor.fetchall()

    def get_repair_action(self, num_odm: str) -> str:
        self.cursor.execute("""
            SELECT repairAction
            FROM defect_types
            WHERE num_ODM = ?
            LIMIT 1
        """, (num_odm,))
        row = self.cursor.fetchone()
        return row[0] if row and row[0] else ""