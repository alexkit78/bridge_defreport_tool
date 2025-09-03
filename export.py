# export.py
import shutil
from docx import Document
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from collections import defaultdict
from constants import TEMPLATE_PATH

def export_to_docx(file_path, report_data):
    try:
        shutil.copyfile(TEMPLATE_PATH, file_path)
        doc = Document(file_path)
    except Exception:
        doc = Document()
        table = doc.add_table(rows=1, cols=6)
    else:
        table = None
        expected_headers = ["№", "Местоположение", "Тип", "Описание",
                            "Категории", "Мероприятия"]
        for t in doc.tables:
            first_row_texts = [c.text.strip().lower() for c in t.rows[0].cells]
            if all(any(h.lower() in ct for ct in first_row_texts) for h in expected_headers):
                table = t
                break
        if table is None:
            table = doc.tables[0]

    # очищаем строки после заголовка
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[1]._tr)

    # группировка дефектов
    grouped = defaultdict(list)
    for rec in report_data:
        grouped[rec['placement']].append(rec)

    placements_sorted = sorted(grouped.keys(), key=lambda x: int(x.split('.', 1)[0]))
    counter = 1

    for placement in placements_sorted:
        entries = grouped[placement]

        # заголовок раздела
        placement_row = table.add_row()
        first_cell = placement_row.cells[0]
        first_cell.merge(placement_row.cells[-1])
        clean_placement = placement.split('.', 1)[-1].strip()
        para = first_cell.paragraphs[0]
        para.text = ""
        run = para.add_run(clean_placement)
        run.bold = True
        run.font.size = Pt(12)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # строки дефектов
        for rec in entries:
            row_cells = table.add_row().cells
            row_cells[0].text = str(counter)
            row_cells[1].text = rec.get('location', '')
            row_cells[2].text = rec.get('name', '')
            row_cells[3].text = rec.get('option', '')

            cats = []
            if rec.get('safety'):
                cats.append(f"Б{rec['safety']}")
            if rec.get('durability'):
                cats.append(f"Д{rec['durability']}")
            if rec.get('repairability'):
                cats.append(f"Р{rec['repairability']}")
            try:
                if rec.get('loadcap') and int(rec['loadcap']) == 1:
                    cats.append("Г")
            except Exception:
                pass

            row_cells[4].text = ", ".join(cats)
            row_cells[5].text = rec.get('action', '')

            for cell in row_cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            counter += 1

    doc.save(file_path)