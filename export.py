# export.py
import shutil
from collections import defaultdict

from docx import Document
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.shared import Pt
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
from docx.oxml.ns import qn

from constants import TEMPLATE_PATH


# ==================================================
# Word helpers: порядок как в документе
# ==================================================

def iter_block_items(parent):
    """
    Идём по документу по порядку: абзацы и таблицы,
    в том виде, как они реально идут в Word.
    """
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if child.tag.endswith('}p'):
            yield Paragraph(child, parent)
        elif child.tag.endswith('}tbl'):
            yield Table(child, parent)



def find_paragraph_with_marker(doc: Document, marker: str) -> Paragraph | None:
    for p in doc.paragraphs:
        if marker in p.text:
            return p
    return None

def insert_element_after(reference, new_element):
    reference._element.addnext(new_element)

def _is_paragraph_elm(elm) -> bool:
    return elm.tag == qn("w:p")

def _is_table_elm(elm) -> bool:
    return elm.tag == qn("w:tbl")

def find_body_element_index_by_text(doc: Document, marker: str) -> int | None:
    body = doc.element.body
    for i, el in enumerate(list(body)):
        if _is_paragraph_elm(el):
            p = Paragraph(el, doc)
            if marker in p.text:
                return i
    return None

def replace_placeholders_in_body_slice(doc: Document, body_elements: list, mapping: dict):
    """
    Заменяет плейсхолдеры ТОЛЬКО в указанных элементах body (абзацы/таблицы).
    """
    for el in body_elements:
        if _is_paragraph_elm(el):
            replace_in_paragraph(Paragraph(el, doc), mapping)
        elif _is_table_elm(el):
            replace_in_table(Table(el, doc), mapping)

def clone_block_between_markers(
    doc: Document,
    start_marker: str,
    end_marker: str,
    items: list,
    prefix: str
):
    """
    Клонирует блок документа, который находится между start_marker и end_marker
    (start_marker и end_marker — это абзацы с этими текстами).
    
    На каждый item создаётся своя копия блока,
    и плейсхолдеры {{prefix.*}} заменяются только внутри этой копии.
    """
    if not items:
        # если данных нет — просто уберём маркер, чтобы не торчал
        remove_marker_everywhere(doc, start_marker)
        return

    body = doc.element.body
    body_list = list(body)

    start_i = find_body_element_index_by_text(doc, start_marker)
    if start_i is None:
        raise ValueError(f"Не найден маркер {start_marker}")

    end_i = find_body_element_index_by_text(doc, end_marker)
    if end_i is None:
        raise ValueError(f"Не найден маркер {end_marker}")

    if end_i <= start_i:
        raise ValueError(f"Маркер {end_marker} должен быть после {start_marker}")

    # Блок формы = элементы между маркерами:
    # включаем абзац с start_marker (чтобы потом его удалить)
    # но end_marker НЕ включаем
    block = body_list[start_i:end_i]

    # Сначала удалим исходный блок
    # (удалять надо из body, не из body_list)
    for el in block:
        body.remove(el)

    # Теперь вставим N копий на место удаления, в том же порядке
    insert_pos = start_i  # куда вставлять в body

    for item in items:
        cloned_block = [deepcopy(el) for el in block]

        # вставляем в body по порядку
        for el in cloned_block:
            body.insert(insert_pos, el)
            insert_pos += 1

        # заменяем плейсхолдеры ТОЛЬКО в этой копии
        mapping = {
            f"{{{{{prefix}.{k}}}}}": v
            for k, v in item.items()
            if k != "uid"
        }
        replace_placeholders_in_body_slice(doc, cloned_block, mapping)
        replace_placeholders_in_body_slice(doc, cloned_block, {start_marker: ""})

    # В итоге в документе не осталось исходного start_marker, потому что он был внутри block

def find_table_after_marker(doc: Document, marker: str) -> Table | None:
    """
    Находит таблицу, которая идёт СРАЗУ ПОСЛЕ абзаца с marker.
    """
    seen_marker = False
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            if marker in block.text:
                seen_marker = True
        elif isinstance(block, Table):
            if seen_marker:
                return block
    return None


def remove_marker_everywhere(doc: Document, marker: str):
    """
    Удаляет marker, НЕ ломая форматирование.
    Проходим по runs в абзацах и таблицах.
    """
    # основной текст
    for p in doc.paragraphs:
        for run in p.runs:
            if marker in run.text:
                run.text = run.text.replace(marker, "")

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if marker in run.text:
                            run.text = run.text.replace(marker, "")

    # колонтитулы (на всякий)
    for section in doc.sections:
        for p in section.header.paragraphs:
            for run in p.runs:
                if marker in run.text:
                    run.text = run.text.replace(marker, "")
        for t in section.header.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if marker in run.text:
                                run.text = run.text.replace(marker, "")

        for p in section.footer.paragraphs:
            for run in p.runs:
                if marker in run.text:
                    run.text = run.text.replace(marker, "")
        for t in section.footer.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if marker in run.text:
                                run.text = run.text.replace(marker, "")


# ==================================================
# PLACEHOLDERS
# ==================================================

def _copy_run_format(src_run, dst_run):
    """
    Копирует основные свойства форматирования run.
    Этого обычно достаточно, чтобы сохранить bold и размер шрифта.
    """
    dst_run.bold = src_run.bold
    dst_run.italic = src_run.italic
    dst_run.underline = src_run.underline

    # font settings
    dst_run.font.name = src_run.font.name
    dst_run.font.size = src_run.font.size
    dst_run.font.all_caps = src_run.font.all_caps
    dst_run.font.small_caps = src_run.font.small_caps
    dst_run.font.strike = src_run.font.strike
    dst_run.font.double_strike = src_run.font.double_strike
    dst_run.font.subscript = src_run.font.subscript
    dst_run.font.superscript = src_run.font.superscript

    # цвет иногда важен (если задан)
    if src_run.font.color and src_run.font.color.rgb:
        dst_run.font.color.rgb = src_run.font.color.rgb


def replace_in_paragraph(paragraph: Paragraph, mapping: dict):
    if not paragraph.runs:
        return

    full_text = "".join(run.text for run in paragraph.runs)
    new_text = full_text

    has_any_key = any(k in new_text for k in mapping.keys())
    if not has_any_key:
        return

    for key, val in mapping.items():
        new_text = new_text.replace(key, "" if val is None else str(val))

    if new_text == full_text:
        return

    # --- ВАЖНО: если где-то в абзаце была жирность, сохраним её ---
    bold_flag = any((r.bold is True) for r in paragraph.runs)

    # выберем run-образец (для размера/шрифта и т.п.)
    sample_run = paragraph.runs[0]

    # очищаем все runs
    for run in paragraph.runs:
        run.text = ""

    # записываем результат
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        _copy_run_format(sample_run, paragraph.runs[0])

        # ключевое: вернуть жирность, если она была в абзаце
        if bold_flag:
            paragraph.runs[0].bold = True
    else:
        r = paragraph.add_run(new_text)
        _copy_run_format(sample_run, r)
        if bold_flag:
            r.bold = True


def replace_in_table(table: Table, mapping: dict):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_in_paragraph(p, mapping)


def replace_placeholders_everywhere(doc: _Document, mapping: dict):
    """
    Заменяет плейсхолдеры:
    - в тексте
    - в таблицах
    - в колонтитулах
    """
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)

    for t in doc.tables:
        replace_in_table(t, mapping)

    for section in doc.sections:
        header = section.header
        footer = section.footer

        for p in header.paragraphs:
            replace_in_paragraph(p, mapping)
        for t in header.tables:
            replace_in_table(t, mapping)

        for p in footer.paragraphs:
            replace_in_paragraph(p, mapping)
        for t in footer.tables:
            replace_in_table(t, mapping)



# ==================================================
# DEFECTS TABLE (Форма 5)
# ==================================================

def fill_defects_table(doc: Document, defects: list):
    marker = "{{DEFECTS_TABLE}}"
    table = find_table_after_marker(doc, marker)

    if table is None:
        raise ValueError(
            "Не найдена таблица дефектов. "
            "Поставь маркер {{DEFECTS_TABLE}} перед нужной таблицей в шаблоне."
        )

    if len(table.columns) < 6:
        raise ValueError(
            f"Таблица дефектов найдена, но в ней {len(table.columns)} колонок. Нужно 6."
        )

    # очищаем строки после заголовка
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[1]._tr)

    # группировка по разделам
    grouped = defaultdict(list)
    for rec in defects:
        grouped[rec.get("placement", "")].append(rec)

    placements_sorted = sorted(
        [p for p in grouped.keys() if p],
        key=lambda x: int(x.split(".", 1)[0]) if x.split(".", 1)[0].isdigit() else 999
    )

    counter = 1

    for placement in placements_sorted:
        entries = grouped[placement]

        # строка-заголовок раздела
        row = table.add_row()
        cell = row.cells[0]
        cell.merge(row.cells[-1])

        clean_name = placement.split(".", 1)[-1].strip()
        p = cell.paragraphs[0]
        p.text = ""
        run = p.add_run(clean_name)
        run.bold = True
        run.font.size = Pt(12)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # строки дефектов
        for rec in entries:
            cells = table.add_row().cells
            if len(cells) < 6:
                raise ValueError(
                    "При добавлении строки таблицы получилось меньше 6 ячеек. "
                    "Проверь объединения ячеек в шаблоне."
                )

            cells[0].text = str(counter)
            cells[1].text = rec.get("location", "")
            cells[2].text = rec.get("name", "")
            cells[3].text = rec.get("option", "")

            cats = []
            if rec.get("safety"):
                cats.append(f"Б{rec['safety']}")
            if rec.get("durability"):
                cats.append(f"Д{rec['durability']}")
            if rec.get("repairability"):
                cats.append(f"Р{rec['repairability']}")
            try:
                if rec.get("loadcap") and int(rec["loadcap"]) == 1:
                    cats.append("Г")
            except Exception:
                pass

            cells[4].text = ", ".join(cats)
            cells[5].text = rec.get("action", "")

            for c in cells:
                c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            counter += 1

    remove_marker_everywhere(doc, marker)


# ==================================================
# MAIN EXPORT
# ==================================================

def export_to_docx(file_path: str, project: dict):
    """
    Экспорт отчёта в DOCX по шаблону.
    """
    if not isinstance(project, dict):
        raise TypeError("export_to_docx ожидает project=dict")

    shutil.copyfile(TEMPLATE_PATH, file_path)
    doc = Document(file_path)

    # ---------- Форма 1 (Основные сведения bridge.*) ----------
    bridge = project.get("bridge", {})
    mapping = {
        f"{{{{bridge.{k}}}}}": v
        for k, v in bridge.items()
    }
    replace_placeholders_everywhere(doc, mapping)

    # ---------- Форма 2 (Пролётные строения spans.*) ----------
    clone_block_between_markers(
        doc,
        start_marker="{{SPAN_FORM}}",
        end_marker="{{PIER_FORM}}",
        items=project.get("spans", []),
        prefix="span"
    )

# ---------- Форма 3 (Опоры piers.*) ----------
    clone_block_between_markers(
        doc,
        start_marker="{{PIER_FORM}}",
        end_marker="{{FORM4_START}}",
        items=project.get("piers", []),
        prefix="pier"
    )
    # ---------- Форма 4  ----------
    replace_placeholders_everywhere(doc, {"{{FORM4_START}}": ""})
    
    # ---------- Форма 5 (дефекты defects.*) ----------
    fill_defects_table(doc, project.get("defects", []))

    doc.save(file_path)