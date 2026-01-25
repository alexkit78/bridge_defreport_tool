# ui.py
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Menu
import ctypes

from database import Database
from export import export_to_docx, export_report_to_docx
from utils import generate_uid
from project_storage import save_json, load_json
from project_model import make_empty_project
from tabs.tab_general import GeneralTabMixin
from tabs.tab_spans import SpansTabMixin
from tabs.tab_piers import PiersTabMixin
from tabs.tab_defects import DefectsTabMixin
from tabs.tab_photos import PhotosTabMixin


class DefectApp(GeneralTabMixin, SpansTabMixin, PiersTabMixin, DefectsTabMixin,PhotosTabMixin):
    def __init__(self, root):
        self.root = root
        self.is_loading = True

        self.root.title("Оценка состояния мостового сооружения")
        self.build_menu()

        # инициализация БД
        self.db = Database()

        # список дефектов, которые пойдут в отчёт
        self.project = make_empty_project()
        self.defect_options_by_numodm = {}
        self.defect_numodm_map = {}
        self.defect_categories_by_option = {}

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        self.tab_general = ttk.Frame(self.notebook)
        self.tab_spans = ttk.Frame(self.notebook)
        self.tab_piers = ttk.Frame(self.notebook)
        self.tab_defects = ttk.Frame(self.notebook)
        self.tab_photos = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_general, text="Общие сведения")
        self.notebook.add(self.tab_spans, text="Пролётные строения")
        self.notebook.add(self.tab_piers, text="Опоры")
        self.notebook.add(self.tab_defects, text="Дефекты")
        self.notebook.add(self.tab_photos, text="Фотографии")


        self.bridge_vars = {}  # ключ поля -> tk.StringVar
        self.span_forms = {}  # uid -> {"frame": Frame, "vars": {key: StringVar}}
        self.pier_forms = {} # uid -> {"frame": Frame, "vars": {key: StringVar}}

        self.build_ui()
        self.is_loading = False
        self.is_dirty = False
        

    def _generate_uid(self):
        return generate_uid()

    def build_menu(self):
        menubar: Menu = tk.Menu(self.root)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Новое сооружение", 
                              command=self.new_project)
        file_menu.add_command(label="Открыть...", command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label="Сохранить паспорт...", command=self.save_project)
        file_menu.add_command(label="Сохранить отчёт...", command=self.save_report)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.on_close)
        
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.root.config(menu=menubar)

    def save_project(self):
        is_empty = (not self.project["bridge"]) and (not self.project[
            "defects"]) and (not self.project["spans"]) and (not
            self.project["piers"])
        if is_empty:
            messagebox.showwarning("Нет данных", "Проект пустой — нечего сохранять.")
            return False

        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx"),
                       ("JSON project", "*.json")]
        )
        if not file_path:
            return False

        try:
            if file_path.lower().endswith(".json"):
                save_json(file_path, self.project)
            else:
                export_to_docx(file_path, self.project)

            self.is_dirty = False
            self.status_label.config(text=f"Файл сохранён: {file_path}")
            self.root.after(2000, lambda: self.status_label.config(text=""))
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить паспорт:\n{e}")
            return False
        
    def save_report(self):
        is_empty = (not self.project["bridge"]) and (not self.project["defects"]) and (not self.project["spans"]) and (not self.project["piers"])
        if is_empty:
            messagebox.showwarning("Нет данных", "Проект пустой — нечего сохранять.")
            return False

        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")]
        )
        if not file_path:
            return False

        try:
            export_report_to_docx(file_path, self.project)
            self.is_dirty = False
            self.status_label.config(text=f"Отчёт сохранён: {file_path}")
            self.root.after(2000, lambda: self.status_label.config(text=""))
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить отчёт:\n{e}")
            return False


    def load_project(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON project", "*.json")]
        )
        if not file_path:
            return

        try:
            project = load_json(file_path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
            return

        self.project = project #Загрузка всего проекта
        self.rebuild_span_tabs()
        self.rebuild_pier_tabs()
        self.refresh_general_tab_from_project()
        self.table.delete(*self.table.get_children())

        for rec in self.project["defects"]:
            uid = rec.get("uid")
            if not uid:
                continue

            
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
            except:
                pass

            self.table.insert(
                "",
                "end",
                iid=uid,
                values=(
                    rec.get("placement", ""),
                    rec.get("location", ""),
                    rec.get("name", ""),
                    rec.get("option", ""),
                    ", ".join(cats),
                    rec.get("action", "")
                )
            )

        self.update_status_bar()
        self.refresh_photos_tab_from_project()
        self.status_label.config(text="Данные загружены")
        self.root.after(2000, lambda: self.status_label.config(text=""))
        self.is_dirty = False

    def new_project(self):

        if (
                self.project.get("bridge") or
                self.project.get("defects") or
                self.project.get("spans") or
                self.project.get("piers")
        ):
            if not messagebox.askyesno(
                "Новый проект",
                "Текущие данные будут потеряны. Продолжить?"
            ):
                return
        # сбрасываем проект
        self.project = make_empty_project()
        self.rebuild_span_tabs()
        self.rebuild_pier_tabs()
        # очищаем таблицу дефектов
        self.table.delete(*self.table.get_children())
        self.update_status_bar()

        # обновить вкладку Ф1 из проекта (он теперь пустой)
        self.refresh_general_tab_from_project()

        self.is_dirty = False

    def build_ui(self):

        self.build_tab_defects()
        self.build_tab_general()
        self.build_tab_spans()
        self.build_tab_piers()
        self.build_tab_photos()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)


    def on_close(self):
        if self.is_dirty:
            answer = messagebox.askyesnocancel(
                "Несохранённые данные",
                "Данные были изменены, но не сохранены.\n\n"
                "Cохранить перед выходом?"
            )

            if answer is None:
                # Cancel
                return
            if answer:
                # Да → сохранить
                saved = self.save_project()
                if not saved:
                    return
        self.root.destroy()

    def export_docx(self):
        # Временно: экспортируем только если есть дефекты
        if not self.project["defects"]:
            messagebox.showwarning("Нет данных", "Сначала добавьте данные в отчёт.")
            return False

        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")]
        )
        if not file_path:
            return False

        export_to_docx(file_path, self.project)
        self.is_dirty = False
        messagebox.showinfo("Готово", f"Файл сохранён: {file_path}")
        return True

if __name__ == "__main__":
    if sys.platform.startswith('win'):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            ctypes.windll.user32.SetProcessDPIAware()
    root = tk.Tk()
    app = DefectApp(root)
    root.mainloop()