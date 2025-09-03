# ui.py
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ctypes

from database import Database
from export import export_to_docx
from utils import generate_uid
from constants import DB_PATH


class DefectApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Оценка состояния мостового сооружения")

        # инициализация БД
        self.db = Database()

        # список дефектов, которые пойдут в отчёт
        self.report_data = []
        self.defect_options_by_numodm = {}
        self.defect_numodm_map = {}
        self.defect_categories_by_option = {}

        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=26)

        ttk.Label(frame, text="Раздел:").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.placement_cb = ttk.Combobox(frame, state="readonly")
        self.placement_cb.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        self.placement_cb.bind("<<ComboboxSelected>>", self.load_defects)

        ttk.Label(frame, text="Местоположение дефекта:").grid(row=2, column=0, sticky="w", padx=6, pady=2)
        self.location_entry = ttk.Entry(frame)
        self.location_entry.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=2)

        ttk.Label(frame, text="Поиск по типу дефекта:").grid(row=4, column=0, sticky="w", padx=6, pady=2)
        self.search_entry = ttk.Entry(frame)
        self.search_entry.grid(row=5, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        self.search_entry.bind("<KeyRelease>", self.filter_defect_names)

        ttk.Label(frame, text="Тип дефекта:").grid(row=6, column=0, sticky="w", padx=6, pady=2)
        self.defect_cb = ttk.Combobox(frame, state="readonly")
        self.defect_cb.grid(row=7, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        self.defect_cb.bind("<<ComboboxSelected>>", self.populate_defect_fields)

        ttk.Label(frame, text="Описание дефекта:").grid(row=8, column=0, sticky="w", padx=6, pady=2)
        self.option_cb = ttk.Combobox(frame, state="readonly")
        self.option_cb.grid(row=9, column=0, columnspan=2, sticky="ew", padx=5, pady=2)
        self.option_cb.bind("<<ComboboxSelected>>", self.populate_category_fields)

        ttk.Label(frame, text="Категории (Б, Д, Р, Г):").grid(row=10, column=0, sticky="w", padx=6, pady=2)
        cat_frame = ttk.Frame(frame)
        cat_frame.grid(row=11, column=0, columnspan=2, sticky="w")

        self.safety_entry = ttk.Entry(cat_frame, width=5)
        self.durability_entry = ttk.Entry(cat_frame, width=5)
        self.repair_entry = ttk.Entry(cat_frame, width=5)
        self.load_entry = ttk.Entry(cat_frame, width=5)

        self.safety_entry.pack(side="left", padx=5)
        self.durability_entry.pack(side="left", padx=5)
        self.repair_entry.pack(side="left", padx=5)
        self.load_entry.pack(side="left", padx=5)

        ttk.Label(frame, text="Мероприятия по устранению дефекта:").grid(row=12, column=0, sticky="w", padx=6, pady=2)
        self.action_entry = ttk.Entry(frame)
        self.action_entry.grid(row=13, column=0, columnspan=2, sticky="ew", padx=5, pady=2)

        ttk.Button(frame, text="Добавить в отчёт", command=self.add_entry).grid(row=14, column=0, pady=10)
        ttk.Button(frame, text="Сохранить в Word", command=self.export_docx).grid(row=14, column=1, pady=10)

        self.load_placements()

        # Таблица с прокруткой
        table_container = ttk.Frame(self.root)
        table_container.pack(fill="both", expand=True)

        self.tree_scroll_y = ttk.Scrollbar(table_container, orient="vertical")
        self.tree_scroll_y.pack(side="right", fill="y")

        self.tree_scroll_x = ttk.Scrollbar(table_container, orient="horizontal")
        self.tree_scroll_x.pack(side="bottom", fill="x")

        self.table = ttk.Treeview(
            table_container,
            columns=("Раздел", "Местоположение", "Тип", "Описание", "Категории", "Мероприятия"),
            show="headings",
            yscrollcommand=self.tree_scroll_y.set,
            xscrollcommand=self.tree_scroll_x.set
        )

        self.tree_scroll_y.config(command=self.table.yview)
        self.tree_scroll_x.config(command=self.table.xview)

        for col in self.table["columns"]:
            self.table.heading(col, text=col)
            self.table.column(col, minwidth=50, width=150, stretch=False)

        self.table.pack(fill="both", expand=True)

        # биндим двойной клик для редактирования и правый клик для контекстного меню
        self.table.bind("<Double-1>", self.start_cell_edit)
        self.table.bind("<Button-3>", self.show_context_menu)
        # self.table.bind("<Control-Button-1>", self.show_context_menu)
        self.table.bind("<Button-2>", self.show_context_menu)  # пробуем
        # контекстное меню удаления строки по правому клику

        # Обработчики колесика мыши — на всякий случай обеспечим корректное скроллирование
        if sys.platform == 'darwin':
            # MacOS
            self.table.bind_all('<MouseWheel>', self._on_mousewheel_mac)
        else:
            # Windows / Linux
            self.table.bind_all('<MouseWheel>', self._on_mousewheel)
            self.table.bind_all('<Button-4>',
                                lambda e: self.table.yview_scroll(-1, 'units'))
            self.table.bind_all('<Button-5>',
                                lambda e: self.table.yview_scroll(1, 'units'))

        # строка статуса
        self.status_label = ttk.Label(self.root, text="",
        foreground="green", anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True, padx=10,
                               pady=0)
        self.count_label = ttk.Label(self.root, text="Всего дефектов: 0", anchor="e")
        self.count_label.pack(side="right", padx=10, pady=0)

        self.root.bind_class("Entry", "<Control-a>",
                             lambda e: e.widget.select_range(0, tk.END))
        self.root.bind_class("Entry", "<Command-a>",
                             lambda e: e.widget.select_range(0, tk.END))
        self.root.bind_class("Entry", "<Control-A>",
                             lambda e: e.widget.select_range(0, tk.END))

        # Вызов стандартного меню для Entry / Text / Combobox
        self.root.bind_class("Entry",
                             "<Button-3>" if sys.platform != 'darwin' else "<Button-2>",
                             self._show_entry_menu)
        self.root.bind_class("TEntry",
                             "<Button-3>" if sys.platform != 'darwin' else "<Button-2>",
                             self._show_entry_menu)
        self.root.bind_class("Text",
                             "<Button-3>" if sys.platform != 'darwin' else "<Button-2>",
                             self._show_entry_menu)
        self.root.bind_class("TCombobox",
                             "<Button-3>" if sys.platform != 'darwin' else "<Button-2>",
                             self._show_entry_menu)

        # Горячие клавиши в любой раскладке
        self.root.bind_all("<Control-KeyPress>", self._global_shortcuts)

    # ---------------- UI callbacks ----------------

    def _global_shortcuts(self, event):
        key = event.keysym.lower()
        widget = self.root.focus_get()
        if key == 'c':
            try:
                widget.event_generate("<<Copy>>")
            except:
                pass
        elif key == 'v':
            try:
                widget.event_generate("<<Paste>>")
            except:
                pass
        elif key == 'x':
            try:
                widget.event_generate("<<Cut>>")
            except:
                pass

    def _show_entry_menu(self, event):
        widget = event.widget
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Копировать",
                         command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Вставить",
                         command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_command(label="Вырезать",
                         command=lambda: widget.event_generate("<<Cut>>"))
        menu.tk_popup(event.x_root, event.y_root)

    # --- скролл-обработчики ---
    def _on_mousewheel(self, event):
        # Windows: event.delta кратно 120
        self.table.yview_scroll(int(-1 * (event.delta / 120)), 'units')

    def _on_mousewheel_mac(self, event):
        # MacOS: event.delta небольшое значение
        self.table.yview_scroll(int(-1 * event.delta), 'units')

    def load_placements(self):
        self.placement_cb['values'] = self.db.get_placements()

    def load_defects(self, event=None):
        placement = self.placement_cb.get()
        rows = self.db.get_defects_by_placement(placement)

        self.defect_options_by_numodm.clear()
        self.defect_numodm_map.clear()
        self.defect_categories_by_option.clear()

        for num_odm, name, option, s, d, r, l in rows:
            self.defect_numodm_map[name] = (num_odm, placement)
            if num_odm not in self.defect_options_by_numodm:
                self.defect_options_by_numodm[num_odm] = []
            if option:
                if option not in self.defect_options_by_numodm[num_odm]:
                    self.defect_options_by_numodm[num_odm].append(option)

                # теперь ключ уникален: (название дефекта, опция)
                self.defect_categories_by_option[(name, option)] = (s, d, r, l)

        self.defect_cb["values"] = sorted(
            [name for name, (num, pl) in self.defect_numodm_map.items() if
             pl == placement]
        )

    def filter_defect_names(self, event=None):
        text = self.search_entry.get().lower()
        placement = self.placement_cb.get()
        filtered = [name for name, (num, pl) in self.defect_numodm_map.items() if pl == placement and text in name.lower()]
        self.defect_cb["values"] = sorted(set(filtered))

    def populate_defect_fields(self, event=None):
        name = self.defect_cb.get()
        num_odm, _ = self.defect_numodm_map.get(name, (None, None))
        options = self.defect_options_by_numodm.get(num_odm, [])
        self.option_cb["values"] = options
        if options:
            self.option_cb.set(options[0])
            self.populate_category_fields()

    def populate_category_fields(self, event=None):
        option = self.option_cb.get()
        defect_name = self.defect_cb.get()
        key = (defect_name, option)

        if key in self.defect_categories_by_option:
            s, d, r, l = self.defect_categories_by_option[key]

            self.safety_entry.delete(0, tk.END)
            self.safety_entry.insert(0, s)

            self.durability_entry.delete(0, tk.END)
            self.durability_entry.insert(0, d)

            self.repair_entry.delete(0, tk.END)
            self.repair_entry.insert(0, r)

            self.load_entry.delete(0, tk.END)
            self.load_entry.insert(0, l)

    def update_status_bar(self):
        total = len(self.report_data)
        self.count_label.config(text=f"Всего дефектов: {total}")
    def add_entry(self):
        placement = self.placement_cb.get()
        location = self.location_entry.get()
        name = self.defect_cb.get()
        option = self.option_cb.get()
        safety = self.safety_entry.get()
        durability = self.durability_entry.get()
        repairability = self.repair_entry.get()
        loadcap = self.load_entry.get()
        action_text = self.action_entry.get()

        if not placement or not name:
            messagebox.showerror("Ошибка", "Выберите раздел и тип дефекта")
            return

        uid = generate_uid()
        rec = {
            'uid': uid,
            'placement': placement,
            'location': location,
            'name': name,
            'option': option,
            'safety': safety,
            'durability': durability,
            'repairability': repairability,
            'loadcap': loadcap,
            'action': action_text
        }
        self.report_data.append(rec)

        cats = []
        if safety: cats.append(f"Б{safety}")
        if durability: cats.append(f"Д{durability}")
        if repairability: cats.append(f"Р{repairability}")
        try:
            if loadcap and int(loadcap) == 1:
                cats.append("Г")
        except ValueError:
            pass

        self.table.insert("", "end", iid=uid, values=(placement, location, name, option, ", ".join(cats), action_text))

        self.status_label.config(text="Строка добавлена в отчёт")
        self.count_label.config(text=f"Всего дефектов: {len(self.report_data)}")
        self.root.after(2000, lambda: self.status_label.config(text=""))
        self.action_entry.delete(0, tk.END)

    def start_cell_edit(self, event):
        # Редактируемые столбцы: 1=location, 2=name, 3=option, 5=action
        region = self.table.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.table.identify_row(event.y)
        col = self.table.identify_column(event.x)
        if not row_id or not col:
            return
        col_index = int(col.replace('#', '')) - 1
        if col_index in (0, 4):
            return  # Раздел и Категории не редактируем

        bbox = self.table.bbox(row_id, col)
        if not bbox:
            return
        x, y, width, height = bbox

        value = self.table.item(row_id, 'values')[col_index]

        # tk.Entry вместо ttk.Entry
        entry = tk.Entry(self.table)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, value)
        entry.focus()

        # Горячие клавиши
        entry.bind("<Control-a>", lambda e: entry.select_range(0, tk.END))
        entry.bind("<Control-c>", lambda e: (self.root.clipboard_clear(),
                                             self.root.clipboard_append(
                                                 entry.selection_get())))
        entry.bind("<Control-v>", lambda e: entry.insert(tk.INSERT,
                                                         self.root.clipboard_get()))
        entry.bind("<Control-x>", lambda e: (self.root.clipboard_clear(),
                                             self.root.clipboard_append(
                                                 entry.selection_get()),
                                             entry.delete("sel.first",
                                                          "sel.last")))

        # Контекстное меню
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Копировать",
                         command=lambda: (self.root.clipboard_clear(),
                                          self.root.clipboard_append(
                                              entry.selection_get())))
        menu.add_command(label="Вставить",
                         command=lambda: entry.insert(tk.INSERT,
                                                      self.root.clipboard_get()))
        menu.add_command(label="Вырезать",
                         command=lambda: (self.root.clipboard_clear(),
                                          self.root.clipboard_append(
                                              entry.selection_get()),
                                          entry.delete("sel.first",
                                                       "sel.last")))

        entry.bind("<Button-3>",
                   lambda e: menu.tk_popup(e.x_root, e.y_root))  # только ПКМ

        def save_edit(event=None):
            new_value = entry.get()
            values = list(self.table.item(row_id, 'values'))
            values[col_index] = new_value
            self.table.item(row_id, values=values)
            entry.destroy()

            # синхронизация report_data
            for rec in self.report_data:
                if rec['uid'] == row_id:
                    if col_index == 1:
                        rec['location'] = new_value
                    elif col_index == 2:
                        rec['name'] = new_value
                    elif col_index == 3:
                        rec['option'] = new_value
                    elif col_index == 5:
                        rec['action'] = new_value
                    break
            self.update_status_bar()

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    def show_context_menu(self, event):
        row_id = self.table.identify_row(event.y)
        if row_id:
            self.table.selection_set(row_id)
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Удалить", command=self.delete_selected_row)
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    def delete_selected_row(self):
        sel = self.table.selection()
        if not sel:
            return
        for iid in sel:
            # удалить из report_data
            self.report_data = [r for r in self.report_data if r['uid'] != iid]
            self.table.delete(iid)
        self.update_status_bar()

    def export_docx(self):
        if not self.report_data:
            messagebox.showwarning("Нет данных", "Сначала добавьте данные в отчёт.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word document", "*.docx")])
        if file_path:
            export_to_docx(file_path, self.report_data)
            messagebox.showinfo("Готово", f"Файл сохранён: {file_path}")


if __name__ == "__main__":
    if sys.platform.startswith('win'):
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            ctypes.windll.user32.SetProcessDPIAware()
    root = tk.Tk()
    app = DefectApp(root)
    root.mainloop()