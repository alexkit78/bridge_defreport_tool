# ui.py
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Menu
import ctypes

from database import Database
from export import export_to_docx
from utils import generate_uid
from project_storage import save_json, load_json
from project_model import make_empty_project


class DefectApp:
    def __init__(self, root):
        self.root = root
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

        self.notebook.add(self.tab_general, text="Общие сведения")
        self.notebook.add(self.tab_spans, text="Пролётные строения")
        self.notebook.add(self.tab_piers, text="Опоры")
        self.notebook.add(self.tab_defects, text="Дефекты")

        self.bridge_vars = {}  # ключ поля -> tk.StringVar
        self.span_forms = {}  # uid -> {"frame": Frame, "vars": {key: StringVar}}
        self.pier_forms = {} # uid -> {"frame": Frame, "vars": {key: StringVar}}

        self.build_ui()
        self.is_dirty = False


    def build_menu(self):
        menubar: Menu = tk.Menu(self.root)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Новое сооружение", 
                              command=self.new_project)
        file_menu.add_command(label="Открыть...", command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label="Сохранить...", command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)
        
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
                export_to_docx(file_path, self.project["defects"])

            self.is_dirty = False
            self.status_label.config(text=f"Файл сохранён: {file_path}")
            self.root.after(2000, lambda: self.status_label.config(text=""))
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
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

            #self.report_data.append(rec)

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
        frame = ttk.Frame(self.tab_defects, padding=10)
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
        ttk.Button(frame, text="Сохранить",
                   command=self.save_project).grid(row=14, column=1, pady=10)

        self.load_placements()

        # Таблица с прокруткой
        table_container = ttk.Frame(self.tab_defects)
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
        self.status_label = ttk.Label(self.tab_defects, text="",
        foreground="green", anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True, padx=10,
                               pady=0)
        self.count_label = ttk.Label(self.tab_defects, text="Всего дефектов: "
                                                           "0", anchor="e")
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

        self.build_tab_general()
        self.build_tab_spans()
        self.build_tab_piers()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_tab_general(self):
        # ----- scroll area ---
        container = ttk.Frame(self.tab_general)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical",
                                  command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        form_frame = ttk.Frame(canvas, padding=10)
        canvas_window = canvas.create_window((0, 0), window=form_frame,
                                             anchor="nw")
        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        form_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # --- form 1 ---
        fields = [
            ("ID моста", "id"),
            ("Тип сооружения", "structure_type"),
            ("Пересекаемое препятствие", "obstacle"),
            ("Дорога / улица", "road"),
            ("Код дороги", "road_code"),
            ("км привязка", "km"),
            ("Код региона", "region_code"),
            ("Координаты", "coord"),
            ("Категория дороги", "road_category"),
            ("Количество полос на мосту", "lanes_bridge"),
            ("Количество полос на подходах", "lanes_approach"),
            ("Год постройки", "year_built"),
            ("Дата обследования (текущее)", "inspection_current"),
            ("Дата предыдущего обследования", "inspection_prev"),
            ("Примечания", "notes"),
        ]
        # небольшой справочник прямо сейчас (потом заменим на внешний lookups.json)
        road_category_values = ["I", "II", "III", "IV", "V"]

        def bind_var(key: str) -> tk.StringVar:
            '''Создаёт StringVar и привязывает его к self.project['bridge']'''
            var = tk.StringVar(value=self.project["bridge"].get(key, ""))

            def on_change(*_):
                value = var.get()
                # записываем в проект
                self.project["bridge"][key] = value
                self.is_dirty = True

            var.trace_add("write", on_change)
            self.bridge_vars[key] = var
            return var

        # --- группировка полей для красоты (потом возможно убрать) ---
        lf_main = ttk.LabelFrame(form_frame, text="Основные сведения",
                                 padding=10)
        lf_main.pack(fill="x", pady=6)

        lf_location = ttk.LabelFrame(form_frame, text="Расположение / "
                                                      "дорога", padding=10)
        lf_location.pack(fill="x", pady=6)

        lf_dates = ttk.LabelFrame(form_frame, text="Даты / примечания",
                                  padding=10)
        lf_dates.pack(fill="x", pady=6)

        # раскладка: 2 колонки (лейбл + поле)
        def add_row(parent, row, label, key, widget_type="entry"):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w",
                                               padx=6, pady=4)

            var = bind_var(key)

            if widget_type == "combo":
                cb = ttk.Combobox(parent, textvariable=var, state="readonly",
                                  values=road_category_values)
                cb.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
                return cb
            else:
                e = ttk.Entry(parent, textvariable=var)
                e.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
                return e

        for frame in (lf_main, lf_location, lf_dates):
            frame.grid_columnconfigure(1, weight=1)

        # --- заполняем группы ---
        r = 0
        add_row(lf_main, r, "ID моста", "id");
        r += 1
        add_row(lf_main, r, "Тип сооружения", "structure_type");
        r += 1
        add_row(lf_main, r, "Пересекаемое препятствие", "obstacle");
        r += 1

        r = 0
        add_row(lf_location, r, "Дорога / улица", "road");
        r += 1
        add_row(lf_location, r, "Код дороги", "road_code");
        r += 1
        add_row(lf_location, r, "км привязка", "km");
        r += 1
        add_row(lf_location, r, "Код региона", "region_code");
        r += 1
        add_row(lf_location, r, "Координаты", "coord");
        r += 1
        add_row(lf_location, r, "Категория дороги", "road_category",
                widget_type="combo");
        r += 1
        add_row(lf_location, r, "Количество полос на мосту", "lanes_bridge");
        r += 1
        add_row(lf_location, r, "Количество полос на подходах",
                "lanes_approach");
        r += 1
        add_row(lf_location, r, "Год постройки", "year_built");
        r += 1

        r = 0
        add_row(lf_dates, r, "Дата обследования (текущее)",
                "inspection_current");
        r += 1
        add_row(lf_dates, r, "Дата предыдущего обследования",
                "inspection_prev");
        r += 1
        add_row(lf_dates, r, "Примечания", "notes");
        r += 1

        # --- кнопка (пока просто показывает, что в проекте) ---
        def show_debug():
            messagebox.showinfo("Bridge data (debug)",
                                str(self.project["bridge"]))

        btns = ttk.Frame(form_frame)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Показать данные Ф1 (отладка)",
                   command=show_debug).pack(side="left")

    def refresh_general_tab_from_project(self):
        # если вкладка ещё не создана или vars ещё нет — просто выходим
        if not hasattr(self, "bridge_vars"):
            return
        bridge = self.project.get("bridge", {})
        for key, var in self.bridge_vars.items():
            var.set(bridge.get(key, ""))

    def build_tab_spans(self):
        root = self.tab_spans

        container = ttk.Frame(root, padding=10)
        container.pack(fill="both", expand=True)

        #кнопки сверху
        btns = ttk.Frame(container)
        btns.pack(fill="x", pady=(0,8))

        ttk.Button(btns, text="Добавить ПС",
                   command=self.add_span_form).pack(side="left")
        ttk.Button(btns, text="Удалить ПС",
                   command=self.delete_current_span_form).pack(side="left",
                                                               padx=(8,0))
        # внутренний notebook = "листы" формы 2
        self.spans_notebook = ttk.Notebook(container)
        self.spans_notebook.pack(fill="both", expand=True)

        # если в проекте уже есть ПС (например после загрузки) — построим их
        self.rebuild_span_tabs()

    def rebuild_span_tabs(self):
        """Перестраивает вкладки ПС из self.project['spans']."""
        # очистить notebook
        for tab_id in self.spans_notebook.tabs():
            self.spans_notebook.forget(tab_id)

        self.span_forms.clear()

        # если пусто — можно сразу создать один лист по умолчанию (по желанию)
        if not self.project.get("spans"):
            # Создадим один лист, чтобы пользователю не было пусто
            self.add_span_form()
            return

        for st in self.project["spans"]:
            uid = st.get("uid")
            if not uid:
                continue
            self._create_span_tab_for_item(st)

    def add_span_form(self):
        """Добавляет новый лист ПС (новый элемент в spans и новую вкладку)."""
        uid = generate_uid()
        index = len(self.project["spans"]) + 1

        item = {
            "uid": uid,
            "title": f"ПС {index}",
            "span_system": "",
            "span_type": "",
            "deck_structure": "",
            "main_beam_material": "",
            "joints_type": "",
            "span_scheme": "",
            "span_loads": "",
            "typical_project": "",
            "bearings": "",
            "span_expansion_joints": "",
            "transverse_conn": "",
            "span_notes": ""
        }

        self.project["spans"].append(item)
        self.is_dirty = True

        self._create_span_tab_for_item(item)
        # переключимся на новую вкладку
        self.spans_notebook.select(self.span_forms[uid]["frame"])

    def delete_current_span_form(self):
        """Удаляет текущий лист ПС."""
        current_tab = self.spans_notebook.select()
        if not current_tab:
            return

        # current_tab — это id вкладки, но нам нужно понять uid
        uid_to_delete = None
        for uid, info in self.span_forms.items():
            if str(info["frame"]) == str(current_tab):
                uid_to_delete = uid
                break

        if not uid_to_delete:
            return

        if not messagebox.askyesno("Удалить ПС",
                                   "Удалить текущий лист пролётного строения?"):
            return

        # удалить из данных проекта
        self.project["spans"] = [x for x in self.project["spans"] if
                                      x.get("uid") != uid_to_delete]
        self.is_dirty = True

        # перестроить вкладки заново (проще и надёжнее)
        self.rebuild_span_tabs()

    def _create_span_tab_for_item(self, item: dict):
        uid = item["uid"]

        frame = ttk.Frame(self.spans_notebook, padding=10)
        title = item.get("title") or "ПС"
        self.spans_notebook.add(frame, text=title)

        frame.grid_columnconfigure(1, weight=1)

        vars_map = {}

        def bind_var(key: str):
            var = tk.StringVar(value=item.get(key, ""))

            def on_change(*_):
                item[key] = var.get()
                self.is_dirty = True

                # если изменили title — обновим текст вкладки
                if key == "title":
                    # индекс текущей вкладки
                    try:
                        tab_index = self.spans_notebook.index(frame)
                        self.spans_notebook.tab(tab_index,
                                                text=var.get() or "ПС")
                    except Exception:
                        pass

            var.trace_add("write", on_change)
            vars_map[key] = var
            return var

        fields = [
            ("Номера пролётных строений", "title"),
            ("Статическая система", "span_system"),
            ("Пролетное строение (тип)", "span_type"),
            ("Конструкция плиты ПЧ", "deck_structure"),
            ("Материал главных балок", "main_beam_material"),
            ("Тип стыков", "joints_type"),
            ("Продольная схема", "span_scheme"),
            ("Нагрузки", "span_loads"),
            ("Типовой проект", "typical_project"),
            ("Опорные части", "bearings"),
            ("Деформационные швы (Ф2)", "span_expansion_joints"),
            ("Поперечное объединение", "transverse_conn"),
            ("Примечания (Ф2)", "span_notes"),
        ]

        for r, (label, key) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w",
                                              padx=6, pady=4)
            e = ttk.Entry(frame, textvariable=bind_var(key))
            e.grid(row=r, column=1, sticky="ew", padx=6, pady=4)

        self.span_forms[uid] = {"frame": frame, "vars": vars_map}

    def build_tab_piers(self):
        root = self.tab_piers

        container = ttk.Frame(root, padding=10)
        container.pack(fill="both", expand=True)

        btns = ttk.Frame(container)
        btns.pack(fill="x", pady=(0, 8))

        ttk.Button(btns, text="Добавить опору",
                   command=self.add_pier_form).pack(side="left")
        ttk.Button(btns, text="Удалить опору",
                   command=self.delete_current_pier_form).pack(side="left",
                                                               padx=(8, 0))

        self.piers_notebook = ttk.Notebook(container)
        self.piers_notebook.pack(fill="both", expand=True)

        self.rebuild_pier_tabs()

    def rebuild_pier_tabs(self):
        """Перестраивает вкладки опор из self.project['piers']."""
        self.project.setdefault("piers", [])

        for tab_id in self.piers_notebook.tabs():
            self.piers_notebook.forget(tab_id)

        self.pier_forms.clear()

        if not self.project["piers"]:
            self.add_pier_form()
            return

        for item in self.project["piers"]:
            uid = item.get("uid")
            if not uid:
                continue
            self._create_pier_tab_for_item(item)

    def add_pier_form(self):
        """Добавляет новый лист опоры (новый элемент в piers и новую вкладку)."""
        self.project.setdefault("piers", [])

        uid = generate_uid()
        index = len(self.project["piers"]) + 1

        item = {
            "uid": uid,
            "title": f"Опора {index}",

            "piers_type": "",
            "foundation_type": "",
            "pier_material": "",
            "pier_height": "",
            "foundation_depth": "",
            "pier_typical_project": "",

            "pier_size_a": "",
            "pier_size_b": "",
            "piles_qty": "",
            "piles_spacing": "",
            "pier_scheme": "",

            "pier_rigel_width": "",
            "pier_rigel_height": "",
            "pier_rigel_length": "",
            "pile_section": "",

            "pier_notes": ""
        }

        self.project["piers"].append(item)
        self.is_dirty = True

        self._create_pier_tab_for_item(item)
        self.piers_notebook.select(self.pier_forms[uid]["frame"])

    def delete_current_pier_form(self):
        """Удаляет текущий лист опоры."""
        current_tab = self.piers_notebook.select()
        if not current_tab:
            return

        uid_to_delete = None
        for uid, info in self.pier_forms.items():
            if str(info["frame"]) == str(current_tab):
                uid_to_delete = uid
                break

        if not uid_to_delete:
            return

        if not messagebox.askyesno("Удалить опору",
                                   "Удалить текущий лист опоры?"):
            return

        self.project["piers"] = [x for x in self.project["piers"] if
                                 x.get("uid") != uid_to_delete]
        self.is_dirty = True

        self.rebuild_pier_tabs()

    def _create_pier_tab_for_item(self, item: dict):
        uid = item["uid"]

        frame = ttk.Frame(self.piers_notebook, padding=10)
        title = item.get("title") or "Опора"
        self.piers_notebook.add(frame, text=title)

        frame.grid_columnconfigure(1, weight=1)

        vars_map = {}

        def bind_var(key: str):
            var = tk.StringVar(value=item.get(key, ""))

            def on_change(*_):
                item[key] = var.get()
                self.is_dirty = True

                if key == "title":
                    try:
                        tab_index = self.piers_notebook.index(frame)
                        self.piers_notebook.tab(tab_index,
                                                text=var.get() or "Опора")
                    except Exception:
                        pass

            var.trace_add("write", on_change)
            vars_map[key] = var
            return var

        fields = [
            ("Номера опор", "title"),

            ("Тип опоры", "piers_type"),
            ("Тип фундамента", "foundation_type"),
            ("Материал опоры", "pier_material"),
            ("Высота опоры", "pier_height"),
            ("Глубина фундамента", "foundation_depth"),
            ("Типовой проект опоры", "pier_typical_project"),

            ("Размер вдоль сооружения", "pier_size_a"),
            ("Размер поперёк сооружения", "pier_size_b"),

            ("Кол-во свай", "piles_qty"),
            ("Шаг свай", "piles_spacing"),
            ("Схема опоры", "pier_scheme"),

            ("Ширина ригеля", "pier_rigel_width"),
            ("Высота ригеля", "pier_rigel_height"),
            ("Длина ригеля", "pier_rigel_length"),

            ("Сечение свай", "pile_section"),

            ("Примечания (Ф3)", "pier_notes"),
        ]

        for r, (label, key) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w",
                                              padx=6, pady=4)
            e = ttk.Entry(frame, textvariable=bind_var(key))
            e.grid(row=r, column=1, sticky="ew", padx=6, pady=4)

        self.pier_forms[uid] = {"frame": frame, "vars": vars_map}

    # ------- UI callbacks ----------------

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

    def load_placements(self):
        self.placement_cb['values'] = self.db.get_placements()

    def load_defects(self, event=None):
        placement = self.placement_cb.get()

        # Сброс предыдущего выбора
        self.defect_cb.set("")
        self.option_cb.set("")

        rows = self.db.get_defects_by_placement(placement)

        self.defect_options_by_numodm.clear()
        self.defect_numodm_map.clear()
        self.defect_categories_by_option.clear()

        for num_odm, name, option, s, d, r, l, localization in rows:
            key = (name, localization or "")
            self.defect_numodm_map[key] = (
            num_odm, placement, localization or "")

            if num_odm not in self.defect_options_by_numodm:
                self.defect_options_by_numodm[num_odm] = []
            if option:
                if option not in self.defect_options_by_numodm[num_odm]:
                    self.defect_options_by_numodm[num_odm].append(option)

            self.defect_categories_by_option[(name, option)] = (s, d, r, l)

        # Combobox с отображением локализации
        self.defect_cb["values"] = sorted(
            [f"{name} ({localization})" if localization else name
             for (name, localization), (num, pl, loc) in
             self.defect_numodm_map.items()
             if pl == placement]
        )

    def filter_defect_names(self, event=None):
        text = self.search_entry.get().lower()
        placement = self.placement_cb.get()

        # фильтруем по name и localization
        filtered_keys = [
            key for key, (num_odm, pl, localization) in
            self.defect_numodm_map.items()
            if pl == placement and (
                        text in key[0].lower() or text in key[1].lower())
        ]

        # отображаем в combobox красиво
        self.defect_cb["values"] = [
            f"{name} ({localization})" if localization else name
            for name, localization in filtered_keys
        ]

    def populate_defect_fields(self, event=None):
        selected_text = self.defect_cb.get()

        # разбиваем отображаемый текст на name и localization
        if "(" in selected_text and selected_text.endswith(")"):
            name, localization = selected_text.rsplit("(", 1)
            name = name.strip()
            localization = localization[:-1]  # убрать закрывающую скобку
        else:
            name = selected_text
            localization = ""

        # получаем num_odm и опции по ключу (name, localization)
        num_odm, _, _ = self.defect_numodm_map.get((name, localization),
                                                   (None, None, None))

        # заполняем описание дефекта
        options = self.defect_options_by_numodm.get(num_odm, [])
        self.option_cb["values"] = options
        if options:
            self.option_cb.set(options[0])
            self.populate_category_fields()

        # --- автоподстановка мероприятия по дефекту ---
        placement = self.placement_cb.get()

        # автоподстановка мероприятия
        if num_odm:
            repair_action = self.db.get_repair_action(num_odm)

            self.action_entry.delete(0, tk.END)
            if repair_action:
                self.action_entry.insert(0, repair_action)

    def populate_category_fields(self, event=None):
        option = self.option_cb.get()
        defect_name = self.defect_cb.get()

        # если в combobox добавлена локализация, нужно её убрать для поиска
        if "(" in defect_name and defect_name.endswith(")"):
            defect_name = defect_name.rsplit("(", 1)[0].strip()

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
        total = len(self.project["defects"])
        self.count_label.config(text=f"Всего дефектов: {total}")
    def add_entry(self):
        placement = self.placement_cb.get()
        location = self.location_entry.get()
        selected_text = self.defect_cb.get()
        option = self.option_cb.get()
        safety = self.safety_entry.get()
        durability = self.durability_entry.get()
        repairability = self.repair_entry.get()
        loadcap = self.load_entry.get()
        action_text = self.action_entry.get()

        if not placement or not selected_text:
            messagebox.showerror("Ошибка", "Выберите раздел и тип дефекта")
            return

        # Разбираем Combobox на name и localization, сохраняем только name
        if "(" in selected_text and selected_text.endswith(")"):
            name, localization = selected_text.rsplit("(", 1)
            name = name.strip()
        else:
            name = selected_text

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
        self.project["defects"].append(rec)

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
        self.count_label.config(text=f"Всего дефектов: "
                                     f"{len(self.project['defects'])}")
        self.root.after(2000, lambda: self.status_label.config(text=""))
        self.action_entry.delete(0, tk.END)
        self.is_dirty = True

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

            # синхронизация project["defects"]
            for rec in self.project["defects"]:
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
            self.is_dirty = True

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
            self.project["defects"] = [r for r in self.project["defects"] if r[
                'uid'] != iid]
            self.table.delete(iid)
        self.update_status_bar()
        self.is_dirty = True


    def export_docx(self):
        if not self.project["defects"]:
            messagebox.showwarning("Нет данных", "Сначала добавьте данные в отчёт.")
            return False

        file_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")]
        )
        if not file_path:
            return False

        export_to_docx(file_path, self.project["defects"])
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