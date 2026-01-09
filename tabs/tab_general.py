# tabs/tab_general.py
import tkinter as tk
from tkinter import ttk, messagebox


class GeneralTabMixin:
    def build_tab_general(self):
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

        road_category_values = ["I", "II", "III", "IV", "V"]

        def bind_var(key: str) -> tk.StringVar:
            var = tk.StringVar(value=self.project["bridge"].get(key, ""))

            def on_change(*_):
                self.project["bridge"][key] = var.get()
                self.is_dirty = True

            var.trace_add("write", on_change)
            self.bridge_vars[key] = var
            return var

        lf_main = ttk.LabelFrame(form_frame, text="Основные сведения",
                                 padding=10)
        lf_main.pack(fill="x", pady=6)

        lf_location = ttk.LabelFrame(form_frame, text="Расположение / дорога",
                                     padding=10)
        lf_location.pack(fill="x", pady=6)

        lf_dates = ttk.LabelFrame(form_frame, text="Даты / примечания",
                                  padding=10)
        lf_dates.pack(fill="x", pady=6)

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

        r = 0
        add_row(lf_main, r, "ID моста", "id"); r += 1
        add_row(lf_main, r, "Тип сооружения", "structure_type"); r += 1
        add_row(lf_main, r, "Пересекаемое препятствие", "obstacle"); r += 1

        r = 0
        add_row(lf_location, r, "Дорога / улица", "road"); r += 1
        add_row(lf_location, r, "Код дороги", "road_code"); r += 1
        add_row(lf_location, r, "км привязка", "km"); r += 1
        add_row(lf_location, r, "Код региона", "region_code"); r += 1
        add_row(lf_location, r, "Координаты", "coord"); r += 1
        add_row(lf_location, r, "Категория дороги", "road_category",
                widget_type="combo"); r += 1
        add_row(lf_location, r, "Количество полос на мосту", "lanes_bridge"); r += 1
        add_row(lf_location, r, "Количество полос на подходах", "lanes_approach"); r += 1
        add_row(lf_location, r, "Год постройки", "year_built"); r += 1

        r = 0
        add_row(lf_dates, r, "Дата обследования (текущее)", "inspection_current"); r += 1
        add_row(lf_dates, r, "Дата предыдущего обследования", "inspection_prev"); r += 1
        add_row(lf_dates, r, "Примечания", "notes"); r += 1

        def show_debug():
            messagebox.showinfo("Bridge data (debug)", str(self.project["bridge"]))

        btns = ttk.Frame(form_frame)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Показать данные Ф1 (отладка)",
                   command=show_debug).pack(side="left")

    def refresh_general_tab_from_project(self):
        if not hasattr(self, "bridge_vars"):
            return
        bridge = self.project.get("bridge", {})
        for key, var in self.bridge_vars.items():
            var.set(bridge.get(key, ""))