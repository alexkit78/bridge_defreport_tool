# tabs/tab_general.py
import tkinter as tk
from tkinter import ttk, messagebox
from dictionary import (
    YES_NO,
    FLOW_DIRECTION,
    STRUCTURE_TYPE,
    ROAD_CATEGORY,
    LOADS,
    LOC_ON_PLANE,
    PAVEMENT,
    DRAINAGE,
    EXPANSION_JOINTS,
    GUARDRAILS,
    SIDEWALKS,
    RAILINGS,
    REG_STRUCTURES,
    CONE_PROTECTION,
    SIGNS_DEFAULT,
    MAINTENANCE_DEFAULT,
    ORGANIZATION_DEFAULT,
)


class GeneralTabMixin:
    def build_tab_general(self):
        container = ttk.Frame(self.tab_general)
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

                # --- mousewheel scroll for canvas (Win/Mac/Linux) ---
        def _on_mousewheel(event):
            # Windows / macOS
            if event.delta:
                # on Windows delta is multiple of 120, on macOS it can be small
                step = -1 if event.delta > 0 else 1
                canvas.yview_scroll(step, "units")
                return "break"

        def _on_mousewheel_linux_up(event):
            canvas.yview_scroll(-1, "units")
            return "break"

        def _on_mousewheel_linux_down(event):
            canvas.yview_scroll(1, "units")
            return "break"

        def _bind_wheel(_event=None):
            # bind globally while pointer is over the form/canvas
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel_linux_up)    # Linux scroll up
            canvas.bind_all("<Button-5>", _on_mousewheel_linux_down)  # Linux scroll down

        def _unbind_wheel(_event=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        form_frame = ttk.Frame(canvas, padding=10)
        canvas_window = canvas.create_window((0, 0), window=form_frame, anchor="nw")

        # включаем прокрутку когда курсор над формой
        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)
        form_frame.bind("<Enter>", _bind_wheel)
        form_frame.bind("<Leave>", _unbind_wheel)

        def _on_frame_configure(event):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        form_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # ---- helpers ----
        def bind_var(key: str) -> tk.StringVar:
            var = tk.StringVar(value=self.project["bridge"].get(key, ""))

            def on_change(*_):
                self.project["bridge"][key] = var.get()
                self.is_dirty = True

            var.trace_add("write", on_change)
            self.bridge_vars[key] = var
            return var

        def add_row(parent, row, label, key, spec):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
            var = bind_var(key)

            if spec == "entry":
                e = ttk.Entry(parent, textvariable=var)
                e.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
                return e

            if isinstance(spec, tuple) and len(spec) >= 2 and spec[0] == "combo":
                values = spec[1]
                editable = bool(spec[2]) if len(spec) >= 3 else False
                cb = ttk.Combobox(
                    parent,
                    textvariable=var,
                    values=values,
                    state=("normal" if editable else "readonly")
                )
                cb.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
                return cb

            # fallback
            e = ttk.Entry(parent, textvariable=var)
            e.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
            return e

        # ---- fields ----
        fields = [
            ("ID моста", "id", "entry"),

            ("Тип сооружения", "structure_type", ("combo", STRUCTURE_TYPE, True)),
            ("Пересекаемое препятствие", "obstacle", "entry"),

            ("Дорога / улица", "road", "entry"),
            ("В каком районе", "district", "entry"),
            ("Код дороги", "road_code", "entry"),
            ("Км привязка", "km", "entry"),
            ("Код региона", "region_code", "entry"),
            ("Координаты", "coord", "entry"),
            ("Категория дороги", "road_category", ("combo", ROAD_CATEGORY, True)),

            ("Количество полос на мосту", "lanes_bridge", "entry"),
            ("Количество полос на подходах", "lanes_approach", "entry"),
            ("Разметка", "marking", ("combo", YES_NO, False)),

            ("Ближайший населённый пункт", "nearest_settlement", "entry"),

            ("Ширина русла В", "hydro_B", "entry"),
            ("Глубина русла H", "hydro_H", "entry"),
            ("Скорость V", "hydro_V", "entry"),
            ("Направление течения", "flow_direction", ("combo", FLOW_DIRECTION, False)),

            ("Подмостовой габарит", "under_clearance", "entry"),
            ("Длина сооружения", "length", "entry"),
            ("Отверстие", "opening", "entry"),
            ("Вертикальный габарит", "height_clearance", "entry"),

            ("В", "width_B", "entry"),
            ("Г", "width_G", "entry"),
            ("С1", "width_C1", "entry"),
            ("С2", "width_C2", "entry"),
            ("Т1", "width_T1", "entry"),
            ("Т2", "width_T2", "entry"),

            ("Год постройки", "year_built", "entry"),
            ("Год реконструкции", "year_recon", "entry"),
            ("Год ремонта", "year_repair", "entry"),

            ("Нормативные нагрузки", "loads", ("combo", LOADS, True)),
            ("Продольная схема", "long_scheme", "entry"),
            ("Угол косины", "skew_angle", "entry"),
            ("Местоположение в плане", "loc_on_plane", ("combo", LOC_ON_PLANE, True)),

            ("Продольные уклоны, ‰", "longitudinal_slopes", "entry"),
            ("Поперечные уклоны, ‰", "slopes", "entry"),

            ("Покрытие на мосту", "pavement_bridge", ("combo", PAVEMENT, True)),
            ("Покрытие на подходах", "pavement_approach", ("combo", PAVEMENT, True)),

            ("Водоотвод", "drainage", ("combo", DRAINAGE, True)),
            ("Деформационные швы", "expansion_joints", ("combo", EXPANSION_JOINTS, True)),

            ("Ограждения на мосту", "guardrails_bridge", ("combo", GUARDRAILS, True)),
            ("Ограждения на подходах", "guardrails_approach", ("combo", GUARDRAILS, True)),
            ("Высота ограждений на мосту", "guardrails_height_bridge", "entry"),
            ("Высота ограждений на подходах", "guardrails_height_approach", "entry"),

            ("Тротуары", "sidewalks", ("combo", SIDEWALKS, True)),
            ("Перильное ограждение", "railings", ("combo", RAILINGS, True)),
            ("Высота п.о.", "railings_height", "entry"),

            ("Ширина подхода 1", "approach_width1", "entry"),
            ("Ширина подхода 2", "approach_width2", "entry"),
            ("Уклон подхода 1", "approach_slope1", "entry"),
            ("Уклон подхода 2", "approach_slope2", "entry"),
            ("Высота насыпи 1", "mound_height1", "entry"),
            ("Высота насыпи 2", "mound_height2", "entry"),

            ("Регуляционные сооружения", "reg_structures", ("combo", REG_STRUCTURES, True)),
            ("Укрепление конусов", "cone_protection", ("combo", CONE_PROTECTION, True)),
            ("Переходные плиты", "transition_slabs", ("combo", YES_NO, False)),

            ("Проектная организация", "design_org", ("combo", ORGANIZATION_DEFAULT, True)),
            ("Строительная организация", "build_org", ("combo", ORGANIZATION_DEFAULT, True)),
            ("Орган управления дорогой", "road_admin", ("combo", ORGANIZATION_DEFAULT, True)),
            ("Эксплуатирующая организация", "operator", ("combo", ORGANIZATION_DEFAULT, True)),

            ("Знаки до моста", "signs_before", ("combo", SIGNS_DEFAULT, True)),
            ("Знаки после моста", "signs_after", ("combo", SIGNS_DEFAULT, True)),

            ("Сведения о ремонтах", "repairs_info", "entry"),
            ("Коммуникации", "communications", ("combo", SIGNS_DEFAULT, True)),
            ("Обустройства", "maintenance", ("combo", MAINTENANCE_DEFAULT, True)),

            ("Дата обследования", "inspection_current", "entry"),
            ("Дата предыдущего обследования", "inspection_prev", "entry"),

            ("Примечания", "notes", "entry"),
        ]

        lf = ttk.LabelFrame(form_frame, text="Форма 1", padding=10)
        lf.pack(fill="both", expand=True, pady=6)
        lf.grid_columnconfigure(1, weight=1)

        for r, (label, key, spec) in enumerate(fields):
            add_row(lf, r, label, key, spec)

        # debug-кнопка по желанию
        def show_debug():
            messagebox.showinfo("Bridge data (debug)", str(self.project["bridge"]))

        btns = ttk.Frame(form_frame)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Показать данные Ф1 (отладка)", command=show_debug).pack(side="left")

    def refresh_general_tab_from_project(self):
        if not hasattr(self, "bridge_vars"):
            return
        bridge = self.project.get("bridge", {})
        for key, var in self.bridge_vars.items():
            var.set(bridge.get(key, ""))