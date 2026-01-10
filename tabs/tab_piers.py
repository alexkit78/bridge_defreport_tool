# tabs/tab_piers.py
import tkinter as tk
from tkinter import ttk, messagebox


class PiersTabMixin:
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
        self.project.setdefault("piers", [])

        uid = self._generate_uid()
        index = len(self.project["piers"]) + 1

        item = {
            "uid": uid,
            "title": f"ОПОРЫ № {index}",

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

        if not messagebox.askyesno("Удалить опору", "Удалить текущий лист опоры?"):
            return

        self.project["piers"] = [x for x in self.project["piers"] if x.get("uid") != uid_to_delete]
        self.is_dirty = True
        self.rebuild_pier_tabs()

    def _create_pier_tab_for_item(self, item: dict):
        uid = item["uid"]

        frame = ttk.Frame(self.piers_notebook, padding=10)
        title = item.get("title") or "ОПОРЫ №"
        self.piers_notebook.add(frame, text=title)

        frame.grid_columnconfigure(1, weight=1)

        def bind_var(key: str):
            var = tk.StringVar(value=item.get(key, ""))

            def on_change(*_):
                item[key] = var.get()
                self.is_dirty = True
                if key == "title":
                    try:
                        tab_index = self.piers_notebook.index(frame)
                        self.piers_notebook.tab(tab_index, text=var.get() or "ОПОРЫ №")
                    except Exception:
                        pass

            var.trace_add("write", on_change)
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

        vars_map = {}
        for r, (label, key) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=4)
            var = bind_var(key)
            vars_map[key] = var
            e = ttk.Entry(frame, textvariable=var)
            e.grid(row=r, column=1, sticky="ew", padx=6, pady=4)

        self.pier_forms[uid] = {"frame": frame, "vars": vars_map}