# tabs/tab_piers.py
import tkinter as tk
from tkinter import ttk, messagebox
from tabs.scrollable import make_scrollable_frame
from dictionary import (
    PIERS_TYPE, FOUNDATION_TYPE, MATERIAL, FOUNDATION_DEPTH_DEFAULT
)

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
            "foundation_depth": "нет сведений",
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
        if not getattr(self, "is_loading", False):
            self.is_dirty = True

        self._create_pier_tab_for_item(item)
        self.piers_notebook.select(self.pier_forms[uid]["tab"])

    def delete_current_pier_form(self):
        current_tab = self.piers_notebook.select()
        if not current_tab:
            return

        uid_to_delete = None
        for uid, info in self.pier_forms.items():
            if str(info["tab"]) == str(current_tab):
                uid_to_delete = uid
                break

        if not uid_to_delete:
            return

        if not messagebox.askyesno("Удалить опору", "Удалить текущий лист опоры?"):
            return

        self.project["piers"] = [x for x in self.project["piers"] if x.get("uid") != uid_to_delete]
        if not getattr(self, "is_loading", False):
            self.is_dirty = True
        self.rebuild_pier_tabs()

    def _create_pier_tab_for_item(self, item: dict):
        uid = item["uid"]

        tab = ttk.Frame(self.piers_notebook, padding=0)
        title = item.get("title") or "Пролёты №"
        self.piers_notebook.add(tab, text=title)

        inner = make_scrollable_frame(tab)
        inner.grid_columnconfigure(1, weight=1)

        def bind_var(key: str):
            var = tk.StringVar(value=item.get(key, ""))

            def on_change(*_):
                item[key] = var.get()
                if not getattr(self, "is_loading", False):
                    self.is_dirty = True
                if key == "title":
                    try:
                        tab_index = self.piers_notebook.index(tab)
                        self.piers_notebook.tab(tab_index, text=var.get() or "ОПОРЫ №")
                    except Exception:
                        pass

            var.trace_add("write", on_change)
            return var
        
        def add_row(parent, row, label, key, spec):
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
            var = bind_var(key)

            if spec == "entry":
                e = ttk.Entry(parent, textvariable=var)
                e.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
                return e

            # spec = ("combo", values, editable)
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

        fields = [
            ("Номера опор", "title", "entry"),

            ("Тип опор", "piers_type", ("combo", PIERS_TYPE, True)),
            ("Тип фундамента", "foundation_type", ("combo", FOUNDATION_TYPE, True)),
            ("Материал опор", "pier_material", ("combo", MATERIAL, True)),
            ("Высота опоры", "pier_height", "entry"),
            ("Глубина фундамента", "foundation_depth", ("combo", FOUNDATION_DEPTH_DEFAULT, True)),
            ("Типовой проект опоры", "pier_typical_project", "entry"),

            ("Размер вдоль сооружения", "pier_size_a", "entry"),
            ("Размер поперёк сооружения", "pier_size_b", "entry"),
            ("Кол-во свай", "piles_qty", "entry"),
            ("Шаг свай", "piles_spacing", "entry"),
            ("Схема опоры", "pier_scheme", "entry"),

            ("Ширина ригеля", "pier_rigel_width", "entry"),
            ("Высота ригеля", "pier_rigel_height", "entry"),
            ("Длина ригеля", "pier_rigel_length", "entry"),
            ("Сечение свай", "pile_section", "entry"),

            ("Примечания", "pier_notes", "entry"),
        ]
        
        for r, (label, key, spec) in enumerate(fields):
                add_row(inner, r, label, key, spec)

        self.pier_forms[uid] = {"tab": tab}