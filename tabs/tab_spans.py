# tabs/tab_spans.py
import tkinter as tk
from tkinter import ttk, messagebox
from tabs.scrollable import make_scrollable_frame
from dictionary import (
    SPAN_SYSTEM, SPAN_TYPE, DECK_STRUCTURE, MATERIAL, JOINTS_TYPE,
    BEARINGS, EXPANSION_JOINTS, TRANSVERSE_CONN, PAVEMENT
)

class SpansTabMixin:
    def build_tab_spans(self):
        root = self.tab_spans

        container = ttk.Frame(root, padding=10)
        container.pack(fill="both", expand=True)

        btns = ttk.Frame(container)
        btns.pack(fill="x", pady=(0, 8))

        ttk.Button(btns, text="Добавить ПС",
                   command=self.add_span_form).pack(side="left")
        ttk.Button(btns, text="Удалить ПС",
                   command=self.delete_current_span_form).pack(side="left",
                                                               padx=(8, 0))

        self.spans_notebook = ttk.Notebook(container)
        self.spans_notebook.pack(fill="both", expand=True)

        self.rebuild_span_tabs()

    def rebuild_span_tabs(self):
        for tab_id in self.spans_notebook.tabs():
            self.spans_notebook.forget(tab_id)

        self.span_forms.clear()

        if not self.project.get("spans"):
            self.add_span_form()
            return

        for st in self.project["spans"]:
            uid = st.get("uid")
            if not uid:
                continue
            self._create_span_tab_for_item(st)

    def add_span_form(self):
        uid = self._generate_uid()
        index = len(self.project["spans"]) + 1

        item = {
            "uid": uid,
            "title": f"Пролёты № {index}",

            "span_system": "",
            "span_type": "",
            "deck_structure": "",
            "main_beam_material": "",
            "joints_type": "",
            "span_scheme": "",

            "span_width_B": "",
            "span_width_G": "",
            "span_width_C1": "",
            "span_width_C2": "",
            "span_width_T1": "",
            "span_width_T2": "",

            "span_year": "",
            "typical_project": "",
            "bearings": "",
            "span_expansion_joints": "",
            "transverse_conn": "",
            "transverse_scheme": "",

            "deck_thickness": "",
            "deck_material": "",

            "pavement_thickness": "",
            "pavement_extrathickness": "",
            "pavement_material": "",

            "main_beams_qty": "",
            "main_beam_h_mid": "",
            "main_beam_h_support": "",

            "cross_beams": "",
            "long_beams": "",
            "extra_loads": "",

            "span_notes": "",
        }

        self.project["spans"].append(item)
        self.is_dirty = True

        self._create_span_tab_for_item(item)
        self.spans_notebook.select(self.span_forms[uid]["tab"])

    def delete_current_span_form(self):
        current_tab = self.spans_notebook.select()
        if not current_tab:
            return

        uid_to_delete = None
        for uid, info in self.span_forms.items():
            if str(info["tab"]) == str(current_tab):
                uid_to_delete = uid
                break

        if not uid_to_delete:
            return

        if not messagebox.askyesno("Удалить ПС",
                                   "Удалить текущий лист пролётного строения?"):
            return

        self.project["spans"] = [
            x for x in self.project["spans"] if x.get("uid") != uid_to_delete
        ]
        self.is_dirty = True
        self.rebuild_span_tabs()

    def _create_span_tab_for_item(self, item: dict):
        uid = item["uid"]

        tab = ttk.Frame(self.spans_notebook, padding=0)
        title = item.get("title") or "Пролёты №"
        self.spans_notebook.add(tab, text=title)

        inner = make_scrollable_frame(tab)
        inner.grid_columnconfigure(1, weight=1)

        
        def bind_var(key: str):
            var = tk.StringVar(value=item.get(key, ""))

            def on_change(*_):
                item[key] = var.get()
                self.is_dirty = True
                if key == "title":
                    try:
                        tab_index = self.spans_notebook.index(tab)
                        self.spans_notebook.tab(tab_index, text=var.get() or "Пролёты №")
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
            ("Номера пролётных строений", "title", "entry"),

            ("Статическая система", "span_system", ("combo", SPAN_SYSTEM, True)),
            ("Пролетное строение (тип)", "span_type", ("combo", SPAN_TYPE, True)),
            ("Конструкция плиты проезжей части", "deck_structure", ("combo", DECK_STRUCTURE, True)),
            ("Материал главных балок", "main_beam_material", ("combo", MATERIAL, True)),
            ("Тип стыков", "joints_type", ("combo", JOINTS_TYPE, True)),
            ("Продольная схема", "span_scheme", "entry"),

            ("Ширина В", "span_width_B", "entry"),
            ("Ширина Г", "span_width_G", "entry"),
            ("Ширина C1", "span_width_C1", "entry"),
            ("Ширина C2", "span_width_C2", "entry"),
            ("Ширина T1", "span_width_T1", "entry"),
            ("Ширина T2", "span_width_T2", "entry"),

            ("Год изготовления", "span_year", "entry"),
            ("Типовой проект", "typical_project", "entry"),

            ("Опорные части", "bearings", ("combo", BEARINGS, True)),
            ("Деформационные швы", "span_expansion_joints", ("combo", EXPANSION_JOINTS, True)),
            ("Поперечное объединение", "transverse_conn", ("combo", TRANSVERSE_CONN, True)),
            ("Поперечная схема", "transverse_scheme", "entry"),

            ("Толщина плиты ПЧ", "deck_thickness", "entry"),
            ("Материал плиты", "deck_material", ("combo", MATERIAL, True)),

            ("Толщина покрытия ПЧ", "pavement_thickness", "entry"),
            ("Толщина доп. слоя покрытия", "pavement_extrathickness", "entry"),
            ("Материал покрытия ПЧ", "pavement_material", ("combo", PAVEMENT, True)),

            ("Кол-во главных балок", "main_beams_qty", "entry"),
            ("Высота балки в середине", "main_beam_h_mid", "entry"),
            ("Высота балки у опоры", "main_beam_h_support", "entry"),

            ("Поперечные балки", "cross_beams", "entry"),
            ("Продольные балки", "long_beams", "entry"),
            ("Доп. нагрузки", "extra_loads", "entry"),

            ("Примечания", "span_notes", "entry"),
        ]
        
        for r, (label, key, spec) in enumerate(fields):
            add_row(inner, r, label, key, spec)
            
        self.span_forms[uid] = {"tab": tab}

