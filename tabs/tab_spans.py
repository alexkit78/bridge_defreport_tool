# tabs/tab_spans.py
import tkinter as tk
from tkinter import ttk, messagebox


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
        self.spans_notebook.select(self.span_forms[uid]["frame"])

    def delete_current_span_form(self):
        current_tab = self.spans_notebook.select()
        if not current_tab:
            return

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

        self.project["spans"] = [
            x for x in self.project["spans"] if x.get("uid") != uid_to_delete
        ]
        self.is_dirty = True
        self.rebuild_span_tabs()

    def _create_span_tab_for_item(self, item: dict):
        uid = item["uid"]

        frame = ttk.Frame(self.spans_notebook, padding=10)
        title = item.get("title") or "Пролёты №"
        self.spans_notebook.add(frame, text=title)

        frame.grid_columnconfigure(1, weight=1)

        def bind_var(key: str):
            var = tk.StringVar(value=item.get(key, ""))

            def on_change(*_):
                item[key] = var.get()
                self.is_dirty = True
                if key == "title":
                    try:
                        tab_index = self.spans_notebook.index(frame)
                        self.spans_notebook.tab(tab_index, text=var.get() or "Пролёты №")
                    except Exception:
                        pass

            var.trace_add("write", on_change)
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
            ("Деформационные швы", "span_expansion_joints"),
            ("Поперечное объединение", "transverse_conn"),
            ("Примечания", "span_notes"),
        ]

        vars_map = {}
        for r, (label, key) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=4)
            var = bind_var(key)
            vars_map[key] = var
            e = ttk.Entry(frame, textvariable=var)
            e.grid(row=r, column=1, sticky="ew", padx=6, pady=4)

        self.span_forms[uid] = {"frame": frame, "vars": vars_map}

