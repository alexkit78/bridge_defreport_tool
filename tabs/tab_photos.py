import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class PhotosTabMixin:
    def build_tab_photos(self):
        frame = ttk.Frame(self.tab_photos, padding=10)
        frame.pack(fill="both", expand=True)

        self.project.setdefault("photos", {
            "folder": "",
            "cover": {"filename": "", "caption": ""},
            "gallery": []
        })

    
        # ---------- ПАПКА С ФОТО ----------
        folder_frame = ttk.LabelFrame(frame, text="Папка с фотографиями", padding=10)
        folder_frame.pack(fill="x", pady=5)

        self.photos_folder_var = tk.StringVar(value=self.project["photos"].get("folder", ""))

        ttk.Entry(folder_frame, textvariable=self.photos_folder_var, state="readonly")\
            .pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(
            folder_frame,
            text="Выбрать папку",
            command=self.select_photos_folder
        ).pack(side="left")

        # ---------- ФОТО ТИТУЛЬНОГО ЛИСТА ----------
        cover_frame = ttk.LabelFrame(frame, text="Фото титульного листа", padding=10)
        cover_frame.pack(fill="x", pady=5)

        self.cover_photo_cb = ttk.Combobox(cover_frame, state="readonly")
        self.cover_photo_cb.pack(fill="x", pady=2)
        self.cover_photo_cb.bind("<<ComboboxSelected>>", self.on_cover_selected)

        self.cover_caption_var = tk.StringVar()
        ttk.Entry(
            cover_frame,
            textvariable=self.cover_caption_var
        ).pack(fill="x", pady=2)

        # ---------- БЛОК ФОТОГРАФИЙ ----------
        gallery_frame = ttk.LabelFrame(frame, text="Фото конструкций и дефектов", padding=10)
        gallery_frame.pack(fill="both", expand=True, pady=5)

        self.photos_table = ttk.Treeview(
            gallery_frame,
            columns=("file", "caption"),
            show="headings",
            height=8
        )
        self.photos_table.heading("file", text="Файл")
        self.photos_table.heading("caption", text="Подпись")
        self.photos_table.column("file", width=200)
        self.photos_table.column("caption", width=400)

        self.photos_table.pack(fill="both", expand=True, pady=5)

        btns = ttk.Frame(gallery_frame)
        btns.pack(fill="x")

        ttk.Button(btns, text="Добавить фото", command=self.add_gallery_photo)\
            .pack(side="left")
        ttk.Button(btns, text="Удалить", command=self.delete_gallery_photo)\
            .pack(side="left", padx=5)

    # =====================================================
    # ЛОГИКА
    # =====================================================
    def _ensure_photos_block(self):
        if "photos" not in self.project:
            self.project["photos"] = {
                "folder": "",
                "cover": {"filename": "", "caption": ""},
                "gallery": []
            }

    def select_photos_folder(self):
        self._ensure_photos_block()
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.project["photos"]["folder"] = folder
        self.photos_folder_var.set(folder)

        files = [
            f for f in os.listdir(folder)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]

        self.cover_photo_cb["values"] = files

        if not getattr(self, "is_loading", False):
            self.is_dirty = True

    def on_cover_selected(self, event=None):
        self._ensure_photos_block()
        filename = self.cover_photo_cb.get()
        self.project["photos"]["cover"]["filename"] = filename
        self.project["photos"]["cover"]["caption"] = self.cover_caption_var.get()

        if not getattr(self, "is_loading", False):
            self.is_dirty = True

    def add_gallery_photo(self):
        self._ensure_photos_block()
        folder = self.project["photos"].get("folder", "")
        if not folder:
            messagebox.showwarning("Нет папки", "Сначала выберите папку с фотографиями")
            return

        files = [
            f for f in os.listdir(folder)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]

        if not files:
            return

        win = tk.Toplevel(self.root)
        win.title("Выбор фотографии")

        cb = ttk.Combobox(win, values=files, state="readonly")
        cb.pack(padx=10, pady=5)
        cb.current(0)

        caption_var = tk.StringVar()
        ttk.Entry(win, textvariable=caption_var).pack(fill="x", padx=10, pady=5)

        def add():
            rec = {
                "filename": cb.get(),
                "caption": caption_var.get()
            }
            self.project["photos"]["gallery"].append(rec)
            self.refresh_gallery_table()
            win.destroy()
            if not getattr(self, "is_loading", False):
                self.is_dirty = True

        ttk.Button(win, text="Добавить", command=add).pack(pady=10)

    def delete_gallery_photo(self):
        sel = self.photos_table.selection()
        if not sel:
            return

        index = int(sel[0])
        del self.project["photos"]["gallery"][index]
        self.refresh_gallery_table()

        if not getattr(self, "is_loading", False):
            self.is_dirty = True

    def refresh_gallery_table(self):
        self._ensure_photos_block()
        self.photos_table.delete(*self.photos_table.get_children())

        for i, rec in enumerate(self.project["photos"]["gallery"]):
            self.photos_table.insert(
                "",
                "end",
                iid=str(i),
                values=(rec.get("filename", ""), rec.get("caption", ""))
            )