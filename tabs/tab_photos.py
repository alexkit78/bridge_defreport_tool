#tabs/tab_photos.py
import os
import tkinter as tk
import requests
import certifi
from tkinter import ttk, filedialog, messagebox
from tabs.photo_viewer import PhotoViewerWindow


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

        ttk.Button(
            folder_frame,
            text="Создать фото карты",
            command=self.create_map_photo
        ).pack(side="left", padx=8)



        # ---------- ФОТО ТИТУЛЬНОГО ЛИСТА ----------
        cover_frame = ttk.LabelFrame(frame, text="Фото титульного листа", padding=10)
        cover_frame.pack(fill="x", pady=5)

        self.cover_photo_cb = ttk.Combobox(cover_frame, state="readonly")
        self.cover_photo_cb.pack(fill="x", pady=2)
        self.cover_photo_cb.bind("<<ComboboxSelected>>", self.on_cover_selected)

        self.cover_caption_var = tk.StringVar()
        self.cover_caption_var.trace_add("write", lambda *_: self._save_cover_caption())
        
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
        self.photos_table.bind("<Double-1>", self._start_edit_caption)

        btns = ttk.Frame(gallery_frame)
        btns.pack(fill="x")

        ttk.Button(btns, text="Добавить фото", command=self.add_gallery_photo)\
            .pack(side="left")
        ttk.Button(btns, text="Удалить", command=self.delete_gallery_photo)\
            .pack(side="left", padx=5)
        def open_viewer():
            folder = self.project.get("photos", {}).get("folder", "")
            if not folder:
                messagebox.showwarning(
                    "Нет папки",
                    "Сначала выберите папку с фотографиями"
                )
                return
            PhotoViewerWindow(self.root, self.project, on_update=self._on_gallery_updated_from_viewer)
        ttk.Button(btns, text="Просмотр фото", command=open_viewer)\
    .pack(side="left", padx=10)

    # =====================================================
    # ЛОГИКА
    # =====================================================
    def _on_gallery_updated_from_viewer(self):
        self.refresh_gallery_table()
        if not getattr(self, "is_loading", False):
            self.is_dirty = True
    
    def _save_cover_caption(self):
            self._ensure_photos_block()
            self.project["photos"]["cover"]["caption"] = self.cover_caption_var.get()
            if not getattr(self, "is_loading", False):
                self.is_dirty = True
    
    def _start_edit_caption(self, event):
        region = self.photos_table.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = self.photos_table.identify_row(event.y)
        col = self.photos_table.identify_column(event.x)
        if not row_id or col != "#2":  # редактируем только "caption"
            return

        bbox = self.photos_table.bbox(row_id, col)
        if not bbox:
            return
        x, y, w, h = bbox

        old = self.photos_table.item(row_id, "values")[1]

        entry = tk.Entry(self.photos_table)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, old)
        entry.focus_set()

        def save(_=None):
            new_text = entry.get()
            entry.destroy()

            idx = int(row_id)
            self.project["photos"]["gallery"][idx]["caption"] = new_text
            self.refresh_gallery_table()

            if not getattr(self, "is_loading", False):
                self.is_dirty = True

        def cancel(_=None):
            entry.destroy()

        entry.bind("<Return>", save)
        entry.bind("<FocusOut>", save)
        entry.bind("<Escape>", cancel)

    def _ensure_photos_block(self):
        if "photos" not in self.project:
            self.project["photos"] = {
                "folder": "",
                "cover": {"filename": "", "caption": ""},
                "gallery": []
            }

    def refresh_photos_tab_from_project(self):
        self._ensure_photos_block()
        folder = self.project["photos"].get("folder", "") or ""
        self.photos_folder_var.set(folder)

        self.refresh_cover_controls()
        self.refresh_gallery_table()

    def select_photos_folder(self):
        self._ensure_photos_block()
        folder = filedialog.askdirectory()
        if not folder:
            return

        self.project["photos"]["folder"] = folder
        self.photos_folder_var.set(folder)

        if not getattr(self, "is_loading", False):
            self.is_dirty = True

        self.refresh_cover_controls()
        self.refresh_gallery_table()

    def refresh_cover_controls(self):
            self._ensure_photos_block()
            folder = self.project["photos"].get("folder", "") or ""

            files = []
            if folder and os.path.isdir(folder):
                files = [f for f in os.listdir(folder) if f.lower().endswith((".jpg", ".jpeg", ".png"))]
                files.sort(key=lambda s: s.lower())

            self.cover_photo_cb["values"] = files

            cover = self.project["photos"].get("cover", {}) or {}
            self.cover_photo_cb.set(cover.get("filename", "") or "")
            self.cover_caption_var.set(cover.get("caption", "") or "")

    def on_cover_selected(self, event=None):
        self._ensure_photos_block()
        filename = self.cover_photo_cb.get()
        self.project["photos"]["cover"]["filename"] = filename
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

        try:
            index = int(sel[0])
        except ValueError:
            return

        gallery = self.project["photos"]["gallery"]
        if index < 0 or index >= len(gallery):
            return

        del gallery[index]
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

    def create_map_photo(self):
        bridge = self.project.get("bridge", {})
        coord = bridge.get("coord", "")
        folder = self.project.get("photos", {}).get("folder", "")

        if not coord or "," not in coord:
            messagebox.showerror(
                "Нет координат",
                "В форме «Общие сведения» не заданы координаты моста"
            )
            return

        if not folder:
            messagebox.showwarning(
                "Нет папки",
                "Сначала выберите папку с фотографиями"
            )
            return

        lat, lon = [c.strip() for c in coord.split(",")]

        url = (
            "https://static-maps.yandex.ru/1.x/"
            f"?ll={lon},{lat}"
            "&z=15"
            "&size=650,450"
            "&l=map"
            f"&pt={lon},{lat},pm2rdm"
        )

        filename = "_map_bridge.png"
        path = os.path.join(folder, filename)

        try:
            r = requests.get(url, timeout=20, verify=certifi.where())
            r.raise_for_status()
            with open(path, "wb") as f:
                f.write(r.content)
        except Exception as e:
            messagebox.showerror("Ошибка карты", str(e))
            return

        caption = (
            "Ситуационный план расположения моста.\n"
            f"Координаты расположения моста ({lat}, {lon})"
        )

        gallery = self.project["photos"]["gallery"]

        # если карта уже есть — обновляем
        for rec in gallery:
            if rec["filename"] == filename:
                rec["caption"] = caption
                break
        else:
            gallery.insert(0, {
                "filename": filename,
                "caption": caption
            })

        self.refresh_gallery_table()

        if not getattr(self, "is_loading", False):
            self.is_dirty = True