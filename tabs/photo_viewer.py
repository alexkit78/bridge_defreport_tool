# tabs/photo_viewer.py
import os
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk


class PhotoViewerWindow(tk.Toplevel):
    def __init__(self, master, project, on_update=None):
        super().__init__(master)
        self.title("Просмотр фотографий")
        self.project = project
        self.photos = project["photos"]
        self.folder = self.photos.get("folder", "")
        self.on_update = on_update

        self.files = [
            f for f in os.listdir(self.folder)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        ]
        self.files.sort()

        self.index = 0
        self.current_image = None
        self.current_pil = None

        self._build_ui()
        # хоткеи
        self.bind_all("<Left>", lambda e: self.prev_photo())
        self.bind_all("<Right>", lambda e: self.next_photo())

        # Enter: toggle include + focus caption
        self.bind_all("<Return>", self._enter_accept_hotkey)

        # Ctrl/Cmd+Space: toggle include всегда
        self.bind_all("<Control-space>", self._toggle_include_hotkey)
        self.bind_all("<Command-space>", self._toggle_include_hotkey)  # macOS

        self.bind_all("<Escape>", lambda e: self.destroy())
        self.after(50, self._load_current)

    # ---------------- UI ----------------

    def _build_ui(self):
        # разумный размер по умолчанию (не на весь экран)
        w = int(self.winfo_screenwidth() * 0.75)
        h = int(self.winfo_screenheight() * 0.80)
        self.geometry(f"{w}x{h}")
        self.minsize(800, 600)

        # grid: верх (кнопки), центр (картинка), низ (подпись/галка)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        # --- верхняя панель ---
        nav = ttk.Frame(self)
        nav.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        nav.columnconfigure(0, weight=1)
        nav.columnconfigure(1, weight=0)
        nav.columnconfigure(2, weight=0)
        nav.columnconfigure(3, weight=1)

        ttk.Button(nav, text="◀ Предыдущее", command=self.prev_photo).grid(row=0, column=1, padx=5)
        ttk.Button(nav, text="Следующее ▶", command=self.next_photo).grid(row=0, column=2, padx=5)
        self.counter_label = ttk.Label(nav, text="")
        self.counter_label.grid(row=0, column=0, sticky="w")

        # --- центральная зона под изображение ---
        self.image_frame = ttk.Frame(self)
        self.image_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.image_frame.columnconfigure(0, weight=1)
        self.image_frame.rowconfigure(0, weight=1)

        # label будет занимать всё поле и центрировать картинку
        self.image_label = ttk.Label(self.image_frame, anchor="center")
        self.image_label.grid(row=0, column=0, sticky="nsew")

        # --- нижняя панель ---
        bottom = ttk.Frame(self)
        bottom.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))
        bottom.columnconfigure(1, weight=1)

        self.filename_label = ttk.Label(bottom, text="")
        self.filename_label.grid(row=0, column=0, columnspan=2, sticky="w")

        self.include_var = tk.BooleanVar()
        self.include_cb = ttk.Checkbutton(
            bottom,
            text="Включить в итог",
            variable=self.include_var,
            command=self.on_toggle_include
        )
        self.include_cb.grid(row=1, column=0, sticky="w", pady=5)

        ttk.Label(bottom, text="Подпись:").grid(row=2, column=0, sticky="w")
        self.caption_var = tk.StringVar()
        self.caption_entry = ttk.Entry(bottom, textvariable=self.caption_var)
        self.caption_entry.grid(row=2, column=1, sticky="ew", padx=(10, 0))

        ttk.Button(bottom, text="Закрыть", command=self.destroy).grid(
            row=3, column=1, sticky="e", pady=(10, 0)
        )

        # на изменение текста подписи (у тебя это уже добавлено ранее — оставь как есть)
        self.caption_var.trace_add("write", self._on_caption_change)

        # важное: перерисовка картинки при изменении размеров окна
        self.image_frame.bind("<Configure>", self._on_resize)

    # ---------------- Logic ----------------

    def _on_caption_change(self, *_):
        # если текущее фото уже "включено" — обновляем caption в gallery
        if not self.files:
            return
        filename = self.files[self.index]

        gallery = self.photos.get("gallery", [])
        for r in gallery:
            if r.get("filename") == filename:
                r["caption"] = self.caption_var.get()
                if self.on_update:
                    self.on_update()
                return
            
    def _toggle_include_hotkey(self, event=None):
        # чтобы пробел в поле подписи печатался как пробел
        if self.focus_get() == self.caption_entry:
            return

        self.include_var.set(not self.include_var.get())
        self.on_toggle_include()
        return "break"
    
    def _enter_accept_hotkey(self, event=None):
        # Enter: включить/выключить и сразу в подпись
        self.include_var.set(not self.include_var.get())
        self.on_toggle_include()

        # если включили — фокус в подпись
        if self.include_var.get():
            self.caption_entry.state(["!disabled"])
            self.caption_entry.focus_set()
            self.caption_entry.icursor(tk.END)
        return "break"

    def _toggle_include_hotkey(self, event=None):
        # Ctrl/Cmd+Space: toggle без конфликтов с вводом пробела
        self.include_var.set(not self.include_var.get())
        self.on_toggle_include()
        return "break"

    def _focus_caption_hotkey(self, event=None):
        # Enter в поле подписи не должен прыгать заново
        if self.focus_get() == self.caption_entry:
            return

        self.caption_entry.focus_set()
        self.caption_entry.icursor(tk.END)
        return "break"
            
    def _on_resize(self, event=None):
        # если картинка уже загружена — перерисуем под новый размер
        if self.current_pil is not None:
            self._render_current_image()

    def _render_current_image(self):
        if self.current_pil is None:
            return

        # размеры области под картинку
        max_w = self.image_frame.winfo_width() - 20
        max_h = self.image_frame.winfo_height() - 20
        if max_w < 50 or max_h < 50:
            return

        img = self.current_pil.copy()
        img.thumbnail((max_w, max_h), Image.LANCZOS)

        self.current_image = ImageTk.PhotoImage(img)
        self.image_label.configure(image=self.current_image)
    
    def _load_current(self):
        if not self.files:
            self.filename_label.config(text="(В папке нет изображений)")
            self.image_label.configure(image="")
            return

        filename = self.files[self.index]
        path = os.path.join(self.folder, filename)

        # читаем исходник и сохраняем
        try:
            self.current_pil = Image.open(path)
        except Exception:
            self.current_pil = None
            self.image_label.configure(image="")
            self.filename_label.config(text=f"{filename} (не удалось открыть)")
            return

        self.filename_label.config(text=filename)
        self.counter_label.config(text=f"{self.index + 1} / {len(self.files)}")

        # отрисовать под текущий размер области
        self._render_current_image()

        # синхронизация чекбокса/подписи
        gallery = self.photos["gallery"]
        rec = next((r for r in gallery if r["filename"] == filename), None)

        if rec:
            self.include_var.set(True)
            self.caption_var.set(rec.get("caption", ""))
            self.caption_entry.state(["!disabled"])
        else:
            self.include_var.set(False)
            self.caption_var.set("")
            self.caption_entry.state(["disabled"])

    def on_toggle_include(self):
        filename = self.files[self.index]
        gallery = self.photos["gallery"]

        if self.include_var.get():
            # если уже есть — не дублируем
            rec = next((r for r in gallery if r.get("filename") == filename), None)
            if not rec:
                rec = {"filename": filename, "caption": ""}
                gallery.append(rec)

            # включили → разрешаем ввод и сохраняем текущий текст
            self.caption_entry.state(["!disabled"])
            rec["caption"] = self.caption_var.get()

        else:
            # выключили → удаляем запись
            gallery[:] = [r for r in gallery if r.get("filename") != filename]
            self.caption_entry.state(["disabled"])
            self.caption_var.set("")

        if self.on_update:
            self.on_update()

    def next_photo(self):
        if self.index < len(self.files) - 1:
            self.index += 1
            self._load_current()

    def prev_photo(self):
        if self.index > 0:
            self.index -= 1
            self._load_current()