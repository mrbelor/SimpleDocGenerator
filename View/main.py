# this file is fully vibecoded

import os
import threading
import traceback

import customtkinter as ctk
from PIL import Image
from tkinter import filedialog as fd, messagebox
from tkinterdnd2 import TkinterDnD, DND_ALL

from Controller.main_controller import MainController
from constants import (
    INVALID_TEMPLATES,
    TEMPLATE_PLACEHOLDER,
    POPUP_ERRORS,
    SHEET_ALL,
    VERSION,
)

# Поддерживаемые форматы при drag-n-drop
_DATA_EXTENSIONS = {".xls", ".xlsx", ".csv"}
_TEMPLATE_EXTENSIONS = {".docx", ".xls", ".xlsx"}
DND_SPLIT = 45  # % высоты окна: выше → данные, ниже → шаблон


class App(ctk.CTk, TkinterDnD.DnDWrapper):
    """Основное окно приложения DocGenerator."""

    def __init__(self) -> None:
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self.controller = MainController()
        self._reset_vis_job = None
        self._task = None  # поток текущей фоновой работы
        self._progress = None  # (готово, всего) от потока генерации

        appearance = self.controller.config.config.get("appearance_mode", "dark")
        ctk.set_appearance_mode(appearance)
        ctk.set_default_color_theme("blue")

        self.title(self.controller.name)
        self.geometry("500x520")
        self.minsize(500, 520)

        icon_path = self.controller.config.resource_path(os.path.join("exe_belly", "app_icon.ico"))
        self.iconbitmap(icon_path)

        self._init_icons()
        self._setup_ui()
        self._setup_overlay()
        self._setup_error_popup()

        self.drop_target_register(DND_ALL)
        self.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.dnd_bind("<<DropLeave>>", self._on_drag_leave)
        self.dnd_bind("<<Drop>>", self._on_drop)

        self.bind_all("<Button-1>", self.hide_error_popup)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Инициализация иконок
    # ------------------------------------------------------------------

    def _init_icons(self) -> None:
        def load(name: str, size=(20, 20)) -> ctk.CTkImage:
            path = self.controller.config.resource_path(os.path.join("exe_belly", "icons", f"{name}.png"))
            img = Image.open(path)
            return ctk.CTkImage(light_image=img, dark_image=img, size=size)

        self.icon_sun = load("sun")
        self.icon_moon = load("moon")
        self.icon_folder = load("folder")
        self.icon_trash = load("trash", (16, 16))
        self.icon_loading = load("loading")
        self.icon_check = load("check")

        logo_path = self.controller.config.resource_path(os.path.join("exe_belly", "logo.png"))
        logo_img = Image.open(logo_path)
        self.icon_logo = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(36, 36))

    # ------------------------------------------------------------------
    # Построение UI
    # ------------------------------------------------------------------

    def _setup_ui(self) -> None:
        self._build_header()
        self._build_main_frame()
        self._build_action_buttons()
        self.progress_bar = ctk.CTkProgressBar(self)  # виден, пока идёт фоновая работа
        self.status_label = ctk.CTkLabel(self, text="", font=("Arial", 12))
        self.status_label.pack(pady=(0, 10))
        self.update_save_info()

    def _build_header(self) -> None:
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(10, 0))

        ctk.CTkLabel(header, image=self.icon_logo, text="").pack(side="left", padx=(0, 8))
        ctk.CTkLabel(header, text="SimpleDocGenerator", font=("Arial", 15, "bold")).pack(side="left")
        ctk.CTkLabel(header, text=VERSION, font=("Arial", 10), text_color="gray").pack(side="left", padx=(6, 0), pady=(4, 0))

        mode = self.controller.config.config.get("appearance_mode", "dark")
        self.theme_btn = ctk.CTkButton(
            header,
            text="",
            image=self.icon_sun if mode == "dark" else self.icon_moon,
            width=40, height=40,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            text_color=("black", "white"),
            command=self.toggle_theme,
        )
        self.theme_btn.pack(side="right")

    def _build_main_frame(self) -> None:
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Секция ДАННЫЕ
        ctk.CTkLabel(frame, text="ДАННЫЕ (Excel / CSV)", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.data_path_label = ctk.CTkLabel(frame, text="Файл не выбран", text_color="gray", wraplength=400)
        self.data_path_label.pack(pady=5)
        ctk.CTkButton(
            frame, text=" Выбрать файл данных",
            image=self.icon_folder, compound="left",
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("black", "white"),
            command=self.select_data_file,
        ).pack(pady=5)

        self._build_sheet_controls(frame)

        self.data_separator = ctk.CTkFrame(frame, height=2, fg_color="gray30")
        self.data_separator.pack(fill="x", padx=20, pady=15)

        # Секция ШАБЛОН
        ctk.CTkLabel(frame, text="ШАБЛОН (Word)", font=("Arial", 14, "bold")).pack(pady=5)
        self._build_template_controls(frame)

        ctk.CTkButton(
            frame, text=" Открыть папку шаблонов",
            image=self.icon_folder, compound="left",
            width=200, height=28,
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("black", "white"),
            command=self.controller.open_templates_folder,
        ).pack(pady=(5, 0))

    def _build_sheet_controls(self, parent) -> None:
        """Выбор листа. Строка скрыта, пока в файле не окажется больше одного листа."""
        self.sheet_row = ctk.CTkFrame(parent, fg_color="transparent")

        ctk.CTkLabel(self.sheet_row, text="Лист:").pack(side="left", padx=(0, 5))

        self.sheet_var = ctk.StringVar(value=SHEET_ALL)
        self.sheet_dropdown = ctk.CTkOptionMenu(
            self.sheet_row,
            values=[SHEET_ALL],
            variable=self.sheet_var,
            width=160,
            text_color=("black", "white"),
            dynamic_resizing=True,
            command=self.on_sheet_change,
        )
        self.sheet_dropdown.pack(side="left")

    def _build_template_controls(self, parent) -> None:
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(pady=5)

        self.template_var = ctk.StringVar()
        self.template_var.trace_add("write", self.update_save_info)
        templates = self.controller.get_templates()

        self.template_dropdown = ctk.CTkOptionMenu(
            row,
            values=templates or [TEMPLATE_PLACEHOLDER],
            variable=self.template_var,
            width=125,
            text_color=("black", "white"),
            dynamic_resizing=True,
        )
        self.template_dropdown.pack(side="left", padx=(0, 5))
        self.template_var.set(templates[0] if templates else TEMPLATE_PLACEHOLDER)

        self.template_dropdown.bind("<Enter>", lambda _: self.update_templates_list())
        self.template_dropdown.bind("<Button-1>", lambda _: self.update_templates_list(), add="+")

        add_btn = ctk.CTkButton(
            row, text="+", width=30, height=30,
            font=("Arial", 16, "bold"), anchor="center",
            fg_color=("gray75", "gray25"), hover_color=("gray65", "gray35"),
            text_color=("black", "white"),
            command=self.select_template_file,
        )
        add_btn.pack(side="left", padx=2)
        add_btn.bind("<Enter>", lambda _: self.show_status("Добавить новый шаблон", "gray"))
        add_btn.bind("<Leave>", lambda _: self.show_status("", "gray"))

        del_btn = ctk.CTkButton(
            row, text="", image=self.icon_trash, width=30, height=30,
            fg_color="#c0392b", hover_color="#e74c3c", anchor="center",
            command=self.delete_template,
        )
        del_btn.pack(side="left", padx=2)
        del_btn.bind("<Enter>", lambda _: self.show_status("Удалить текущий шаблон", "red"))
        del_btn.bind("<Leave>", lambda _: self.show_status("", "gray"))

    def _build_action_buttons(self) -> None:
        outer = ctk.CTkFrame(self, fg_color="transparent")
        outer.pack(fill="x", padx=40, pady=(10, 0))

        left_col = ctk.CTkFrame(outer, fg_color="transparent")
        left_col.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.gen_btn = ctk.CTkButton(
            left_col, text="Сформировать", height=45,
            font=("Arial", 14, "bold"), fg_color="#27ae60", hover_color="#2ecc71",
            text_color=("black", "white"),
            command=self.generate,
        )
        self.gen_btn.pack(fill="x")
        self.gen_btn.bind("<Enter>", lambda _: self.show_status("Сохранить результат в текущую папку", "gray"))
        self.gen_btn.bind("<Leave>", lambda _: self.show_status("", "gray"))

        self.save_info_label = ctk.CTkLabel(
            left_col, text=f"Папка: {self.controller.get_save_folder_name()}",
            font=("Arial", 12), text_color="gray", height=12,
        )
        self.save_info_label.pack(fill="x")

        self.save_name_label = ctk.CTkLabel(
            left_col, text="Имя: не определено",
            font=("Arial", 12), text_color="gray", height=12,
        )
        self.save_name_label.pack(fill="x")

        self.gen_as_btn = ctk.CTkButton(
            outer, text="Сформировать как...", height=45,
            font=("Arial", 14, "bold"), fg_color="#2980b9", hover_color="#3498db",
            text_color=("black", "white"),
            command=self.generate_as,
        )
        self.gen_as_btn.pack(side="left", expand=True, fill="x", padx=(5, 0), anchor="n")
        self.gen_as_btn.bind("<Enter>", lambda _: self.show_status("Выбрать имя и папку для сохранения", "gray"))
        self.gen_as_btn.bind("<Leave>", lambda _: self.show_status("", "gray"))

    def _setup_overlay(self) -> None:
        self.overlay_frame = ctk.CTkFrame(self, fg_color="transparent")

        top_h = (DND_SPLIT - 2) / 100
        bot_y = (DND_SPLIT + 2) / 100
        bot_h = (96 - DND_SPLIT) / 100

        ov_top = ctk.CTkFrame(
            self.overlay_frame, fg_color=("gray95", "gray5"),
            border_width=4, border_color="#3498db", corner_radius=16,
        )
        ctk.CTkLabel(
            ov_top, text="Данные (Excel / CSV)",
            font=("Arial", 20, "bold"), text_color="#3498db",
        ).place(relx=0.5, rely=0.5, anchor="center")
        ov_top.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=top_h)

        ov_bot = ctk.CTkFrame(
            self.overlay_frame, fg_color=("gray95", "gray5"),
            border_width=4, border_color="#e67e22", corner_radius=16,
        )
        ctk.CTkLabel(
            ov_bot, text="Шаблон (Word / Excel)",
            font=("Arial", 20, "bold"), text_color="#e67e22",
        ).place(relx=0.5, rely=0.5, anchor="center")
        ov_bot.place(relx=0.02, rely=bot_y, relwidth=0.96, relheight=bot_h)

    def _setup_error_popup(self) -> None:
        self.error_frame = ctk.CTkFrame(self, fg_color="#e74c3c", corner_radius=0)
        self.error_label = ctk.CTkLabel(
            self.error_frame, text="Неверный формат файла!",
            text_color="white", font=("Arial", 14, "bold"),
        )
        self.error_label.pack(padx=20, pady=15)

    # ------------------------------------------------------------------
    # Ошибки и статус
    # ------------------------------------------------------------------

    def show_error_popup(self, message: str = "Неверный формат файла!") -> None:
        self.error_label.configure(text=message)
        self.error_frame.place(relx=0.5, rely=0.5, anchor="center")
        self.error_frame.lift()

    def hide_error_popup(self, _=None) -> None:
        if hasattr(self, "error_frame") and self.error_frame.winfo_ismapped():
            self.error_frame.place_forget()

    def show_status(self, text: str, color: str) -> None:
        color_map = {
            "red": "#e74c3c",
            "green": "#2ecc71",
            "orange": "#f39c12",
            "gray": ("gray60", "gray40"),
        }
        self.status_label.configure(text=text, text_color=color_map.get(color, ("black", "white")))

    # ------------------------------------------------------------------
    # Drag-n-drop
    # ------------------------------------------------------------------

    def _on_drag_enter(self, event) -> None:
        self.hide_error_popup()
        self.overlay_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.overlay_frame.lift()
        self.update_idletasks()

    def _on_drag_leave(self, event) -> None:
        self.overlay_frame.place_forget()

    def _on_drop(self, event) -> None:
        self.overlay_frame.place_forget()
        self.focus_force()
        self.hide_error_popup()

        path = event.data.strip("{}")
        ext = os.path.splitext(path)[1].lower()

        rel_y = event.y_root - self.winfo_rooty()
        drop_top = rel_y < self.winfo_height() * DND_SPLIT / 100

        if drop_top:
            if ext in _DATA_EXTENSIONS:
                self.process_data_file(path)
            else:
                self.show_error_popup()
        else:
            if ext in _TEMPLATE_EXTENSIONS:
                self.process_template_file(path)
            else:
                self.show_error_popup()

    # ------------------------------------------------------------------
    # Тема
    # ------------------------------------------------------------------

    def toggle_theme(self) -> None:
        if ctk.get_appearance_mode() == "Dark":
            new_mode = "light"
            self.theme_btn.configure(image=self.icon_moon)
        else:
            new_mode = "dark"
            self.theme_btn.configure(image=self.icon_sun)
        ctk.set_appearance_mode(new_mode)
        self.controller.set_appearance_mode(new_mode)

    # ------------------------------------------------------------------
    # Данные
    # ------------------------------------------------------------------

    def select_data_file(self) -> None:
        self.hide_error_popup()
        path = fd.askopenfilename(
            initialdir=self.controller.get_last_data_folder(),
            filetypes=[("Данные", "*.xlsx *.xls *.csv")],
        )
        if path:
            self.process_data_file(path)

    def process_data_file(self, path: str) -> None:
        self._in_background(lambda: self.controller.load_source(path), self._data_loaded, f"Читаю {os.path.basename(path)}")

    def _data_loaded(self, success: bool, msg: str) -> None:
        if success:
            self._flash_data_label(msg)
            self.update_sheets_list()
            self.update_save_info()
        else:
            self.show_status(msg, "red")

    def _flash_data_label(self, msg: str) -> None:
        self.data_path_label.configure(text=msg, text_color="#2ecc71")
        self.after(2000, lambda: self.data_path_label.configure(text_color=("black", "white")))

    def update_sheets_list(self) -> None:
        """Показывает выбор листа только когда листов больше одного."""
        sheets = self.controller.get_sheets()
        if len(sheets) > 1:
            self.sheet_dropdown.configure(values=[SHEET_ALL] + sheets)
            self.sheet_var.set(SHEET_ALL)
            self.sheet_row.pack(pady=(0, 5), before=self.data_separator)
        else:
            self.sheet_row.pack_forget()

    def on_sheet_change(self, choice: str) -> None:
        """Переключает лист без перечитывания файла."""
        self._flash_data_label(self.controller.select_sheet(None if choice == SHEET_ALL else choice))

    # ------------------------------------------------------------------
    # Шаблоны
    # ------------------------------------------------------------------

    def select_template_file(self) -> None:
        self.hide_error_popup()
        path = fd.askopenfilename(
            initialdir=self.controller.get_last_template_folder(),
            filetypes=[("Шаблон", "*.docx *.xls *.xlsx")],
        )
        if path:
            self.process_template_file(path)

    def process_template_file(self, path: str) -> None:
        success, name = self.controller.add_template(path)
        if success:
            self.update_templates_list(select_name=name)

    def delete_template(self) -> None:
        self.hide_error_popup()
        name = self.template_var.get()
        if name and name not in INVALID_TEMPLATES:
            if self.controller.remove_template(name):
                self.update_templates_list(select_name=TEMPLATE_PLACEHOLDER)
                self.show_status(f"Удалено: {name}", "orange")

    def update_templates_list(self, select_name: str | None = None) -> None:
        templates = self.controller.get_templates()
        new_values = templates or [TEMPLATE_PLACEHOLDER]

        if list(self.template_dropdown.cget("values")) != new_values:
            self.template_dropdown.configure(values=new_values)

        current = self.template_var.get()
        if select_name:
            self.template_var.set(select_name)
        elif current not in new_values:
            self.template_var.set(TEMPLATE_PLACEHOLDER)

    # ------------------------------------------------------------------
    # Генерация
    # ------------------------------------------------------------------

    def _is_template_selected(self) -> bool:
        return self.template_var.get() not in INVALID_TEMPLATES

    def generate(self) -> None:
        self.hide_error_popup()
        if not self._is_template_selected():
            self.show_error_popup("Шаблон не выбран!")
            return
        self.update_templates_list()
        if not self._is_template_selected():
            self.show_error_popup("Шаблон не выбран!")
            return
        self._generate(self.template_var.get())

    def generate_as(self) -> None:
        self.hide_error_popup()
        if not self._is_template_selected():
            self.show_error_popup("Шаблон не выбран!")
            return
        self.update_templates_list()
        template = self.template_var.get()
        if not template or template in INVALID_TEMPLATES:
            self.show_error_popup("Шаблон не выбран!")
            return

        tmpl_suffix = os.path.splitext(template)[1].lower()
        if tmpl_suffix in (".xls", ".xlsx"):
            out_ext, filetypes = tmpl_suffix, [("Excel", f"*{tmpl_suffix}")]  # .xls даёт .xls
        else:
            out_ext, filetypes = ".docx", [("Word Document", "*.docx")]

        path = fd.asksaveasfilename(
            initialdir=self.controller.get_last_output_folder(),
            defaultextension=out_ext,
            initialfile=self.controller.get_suggested_filename(template),
            filetypes=filetypes,
        )
        if path:
            self._generate(template, path)

    def _generate(self, template: str, path: str | None = None) -> None:
        work = lambda: self.controller.generate_document(template, custom_path=path, progress=self._set_progress)
        name = os.path.basename(path) if path else self.controller.get_output_name()
        if self._in_background(work, self._finish_generation, f"Формирую {name}"):
            self._start_loading()

    def _set_progress(self, done: int, total: int) -> None:
        """Зовётся из потока генерации: только запоминаем, рисует _poll_task."""
        self._progress = (done, total)

    def _finish_generation(self, success: bool, msg: str) -> None:
        if success:
            self._show_success()
        else:
            self._stop_loading()
            if msg in POPUP_ERRORS:
                self.show_error_popup(msg)
            else:
                self.show_status(msg, "red")
        self.update_save_info()

    # ------------------------------------------------------------------
    # Фоновая работа
    # ------------------------------------------------------------------

    def _in_background(self, work, done, text: str) -> bool:
        """work() в отдельном потоке, done(*результат) в главном, когда поток закончит.

        Окно не замирает. Виджеты трогает только главный поток (tkinter не потокобезопасен):
        поток кладёт результат и self._progress, главный забирает их опросом в _poll_task.
        Задача одна за раз: пока идёт прошлая, новая не стартует (возвращаем False).
        """
        if self._task and self._task.is_alive():
            self.show_status("Подождите, идёт обработка...", "orange")
            return False
        box = []

        def target():
            try:
                box.append(work())
            except Exception as e:  # иначе окно навсегда останется в «Загрузке»
                traceback.print_exc()
                box.append((False, f"Ошибка: {e}"))

        self._progress = None
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        # полоска у нижнего края окна, под статусом; place не двигает остальную вёрстку
        self.progress_bar.place(relx=0.5, rely=1, y=-2, anchor="s", relwidth=0.84)
        self._task = threading.Thread(target=target, daemon=True)
        self._task.start()
        # первый опрос - после выхода из обработчика: даже мгновенная работа
        # завершится (done) позже того, что вызывающий сделает после старта
        self._task_text = text  # для вопроса при закрытии окна
        self.after(0, lambda: self._poll_task(box, done, text))
        return True

    def _poll_task(self, box: list, done, text: str) -> None:
        if not self._task.is_alive():
            self.progress_bar.stop()
            self.progress_bar.place_forget()
            self.show_status("", "gray")
            done(*box[0])
            return
        status = f"{text}..."
        if self._progress:  # процент знает генерация; чтение файла крутится без него
            ready, total = self._progress
            if ready < total:
                self.progress_bar.stop()
                self.progress_bar.configure(mode="determinate")
                self.progress_bar.set(ready / total)
                status = f"{text}: {ready} из {total}"
            else:  # записи готовы, файл пишется на диск - сколько осталось, неизвестно
                self.progress_bar.configure(mode="indeterminate")
                self.progress_bar.start()
                status = "Сохраняю файл..."
        self.show_status(status, "gray")
        self.after(100, lambda: self._poll_task(box, done, text))

    def _on_close(self) -> None:
        """Закрытие окна убивает фоновый поток, файл не допишется - поэтому спрашиваем."""
        if self._task and self._task.is_alive() and not messagebox.askyesno(
                "Идёт обработка", f"{self._task_text}...\n\nЗакрыть программу? Работа прервётся, файл не сохранится.",
                icon="warning"):
            return
        self.destroy()

    # ------------------------------------------------------------------
    # Анимация кнопок
    # ------------------------------------------------------------------

    def _start_loading(self) -> None:
        self._cancel_reset_job()
        self.gen_btn.configure(text=" Загрузка...", image=self.icon_loading, compound="left", state="disabled")
        self.gen_as_btn.configure(state="disabled")

    def _stop_loading(self) -> None:
        self._cancel_reset_job()
        self.gen_btn.configure(text="Сформировать", image=None, state="normal")
        self.gen_as_btn.configure(state="normal")

    def _show_success(self) -> None:
        self.gen_btn.configure(text="", image=self.icon_check, state="disabled")
        self.after(1500, lambda: self.gen_btn.configure(state="normal"))
        self.after(1500, lambda: self.gen_as_btn.configure(state="normal"))

        def reset_vis():
            if self.gen_btn.cget("text") == "":
                self.gen_btn.configure(text="Сформировать", image=None)

        self._reset_vis_job = self.after(2000, reset_vis)
        self.show_status("Готово!", "green")

    def _cancel_reset_job(self) -> None:
        if self._reset_vis_job:
            self.after_cancel(self._reset_vis_job)
            self._reset_vis_job = None

    # ------------------------------------------------------------------
    # Информация о сохранении
    # ------------------------------------------------------------------

    def update_save_info(self, *_) -> None:
        if not hasattr(self, "save_info_label"):
            return
        self.save_info_label.configure(text=f"Папка: {self.controller.get_save_folder_name()}")
        name = self.controller.get_output_name()
        self.save_name_label.configure(text=f"Имя: {name or 'не определено'}")
