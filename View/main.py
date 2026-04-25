import os
import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_ALL
from tkinter import filedialog as fd
from Controller.main_controller import MainController

class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)
        self.controller = MainController()
        
        self.appearance_mode = self.controller.config.config.get("appearance_mode", "dark")
        ctk.set_appearance_mode(self.appearance_mode)
        ctk.set_default_color_theme("blue")

        self.title(self.controller.name)
        self.geometry("500x520")
        self.minsize(500, 520)
        
        self._setup_ui()
        self._setup_overlay()
        self._setup_error_popup()
        
        self.drop_target_register(DND_ALL)
        self.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.dnd_bind("<<DropLeave>>", self._on_drag_leave)
        self.dnd_bind("<<Drop>>", self.handle_drop)

        # Любой клик в окне скрывает ошибку
        self.bind_all("<Button-1>", self.hide_error_popup)

    def _setup_ui(self):
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        # Кнопка переключения темы
        self.theme_btn = ctk.CTkButton(
            self.header_frame, 
            text="☀️" if self.appearance_mode == "dark" else "🌙", 
            width=40, 
            height=40,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            text_color=("black", "white"),
            command=self.toggle_theme
        )
        self.theme_btn.pack(side="right")

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # ДАННЫЕ
        ctk.CTkLabel(self.main_frame, text="ДАННЫЕ (Excel / CSV)", font=("Arial", 14, "bold")).pack(pady=(10, 5))
        self.data_path_label = ctk.CTkLabel(self.main_frame, text="Файл не выбран", text_color="gray", wraplength=400)
        self.data_path_label.pack(pady=5)
        
        self.select_data_btn = ctk.CTkButton(
            self.main_frame, text="📁 Выбрать файл данных", 
            fg_color=("gray75", "gray25"),
            hover_color=("gray65", "gray35"),
            text_color=("black", "white"),
            command=self.select_data_file
        )
        self.select_data_btn.pack(pady=5)

        ctk.CTkFrame(self.main_frame, height=2, fg_color="gray30").pack(fill="x", padx=20, pady=15)

        # ШАБЛОН
        ctk.CTkLabel(self.main_frame, text="ШАБЛОН (Word)", font=("Arial", 14, "bold")).pack(pady=5)
        
        template_controls = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        template_controls.pack(pady=5)
        
        self.template_var = ctk.StringVar()
        templates = self.controller.get_templates()
        self.template_dropdown = ctk.CTkOptionMenu(
            template_controls, 
            values=templates if templates else ["Шаблон не выбран"],
            variable=self.template_var,
            width=125,
            text_color=("black", "white"),
            dynamic_resizing=True
        )
        self.template_dropdown.pack(side="left", padx=(0, 5))
        self.template_var.set(templates[0] if templates else "Шаблон не выбран")
        
        # Кнопка Плюс
        self.add_temp_btn = ctk.CTkButton(
            template_controls, text="+", width=30, height=30,
            fg_color=("gray75", "gray25"),
            hover_color=("gray65", "gray35"),
            text_color=("black", "white"),
            command=self.select_template_file
        )
        self.add_temp_btn.pack(side="left", padx=2)
        self.add_temp_btn.bind("<Enter>", lambda e: self.show_status("Добавить новый шаблон", "gray"))
        self.add_temp_btn.bind("<Leave>", lambda e: self.show_status("", "gray"))

        # Кнопка Корзина
        self.delete_btn = ctk.CTkButton(
            template_controls, text="🗑️", width=30, height=30,
            fg_color="#c0392b", hover_color="#e74c3c",
            command=self.delete_template
        )
        self.delete_btn.pack(side="left", padx=2)
        self.delete_btn.bind("<Enter>", lambda e: self.show_status("Удалить текущий шаблон", "red"))
        self.delete_btn.bind("<Leave>", lambda e: self.show_status("", "gray"))

        self.folder_btn = ctk.CTkButton(
            self.main_frame, text="📂 Открыть папку шаблонов", 
            width=200, height=28, 
            fg_color=("gray75", "gray25"),
            hover_color=("gray65", "gray35"),
            text_color=("black", "white"),
            command=self.controller.open_templates_folder
        )
        self.folder_btn.pack(pady=(5, 0))

        # КНОПКИ СФОРМИРОВАТЬ
        btns_outer_frame = ctk.CTkFrame(self, fg_color="transparent")
        btns_outer_frame.pack(fill="x", padx=40, pady=(10, 0))

        left_column = ctk.CTkFrame(btns_outer_frame, fg_color="transparent")
        left_column.pack(side="left", expand=True, fill="x", padx=(0, 5))

        self.gen_btn = ctk.CTkButton(
            left_column, text="Сформировать", height=45,
            font=("Arial", 14, "bold"), fg_color="#27ae60", hover_color="#2ecc71",
            text_color=("black", "white"),
            command=self.generate
        )
        self.gen_btn.pack(fill="x")
        self.gen_btn.bind("<Enter>", lambda e: self.show_status("Сохранить результат в текущую папку", "gray"))
        self.gen_btn.bind("<Leave>", lambda e: self.show_status("", "gray"))

        self.save_info_label = ctk.CTkLabel(
            left_column, text=f"Папка: {self.controller.get_save_folder_name()}",
            font=("Arial", 10), text_color="gray"
        )
        self.save_info_label.pack(fill="x", pady=(2, 0))

        self.gen_as_btn = ctk.CTkButton(
            btns_outer_frame, text="Сформировать как...", height=45,
            font=("Arial", 14, "bold"), fg_color="#2980b9", hover_color="#3498db",
            text_color=("black", "white"),
            command=self.generate_as
        )
        self.gen_as_btn.pack(side="left", expand=True, fill="x", padx=(5, 0), anchor="n")
        self.gen_as_btn.bind("<Enter>", lambda e: self.show_status("Выбрать имя и папку для сохранения", "gray"))
        self.gen_as_btn.bind("<Leave>", lambda e: self.show_status("", "gray"))

        self.status_label = ctk.CTkLabel(self, text="", font=("Arial", 12))
        self.status_label.pack(pady=(0, 10))

    def _setup_overlay(self):
        self.overlay_frame = ctk.CTkFrame(
            self, fg_color=("gray95", "gray5"), border_width=5, 
            border_color="#3498db", corner_radius=20
        )
        self.overlay_label = ctk.CTkLabel(
            self.overlay_frame, text="Переместите свои файлы сюда\n(Excel, CSV или Word)", 
            font=("Arial", 22, "bold"), text_color="#3498db"
        )
        self.overlay_label.place(relx=0.5, rely=0.5, anchor="center")

    def _setup_error_popup(self):
        self.error_frame = ctk.CTkFrame(self, fg_color="#e74c3c", corner_radius=0)
        self.error_label = ctk.CTkLabel(
            self.error_frame, text="Неверный формат файла!", 
            text_color="white", font=("Arial", 14, "bold")
        )
        self.error_label.pack(padx=20, pady=15)

    def show_error_popup(self):
        self.error_frame.place(relx=0.5, rely=0.5, anchor="center")
        self.error_frame.lift()

    def hide_error_popup(self, _=None):
        if hasattr(self, 'error_frame') and self.error_frame.winfo_ismapped():
            self.error_frame.place_forget()

    def _on_drag_enter(self, event):
        self.hide_error_popup()
        self.overlay_frame.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)
        self.overlay_frame.lift()
        self.update_idletasks()

    def _on_drag_leave(self, event):
        self.overlay_frame.place_forget()

    def handle_drop(self, event):
        self.overlay_frame.place_forget()
        self.focus_force()
        self.hide_error_popup()
        
        path = event.data.strip("{}")
        ext = os.path.splitext(path)[1].lower()
        if ext in ['.xls', '.xlsx', '.csv']: 
            self.process_data_file(path)
        elif ext == '.docx': 
            self.process_template_file(path)
        else:
            self.show_error_popup()

    def toggle_theme(self):
        if ctk.get_appearance_mode() == "Dark":
            new_mode = "light"
            self.theme_btn.configure(text="🌙")
        else:
            new_mode = "dark"
            self.theme_btn.configure(text="☀️")
            
        ctk.set_appearance_mode(new_mode)
        self.controller.set_appearance_mode(new_mode)

    def select_data_file(self):
        self.hide_error_popup()
        path = fd.askopenfilename(
            initialdir=self.controller.get_last_data_folder(),
            filetypes=[("Данные", "*.xlsx *.xls *.csv")]
        )
        if path: self.process_data_file(path)

    def process_data_file(self, path):
        success, msg = self.controller.load_source(path)
        if success:
            self.data_path_label.configure(text=msg, text_color="#2ecc71")
            self.after(2000, lambda: self.data_path_label.configure(text_color="gray"))
        else:
            self.show_status(msg, "red")

    def select_template_file(self):
        self.hide_error_popup()
        path = fd.askopenfilename(
            initialdir=self.controller.get_last_template_folder(),
            filetypes=[("Шаблон Word", "*.docx *.doc")]
        )
        if path: self.process_template_file(path)

    def process_template_file(self, path):
        success, name = self.controller.add_template(path)
        if success: self.update_templates_list(select_name=name)

    def delete_template(self):
        self.hide_error_popup()
        name = self.template_var.get()
        if name and name not in ["Шаблон не выбран", "Файл не найден"]:
            if self.controller.remove_template(name):
                self.update_templates_list(select_name="Шаблон не выбран")
                self.show_status(f"Удалено: {name}", "orange")

    def update_templates_list(self, select_name=None):
        templates = self.controller.get_templates()
        new_values = templates if templates else ["Шаблон не выбран"]
        
        # Обновляем только если список реально изменился, 
        # чтобы не "сбрасывать" состояние виджета во время клика
        if list(self.template_dropdown.cget("values")) != new_values:
            self.template_dropdown.configure(values=new_values)
        
        current = self.template_var.get()
        
        if select_name:
            self.template_var.set(select_name)
        elif current not in new_values:
            # Если текущего файла больше нет в списке
            if templates:
                self.template_var.set(templates[0])
            else:
                self.template_var.set("Шаблон не выбран")

    def generate(self):
        self.hide_error_popup()
        self.update_templates_list()
        if self.template_var.get() in ["Шаблон не выбран", "Файл не найден"]:
            self.show_status("Ошибка: Выберите существующий шаблон!", "red")
            return
        self._start_loading()
        self.after(100, self._run_generate)

    def _run_generate(self):
        success, msg = self.controller.generate_document(self.template_var.get())
        if success: self._show_success()
        else: self._stop_loading(); self.show_status(msg, "red")
        self.update_save_info()

    def generate_as(self):
        self.hide_error_popup()
        self.update_templates_list()
        template = self.template_var.get()
        if not template or template in ["Шаблон не выбран", "Файл не найден"]:
            self.show_status("Сначала выберите существующий шаблон!", "red")
            return
            
        path = fd.asksaveasfilename(
            initialdir=self.controller.get_last_output_folder(),
            defaultextension=".docx",
            initialfile=f"Результат_{template}",
            filetypes=[("Word Document", "*.docx")]
        )
        if path:
            self._start_loading()
            self.after(100, lambda: self._run_generate_as(template, path))

    def _run_generate_as(self, template, path):
        success, msg = self.controller.generate_document(template, custom_path=path)
        if success: self._show_success()
        else: self._stop_loading(); self.show_status(msg, "red")
        self.update_save_info()

    def _start_loading(self):
        self.gen_btn.configure(text="⌛ Загрузка...", state="disabled")
        self.gen_as_btn.configure(state="disabled")

    def _stop_loading(self):
        self.gen_btn.configure(text="Сформировать", state="normal")
        self.gen_as_btn.configure(state="normal")

    def _show_success(self):
        self.gen_btn.configure(text="✅", state="disabled")
        self.after(3000, self._stop_loading)
        self.show_status("Готово!", "green")

    def update_save_info(self):
        self.save_info_label.configure(text=f"Папка: {self.controller.get_save_folder_name()}")

    def show_status(self, text, color):
        colors = {"red": "#e74c3c", "green": "#2ecc71", "orange": "#f39c12", "gray": "gray"}
        self.status_label.configure(text=text, text_color=colors.get(color, "gray"))

if __name__ == "__main__":
    app = App()
    app.mainloop()
