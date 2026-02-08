"""
GUI приложение для валидации email адресов — Валидатор
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from datetime import datetime
from email_validator import process_excel_file_advanced
import logging
import math

# Настройка логирования для GUI
logging.basicConfig(level=logging.WARNING)


class RoundedButton(tk.Canvas):
    """Кастомная кнопка со скругленными углами"""
    def __init__(self, parent, text, command, width=200, height=40, 
                 bg_color="#1976D2", hover_color="#1565C0", text_color="white",
                 corner_radius=20, font=("Arial", 11, "bold")):
        super().__init__(parent, width=width, height=height,
                        highlightthickness=0, relief=tk.FLAT, bg="#F8FAFC")
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.text_color = text_color
        self.corner_radius = corner_radius
        self.font = font
        self.text = text
        self.enabled = True
        
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        
        self.draw_button()
        
    def draw_button(self, color=None):
        """Отрисовка кнопки"""
        if color is None:
            color = self.bg_color
            
        self.delete("all")
        
        # Рисуем скругленный прямоугольник
        self.create_rounded_rectangle(0, 0, self.winfo_reqwidth(), 
                                     self.winfo_reqheight(), 
                                     radius=self.corner_radius,
                                     fill=color, outline=color)
        
        # Текст
        self.create_text(self.winfo_reqwidth() // 2,
                        self.winfo_reqheight() // 2,
                        text=self.text,
                        fill=self.text_color,
                        font=self.font)
    
    def create_rounded_rectangle(self, x1, y1, x2, y2, radius=20, **kwargs):
        """Создание скругленного прямоугольника"""
        points = []
        # Верхний левый угол
        points.extend([x1 + radius, y1])
        points.extend([x2 - radius, y1])
        # Верхний правый угол
        points.extend([x2, y1])
        points.extend([x2, y1 + radius])
        # Нижний правый угол
        points.extend([x2, y2 - radius])
        points.extend([x2, y2])
        points.extend([x2 - radius, y2])
        # Нижний левый угол
        points.extend([x1 + radius, y2])
        points.extend([x1, y2])
        points.extend([x1, y2 - radius])
        # Верхний левый угол
        points.extend([x1, y1 + radius])
        points.extend([x1, y1])
        
        return self.create_polygon(points, smooth=True, **kwargs)
    
    def _on_click(self, event):
        if self.command and self.enabled:
            self.command()
    
    def _on_enter(self, event):
        if self.enabled:
            self.draw_button(self.hover_color)
    
    def _on_leave(self, event):
        if self.enabled:
            self.draw_button(self.bg_color)
    
    def disable(self):
        """Отключить кнопку"""
        self.enabled = False
        self.draw_button("#9E9E9E")
        self.unbind("<Enter>")
        self.unbind("<Leave>")
    
    def enable(self, bg_color=None):
        """Включить кнопку"""
        self.enabled = True
        if bg_color:
            self.bg_color = bg_color
        self.draw_button(self.bg_color)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)


class EmailValidatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Валидатор V3")
        self.root.geometry("880x820")
        self.root.minsize(600, 620)
        self.root.resizable(True, True)
        
        # Цветовая схема: спокойные тона
        self.colors = {
            'primary': '#2563EB',
            'secondary': '#3B82F6',
            'accent': '#60A5FA',
            'light': '#F8FAFC',
            'panel': '#F1F5F9',
            'dark': '#1E293B',
            'success': '#22C55E',
            'warning': '#F59E0B',
            'error': '#EF4444',
        }
        
        self.root.configure(bg=self.colors['light'])
        
        # Переменные
        self.input_files = []  # список путей к файлам
        self.output_file = tk.StringVar()
        self.validation_mode = tk.StringVar(value="strict")
        self.max_emails = tk.StringVar()
        self.timeout = tk.StringVar(value="10")
        # Опции результата
        self.include_full_results_sheet = tk.BooleanVar(value=True)
        self.only_valid_emails_sheet = tk.BooleanVar(value=False)
        # Результат обработки
        self.result_files = []
        self.is_processing = False
        self.stop_flag = {}
        
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок: компактный, без триколора
        header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=72, pady=16)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="Валидатор V3",
            font=("Segoe UI", 24, "bold"),
            bg=self.colors['primary'],
            fg="white",
        )
        title_label.pack(expand=True)
        
        # Основной контейнер со скроллом
        canvas = tk.Canvas(self.root, bg=self.colors['light'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=canvas.yview)
        
        main_frame = tk.Frame(canvas, bg=self.colors['light'])
        self._canvas = canvas
        self._main_frame = main_frame
        
        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def _on_canvas_configure(e):
            canvas.itemconfig(self._canvas_window_id, width=e.width)
        
        main_frame.bind("<Configure>", _on_frame_configure)
        self._canvas_window_id = canvas.create_window((0, 0), window=main_frame, anchor=tk.NW)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(20, 0), pady=(16, 16))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=(16, 16))
        
        # Фрейм для загрузки файлов
        file_frame = tk.LabelFrame(
            main_frame,
            text="📁 Файлы с email адресами (можно несколько)",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            padx=12,
            pady=12,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground="#E2E8F0"
        )
        file_frame.pack(fill=tk.X, pady=(0, 12))
        
        file_list_frame = tk.Frame(file_frame, bg=self.colors['panel'])
        file_list_frame.pack(fill=tk.BOTH, expand=True)
        
        list_scroll = tk.Scrollbar(file_list_frame)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            file_list_frame,
            height=4,
            font=("Segoe UI", 9),
            selectmode=tk.EXTENDED,
            yscrollcommand=list_scroll.set,
            bg="white",
            fg=self.colors['dark']
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        list_scroll.config(command=self.file_listbox.yview)
        
        btn_frame = tk.Frame(file_frame, bg=self.colors['panel'])
        btn_frame.pack(fill=tk.X, pady=(8, 0))
        
        select_btn = RoundedButton(
            btn_frame,
            text="Выбрать файлы",
            command=self.select_input_files,
            width=150,
            height=35,
            bg_color=self.colors['secondary'],
            hover_color=self.colors['primary'],
            corner_radius=15
        )
        select_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_btn = RoundedButton(
            btn_frame,
            text="Очистить список",
            command=self.clear_input_files,
            width=130,
            height=35,
            bg_color=self.colors['accent'],
            hover_color=self.colors['secondary'],
            corner_radius=15
        )
        clear_btn.pack(side=tk.LEFT)
        
        # Фрейм настроек
        settings_frame = tk.LabelFrame(
            main_frame,
            text="⚙️ Настройки проверки",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            padx=12,
            pady=12,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground="#E2E8F0"
        )
        settings_frame.pack(fill=tk.X, pady=(0, 12))
        
        mode_label = tk.Label(
            settings_frame,
            text="Режим валидации:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        mode_label.grid(row=0, column=0, sticky=tk.W, pady=8)
        
        mode_frame = tk.Frame(settings_frame, bg=self.colors['panel'])
        mode_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=10)
        
        strict_radio = tk.Radiobutton(
            mode_frame,
            text="🔒 Строгий (максимальная точность)",
            variable=self.validation_mode,
            value="strict",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            selectcolor=self.colors['panel'],
            activebackground=self.colors['panel'],
            activeforeground=self.colors['primary']
        )
        strict_radio.pack(side=tk.LEFT, padx=10)
        
        lenient_radio = tk.Radiobutton(
            mode_frame,
            text="✨ Лояльный (+15-20% валидных)",
            variable=self.validation_mode,
            value="lenient",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            selectcolor=self.colors['panel'],
            activebackground=self.colors['panel'],
            activeforeground=self.colors['accent']
        )
        lenient_radio.pack(side=tk.LEFT, padx=10)
        
        opts_label = tk.Label(
            settings_frame,
            text="Формат результата:",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        opts_label.grid(row=1, column=0, sticky=tk.W, pady=(12, 4))
        
        full_sheet_check = tk.Checkbutton(
            settings_frame,
            text="✓ Добавить лист с полными результатами проверки",
            variable=self.include_full_results_sheet,
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            selectcolor=self.colors['panel'],
            activebackground=self.colors['panel'],
            activeforeground=self.colors['primary']
        )
        full_sheet_check.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        only_valid_check = tk.Checkbutton(
            settings_frame,
            text="✓ Получить только список валидных почт (доп. лист)",
            variable=self.only_valid_emails_sheet,
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            selectcolor=self.colors['panel'],
            activebackground=self.colors['panel'],
            activeforeground=self.colors['primary']
        )
        only_valid_check.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        params_frame = tk.Frame(settings_frame, bg=self.colors['panel'])
        params_frame.grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        timeout_label = tk.Label(
            params_frame,
            text="Таймаут (сек):",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        timeout_label.pack(side=tk.LEFT, padx=(0, 5))
        
        timeout_entry = tk.Entry(
            params_frame,
            textvariable=self.timeout,
            width=10,
            font=("Segoe UI", 10),
            relief=tk.SOLID,
            bd=1
        )
        timeout_entry.pack(side=tk.LEFT, padx=5)
        
        max_label = tk.Label(
            params_frame,
            text="Макс. email:",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        max_label.pack(side=tk.LEFT, padx=(20, 5))
        
        max_entry = tk.Entry(
            params_frame,
            textvariable=self.max_emails,
            width=10,
            font=("Segoe UI", 10),
            relief=tk.SOLID,
            bd=1
        )
        max_entry.pack(side=tk.LEFT, padx=5)
        
        hint_label = tk.Label(
            params_frame,
            text="(оставьте пустым для всех)",
            font=("Segoe UI", 9),
            bg=self.colors['panel'],
            fg="#64748B"
        )
        hint_label.pack(side=tk.LEFT, padx=10)
        
        button_frame = tk.Frame(main_frame, bg=self.colors['light'])
        button_frame.pack(pady=16)
        
        self.start_button = RoundedButton(
            button_frame,
            text="🚀 Начать проверку",
            command=self.start_validation,
            width=250,
            height=50,
            bg_color=self.colors['primary'],
            hover_color=self.colors['dark'],
            corner_radius=25,
            font=("Arial", 13, "bold")
        )
        self.start_button.pack()
        
        # Блок прогресса: компактная сетка 2x2 + полоса, всегда виден
        progress_outer = tk.Frame(main_frame, bg=self.colors['light'])
        progress_outer.pack(fill=tk.X, pady=(0, 12))
        
        progress_frame = tk.LabelFrame(
            progress_outer,
            text="  Прогресс проверки  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            padx=16,
            pady=14,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground="#E2E8F0"
        )
        progress_frame.pack(fill=tk.X)
        
        # Сетка: строка 0 — Файл | Обработано; строка 1 — Процент | Осталось
        self.progress_file_label = tk.Label(
            progress_frame,
            text="Файл: —",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        self.progress_file_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 24), pady=(0, 6))
        
        self.progress_count_label = tk.Label(
            progress_frame,
            text="Обработано: 0 из 0",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        self.progress_count_label.grid(row=0, column=1, sticky=tk.W, pady=(0, 6))
        
        self.progress_percent_label = tk.Label(
            progress_frame,
            text="Процент: 0%",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        self.progress_percent_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 24), pady=(0, 10))
        
        self.progress_eta_label = tk.Label(
            progress_frame,
            text="Осталось примерно: —",
            font=("Segoe UI", 10),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        self.progress_eta_label.grid(row=1, column=1, sticky=tk.W, pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        self.progress_bar.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(0, 0))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.columnconfigure(1, weight=1)
        
        # Статус и кнопка скачивания
        self.status_label = tk.Label(
            main_frame,
            text="✓ Готов к работе",
            font=("Segoe UI", 11, "bold"),
            bg=self.colors['light'],
            fg=self.colors['success']
        )
        self.status_label.pack(pady=(8, 4))
        
        self.download_button = RoundedButton(
            main_frame,
            text="📥 Скачать результат",
            command=self.download_result,
            width=200,
            height=40,
            bg_color=self.colors['accent'],
            hover_color=self.colors['secondary'],
            corner_radius=20
        )
        self.download_button.pack(pady=(0, 16))
        self.download_button.disable()
        
        # Информационная панель (без expand — не отжимает прогресс)
        info_frame = tk.LabelFrame(
            main_frame,
            text="  ℹ️ Информация  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.colors['panel'],
            fg=self.colors['dark'],
            padx=14,
            pady=10,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground="#E2E8F0"
        )
        info_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = """🔒 Строгий режим:
   • Только адреса с высокой надежностью
   • Обязательная активность email
   • Обязательная доставляемость
   • Проверка на подозрительные домены
   • Максимальная точность (95-98%)

✨ Лояльный режим:
   • Адреса с высокой и средней надежностью
   • Мягкие требования к активности
   • Мягкие требования к доставляемости
   • На 15-20% больше валидных адресов
   • Подходит для массовых рассылок"""
        
        info_label = tk.Label(
            info_frame,
            text=info_text.strip(),
            justify=tk.LEFT,
            font=("Segoe UI", 9),
            bg=self.colors['panel'],
            fg=self.colors['dark']
        )
        info_label.pack(anchor=tk.W, padx=6, pady=6)
    
    def select_input_files(self):
        filenames = filedialog.askopenfilenames(
            title="Выберите файлы с email адресами (можно несколько)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if filenames:
            for f in filenames:
                if f and f not in self.input_files:
                    self.input_files.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
    
    def clear_input_files(self):
        self.input_files.clear()
        self.file_listbox.delete(0, tk.END)
    
    def start_validation(self):
        if not self.input_files:
            messagebox.showerror("Ошибка", "Пожалуйста, выберите один или несколько файлов с email адресами")
            return
        
        for f in self.input_files:
            if not os.path.exists(f):
                messagebox.showerror("Ошибка", f"Файл не существует:\n{f}")
                return
        
        if self.is_processing:
            messagebox.showwarning("Внимание", "Проверка уже выполняется")
            return
        
        try:
            timeout_val = int(self.timeout.get()) if self.timeout.get() else 10
        except ValueError:
            timeout_val = 10
        
        try:
            max_emails_val = int(self.max_emails.get()) if self.max_emails.get() else None
        except ValueError:
            max_emails_val = None
        
        self.stop_flag = {}
        self.is_processing = True
        self.start_button.config(state=tk.DISABLED)
        self._reset_progress_display()
        self.status_label.config(text="⏳ Проверка выполняется...", fg=self.colors['accent'])
        self.download_button.disable()
        
        thread = threading.Thread(
            target=self.run_validation,
            args=(timeout_val, max_emails_val),
            daemon=True
        )
        thread.start()
    
    def _reset_progress_display(self):
        """Сброс блока прогресса к начальному виду."""
        self.progress_file_label.config(text="Файл: —")
        self.progress_count_label.config(text="Обработано: 0 из 0")
        self.progress_percent_label.config(text="Процент: 0%")
        self.progress_eta_label.config(text="Осталось примерно: —")
        self.progress_bar['value'] = 0

    def run_validation(self, timeout, max_emails):
        try:
            result_paths = []
            total_files = len(self.input_files)

            for file_idx, input_path in enumerate(self.input_files):
                if self.stop_flag.get('stop'):
                    break
                base_name = os.path.splitext(os.path.basename(input_path))[0]
                file_display_name = os.path.basename(input_path)
                output_dir = os.path.dirname(input_path)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = os.path.join(output_dir, f"{base_name}_validated_{timestamp}.xlsx")

                def make_cb(fname, idx, total_f):
                    def _cb(cur, tot, msg, percent, eta_seconds):
                        def update():
                            self.progress_file_label.config(text=f"Файл: {fname}")
                            self.progress_count_label.config(text=f"Обработано: {cur} из {tot}")
                            self.progress_percent_label.config(text=f"Процент: {percent:.1f}%")
                            if eta_seconds is not None and eta_seconds > 0:
                                if eta_seconds >= 60:
                                    eta_str = f"{int(eta_seconds // 60)} мин {int(eta_seconds % 60)} сек"
                                else:
                                    eta_str = f"{int(eta_seconds)} сек"
                                self.progress_eta_label.config(text=f"Осталось примерно: {eta_str}")
                            else:
                                self.progress_eta_label.config(text="Осталось примерно: —")
                            self.progress_bar['value'] = percent
                        try:
                            self.root.after(0, update)
                        except Exception:
                            pass
                    return _cb

                process_excel_file_advanced(
                    input_file=input_path,
                    output_file=output_path,
                    check_smtp=True,  # SMTP всегда включён
                    timeout=timeout,
                    accept_catch_all=False,
                    max_emails=max_emails,
                    validation_mode=self.validation_mode.get(),
                    include_full_results_sheet=self.include_full_results_sheet.get(),
                    only_valid_emails_sheet=self.only_valid_emails_sheet.get(),
                    progress_callback=make_cb(file_display_name, file_idx, total_files),
                    stop_flag=self.stop_flag,
                )
                result_paths.append(output_path)
            
            self.result_files = result_paths
            msg = f"✓ Проверка завершена! Обработано файлов: {len(result_paths)}"
            self.root.after(0, self.validation_complete, True, msg)
            
        except Exception as e:
            self.root.after(0, self.validation_complete, False, f"✗ Ошибка: {str(e)}")
    
    def validation_complete(self, success, message):
        self.is_processing = False
        self._reset_progress_display()
        self.start_button.config(state=tk.NORMAL)
        
        if success:
            self.status_label.config(text=message, fg=self.colors['success'])
            self.download_button.enable(self.colors['accent'])
            detail = "\n".join(self.result_files) if self.result_files else ""
            messagebox.showinfo("Успех", f"Проверка завершена!\n\nРезультаты сохранены:\n{detail}")
        else:
            self.status_label.config(text=message, fg=self.colors['error'])
            messagebox.showerror("Ошибка", message)
    
    def download_result(self):
        if not self.result_files:
            messagebox.showerror("Ошибка", "Файлы результата не найдены")
            return
        first = self.result_files[0]
        if not os.path.exists(first):
            messagebox.showerror("Ошибка", "Файл результата не найден")
            return
        
        import shutil
        
        folder = filedialog.askdirectory(
            title="Выберите папку для сохранения результатов",
            initialdir=os.path.dirname(os.path.abspath(first))
        )
        if not folder:
            return
        
        copied = []
        errors = []
        for src in self.result_files:
            if not os.path.exists(src):
                errors.append(f"Не найден: {os.path.basename(src)}")
                continue
            name = os.path.basename(src)
            dest = os.path.join(folder, name)
            try:
                shutil.copy2(src, dest)
                copied.append(name)
            except Exception as e:
                errors.append(f"{name}: {e}")
        
        if copied:
            msg = f"Скопировано в папку:\n{folder}\n\nФайлы:\n" + "\n".join(copied)
            if errors:
                msg += "\n\nОшибки:\n" + "\n".join(errors)
            messagebox.showinfo("Готово", msg)
        elif errors:
            messagebox.showerror("Ошибка", "\n".join(errors))


def main():
    root = tk.Tk()
    app = EmailValidatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
