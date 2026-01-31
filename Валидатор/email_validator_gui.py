"""
GUI приложение для валидации email адресов - Валидатор 3000
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from datetime import datetime
from email_validator import process_excel_file
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
                        highlightthickness=0, relief=tk.FLAT, bg="#E3F2FD")
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


class SnakeProgressBar(tk.Canvas):
    """Кастомный прогресс-бар в виде ползущей змейки"""
    def __init__(self, parent, width=600, height=30, **kwargs):
        super().__init__(parent, width=width, height=height, 
                        highlightthickness=0, bg="#E3F2FD", **kwargs)
        self.width = width
        self.height = height
        self.position = 0
        self.animation_id = None
        self.segment_count = 8
        self.segment_width = 40
        self.segment_spacing = 20
        
    def start(self):
        """Запуск анимации"""
        self.position = 0
        self.animate()
        
    def stop(self):
        """Остановка анимации"""
        if self.animation_id:
            self.after_cancel(self.animation_id)
            self.animation_id = None
        self.delete("all")
        # Очищаем фон
        self.create_rectangle(0, 0, self.width, self.height, 
                            fill="#E3F2FD", outline="#BBDEFB", width=2)
        
    def animate(self):
        """Анимация движения змейки"""
        self.delete("all")
        
        # Рисуем фон
        self.create_rectangle(0, 0, self.width, self.height, 
                            fill="#E3F2FD", outline="#BBDEFB", width=2)
        
        # Рисуем змейку
        for i in range(self.segment_count):
            x = (self.position + i * (self.segment_width + self.segment_spacing)) % (self.width + self.segment_width)
            
            # Градиент цветов от темно-синего к голубому
            if i == 0:
                color = "#1976D2"  # Темно-синий (голова)
            elif i < self.segment_count // 2:
                # Переход от темно-синего к синему
                ratio = i / (self.segment_count // 2)
                r = int(25 + (66 - 25) * ratio)
                g = int(118 + (165 - 118) * ratio)
                b = int(210 + (245 - 210) * ratio)
                color = f"#{r:02x}{g:02x}{b:02x}"
            else:
                # Переход к голубому
                ratio = (i - self.segment_count // 2) / (self.segment_count // 2)
                r = int(66 + (3 - 66) * ratio)
                g = int(165 + (169 - 165) * ratio)
                b = int(245 + (244 - 245) * ratio)
                color = f"#{r:02x}{g:02x}{b:02x}"
            
            # Скругленный овал для сегмента
            self.create_oval(x, 5, x + self.segment_width, self.height - 5,
                           fill=color, outline="#0D47A1", width=2)
        
        # Обновляем позицию
        self.position += 3
        if self.position > self.width:
            self.position = -self.segment_width * self.segment_count
        
        self.animation_id = self.after(20, self.animate)


class EmailValidatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Валидатор 3000")
        self.root.geometry("850x750")
        self.root.resizable(False, False)
        
        # Сине-голубая цветовая схема
        self.colors = {
            'primary': '#1976D2',      # Темно-синий
            'secondary': '#42A5F5',     # Синий
            'accent': '#03A9F4',       # Голубой
            'light': '#E3F2FD',        # Светло-голубой
            'dark': '#0D47A1',         # Очень темно-синий
            'success': '#4CAF50',      # Зеленый для успеха
            'warning': '#FF9800',       # Оранжевый для предупреждений
            'error': '#F44336'         # Красный для ошибок
        }
        
        # Настройка фона окна
        self.root.configure(bg=self.colors['light'])
        
        # Переменные
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.check_smtp = tk.BooleanVar(value=True)
        self.accept_catch_all = tk.BooleanVar(value=False)
        self.validation_mode = tk.StringVar(value="strict")
        self.max_emails = tk.StringVar()
        self.timeout = tk.StringVar(value="10")
        
        # Результат обработки
        self.result_file = None
        self.is_processing = False
        
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок с российским триколором
        header_frame = tk.Canvas(self.root, height=100, highlightthickness=0, bg="#E3F2FD")
        header_frame.pack(fill=tk.X)
        
        # Функция для отрисовки триколора
        def draw_tricolor(event=None):
            width = header_frame.winfo_width() if header_frame.winfo_width() > 1 else 850
            stripe_height = 100 // 3
            
            header_frame.delete("tricolor")
            
            # Белая полоса (верхняя)
            header_frame.create_rectangle(0, 0, width, stripe_height, fill="white", outline="", tags="tricolor")
            # Синяя полоса (средняя)
            header_frame.create_rectangle(0, stripe_height, width, stripe_height * 2, fill="#0039A6", outline="", tags="tricolor")
            # Красная полоса (нижняя)
            header_frame.create_rectangle(0, stripe_height * 2, width, 100, fill="#D52B1E", outline="", tags="tricolor")
            
            # Обновляем позицию текста
            header_frame.coords("title_text", width // 2, 50)
        
        header_frame.bind("<Configure>", draw_tricolor)
        
        # Название поверх триколора с тенью для читаемости
        title_label = tk.Label(
            header_frame, 
            text="Валидатор 3000",
            font=("Arial", 28, "bold"),
            bg="#0039A6",  # Синий фон для лучшей читаемости
            fg="white",
            padx=30,
            pady=10,
            relief=tk.RAISED,
            bd=2
        )
        header_frame.create_window(425, 50, window=title_label, tags="title_text")
        
        # Вызываем отрисовку после создания
        self.root.after(100, draw_tricolor)
        
        # Основной контейнер
        main_frame = tk.Frame(self.root, bg=self.colors['light'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Фрейм для загрузки файла
        file_frame = tk.LabelFrame(
            main_frame, 
            text="📁 Файл с email адресами", 
            font=("Arial", 11, "bold"),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            padx=15,
            pady=15,
            relief=tk.RAISED,
            bd=2
        )
        file_frame.pack(fill=tk.X, pady=10)
        
        file_entry_frame = tk.Frame(file_frame, bg=self.colors['light'])
        file_entry_frame.pack(fill=tk.X)
        
        file_entry = tk.Entry(
            file_entry_frame, 
            textvariable=self.input_file, 
            width=50, 
            state="readonly",
            font=("Arial", 10),
            relief=tk.SUNKEN,
            bd=2
        )
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        select_btn = RoundedButton(
            file_entry_frame,
            text="Выбрать файл",
            command=self.select_input_file,
            width=150,
            height=35,
            bg_color=self.colors['secondary'],
            hover_color=self.colors['primary'],
            corner_radius=15
        )
        select_btn.pack(side=tk.RIGHT)
        
        # Фрейм для настроек
        settings_frame = tk.LabelFrame(
            main_frame, 
            text="⚙️ Настройки проверки", 
            font=("Arial", 11, "bold"),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            padx=15,
            pady=15,
            relief=tk.RAISED,
            bd=2
        )
        settings_frame.pack(fill=tk.X, pady=10)
        
        # Режим валидации
        mode_label = tk.Label(
            settings_frame, 
            text="Режим валидации:",
            font=("Arial", 10, "bold"),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        mode_label.grid(row=0, column=0, sticky=tk.W, pady=8)
        
        mode_frame = tk.Frame(settings_frame, bg=self.colors['light'])
        mode_frame.grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=10)
        
        strict_radio = tk.Radiobutton(
            mode_frame,
            text="🔒 Строгий (максимальная точность)",
            variable=self.validation_mode,
            value="strict",
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            selectcolor=self.colors['light'],
            activebackground=self.colors['light'],
            activeforeground=self.colors['primary']
        )
        strict_radio.pack(side=tk.LEFT, padx=10)
        
        lenient_radio = tk.Radiobutton(
            mode_frame,
            text="✨ Лояльный (+15-20% валидных)",
            variable=self.validation_mode,
            value="lenient",
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            selectcolor=self.colors['light'],
            activebackground=self.colors['light'],
            activeforeground=self.colors['accent']
        )
        lenient_radio.pack(side=tk.LEFT, padx=10)
        
        # SMTP проверка
        smtp_check = tk.Checkbutton(
            settings_frame,
            text="✓ Выполнять SMTP проверку",
            variable=self.check_smtp,
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            selectcolor=self.colors['light'],
            activebackground=self.colors['light'],
            activeforeground=self.colors['primary']
        )
        smtp_check.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Catch-all
        catchall_check = tk.Checkbutton(
            settings_frame,
            text="✓ Считать валидными catch-all адреса",
            variable=self.accept_catch_all,
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            selectcolor=self.colors['light'],
            activebackground=self.colors['light'],
            activeforeground=self.colors['primary']
        )
        catchall_check.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Таймаут и количество
        params_frame = tk.Frame(settings_frame, bg=self.colors['light'])
        params_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=10)
        
        timeout_label = tk.Label(
            params_frame,
            text="Таймаут (сек):",
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        timeout_label.pack(side=tk.LEFT, padx=(0, 5))
        
        timeout_entry = tk.Entry(
            params_frame,
            textvariable=self.timeout,
            width=10,
            font=("Arial", 10),
            relief=tk.SUNKEN,
            bd=2
        )
        timeout_entry.pack(side=tk.LEFT, padx=5)
        
        max_label = tk.Label(
            params_frame,
            text="Макс. email:",
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        max_label.pack(side=tk.LEFT, padx=(20, 5))
        
        max_entry = tk.Entry(
            params_frame,
            textvariable=self.max_emails,
            width=10,
            font=("Arial", 10),
            relief=tk.SUNKEN,
            bd=2
        )
        max_entry.pack(side=tk.LEFT, padx=5)
        
        hint_label = tk.Label(
            params_frame,
            text="(оставьте пустым для всех)",
            font=("Arial", 9),
            bg=self.colors['light'],
            fg="#757575"
        )
        hint_label.pack(side=tk.LEFT, padx=10)
        
        # Кнопка запуска
        button_frame = tk.Frame(main_frame, bg=self.colors['light'])
        button_frame.pack(pady=20)
        
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
        
        # Прогресс бар (змейка)
        progress_frame = tk.Frame(main_frame, bg=self.colors['light'])
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress = SnakeProgressBar(progress_frame, width=750, height=35)
        self.progress.pack()
        
        # Статус
        self.status_label = tk.Label(
            main_frame,
            text="✓ Готов к работе",
            font=("Arial", 11, "bold"),
            bg=self.colors['light'],
            fg=self.colors['success']
        )
        self.status_label.pack(pady=10)
        
        # Метка для времени до завершения
        self.time_label = tk.Label(
            main_frame,
            text="",
            font=("Arial", 10),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        self.time_label.pack(pady=5)
        
        # Кнопка скачивания результата
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
        self.download_button.pack(pady=10)
        self.download_button.disable()  # Начинаем с отключенной кнопки
        
        # Информационная панель
        info_frame = tk.LabelFrame(
            main_frame,
            text="ℹ️ Информация",
            font=("Arial", 10, "bold"),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            padx=15,
            pady=10,
            relief=tk.RAISED,
            bd=2
        )
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        info_text = """🔒 Строгий режим:
   • Только адреса с высокой надежностью
   • Обязательная активность email
   • Проверка репутации домена
   • Максимальная точность (95-98%)

✨ Лояльный режим:
   • Адреса с высокой и средней надежностью
   • Мягкие требования к активности
   • На 15-20% больше валидных адресов
   • Подходит для массовых рассылок"""
        
        info_label = tk.Label(
            info_frame,
            text=info_text.strip(),
            justify=tk.LEFT,
            font=("Arial", 9),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        info_label.pack(anchor=tk.W, padx=10, pady=5)
    
    def select_input_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите файл с email адресами",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            base_name = os.path.splitext(os.path.basename(filename))[0]
            output_dir = os.path.dirname(filename)
            self.output_file.set(os.path.join(output_dir, f"{base_name}_validated.xlsx"))
    
    def start_validation(self):
        if not self.input_file.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите файл с email адресами")
            return
        
        if not os.path.exists(self.input_file.get()):
            messagebox.showerror("Ошибка", "Выбранный файл не существует")
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
        
        # Подсчитываем количество email для оценки времени
        try:
            import pandas as pd
            df = pd.read_excel(self.input_file.get())
            
            # Находим столбец с email
            email_column = None
            for col in df.columns:
                col_lower = str(col).lower()
                if any(keyword in col_lower for keyword in ['email', 'e-mail', 'почта', 'mail', 'адрес']):
                    email_column = col
                    break
            
            if email_column is None:
                email_column = df.columns[0]
            
            # Подсчитываем уникальные email
            emails = []
            seen_emails = set()
            for value in df[email_column]:
                if pd.notna(value):
                    email_str = str(value).strip()
                    if email_str and email_str.lower() not in ['nan', 'none', '']:
                        email_lower = email_str.lower()
                        if email_lower not in seen_emails:
                            seen_emails.add(email_lower)
                            emails.append(email_str)
            
            total_count = len(emails)
            if max_emails_val and max_emails_val > 0:
                total_count = min(total_count, max_emails_val)
            
            # Оценка времени (с SMTP: ~1.5 сек/email, без SMTP: ~0.7 сек/email)
            avg_time = 1.5 if self.check_smtp.get() else 0.7
            estimated_seconds = total_count * avg_time
            
            if estimated_seconds > 60:
                time_str = f"{int(estimated_seconds // 60)} мин {int(estimated_seconds % 60)} сек"
            else:
                time_str = f"{int(estimated_seconds)} сек"
            
            self.time_label.config(text=f"Примерное время до завершения: ~{time_str}")
        except Exception as e:
            self.time_label.config(text="")
        
        self.is_processing = True
        self.start_button.config(state=tk.DISABLED)
        self.progress.start()
        self.status_label.config(text="⏳ Проверка выполняется...", fg=self.colors['accent'])
        self.download_button.disable()
        
        thread = threading.Thread(
            target=self.run_validation,
            args=(timeout_val, max_emails_val),
            daemon=True
        )
        thread.start()
    
    def run_validation(self, timeout, max_emails):
        try:
            if not self.output_file.get():
                base_name = os.path.splitext(os.path.basename(self.input_file.get()))[0]
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = os.path.join(
                    os.path.dirname(self.input_file.get()),
                    f"{base_name}_{timestamp}.xlsx"
                )
            else:
                output_path = self.output_file.get()
            
            process_excel_file(
                input_file=self.input_file.get(),
                output_file=output_path,
                check_smtp=self.check_smtp.get(),
                timeout=timeout,
                accept_catch_all=self.accept_catch_all.get(),
                max_emails=max_emails,
                validation_mode=self.validation_mode.get()
            )
            
            self.result_file = output_path
            
            self.root.after(0, self.validation_complete, True, "✓ Проверка завершена успешно!")
            
        except Exception as e:
            self.root.after(0, self.validation_complete, False, f"✗ Ошибка: {str(e)}")
    
    def validation_complete(self, success, message):
        self.is_processing = False
        self.progress.stop()
        self.start_button.config(state=tk.NORMAL)
        self.time_label.config(text="")  # Очищаем информацию о времени
        
        if success:
            self.status_label.config(text=message, fg=self.colors['success'])
            self.download_button.enable(self.colors['accent'])
            messagebox.showinfo("Успех", f"Проверка завершена!\n\nРезультат сохранен в:\n{self.result_file}")
        else:
            self.status_label.config(text=message, fg=self.colors['error'])
            messagebox.showerror("Ошибка", message)
    
    def download_result(self):
        if not self.result_file or not os.path.exists(self.result_file):
            messagebox.showerror("Ошибка", "Файл результата не найден")
            return
        
        import subprocess
        import platform
        
        if platform.system() == "Windows":
            os.startfile(os.path.dirname(os.path.abspath(self.result_file)))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", os.path.dirname(os.path.abspath(self.result_file))])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(os.path.abspath(self.result_file))])
        
        messagebox.showinfo("Информация", f"Файл находится в:\n{self.result_file}")


def main():
    root = tk.Tk()
    app = EmailValidatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
