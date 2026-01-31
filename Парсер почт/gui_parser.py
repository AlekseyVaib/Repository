#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI приложение для парсера e-mail адресов компаний
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from pathlib import Path
import sys
import importlib.util

# Импортируем парсер (избегаем конфликта с встроенным модулем parser)
spec = importlib.util.spec_from_file_location("email_parser", "parser.py")
email_parser_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(email_parser_module)

EmailParser = email_parser_module.EmailParser
save_results = email_parser_module.save_results
logger = email_parser_module.logger


class ParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер E-mail адресов компаний")
        self.root.geometry("800x750")
        self.root.resizable(False, False)
        
        # Цветовая схема (голубо-синяя)
        self.colors = {
            'bg_primary': '#1E3A5F',      # Темно-синий фон
            'bg_secondary': '#2E4A6F',    # Средний синий
            'bg_light': '#3E5A7F',        # Светлый синий
            'accent': '#4A90E2',          # Голубой акцент
            'accent_hover': '#5BA0F2',    # Голубой при наведении
            'text_primary': '#FFFFFF',     # Белый текст
            'text_secondary': '#E0E0E0',  # Светло-серый текст
            'entry_bg': '#FFFFFF',        # Белый фон полей
            'entry_fg': '#1E3A5F',        # Темно-синий текст в полях
        }
        
        # Настройка стиля
        self.setup_style()
        
        # Создание интерфейса
        self.create_widgets()
        
        # Переменные
        self.parsing = False
        self.parser = None
    
    def setup_style(self):
        """Настройка стилей для виджетов"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Настройка стилей для кнопок
        style.configure('Accent.TButton',
                       background=self.colors['accent'],
                       foreground=self.colors['text_primary'],
                       borderwidth=0,
                       focuscolor='none',
                       padding=10)
        style.map('Accent.TButton',
                 background=[('active', self.colors['accent_hover']),
                            ('pressed', self.colors['accent'])])
        
        style.configure('Secondary.TButton',
                       background=self.colors['bg_light'],
                       foreground=self.colors['text_primary'],
                       borderwidth=0,
                       focuscolor='none',
                       padding=10)
    
    def create_widgets(self):
        """Создание виджетов интерфейса"""
        # Главный контейнер
        main_frame = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Заголовок
        header_frame = tk.Frame(main_frame, bg=self.colors['bg_primary'])
        header_frame.pack(fill=tk.X, pady=(20, 30))
        
        title_label = tk.Label(header_frame,
                               text="Парсер E-mail адресов",
                               font=('Arial', 24, 'bold'),
                               bg=self.colors['bg_primary'],
                               fg=self.colors['text_primary'])
        title_label.pack()
        
        subtitle_label = tk.Label(header_frame,
                                 text="Автоматический сбор контактной информации",
                                 font=('Arial', 11),
                                 bg=self.colors['bg_primary'],
                                 fg=self.colors['text_secondary'])
        subtitle_label.pack(pady=(5, 0))
        
        # Контейнер для полей ввода
        form_frame = tk.Frame(main_frame, bg=self.colors['bg_primary'])
        form_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=(0, 20))
        
        # Название сайта
        self.create_labeled_entry(form_frame, "Название сайта:", "site_name", 0)
        
        # Ссылка на сайт
        self.create_labeled_entry(form_frame, "Ссылка на сайт:", "site_url", 1)
        
        # Глубина поиска
        depth_frame = tk.Frame(form_frame, bg=self.colors['bg_primary'])
        depth_frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=15)
        
        depth_label = tk.Label(depth_frame,
                               text="Глубина поиска:",
                               font=('Arial', 11, 'bold'),
                               bg=self.colors['bg_primary'],
                               fg=self.colors['text_primary'],
                               anchor='w')
        depth_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.depth_var = tk.IntVar(value=0)
        depth_info = tk.Label(depth_frame,
                             text="(0 - сразу собираем почты, 1+ - собираем ссылки и переходим по ним)",
                             font=('Arial', 9),
                             bg=self.colors['bg_primary'],
                             fg=self.colors['text_secondary'])
        depth_info.pack(side=tk.LEFT)
        
        depth_spinbox = tk.Spinbox(depth_frame,
                                  from_=0,
                                  to=3,
                                  textvariable=self.depth_var,
                                  font=('Arial', 11),
                                  width=5,
                                  bg=self.colors['entry_bg'],
                                  fg=self.colors['entry_fg'],
                                  relief=tk.FLAT,
                                  bd=0)
        depth_spinbox.pack(side=tk.RIGHT)
        
        # Таймаут
        self.create_labeled_entry(form_frame, "Таймаут запроса (сек):", "timeout", 3, default="15")
        
        # Задержка между запросами
        self.create_labeled_entry(form_frame, "Задержка между запросами (сек):", "delay", 4, default="1.0")
        
        # Опция выгрузки
        export_frame = tk.Frame(form_frame, bg=self.colors['bg_primary'])
        export_frame.grid(row=5, column=0, columnspan=2, sticky='ew', pady=15)
        
        self.export_sites_only_var = tk.BooleanVar(value=False)
        export_checkbox = tk.Checkbutton(export_frame,
                                         text="Выгрузить только сайты с названиями (без email)",
                                         font=('Arial', 10, 'bold'),
                                         bg=self.colors['bg_primary'],
                                         fg=self.colors['accent'],
                                         selectcolor='#FFFFFF',  # Белый цвет для галочки
                                         activebackground=self.colors['bg_primary'],
                                         activeforeground=self.colors['accent'],
                                         variable=self.export_sites_only_var,
                                         anchor='w',
                                         cursor='hand2',
                                         relief=tk.FLAT,
                                         bd=0,
                                         highlightthickness=0,
                                         indicatoron=True)
        export_checkbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Выходной файл
        output_frame = tk.Frame(form_frame, bg=self.colors['bg_primary'])
        output_frame.grid(row=6, column=0, columnspan=2, sticky='ew', pady=15)
        
        output_label = tk.Label(output_frame,
                               text="Выходной файл:",
                               font=('Arial', 11, 'bold'),
                               bg=self.colors['bg_primary'],
                               fg=self.colors['text_primary'],
                               anchor='w')
        output_label.pack(side=tk.LEFT, padx=(0, 10))
        
        self.output_file_var = tk.StringVar(value="results.xlsx")
        output_entry = tk.Entry(output_frame,
                               textvariable=self.output_file_var,
                               font=('Arial', 11),
                               bg=self.colors['entry_bg'],
                               fg=self.colors['entry_fg'],
                               relief=tk.FLAT,
                               bd=0,
                               insertbackground=self.colors['entry_fg'])
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_btn = tk.Button(output_frame,
                              text="Обзор",
                              font=('Arial', 10),
                              bg=self.colors['bg_light'],
                              fg=self.colors['text_primary'],
                              relief=tk.FLAT,
                              bd=0,
                              cursor='hand2',
                              command=self.browse_output_file,
                              padx=15,
                              pady=5)
        browse_btn.pack(side=tk.RIGHT)
        
        # Прогресс бар
        self.progress_var = tk.StringVar(value="Готов к работе")
        progress_label = tk.Label(form_frame,
                                 textvariable=self.progress_var,
                                 font=('Arial', 10),
                                 bg=self.colors['bg_primary'],
                                 fg=self.colors['text_secondary'],
                                 anchor='w')
        progress_label.grid(row=7, column=0, columnspan=2, sticky='ew', pady=(20, 10))
        
        self.progress_bar = ttk.Progressbar(form_frame,
                                           mode='indeterminate',
                                           length=400)
        self.progress_bar.grid(row=8, column=0, columnspan=2, sticky='ew', pady=(0, 20))
        
        # Кнопки
        button_frame = tk.Frame(main_frame, bg=self.colors['bg_primary'])
        button_frame.pack(fill=tk.X, padx=40, pady=(0, 30))
        
        self.start_btn = tk.Button(button_frame,
                                  text="Начать парсинг",
                                  font=('Arial', 12, 'bold'),
                                  bg=self.colors['accent'],
                                  fg=self.colors['text_primary'],
                                  relief=tk.FLAT,
                                  bd=0,
                                  cursor='hand2',
                                  command=self.start_parsing,
                                  padx=30,
                                  pady=12)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_btn = tk.Button(button_frame,
                                 text="Остановить",
                                 font=('Arial', 12),
                                 bg=self.colors['bg_light'],
                                 fg=self.colors['text_primary'],
                                 relief=tk.FLAT,
                                 bd=0,
                                 cursor='hand2',
                                 command=self.stop_parsing,
                                 state=tk.DISABLED,
                                 padx=30,
                                 pady=12)
        self.stop_btn.pack(side=tk.LEFT)
        
        # Настройка grid
        form_frame.columnconfigure(1, weight=1)
    
    def create_labeled_entry(self, parent, label_text, var_name, row, default=""):
        """Создание поля ввода с меткой"""
        label = tk.Label(parent,
                        text=label_text,
                        font=('Arial', 11, 'bold'),
                        bg=self.colors['bg_primary'],
                        fg=self.colors['text_primary'],
                        anchor='w')
        label.grid(row=row, column=0, sticky='w', pady=15)
        
        var = tk.StringVar(value=default)
        setattr(self, var_name + '_var', var)
        
        entry = tk.Entry(parent,
                        textvariable=var,
                        font=('Arial', 11),
                        bg=self.colors['entry_bg'],
                        fg=self.colors['entry_fg'],
                        relief=tk.FLAT,
                        bd=0,
                        insertbackground=self.colors['entry_fg'])
        entry.grid(row=row, column=1, sticky='ew', padx=(10, 0), pady=15)
    
    def browse_output_file(self):
        """Выбор выходного файла"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file_var.set(filename)
    
    def validate_inputs(self):
        """Проверка введенных данных"""
        if not self.site_url_var.get().strip():
            messagebox.showerror("Ошибка", "Введите ссылку на сайт!")
            return False
        
        try:
            timeout = int(self.timeout_var.get())
            if timeout <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Таймаут должен быть положительным числом!")
            return False
        
        try:
            delay = float(self.delay_var.get())
            if delay < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Задержка должна быть неотрицательным числом!")
            return False
        
        return True
    
    def start_parsing(self):
        """Запуск парсинга"""
        if not self.validate_inputs():
            return
        
        if self.parsing:
            return
        
        self.parsing = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar.start()
        self.progress_var.set("Парсинг начат...")
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.parse_thread, daemon=True)
        thread.start()
    
    def stop_parsing(self):
        """Остановка парсинга"""
        self.parsing = False
        self.progress_var.set("Остановка...")
        # Парсер будет остановлен при следующей проверке
    
    def parse_thread(self):
        """Поток для парсинга"""
        try:
            site_name = self.site_name_var.get().strip() or "Сайт"
            site_url = self.site_url_var.get().strip()
            depth = self.depth_var.get()
            timeout = int(self.timeout_var.get())
            delay = float(self.delay_var.get())
            output_file = self.output_file_var.get().strip() or "results.xlsx"
            export_sites_only = self.export_sites_only_var.get()
            
            # Создание парсера
            self.parser = EmailParser(timeout=timeout, delay=delay)
            
            # Парсинг
            self.progress_var.set(f"Обработка сайта: {site_url}")
            
            if export_sites_only:
                # Режим: только сайты с названиями (без парсинга email)
                results = self.parser.parse_sites_only(site_url, depth=depth)
            else:
                # Обычный режим: парсинг email
                results = self.parser.parse_with_depth(site_url, depth=depth)
            
            if not self.parsing:
                self.progress_var.set("Парсинг остановлен")
                self.root.after(0, self.parsing_finished)
                return
            
            if not results:
                self.progress_var.set("Результаты не найдены")
                messagebox.showwarning("Предупреждение", "Не удалось найти данные на сайте")
                self.root.after(0, self.parsing_finished)
                return
            
            # Сохранение результатов
            self.progress_var.set("Сохранение результатов...")
            save_results(results, output_file, mode='companies', sites_only=export_sites_only)
            
            # Статистика
            total = len(results)
            
            if export_sites_only:
                self.progress_var.set(f"Готово! Найдено {total} сайтов")
                messagebox.showinfo("Успех", 
                                  f"Парсинг завершён!\n\n"
                                  f"Всего сайтов: {total}\n\n"
                                  f"Результаты сохранены в: {output_file}")
            else:
                with_email = sum(1 for r in results if r.get('Email') and r.get('Email') != '-')
                self.progress_var.set(f"Готово! Найдено {total} записей, {with_email} с email")
                messagebox.showinfo("Успех", 
                                  f"Парсинг завершён!\n\n"
                                  f"Всего записей: {total}\n"
                                  f"С email: {with_email}\n"
                                  f"Без email: {total - with_email}\n\n"
                                  f"Результаты сохранены в: {output_file}")
            
        except Exception as e:
            logger.error(f"Ошибка при парсинге: {e}", exc_info=True)
            self.progress_var.set(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка при парсинге:\n{str(e)}")
        finally:
            self.root.after(0, self.parsing_finished)
    
    def parsing_finished(self):
        """Завершение парсинга"""
        self.parsing = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress_bar.stop()


def main():
    """Главная функция"""
    root = tk.Tk()
    app = ParserGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()
