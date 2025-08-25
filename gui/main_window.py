import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import logging
from io import StringIO
import sys
import threading
from core.report_updater import update_reports
from core.config_manager import load_config, save_config, validate_config
from .config_editor import ConfigEditor
from .mapping_editor import MappingEditor

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Report Updater")
        self.geometry("1000x700")
        self.config = load_config()
        
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=("Arial", 10))
        self.style.configure("TFrame", padding=10)
        self.style.configure("TLabel", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 11, "bold"))
        
        self.container = ttk.Frame(self)
        self.container.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.notebook = ttk.Notebook(self.container)
        self.notebook.pack(fill='both', expand=True)
        
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Главная")
        
        self.create_widgets()
        self.setup_layout()
        
        # Горячие клавиши
        self.bind('<Control-s>', self.save_config_shortcut)
        self.bind('<Control-q>', lambda e: self.destroy())
        
    def save_config_shortcut(self, event=None):
        if save_config(self.config):
            messagebox.showinfo("Успех", "Конфигурация сохранена")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить конфигурацию")
    
    def create_widgets(self):
        self.input_file_var = tk.StringVar()
        self.report_file_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar(value="Sheet1")
        self.day_var = tk.IntVar(value=1)
        
        # Группировка элементов
        input_frame = ttk.LabelFrame(self.main_frame, text="Источник данных")
        input_frame.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        
        ttk.Label(input_frame, text="Входной файл:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_input = ttk.Entry(input_frame, textvariable=self.input_file_var, width=60)
        self.entry_input.grid(row=0, column=1, padx=5, pady=5, sticky='we')
        self.btn_browse_input = ttk.Button(
            input_frame, text="Обзор", 
            command=lambda: self.browse_file(self.input_file_var, "Выберите входной файл"),
            width=10
        )
        self.btn_browse_input.grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(input_frame, text="Файл отчета:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_report = ttk.Entry(input_frame, textvariable=self.report_file_var, width=60)
        self.entry_report.grid(row=1, column=1, padx=5, pady=5, sticky='we')
        self.btn_browse_report = ttk.Button(
            input_frame, text="Обзор", 
            command=lambda: self.browse_file(self.report_file_var, "Выберите файл отчета"),
            width=10
        )
        self.btn_browse_report.grid(row=1, column=2, padx=5, pady=5)
        
        # Настройки отчета
        settings_frame = ttk.LabelFrame(self.main_frame, text="Настройки отчета")
        settings_frame.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky='ew')
        
        ttk.Label(settings_frame, text="Имя листа:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_sheet = ttk.Entry(settings_frame, textvariable=self.sheet_name_var, width=20)
        self.entry_sheet.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        ttk.Label(settings_frame, text="День:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.entry_day = ttk.Spinbox(settings_frame, textvariable=self.day_var, from_=1, to=31, width=5)
        self.entry_day.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        
        # Кнопки действий
        self.btn_run = ttk.Button(self.main_frame, text="Обновить отчеты", command=self.run_update, width=20)
        self.btn_config = ttk.Button(self.main_frame, text="Редактировать конфиг", 
                                   command=lambda: self.open_editor("config"), width=20)
        self.btn_mappings = ttk.Button(self.main_frame, text="Редактировать маппинги", 
                                     command=lambda: self.open_editor("mapping"), width=22)
        
        # Лог выполнения
        self.log_frame = ttk.LabelFrame(self.main_frame, text="Лог выполнения")
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=15, state='disabled', font=("Consolas", 9))
        
        # Кнопки для лога
        log_btn_frame = ttk.Frame(self.log_frame)
        ttk.Button(log_btn_frame, text="Очистить лог", command=self.clear_log).pack(side=tk.LEFT, padx=2)
        ttk.Button(log_btn_frame, text="Экспорт в файл", command=self.export_log).pack(side=tk.LEFT, padx=2)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(self.main_frame, mode='indeterminate')
        
        self.status_var = tk.StringVar(value="Готов к работе")
        self.status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
    
    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def export_log(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить лог",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(self.log_text.get(1.0, tk.END))
            messagebox.showinfo("Успех", "Лог успешно экспортирован")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать лог: {str(e)}")
    
    def setup_layout(self):
        self.btn_run.grid(row=2, column=0, padx=5, pady=10, sticky='w')
        self.btn_config.grid(row=2, column=1, padx=5, pady=10)
        self.btn_mappings.grid(row=2, column=2, padx=5, pady=10)
        
        self.log_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky='nsew')
        log_btn_frame = ttk.Frame(self.log_frame)
        log_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(log_btn_frame, text="Очистить лог", command=self.clear_log).pack(side=tk.LEFT, padx=2)
        ttk.Button(log_btn_frame, text="Экспорт в файл", command=self.export_log).pack(side=tk.LEFT, padx=2)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.progress.grid(row=4, column=0, columnspan=3, sticky='ew', padx=5, pady=5)
        
        self.status_bar.grid(row=5, column=0, columnspan=3, sticky='ew')
        
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=1)
    
    def open_editor(self, editor_type):
        for tab in self.notebook.tabs()[1:]:
            self.notebook.forget(tab)
        
        if editor_type == "config":
            editor_frame = ttk.Frame(self.notebook)
            ConfigEditor(editor_frame, self.config)
            self.notebook.add(editor_frame, text="Редактор конфигурации")
        else:
            editor_frame = ttk.Frame(self.notebook)
            MappingEditor(editor_frame, self.config, on_save=self.save_config_callback)
            self.notebook.add(editor_frame, text="Редактор маппингов")
        
        self.notebook.select(len(self.notebook.tabs())-1)
    
    def save_config_callback(self):
        """Callback для сохранения конфигурации после изменений в редакторе маппингов"""
        if save_config(self.config):
            messagebox.showinfo("Успех", "Конфигурация сохранена")
        else:
            messagebox.showerror("Ошибка", "Не удалось сохранить конфигурацию")
    
    def browse_file(self, target_var, title):
        file = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            target_var.set(file)
            
    def run_update(self):
        if not self.validate_inputs():
            return
            
        self.btn_run.config(state=tk.DISABLED)
        self.btn_config.config(state=tk.DISABLED)
        self.status_var.set("Выполняется обновление...")
        self.progress.start(10)
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self._run_update_thread)
        thread.daemon = True
        thread.start()
        
    def _run_update_thread(self):
        try:
            log_capture = StringIO()
            handler = logging.StreamHandler(log_capture)
            handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logging.getLogger().addHandler(handler)
            logging.getLogger().setLevel(logging.INFO)
            
            success = update_reports(
                self.input_file_var.get(),
                self.report_file_var.get(),
                self.sheet_name_var.get(),
                self.day_var.get()
            )
            
            logging.getLogger().removeHandler(handler)
            log_content = log_capture.getvalue()
            
            self.after(0, self._update_ui_after_run, log_content, success)
        except Exception as e:
            self.after(0, self._handle_error, str(e))
            
    def _update_ui_after_run(self, log_content, success):
        self.btn_run.config(state=tk.NORMAL)
        self.btn_config.config(state=tk.NORMAL)
        self.progress.stop()
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, log_content)
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        
        if success:
            self.status_var.set("Обновление завершено успешно")
            messagebox.showinfo("Успех", "Отчеты успешно обновлены")
        else:
            self.status_var.set("Обновление завершено с ошибками")
            messagebox.showerror("Ошибка", "При обновлении отчетов произошли ошибки. Проверьте лог для деталей.")
    
    def _handle_error(self, error):
        self.btn_run.config(state=tk.NORMAL)
        self.btn_config.config(state=tk.NORMAL)
        self.progress.stop()
        self.status_var.set("Ошибка выполнения")
        messagebox.showerror("Ошибка", f"Произошла ошибка: {error}")
        logging.error(f"Ошибка в потоке выполнения: {error}")
            
    def validate_inputs(self):
        errors = []
        
        if not self.input_file_var.get():
            errors.append("Не выбран входной файл")
        if not self.report_file_var.get():
            errors.append("Не выбран файл отчета")
        if not self.sheet_name_var.get():
            errors.append("Не указано имя листа")
        
        try:
            day = int(self.day_var.get())
            if day < 1 or day > 31:
                errors.append("День должен быть числом от 1 до 31")
        except ValueError:
            errors.append("День должен быть числом")
        
        config_errors = validate_config(self.config)
        if config_errors:
            errors.extend(config_errors)
        
        if errors:
            messagebox.showerror("Ошибки ввода", "\n".join(errors))
            return False
        return True