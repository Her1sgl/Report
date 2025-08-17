import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

class MappingEditor:
    def __init__(self, parent, config):
        self.parent = parent
        self.config = config
        self.create_widgets()
        self.load_data()
        # Привязываем обработчик клика к таблицам
        self.manager_table.bind('<Button-1>', self.toggle_selection)
        self.region_table.bind('<Button-1>', self.toggle_selection)

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.parent)
        self.notebook.pack(fill='both', expand=True)

        # Менеджеры
        self.manager_frame = ttk.Frame(self.notebook)
        self.create_manager_tab(self.manager_frame)
        self.notebook.add(self.manager_frame, text="Менеджеры")

        # Регионы
        self.region_frame = ttk.Frame(self.notebook)
        self.create_region_tab(self.region_frame)
        self.notebook.add(self.region_frame, text="Регионы")

        # Кнопки
        btn_frame = ttk.Frame(self.parent)
        ttk.Button(btn_frame, text="Сохранить", command=self.save, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Закрыть", command=self.close, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Импорт", command=self.import_data, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Экспорт", command=self.export_data, width=15).pack(side=tk.LEFT, padx=5)
        btn_frame.pack(pady=10)

    def create_manager_tab(self, frame):
        # Панель поиска
        search_frame = ttk.Frame(frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT, padx=5)
        self.manager_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.manager_search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(search_frame, text="Найти", command=self.search_manager).pack(side=tk.LEFT, padx=5)
        search_entry.bind('<Return>', lambda e: self.search_manager())

        columns = ("manager", "sector")
        self.manager_table = ttk.Treeview(frame, columns=columns, show='headings', height=15)
        self.manager_table.heading("manager", text="Менеджер")
        self.manager_table.heading("sector", text="Сектор")
        self.manager_table.column("manager", width=250, anchor=tk.W)
        self.manager_table.column("sector", width=250, anchor=tk.W)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.manager_table.yview)
        self.manager_table.configure(yscroll=scrollbar.set)
        self.manager_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = ttk.Frame(frame)
        ttk.Button(btn_frame, text="Добавить", command=self.add_manager_row, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_manager_row, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Редактировать", command=self.edit_manager_row, width=16).pack(side=tk.LEFT, padx=2)
        btn_frame.pack(pady=5)

    def create_region_tab(self, frame):
        # Панель поиска
        search_frame = ttk.Frame(frame)
        search_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT, padx=5)
        self.region_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.region_search_var, width=30)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(search_frame, text="Найти", command=self.search_region).pack(side=tk.LEFT, padx=5)
        search_entry.bind('<Return>', lambda e: self.search_region())

        columns = ("region", "sector")
        self.region_table = ttk.Treeview(frame, columns=columns, show='headings', height=15)
        self.region_table.heading("region", text="Регион")
        self.region_table.heading("sector", text="Сектор")
        self.region_table.column("region", width=250, anchor=tk.W)
        self.region_table.column("sector", width=250, anchor=tk.W)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.region_table.yview)
        self.region_table.configure(yscroll=scrollbar.set)
        self.region_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        btn_frame = ttk.Frame(frame)
        ttk.Button(btn_frame, text="Добавить", command=self.add_region_row, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_region_row, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Редактировать", command=self.edit_region_row, width=16).pack(side=tk.LEFT, padx=2)
        btn_frame.pack(pady=5)

    def load_data(self):
        for manager, sector in self.config["manager_mapping"].items():
            self.manager_table.insert("", tk.END, values=(manager, sector))
        for region, sector in self.config["region_mapping"].items():
            self.region_table.insert("", tk.END, values=(region, sector))

    def search_manager(self):
        term = self.manager_search_var.get().lower()
        for item in self.manager_table.get_children():
            values = self.manager_table.item(item, "values")
            if term and term not in values[0].lower() and term not in values[1].lower():
                self.manager_table.detach(item)
            else:
                self.manager_table.reattach(item, "", "end")

    def search_region(self):
        term = self.region_search_var.get().lower()
        for item in self.region_table.get_children():
            values = self.region_table.item(item, "values")
            if term and term not in values[0].lower() and term not in values[1].lower():
                self.region_table.detach(item)
            else:
                self.region_table.reattach(item, "", "end")

    def save(self):
        try:
            self.config["manager_mapping"] = {}
            for item in self.manager_table.get_children():
                values = self.manager_table.item(item, "values")
                if values and len(values) >= 2:
                    self.config["manager_mapping"][values[0]] = values[1]
            self.config["region_mapping"] = {}
            for item in self.region_table.get_children():
                values = self.region_table.item(item, "values")
                if values and len(values) >= 2:
                    self.config["region_mapping"][values[0]] = values[1]
            messagebox.showinfo("Успех", "Маппинги сохранены")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

    def add_manager_row(self):
        self.open_manager_editor()

    def edit_manager_row(self):
        selected = self.manager_table.selection()
        if selected:
            item_id = selected[0]
            values = self.manager_table.item(item_id, "values")
            self.open_manager_editor(values, item_id)

    def open_manager_editor(self, values=None, item_id=None):
        editor = tk.Toplevel(self.parent)
        title = "Редактор менеджера" if values else "Добавление менеджера"
        editor.title(title)
        editor.geometry("400x150")
        editor.grab_set()

        ttk.Label(editor, text="Менеджер:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        manager_entry = ttk.Entry(editor, width=30)
        manager_entry.grid(row=0, column=1, padx=10, pady=10, sticky='we')

        ttk.Label(editor, text="Сектор:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        sector_entry = ttk.Entry(editor, width=30)
        sector_entry.grid(row=1, column=1, padx=10, pady=10, sticky='we')

        if values:
            manager_entry.insert(0, values[0])
            sector_entry.insert(0, values[1])

        def save_manager():
            manager = manager_entry.get()
            sector = sector_entry.get()
            if not manager or not sector:
                messagebox.showerror("Ошибка", "Оба поля должны быть заполнены")
                return
            if item_id:
                self.manager_table.item(item_id, values=(manager, sector))
            else:
                self.manager_table.insert("", tk.END, values=(manager, sector))
            self.save()  # Автоматическое сохранение
            editor.destroy()

        btn_frame = ttk.Frame(editor)
        ttk.Button(btn_frame, text="Сохранить", command=save_manager, width=15).grid(row=2, column=0, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=editor.destroy, width=15).grid(row=2, column=1, padx=5)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        editor.grid_columnconfigure(1, weight=1)

    def add_region_row(self):
        self.open_region_editor()

    def edit_region_row(self):
        selected = self.region_table.selection()
        if selected:
            item_id = selected[0]
            values = self.region_table.item(item_id, "values")
            self.open_region_editor(values, item_id)

    def open_region_editor(self, values=None, item_id=None):
        editor = tk.Toplevel(self.parent)
        title = "Редактор региона" if values else "Добавление региона"
        editor.title(title)
        editor.geometry("400x150")
        editor.grab_set()

        ttk.Label(editor, text="Регион:").grid(row=0, column=0, padx=10, pady=10, sticky='w')
        region_entry = ttk.Entry(editor, width=30)
        region_entry.grid(row=0, column=1, padx=10, pady=10, sticky='we')

        ttk.Label(editor, text="Сектор:").grid(row=1, column=0, padx=10, pady=10, sticky='w')
        sector_entry = ttk.Entry(editor, width=30)
        sector_entry.grid(row=1, column=1, padx=10, pady=10, sticky='we')

        if values:
            region_entry.insert(0, values[0])
            sector_entry.insert(0, values[1])

        def save_region():
            region = region_entry.get()
            sector = sector_entry.get()
            if not region or not sector:
                messagebox.showerror("Ошибка", "Оба поля должны быть заполнены")
                return
            if item_id:
                self.region_table.item(item_id, values=(region, sector))
            else:
                self.region_table.insert("", tk.END, values=(region, sector))
            self.save()  # Автоматическое сохранение
            editor.destroy()

        btn_frame = ttk.Frame(editor)
        ttk.Button(btn_frame, text="Сохранить", command=save_region, width=15).grid(row=2, column=0, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=editor.destroy, width=15).grid(row=2, column=1, padx=5)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        editor.grid_columnconfigure(1, weight=1)

    def remove_manager_row(self):
        selected = self.manager_table.selection()
        if selected:
            self.manager_table.delete(selected)

    def remove_region_row(self):
        selected = self.region_table.selection()
        if selected:
            self.region_table.delete(selected)

    def import_data(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл для импорта",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        try:
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='windows-1251', sep=';')
            else:
                df = pd.read_excel(file_path)
            if "manager" in df.columns and "sector" in df.columns:
                self.manager_table.delete(*self.manager_table.get_children())
                for _, row in df.iterrows():
                    self.manager_table.insert("", tk.END, values=(row["manager"], row["sector"]))
            if "region" in df.columns and "sector" in df.columns:
                self.region_table.delete(*self.region_table.get_children())
                for _, row in df.iterrows():
                    self.region_table.insert("", tk.END, values=(row["region"], row["sector"]))
            messagebox.showinfo("Успех", "Данные успешно импортированы")
        except Exception as e:
            messagebox.showerror("Ошибка импорта", str(e))

    def export_data(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить как",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if not file_path:
            return
        try:
            managers = []
            for item in self.manager_table.get_children():
                values = self.manager_table.item(item, "values")
                if values and len(values) >= 2:
                    managers.append({"manager": values[0], "sector": values[1]})
            regions = []
            for item in self.region_table.get_children():
                values = self.region_table.item(item, "values")
                if values and len(values) >= 2:
                    regions.append({"region": values[0], "sector": values[1]})
            if managers:
                pd.DataFrame(managers).to_csv(
                    file_path.replace('.csv', '_managers.csv'), 
                    index=False, 
                    encoding='windows-1251',
                    sep=';'
                )
            if regions:
                pd.DataFrame(regions).to_csv(
                    file_path.replace('.csv', '_regions.csv'), 
                    index=False,
                    encoding='windows-1251',
                    sep=';'
                )
            messagebox.showinfo("Успех", "Данные успешно экспортированы")
        except Exception as e:
            messagebox.showerror("Ошибка экспорта", str(e))

    def close(self):
        self.parent.destroy()

    def toggle_selection(self, event):
        """Обработчик клика мыши для переключения выделения строки"""
        widget = event.widget
        item = widget.identify_row(event.y)
        current_selection = widget.selection()
        
        if item:
            if item in current_selection:
                # Снимаем выделение при повторном нажатии
                widget.selection_remove(item)
            else:
                # Выделяем элемент при первом нажатии
                widget.selection_set(item)
        return "break"  # Предотвращаем стандартное поведение