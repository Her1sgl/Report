import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from core.config_manager import save_config

class ConfigEditor:
    def __init__(self, parent, config):
        self.parent = parent
        self.config = config
        self.points_lists = {}
        self.create_widgets()
        self.load_data()
        # Привязываем обработчик клика к таблицам
        self.region_table.bind('<Button-1>', self.toggle_selection)
        self.points_table.bind('<Button-1>', self.toggle_selection)

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.parent)
        self.notebook.pack(fill='both', expand=True)

        # Основные настройки
        self.general_frame = ttk.Frame(self.notebook)
        self.create_general_tab(self.general_frame)
        self.notebook.add(self.general_frame, text="Основные настройки")

        # Таблицы отчета
        self.tables_frame = ttk.Frame(self.notebook)
        self.create_tables_tab(self.tables_frame)
        self.notebook.add(self.tables_frame, text="Таблицы отчета")

        # Кнопки
        btn_frame = ttk.Frame(self.parent)
        ttk.Button(btn_frame, text="Сохранить", command=self.save, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Закрыть", command=self.close, width=15).pack(side=tk.LEFT, padx=5)
        btn_frame.pack(pady=10)

    def create_general_tab(self, frame):
        fields = [
            ("Столбец региона:", "region_col"),
            ("Столбец менеджера:", "manager_col"),
            ("Столбец пункта:", "point_col"),
            ("Столбец ЛЧМ продаж:", "bms_col"),
            ("Столбец ЛЦМ продаж:", "fms_col")
        ]
        self.entries = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky='w', padx=10, pady=5)
            entry = ttk.Entry(frame, width=30)
            entry.grid(row=i, column=1, sticky='we', padx=5, pady=5)
            self.entries[key] = entry

        ttk.Label(frame, text="Метод группировки:").grid(row=5, column=0, sticky='w', padx=10, pady=5)
        self.grouping_method = ttk.Combobox(frame, values=["manager", "region"], width=10)
        self.grouping_method.grid(row=5, column=1, sticky='w', padx=5, pady=5)

        frame.grid_columnconfigure(1, weight=1)
        frame.pack_propagate(False)

    def create_tables_tab(self, frame):
        # Таблицы регионов
        region_tab_frame = ttk.LabelFrame(frame, text="Таблицы по регионам")
        region_tab_frame.pack(fill='both', expand=True, padx=5, pady=5)

        columns = ("name", "type", "day_row", "data_start_row", "data_end_row", "region_col", "day_start_col", "day_end_col")
        self.region_table = ttk.Treeview(region_tab_frame, columns=columns, show='headings', height=5)
        # Настройка колонок
        col_widths = [120, 70, 70, 120, 120, 90, 120, 120]
        for i, col in enumerate(columns):
            col_name = col.replace('_', ' ').title()
            self.region_table.heading(col, text=col_name)
            self.region_table.column(col, width=col_widths[i], anchor=tk.CENTER)

        self.region_table.pack(fill='both', expand=True, padx=5, pady=5)

        btn_frame = ttk.Frame(region_tab_frame)
        ttk.Button(btn_frame, text="Добавить", command=self.add_region_table, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_region_table, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Редактировать", command=self.edit_region_table, width=15).pack(side=tk.LEFT, padx=2)
        btn_frame.pack(pady=5)

        # Таблицы пунктов
        points_tab_frame = ttk.LabelFrame(frame, text="Таблицы по пунктам")
        points_tab_frame.pack(fill='both', expand=True, padx=5, pady=5)

        columns = ("name", "type", "start_row", "end_row", "point_col", "data_col")
        self.points_table = ttk.Treeview(points_tab_frame, columns=columns, show='headings', height=5)
        # Настройка колонок
        col_widths = [120, 70, 80, 80, 80, 80]
        for i, col in enumerate(columns):
            col_name = col.replace('_', ' ').title()
            self.points_table.heading(col, text=col_name)
            self.points_table.column(col, width=col_widths[i], anchor=tk.CENTER)

        self.points_table.pack(fill='both', expand=True, padx=5, pady=5)

        btn_frame = ttk.Frame(points_tab_frame)
        ttk.Button(btn_frame, text="Добавить", command=self.add_points_table, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_points_table, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Редактировать", command=self.edit_points_table, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Редактировать пункты", command=self.edit_points_list, width=19).pack(side=tk.LEFT, padx=2)
        btn_frame.pack(pady=5)

    def load_data(self):
        source_cols = self.config["source_columns"]
        self.entries["region_col"].insert(0, source_cols["region"])
        self.entries["manager_col"].insert(0, source_cols["manager"])
        self.entries["point_col"].insert(0, source_cols["point"])
        self.entries["bms_col"].insert(0, source_cols["bms_sales"])
        self.entries["fms_col"].insert(0, source_cols["fms_sales"])
        self.grouping_method.set(self.config["grouping_method"])

        # Загрузка таблиц регионов
        for table in self.config["region_tables"]:
            self.region_table.insert("", tk.END, values=(
                table["name"],
                table["type"],
                table["day_row"],
                table["data_start_row"],
                table["data_end_row"],
                table["region_col"],
                table["day_start_col"],
                table["day_end_col"]
            ))

        # Загрузка таблиц пунктов
        for table in self.config["new_points_tables"]:
            item_id = self.points_table.insert("", tk.END, values=(
                table["name"],
                table["type"],
                table["start_row"],
                table["end_row"],
                table["point_col"],
                table["data_col"]
            ))
            self.points_lists[item_id] = table.get("point_names", [])

    def save(self):
        try:
            self.config["source_columns"] = {
                "region": self.entries["region_col"].get(),
                "manager": self.entries["manager_col"].get(),
                "point": self.entries["point_col"].get(),
                "bms_sales": self.entries["bms_col"].get(),
                "fms_sales": self.entries["fms_col"].get()
            }
            self.config["grouping_method"] = self.grouping_method.get()

            # Сохранение таблиц регионов
            self.config["region_tables"] = []
            for item in self.region_table.get_children():
                values = self.region_table.item(item, "values")
                self.config["region_tables"].append({
                    "name": values[0],
                    "type": values[1],
                    "day_row": int(values[2]),
                    "data_start_row": int(values[3]),
                    "data_end_row": int(values[4]),
                    "region_col": values[5],
                    "day_start_col": values[6],
                    "day_end_col": values[7]
                })

            # Сохранение таблиц пунктов
            self.config["new_points_tables"] = []
            for item in self.points_table.get_children():
                values = self.points_table.item(item, "values")
                point_names = self.points_lists.get(item, [])
                self.config["new_points_tables"].append({
                    "name": values[0],
                    "type": values[1],
                    "point_names": point_names,
                    "start_row": int(values[2]),
                    "end_row": int(values[3]),
                    "point_col": values[4],
                    "data_col": values[5]
                })

            if save_config(self.config):
                messagebox.showinfo("Успех", "Конфигурация сохранена")
            else:
                messagebox.showerror("Ошибка", "Не удалось сохранить конфигурацию")
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

    def add_region_table(self):
        self.open_region_table_editor()

    def edit_region_table(self):
        selected = self.region_table.selection()
        if selected:
            item_id = selected[0]
            values = self.region_table.item(item_id, "values")
            self.open_region_table_editor(values, item_id)

    def open_region_table_editor(self, values=None, item_id=None):
        editor = tk.Toplevel(self.parent)
        editor.title("Редактор таблицы по регионам" if values else "Добавление таблицы по регионам")
        editor.geometry("500x400")
        editor.grab_set()

        fields = [
            ("Название таблицы", "name"),
            ("Тип данных (bms/fms)", "type"),
            ("Строка дня", "day_row"),
            ("Начальная строка данных", "data_start_row"),
            ("Конечная строка данных", "data_end_row"),
            ("Столбец региона", "region_col"),
            ("Начальный столбец дня", "day_start_col"),
            ("Конечный столбец дня", "day_end_col")
        ]
        entries = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(editor, text=label).grid(row=i, column=0, padx=10, pady=5, sticky='w')
            entry = ttk.Entry(editor, width=30)
            # Валидация для числовых полей
            if key in ["day_row", "data_start_row", "data_end_row"]:
                validate_cmd = (editor.register(lambda s: s.isdigit() or s == ""), '%P')
                entry.config(validate="key", validatecommand=validate_cmd)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')
            if values:
                entry.insert(0, values[i])
            elif key == "type":
                entry.insert(0, "bms")
            else:
                entry.insert(0, "")
            entries[key] = entry

        def save_table():
            new_values = [entries[key].get() for key in [f[1] for f in fields]]
            # Проверка обязательных полей
            if not new_values[0] or not new_values[1]:
                messagebox.showerror("Ошибка", "Название и тип таблицы обязательны")
                return
            try:
                # Проверка числовых полей
                int(new_values[2])  # day_row
                int(new_values[3])  # data_start_row
                int(new_values[4])  # data_end_row
            except ValueError:
                messagebox.showerror("Ошибка", "Поля строк должны быть числами")
                return
            if item_id:
                self.region_table.item(item_id, values=new_values)
            else:
                self.region_table.insert("", tk.END, values=new_values)
            self.save()  # Автоматическое сохранение
            editor.destroy()

        btn_frame = ttk.Frame(editor)
        ttk.Button(btn_frame, text="Сохранить", command=save_table, width=15).grid(row=len(fields), column=0, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=editor.destroy, width=15).grid(row=len(fields), column=1, padx=5)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)
        editor.grid_columnconfigure(1, weight=1)

    def add_points_table(self):
        self.open_points_table_editor()

    def edit_points_table(self):
        selected = self.points_table.selection()
        if selected:
            item_id = selected[0]
            values = self.points_table.item(item_id, "values")
            self.open_points_table_editor(values, item_id)

    def open_points_table_editor(self, values=None, item_id=None):
        editor = tk.Toplevel(self.parent)
        editor.title("Редактор таблицы по пунктам" if values else "Добавление таблицы по пунктам")
        editor.geometry("500x300")
        editor.grab_set()

        fields = [
            ("Название таблицы", "name"),
            ("Тип данных (bms/fms)", "type"),
            ("Начальная строка", "start_row"),
            ("Конечная строка", "end_row"),
            ("Столбец пункта", "point_col"),
            ("Столбец данных", "data_col")
        ]
        entries = {}
        for i, (label, key) in enumerate(fields):
            ttk.Label(editor, text=label).grid(row=i, column=0, padx=10, pady=5, sticky='w')
            entry = ttk.Entry(editor, width=30)
            # Валидация для числовых полей
            if key in ["start_row", "end_row"]:
                validate_cmd = (editor.register(lambda s: s.isdigit() or s == ""), '%P')
                entry.config(validate="key", validatecommand=validate_cmd)
            entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')
            if values:
                entry.insert(0, values[i])
            elif key == "type":
                entry.insert(0, "bms")
            else:
                entry.insert(0, "")
            entries[key] = entry

        def save_table():
            new_values = [entries[key].get() for key in [f[1] for f in fields]]
            # Проверка обязательных полей
            if not new_values[0] or not new_values[1]:
                messagebox.showerror("Ошибка", "Название и тип таблицы обязательны")
                return
            try:
                int(new_values[2])  # start_row
                int(new_values[3])  # end_row
            except ValueError:
                messagebox.showerror("Ошибка", "Поля строк должны быть числами")
                return
            if item_id:
                self.points_table.item(item_id, values=new_values)
            else:
                self.points_table.insert("", tk.END, values=new_values)
            self.save()  # Автоматическое сохранение
            editor.destroy()

        btn_frame = ttk.Frame(editor)
        ttk.Button(btn_frame, text="Сохранить", command=save_table, width=15).grid(row=len(fields), column=0, padx=5)
        ttk.Button(btn_frame, text="Отмена", command=editor.destroy, width=15).grid(row=len(fields), column=1, padx=5)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=10)
        editor.grid_columnconfigure(1, weight=1)

    def remove_region_table(self):
        selected = self.region_table.selection()
        if selected:
            self.region_table.delete(selected)

    def remove_points_table(self):
        selected = self.points_table.selection()
        if selected:
            self.points_table.delete(selected)

    def edit_points_list(self):
        selected = self.points_table.selection()
        if not selected:
            messagebox.showinfo("Информация", "Выберите таблицу для редактирования списка пунктов")
            return
        item_id = selected[0]
        table_name = self.points_table.item(item_id, "values")[0]
        current_points = self.points_lists.get(item_id, [])

        # Создаем окно редактора
        editor = tk.Toplevel(self.parent)
        editor.title(f"Редактор пунктов: {table_name}")
        editor.geometry("600x500")
        editor.grab_set()

        # Основной фрейм
        main_frame = ttk.Frame(editor)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Фрейм для поиска
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill=tk.X, pady=5)
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(search_frame, text="Найти", command=self.search_points).pack(side=tk.LEFT, padx=5)
        search_entry.bind('<Return>', lambda event: self.search_points())

        # Список пунктов
        list_frame = ttk.LabelFrame(main_frame, text="Список пунктов")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.points_listbox = tk.Listbox(
            list_frame, 
            yscrollcommand=scrollbar.set, 
            height=15, 
            selectmode=tk.EXTENDED,
            font=("Arial", 10)
        )
        self.points_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.config(command=self.points_listbox.yview)
        
        # Привязываем обработчик клика к списку пунктов
        self.points_listbox.bind('<Button-1>', self.toggle_listbox_selection)

        # Заполняем список
        for point in current_points:
            self.points_listbox.insert(tk.END, point)

        # Кнопки управления
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="Добавить", command=self.add_point_to_list, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_point_from_list, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="Редактировать", command=self.edit_point_in_list, width=15).pack(side=tk.LEFT, padx=2)

        # Кнопки сохранения/отмены
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=5)
        ttk.Button(bottom_frame, text="Сохранить", 
                  command=lambda: self.save_points_list(item_id, editor), width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="Отмена", command=editor.destroy, width=15).pack(side=tk.LEFT, padx=5)

    def search_points(self):
        search_term = self.search_var.get().lower()
        if not search_term:
            return
        # Снимаем выделение
        self.points_listbox.selection_clear(0, tk.END)
        # Ищем совпадения
        found = False
        for i in range(self.points_listbox.size()):
            point = self.points_listbox.get(i).lower()
            if search_term in point:
                self.points_listbox.selection_set(i)
                self.points_listbox.see(i)
                found = True
        if not found:
            messagebox.showinfo("Поиск", "Совпадений не найдено")

    def add_point_to_list(self):
        point = simpledialog.askstring("Добавить пункт", "Введите название пункта:")
        if point:
            self.points_listbox.insert(tk.END, point)

    def remove_point_from_list(self):
        selected = self.points_listbox.curselection()
        if selected:
            # Удаляем все выделенные элементы, начиная с последнего
            for index in reversed(selected):
                self.points_listbox.delete(index)

    def edit_point_in_list(self):
        selected = self.points_listbox.curselection()
        if not selected:
            return
        # Редактируем только первый выделенный элемент
        index = selected[0]
        old_point = self.points_listbox.get(index)
        new_point = simpledialog.askstring("Редактировать пункт", "Введите новое название:", initialvalue=old_point)
        if new_point:
            self.points_listbox.delete(index)
            self.points_listbox.insert(index, new_point)

    def save_points_list(self, item_id, editor):
        points = list(self.points_listbox.get(0, tk.END))
        self.points_lists[item_id] = points
        # Сохраняем всю конфигурацию
        self.save()  # Автоматическое сохранение
        messagebox.showinfo("Успех", "Список пунктов и конфигурация сохранены")
        editor.destroy()

    def close(self):
        self.parent.destroy()

    def toggle_selection(self, event):
        """Обработчик клика мыши для переключения выделения строки в таблицах"""
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

    def toggle_listbox_selection(self, event):
        """Обработчик клика мыши для переключения выделения элемента в списке пунктов"""
        widget = event.widget
        index = widget.nearest(event.y)
        current_selection = widget.curselection()
        
        if index >= 0:
            if index in current_selection:
                # Снимаем выделение при повторном нажатии
                widget.selection_clear(index)
            else:
                # Выделяем элемент при первом нажатии
                widget.selection_set(index)
        return "break"  # Предотвращаем стандартное поведение