import os
import json
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog, filedialog
from datetime import datetime
import pythoncom
import win32com.client as wcomcli
from win32com.shell import shell, shellcon

# Константы
CLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
IID_IFolderView = "{CDE725B0-CCC9-4519-917E-325D72FAB4CE}"
SWC_DESKTOP = 0x08
SWFO_NEEDDISPATCH = 0x01

# Типы иконок
ICON_TYPES = {
    "неопознано": {"color": "#888888", "description": "Неизвестный тип"},
    "игра": {"color": "#FF6B6B", "description": "Игровые приложения"},
    "работа": {"color": "#4ECDC4", "description": "Рабочие приложения"},
    "папка": {"color": "#45B7D1", "description": "Папки и директории"},
    "документ": {"color": "#96CEB4", "description": "Документы и файлы"},
    "мультимедиа": {"color": "#FECA57", "description": "Медиа файлы"},
    "система": {"color": "#FF9FF3", "description": "Системные приложения"},
    "интернет": {"color": "#54A0FF", "description": "Браузеры и интернет"}
}



class Shortcut:
    def __init__(self, name, position, pidl, icon_type="неопознано", tags=None,
                 description="", custom_color=None, importance=1):
        self.name = name
        self.position = position
        self.pidl = pidl
        self.icon_type = icon_type
        self.tags = tags or []
        self.description = description
        self.custom_color = custom_color
        self.importance = importance  # 1-5, где 5 - максимальная важность
        self.created = datetime.now().isoformat()
        self.modified = self.created

    def to_dict(self):
        return {
            'name': self.name,
            'position': self.position,
            'pidl': self.pidl,
            'icon_type': self.icon_type,
            'tags': self.tags,
            'description': self.description,
            'custom_color': self.custom_color,
            'importance': self.importance,
            'created': self.created,
            'modified': self.modified
        }

    @classmethod
    def from_dict(cls, data):
        shortcut = cls(
            data['name'],
            data['position'],
            data['pidl'],
            data.get('icon_type', 'неопознано'),
            data.get('tags', []),
            data.get('description', ''),
            data.get('custom_color'),
            data.get('importance', 1)
        )
        shortcut.created = data.get('created', datetime.now().isoformat())
        shortcut.modified = data.get('modified', shortcut.created)
        return shortcut

    def update(self, **kwargs):
        for key, value in kwargs.items():
            if hasattr(self, key):
                setattr(self, key, value)
        self.modified = datetime.now().isoformat()


class DesktopLayout:
    def __init__(self, name, description=""):
        self.name = name
        self.description = description
        self.created = datetime.now().isoformat()
        self.modified = self.created
        self.shortcuts = []
        self.version = "1.0"

    def add_shortcut(self, shortcut):
        self.shortcuts.append(shortcut)
        self.modified = datetime.now().isoformat()

    def remove_shortcut(self, pidl):
        self.shortcuts = [s for s in self.shortcuts if s.pidl != pidl]
        self.modified = datetime.now().isoformat()

    def get_shortcut(self, pidl):
        for shortcut in self.shortcuts:
            if shortcut.pidl == pidl:
                return shortcut
        return None

    def to_dict(self):
        return {
            'name': self.name,
            'description': self.description,
            'created': self.created,
            'modified': self.modified,
            'version': self.version,
            'shortcuts': [s.to_dict() for s in self.shortcuts]
        }

    @classmethod
    def from_dict(cls, data):
        layout = cls(data['name'], data.get('description', ''))
        layout.created = data.get('created', datetime.now().isoformat())
        layout.modified = data.get('modified', layout.created)
        layout.version = data.get('version', '1.0')

        for shortcut_data in data.get('shortcuts', []):
            layout.add_shortcut(Shortcut.from_dict(shortcut_data))

        return layout


class DesktopIconManager:
    def __init__(self):
        self.layouts_dir = "desktop_layouts"
        self.current_layout = None
        self.shell_windows = None
        self.folder_view = None
        self.initialize_com()
        self.ensure_directories()

    def ensure_directories(self):
        """Создает необходимые директории"""
        if not os.path.exists(self.layouts_dir):
            os.makedirs(self.layouts_dir)

    def initialize_com(self):
        """Инициализация COM объектов"""
        try:
            pythoncom.CoInitialize()
            self.shell_windows = wcomcli.Dispatch(CLSID_ShellWindows)
            hwnd = 0
            dispatch = self.shell_windows.FindWindowSW(
                wcomcli.VARIANT(pythoncom.VT_I4, shellcon.CSIDL_DESKTOP),
                wcomcli.VARIANT(pythoncom.VT_EMPTY, None),
                SWC_DESKTOP, hwnd, SWFO_NEEDDISPATCH,
            )
            service_provider = dispatch._oleobj_.QueryInterface(pythoncom.IID_IServiceProvider)
            browser = service_provider.QueryService(shell.SID_STopLevelBrowser, shell.IID_IShellBrowser)
            shell_view = browser.QueryActiveShellView()
            self.folder_view = shell_view.QueryInterface(IID_IFolderView)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось инициализировать COM: {e}")

    def get_desktop_items(self):
        """Получить все элементы рабочего стола"""
        if not self.folder_view:
            self.initialize_com()

        items_data = []
        try:
            items_len = self.folder_view.ItemCount(shellcon.SVGIO_ALLVIEW)
            for i in range(items_len):
                item = self.folder_view.Item(i)
                name = self.get_item_name(item, i)  # передаем индекс
                position = self.folder_view.GetItemPosition(item)
                items_data.append({
                    'index': i,
                    'name': name,
                    'position': (position[0], position[1]),
                    'pidl': str(item)
                })
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось получить элементы: {e}")

        return items_data

    def get_item_name(self, item, index=None):
        """Получить имя элемента"""
        try:
            shell_app = wcomcli.Dispatch("Shell.Application")
            desktop = shell_app.NameSpace(0)  # Desktop

            if index is not None and index < len(list(desktop.Items())):
                return list(desktop.Items())[index].Name
            else:
                # Перебор всех элементов
                for i, shell_item in enumerate(desktop.Items()):
                    if i == index:
                        return shell_item.Name
        except Exception as e:
            print(f"Error getting name for index {index}: {e}")

        return f"Item_{index if index is not None else hash(item)}"

    def create_layout(self, name, description=""):
        """Создать новый layout"""
        layout = DesktopLayout(name, description)
        items = self.get_desktop_items()

        for item in items:
            shortcut = Shortcut(
                name=item['name'],
                position=item['position'],
                pidl=item['pidl']
            )
            layout.add_shortcut(shortcut)

        self.current_layout = layout
        return layout

    def save_layout(self, layout):
        """Сохранить layout в файл"""
        try:
            filename = f"{layout.name.replace(' ', '_')}.json"
            filepath = os.path.join(self.layouts_dir, filename)

            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(layout.to_dict(), f, indent=2, ensure_ascii=False)

            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить layout: {e}")
            return False

    def load_layout(self, filename):
        """Загрузить layout из файлa"""
        try:
            filepath = os.path.join(self.layouts_dir, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self.current_layout = DesktopLayout.from_dict(data)
            return self.current_layout
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить layout: {e}")
            return None

    def get_saved_layouts(self):
        """Получить список сохраненных layouts"""
        layouts = []
        try:
            for filename in os.listdir(self.layouts_dir):
                if filename.endswith('.json'):
                    layouts.append(filename)
        except:
            pass
        return layouts

    def delete_layout(self, filename):
        """Удалить layout"""
        try:
            filepath = os.path.join(self.layouts_dir, filename)
            if os.path.exists(filepath):
                os.remove(filepath)
                return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось удалить layout: {e}")
        return False

    def restore_layout(self, layout):
        """Восстановить layout на рабочем столе"""
        if not layout:
            return 0

        restored_count = 0
        current_items = self.get_desktop_items()

        for shortcut in layout.shortcuts:
            for current_item in current_items:
                if (shortcut.name == current_item['name'] or
                        shortcut.pidl == current_item['pidl']):
                    try:
                        self.folder_view.SelectAndPositionItem(
                            current_item['index'],
                            shortcut.position,
                            shellcon.SVSI_POSITIONITEM
                        )
                        restored_count += 1
                        break
                    except Exception as e:
                        print(f"Ошибка при восстановлении {shortcut.name}: {e}")

        return restored_count


class IconEditorDialog(simpledialog.Dialog):
    def __init__(self, parent, shortcut, title="Редактирование ярлыка"):
        self.shortcut = shortcut
        super().__init__(parent, title)

    def body(self, frame):
        ttk.Label(frame, text="Название:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.name_var = tk.StringVar(value=self.shortcut.name)
        ttk.Entry(frame, textvariable=self.name_var, width=30).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Тип:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.type_var = tk.StringVar(value=self.shortcut.icon_type)
        type_combo = ttk.Combobox(frame, textvariable=self.type_var, values=list(ICON_TYPES.keys()))
        type_combo.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Важность (1-5):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.importance_var = tk.IntVar(value=self.shortcut.importance)
        ttk.Spinbox(frame, from_=1, to=5, textvariable=self.importance_var, width=5).grid(row=2, column=1, padx=5,
                                                                                          pady=5)

        ttk.Label(frame, text="Описание:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.desc_var = tk.StringVar(value=self.shortcut.description)
        ttk.Entry(frame, textvariable=self.desc_var, width=30).grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(frame, text="Теги (через запятую):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.tags_var = tk.StringVar(value=", ".join(self.shortcut.tags))
        ttk.Entry(frame, textvariable=self.tags_var, width=30).grid(row=4, column=1, padx=5, pady=5)

        return frame

    def apply(self):
        self.shortcut.update(
            name=self.name_var.get(),
            icon_type=self.type_var.get(),
            importance=self.importance_var.get(),
            description=self.desc_var.get(),
            tags=[tag.strip() for tag in self.tags_var.get().split(',') if tag.strip()]
        )


class DesktopIconApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Продвинутый менеджер рабочего стола")
        self.root.geometry("800x600")

        self.manager = DesktopIconManager()
        self.current_layout = None

        self.setup_ui()
        self.refresh_layouts_list()

    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        # Основные фреймы
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Левая панель - управление layouts
        left_frame = ttk.LabelFrame(main_frame, text="Управление сохранениями", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # Правая панель - редактирование
        right_frame = ttk.LabelFrame(main_frame, text="Редактирование ярлыков", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Левая панель - элементы
        ttk.Button(left_frame, text="Создать новое сохранение",
                   command=self.create_new_layout).pack(pady=5, fill=tk.X)

        ttk.Label(left_frame, text="Сохраненные layouts:").pack(pady=(10, 5), anchor="w")

        self.layouts_listbox = tk.Listbox(left_frame, height=10)
        self.layouts_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        self.layouts_listbox.bind('<<ListboxSelect>>', self.on_layout_select)

        button_frame = ttk.Frame(left_frame)
        button_frame.pack(fill=tk.X, pady=5)

        ttk.Button(button_frame, text="Загрузить",
                   command=self.load_selected_layout).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(button_frame, text="Удалить",
                   command=self.delete_selected_layout).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=2)

        ttk.Button(left_frame, text="Восстановить на рабочий стол",
                   command=self.restore_current_layout).pack(pady=5, fill=tk.X)

        # Правая панель - Treeview для ярлыков
        self.setup_shortcuts_tree(right_frame)

        # Статусная строка
        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = ttk.Label(main_frame, textvariable=self.status_var,
                                 foreground="green", font=('Arial', 9))
        status_label.pack(side=tk.BOTTOM, fill=tk.X, pady=5)

    def setup_shortcuts_tree(self, parent):
        """Настройка TreeView для ярлыков"""
        # Создаем фрейм для TreeView и скроллбара
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Создаем TreeView с правильными колонками
        self.shortcuts_tree = ttk.Treeview(tree_frame, columns=('name', 'type', 'importance', 'description'),
                                           show='headings', height=15)

        # Настраиваем заголовки колонок
        self.shortcuts_tree.heading('name', text='Название')
        self.shortcuts_tree.heading('type', text='Тип')
        self.shortcuts_tree.heading('importance', text='Важность')
        self.shortcuts_tree.heading('description', text='Описание')

        # Настраиваем ширину колонок
        self.shortcuts_tree.column('name', width=150)
        self.shortcuts_tree.column('type', width=100)
        self.shortcuts_tree.column('importance', width=80)
        self.shortcuts_tree.column('description', width=200)

        # Добавляем скроллбар
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.shortcuts_tree.yview)
        self.shortcuts_tree.configure(yscrollcommand=scrollbar.set)

        # Размещаем TreeView и скроллбар
        self.shortcuts_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Биндим события
        self.shortcuts_tree.bind('<Double-1>', self.on_shortcut_double_click)

        # Кнопки редактирования
        edit_frame = ttk.Frame(parent)
        edit_frame.pack(fill=tk.X, pady=5)

        ttk.Button(edit_frame, text="Редактировать",
                   command=self.edit_selected_shortcut).pack(side=tk.LEFT, padx=2)
        ttk.Button(edit_frame, text="Сохранить изменения",
                   command=self.save_current_layout).pack(side=tk.RIGHT, padx=2)

    def refresh_layouts_list(self):
        """Обновить список layouts"""
        self.layouts_listbox.delete(0, tk.END)
        for layout_file in self.manager.get_saved_layouts():
            self.layouts_listbox.insert(tk.END, layout_file)

    def create_new_layout(self):
        """Создать новое сохранение"""
        name = simpledialog.askstring("Новое сохранение", "Введите название сохранения:")
        if name:
            description = simpledialog.askstring("Описание", "Введите описание (необязательно):")
            layout = self.manager.create_layout(name, description or "")
            if self.manager.save_layout(layout):
                self.current_layout = layout
                self.refresh_layouts_list()
                self.update_shortcuts_tree()
                self.status_var.set(f"Создано сохранение: {name}")

    def load_selected_layout(self):
        """Загрузить выбранное сохранение"""
        selection = self.layouts_listbox.curselection()
        if selection:
            filename = self.layouts_listbox.get(selection[0])
            layout = self.manager.load_layout(filename)
            if layout:
                self.current_layout = layout
                self.update_shortcuts_tree()
                self.status_var.set(f"Загружено: {layout.name}")

    def delete_selected_layout(self):
        """Удалить выбранное сохранение"""
        selection = self.layouts_listbox.curselection()
        if selection:
            filename = self.layouts_listbox.get(selection[0])
            if messagebox.askyesno("Подтверждение", f"Удалить сохранение {filename}?"):
                if self.manager.delete_layout(filename):
                    self.refresh_layouts_list()
                    self.status_var.set(f"Удалено: {filename}")

    def restore_current_layout(self):
        """Восстановить текущий layout на рабочий стол"""
        if self.current_layout:
            count = self.manager.restore_layout(self.current_layout)
            self.status_var.set(f"Восстановлено {count} иконок")
        else:
            messagebox.showwarning("Внимание", "Сначала загрузите сохранение!")

    def update_shortcuts_tree(self):
        """Обновить дерево ярлыков"""
        self.shortcuts_tree.delete(*self.shortcuts_tree.get_children())
        if self.current_layout:
            for shortcut in self.current_layout.shortcuts:
                self.shortcuts_tree.insert('', 'end', values=(
                    shortcut.name,
                    shortcut.icon_type,
                    shortcut.importance,
                    shortcut.description
                ))

    def on_shortcut_double_click(self, event):
        """Обработчик двойного клика по ярлыку"""
        self.edit_selected_shortcut()

    def edit_selected_shortcut(self):
        """Редактировать выбранный ярлык"""
        selection = self.shortcuts_tree.selection()
        if selection and self.current_layout:
            item = self.shortcuts_tree.item(selection[0])
            shortcut_name = item['values'][0]

            # Найти shortcut по имени
            for shortcut in self.current_layout.shortcuts:
                if shortcut.name == shortcut_name:
                    dialog = IconEditorDialog(self.root, shortcut)
                    if dialog.result:
                        self.update_shortcuts_tree()
                        self.status_var.set(f"Обновлен: {shortcut.name}")
                    break

    def save_current_layout(self):
        """Сохранить текущий layout"""
        if self.current_layout:
            if self.manager.save_layout(self.current_layout):
                self.status_var.set(f"Сохранено: {self.current_layout.name}")
            else:
                self.status_var.set("Ошибка сохранения")
        else:
            messagebox.showwarning("Внимание", "Нет активного сохранения!")

    def on_layout_select(self, event):
        """Обработчик выбора layout"""
        self.load_selected_layout()


def main():
    """Основная функция"""
    pythoncom.CoInitialize()

    try:
        root = tk.Tk()
        app = DesktopIconApp(root)
        root.mainloop()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()