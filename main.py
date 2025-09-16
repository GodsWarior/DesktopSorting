import os
import json
import tkinter as tk
from tkinter import messagebox, ttk
import pythoncom
import win32com.client as wcomcli
from win32com.shell import shell, shellcon

# Константы
CLSID_ShellWindows = "{9BA05972-F6A8-11CF-A442-00A0C90A8F39}"
IID_IFolderView = "{CDE725B0-CCC9-4519-917E-325D72FAB4CE}"
SWC_DESKTOP = 0x08
SWFO_NEEDDISPATCH = 0x01


class DesktopIconManager:
    def __init__(self):
        self.positions_file = "desktop_positions.json"
        self.shell_windows = None
        self.folder_view = None
        self.initialize_com()

    def initialize_com(self):
        """Инициализация COM объектов"""
        try:
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
                name = self.get_item_name(item)
                position = self.folder_view.GetItemPosition(item)
                items_data.append({
                    'index': i,
                    'name': name,
                    'position': (position[0], position[1]),
                    'pidl': str(item)  # Уникальный идентификатор элемента
                })
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось получить элементы: {e}")

        return items_data

    def get_item_name(self, item):
        """Получить имя элемента"""
        try:
            # Попробуем получить имя через разные методы
            desktop_folder = shell.SHGetDesktopFolder()
            return desktop_folder.GetDisplayNameOf(item, shellcon.SHGDN_NORMAL)
        except:
            return f"Item_{hash(item)}"

    def save_positions(self):
        """Сохранить текущие позиции иконок"""
        try:
            items = self.get_desktop_items()
            with open(self.positions_file, 'w', encoding='utf-8') as f:
                json.dump(items, f, indent=2, ensure_ascii=False)
            return len(items)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить позиции: {e}")
            return 0

    def restore_positions(self):
        """Восстановить сохраненные позиции иконок"""
        try:
            if not os.path.exists(self.positions_file):
                messagebox.showwarning("Внимание", "Файл с позициями не найден!")
                return 0

            with open(self.positions_file, 'r', encoding='utf-8') as f:
                saved_items = json.load(f)

            current_items = self.get_desktop_items()
            restored_count = 0

            for saved_item in saved_items:
                for current_item in current_items:
                    # Сравниваем по имени и PIDL для надежности
                    if (saved_item['name'] == current_item['name'] or
                            saved_item['pidl'] == current_item['pidl']):
                        try:
                            self.folder_view.SelectAndPositionItem(
                                current_item['index'],  # Можно использовать индекс
                                saved_item['position'],
                                shellcon.SVSI_POSITIONITEM
                            )
                            restored_count += 1
                            break
                        except Exception as e:
                            print(f"Ошибка при восстановлении {saved_item['name']}: {e}")

            return restored_count

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось восстановить позиции: {e}")
            return 0


class DesktopIconApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Менеджер иконок рабочего стола")
        self.root.geometry("400x200")
        self.root.resizable(False, False)

        self.manager = DesktopIconManager()

        self.setup_ui()

    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        # Стиль
        style = ttk.Style()
        style.configure('TButton', font=('Arial', 12), padding=10)
        style.configure('TLabel', font=('Arial', 10))

        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = ttk.Label(main_frame, text="Управление иконками рабочего стола",
                                font=('Arial', 14, 'bold'))
        title_label.pack(pady=(0, 20))

        # Кнопка сохранения
        save_btn = ttk.Button(main_frame, text="Сохранить текущее расположение",
                              command=self.on_save)
        save_btn.pack(pady=10, fill=tk.X)

        # Кнопка восстановления
        restore_btn = ttk.Button(main_frame, text="Восстановить расположение",
                                 command=self.on_restore)
        restore_btn.pack(pady=10, fill=tk.X)

        # Статусная метка
        self.status_label = ttk.Label(main_frame, text="Готов к работе",
                                      foreground="green")
        self.status_label.pack(pady=10)

        # Информация
        info_label = ttk.Label(main_frame, text="Для работы требуются права администратора",
                               foreground="gray", font=('Arial', 8))
        info_label.pack(side=tk.BOTTOM, pady=5)

    def on_save(self):
        """Обработчик сохранения позиций"""
        count = self.manager.save_positions()
        if count > 0:
            self.status_label.config(
                text=f"Сохранено {count} иконок",
                foreground="green"
            )
            messagebox.showinfo("Успех", f"Позиции {count} иконок сохранены!")
        else:
            self.status_label.config(
                text="Ошибка сохранения",
                foreground="red"
            )

    def on_restore(self):
        """Обработчик восстановления позиций"""
        count = self.manager.restore_positions()
        if count > 0:
            self.status_label.config(
                text=f"Восстановлено {count} иконок",
                foreground="green"
            )
            messagebox.showinfo("Успех", f"Позиции {count} иконок восстановлены!")
        else:
            self.status_label.config(
                text="Не удалось восстановить",
                foreground="red"
            )


def main():
    """Основная функция"""
    # Инициализация COM
    pythoncom.CoInitialize()

    try:
        root = tk.Tk()
        app = DesktopIconApp(root)
        root.mainloop()
    finally:
        # Освобождение COM
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()