import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import threading
import subprocess
from datetime import datetime


class DashboardAssemblyWBGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dashboard Assembly WB - Графический интерфейс")
        self.root.geometry("1000x700")
        self.root.resizable(False, False)
        
        # Переменные для хранения путей
        self.folder_path_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!")
        self.folder_path_features_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Дашбоард по рекламным кампаниям")
        self.folder_path_for_db_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям")
        self.folder_path_dudl_var = tk.StringVar(value=r"\\kari.local\public\all\Агрегаторы\Дашборд реклама WB_OZ")
        
        # Создание интерфейса
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок
        title_frame = tk.Frame(self.root, bg="#8e44ad", pady=15)
        title_frame.pack(fill=tk.X)
        
        # Контейнер для заголовка и кнопки справки
        title_container = tk.Frame(title_frame, bg="#8e44ad")
        title_container.pack(fill=tk.X, padx=20)
        
        title_label = tk.Label(
            title_container,
            text="🎛️ Dashboard Assembly WB - Панель управления",
            font=("Arial", 16, "bold"),
            bg="#8e44ad",
            fg="white"
        )
        title_label.pack(side=tk.LEFT, expand=True)
        
        # Кнопка справки
        help_button = tk.Button(
            title_container,
            text="❓ Справка",
            command=self.show_help,
            font=("Arial", 10, "bold"),
            bg="#9b59b6",
            fg="white",
            activebackground="#8e44ad",
            activeforeground="white",
            cursor="hand2",
            padx=15,
            pady=5
        )
        help_button.pack(side=tk.RIGHT)
        
        # Основная рамка для содержимого
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Информация о скрипте
        info_frame = tk.LabelFrame(main_frame, text="О скрипте", font=("Arial", 12, "bold"), padx=10, pady=10)
        info_frame.pack(fill=tk.X, pady=(0, 15))
        
        info_text = tk.Label(
            info_frame,
            text="Скрипт обрабатывает данные Wildberries: воронки, затраты, справочники, остатки и цены.\nСобирает полную базу данных и обновляет Excel файлы.",
            font=("Arial", 9),
            justify=tk.LEFT,
            fg="#555"
        )
        info_text.pack(anchor=tk.W)
        
        # --- Раздел настройки путей ---
        paths_frame = tk.LabelFrame(main_frame, text="Настройка путей к папкам", font=("Arial", 12, "bold"), padx=10, pady=10)
        paths_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # FOLDER_PATH
        self.create_path_row(paths_frame, "FOLDER_PATH (основная папка исходников):", self.folder_path_var, 0)
        
        # FOLDER_PATH_FEATURES
        self.create_path_row(paths_frame, "FOLDER_PATH_FEATURES (папка для сохранения дашборда):", self.folder_path_features_var, 1)
        
        # FOLDER_PATH_FOR_DB
        self.create_path_row(paths_frame, "FOLDER_PATH_FOR_DB (папка для сохранения базы данных):", self.folder_path_for_db_var, 2)
        
        # FOLDER_PATH_DUDL
        self.create_path_row(paths_frame, "FOLDER_PATH_DUDL (вторая папка для сохранения дашборда):", self.folder_path_dudl_var, 3)
        
        # --- Кнопка запуска ---
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        run_button = tk.Button(
            button_frame,
            text="▶️ Запустить сборку базы данных WB",
            command=self.run_assembly,
            font=("Arial", 12, "bold"),
            bg="#8e44ad",
            fg="white",
            activebackground="#8e44ad",
            activeforeground="white",
            pady=10,
            cursor="hand2"
        )
        run_button.pack(fill=tk.X)
        
        # --- Лог-консоль ---
        log_frame = tk.LabelFrame(main_frame, text="Лог выполнения", font=("Arial", 12, "bold"), padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Текстовое поле для логов
        self.log_text = tk.Text(log_frame, height=12, wrap=tk.WORD, state=tk.DISABLED, bg="#e8f8f5", font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Скроллбар
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
    def create_path_row(self, parent, label_text, variable, row):
        """Создает строку с меткой, полем ввода и кнопкой выбора папки"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        label = tk.Label(frame, text=label_text, font=("Arial", 8), width=55, anchor=tk.W)
        label.pack(side=tk.LEFT, padx=(0, 5))
        
        entry = tk.Entry(frame, textvariable=variable, font=("Arial", 8))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_button = tk.Button(
            frame,
            text="📁 Обзор",
            command=lambda: self.browse_folder(variable),
            font=("Arial", 8),
            bg="#8e44ad",
            fg="white",
            activebackground="#8e44ad",
            activeforeground="white",
            cursor="hand2"
        )
        browse_button.pack(side=tk.RIGHT)
        
    def browse_folder(self, variable):
        """Открывает диалог выбора папки"""
        folder_path = filedialog.askdirectory()
        if folder_path:
            variable.set(folder_path)
    
    def show_help(self):
        """Открывает окно с подробной справкой"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Справка - Dashboard Assembly WB")
        help_window.geometry("700x600")
        help_window.resizable(True, True)
        
        # Заголовок окна справки
        header_frame = tk.Frame(help_window, bg="#8e44ad", pady=10)
        header_frame.pack(fill=tk.X)
        
        header_label = tk.Label(
            header_frame,
            text="📖 Справка по использованию",
            font=("Arial", 14, "bold"),
            bg="#8e44ad",
            fg="white"
        )
        header_label.pack()
        
        # Текстовое поле с прокруткой
        text_frame = tk.Frame(help_window, padx=20, pady=20)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        help_text = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Arial", 10),
            bg="#f9f9f9",
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(text_frame, command=help_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        help_text.config(yscrollcommand=scrollbar.set)
        
        # Содержимое справки
        help_content = """🔹 ОПИСАНИЕ ПРИЛОЖЕНИЯ

Это приложение запускает скрипт DashboardAssemblyWB.py для сборки полной базы данных Wildberries.
Скрипт обрабатывает воронки, затраты, справочники, остатки и цены.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ ПОРЯДОК РАБОТЫ

1. Подготовьте файлы дашборда с помощью утилиты Dashboard_Preprocessing
2. Проверьте все 4 пути к папкам (при необходимости измените)
3. Закройте все файлы Excel в этих папках
4. Нажмите "▶️ Запустить сборку базы данных WB"
5. Откроется окно консоли - следите за прогрессом
6. Дождитесь завершения (может занять 30-50 минут)
7. Проверьте результаты в папках FOLDER_FEATURES, FOLDER_PATH_FOR_DB и FOLDER_PATH_DUDL

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📂 ОПИСАНИЕ ПУТЕЙ К ПАПКАМ

1️⃣ FOLDER_PATH (Основная папка исходников)
   Назначение: Главная папка с исходными файлами для дашборда.
   Что хранится: Все основные файлы Excel для сборки базы данных.
   Пример: \\\\kari.local\\...\\Дашбоард\\!!!ИСХОДНИКИ...

2️⃣ FOLDER_PATH_FEATURES (Папка для сохранения дашборда)
   Назначение: Папка для сохранения дашборда.
   Что хранится: Готовый дашборд.
   Пример: \\\\kari.local\\...\\Дашбоард по рекламным кампаниям

3️⃣ FOLDER_PATH_FOR_DB (Папка для базы данных)
   Назначение: Папка для хранения итоговых файлов базы данных.
   Что хранится: Готовые файлы Excel с обработанными данными.
   Пример: \\\\kari.local\\...\\Дашбоард по рекламным кампаниям

4️⃣ FOLDER_PATH_DUDL (Дополнительная папка для сохранения дашборда)
   Назначение: Дополнительная папка для сохранения дашборда.
   Что хранится: Финальные файлы для визуализации.
   Пример: \\\\kari.local\\...\\Дашборд реклама WB_OZ

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔧 ЧТО ДЕЛАЕТ СКРИПТ

1. Загружает данные из "Затраты ВБ_2.xlsx"
2. Получает историю затрат из "История-затрат-Все.xlsx"
3. Создает широкую таблицу по типам активности
4. Получает цены из MS SQL базы данных
5. Загружает справочник кампаний
6. Обрабатывает воронку из "Показатели по неделям ВБ"
7. Объединяет воронку с затратами (OUTER merge)
8. Обновляет итоговые Excel файлы

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚠️ ВАЖНЫЕ ПРЕДУПРЕЖДЕНИЯ

❗ Перед запуском:
   • Убедитесь, что вы запустили утилиту Dashboard_Preprocessing для WB файлов, иначе в дашборде не будет данных на желаемые дни
   • Убедитесь, что все пути указаны правильно
   • Закройте все файлы Excel (особенно исходные файлы)
   • Проверьте доступ к SQL серверу (cl01sql)
   • Убедитесь, что есть минимум 1.5 GB свободного места

❗ Время выполнения:
   • Полный цикл: 30-50 минут в зависимости от объема данных
   • Скрипт запускается в отдельном окне консоли
   • НЕ закрывайте окно консоли до завершения!

❗ Особенности:
   • Скрипт работает с большими объемами данных
   • Использует SQL Server Integration Services
   • Может потребовать значительное количество RAM (около 10GB)
   • При ошибках Excel повторяет попытку сохранения, но иногда может потребоваться ручное вмешательство. Обращайтесь к разработчику, или попытайтесь обновить Excel файл дашборда сами ("Данные"->"Обновить все").

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        
        help_text.insert("1.0", help_content)
        help_text.config(state=tk.DISABLED)
        
        # Кнопка закрытия
        close_button = tk.Button(
            help_window,
            text="Закрыть",
            command=help_window.destroy,
            font=("Arial", 10, "bold"),
            bg="#8e44ad",
            fg="white",
            activebackground="#9b59b6",
            activeforeground="white",
            cursor="hand2",
            pady=8
        )
        close_button.pack(pady=(0, 20), padx=20, fill=tk.X)
            
    def log_message(self, message):
        """Добавляет сообщение в лог-консоль"""
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
        
    def run_assembly(self):
        """Запускает скрипт сборки дашборда в отдельном потоке"""
        # Очистка лога
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Запускаем в отдельном потоке чтобы не блокировать GUI
        thread = threading.Thread(target=self._run_assembly_thread, daemon=True)
        thread.start()
        
    def _run_assembly_thread(self):
        """Внутренний метод для запуска Assembly скрипта в потоке"""
        assembly_file = "DashboardAssemblyWB.py"
        assembly_path = os.path.join(os.path.dirname(__file__), assembly_file)
        
        if not os.path.exists(assembly_path):
            self.log_message(f"❌ ОШИБКА: Файл {assembly_file} не найден!")
            messagebox.showerror("Ошибка", f"Файл {assembly_file} не найден!")
            return
            
        self.log_message(f"▶️ Запуск скрипта: {assembly_file}")
        self.log_message("📦 Сборка дашборда WB...")
        self.log_message("─" * 60)
        
        try:
            f'FOLDER_PATH_FEATURES = r"{self.folder_path_features_var.get()}"'
            
            # Заменяем FOLDER_PATH_FOR_DB
            modified_script = modified_script.replace(
                'FOLDER_PATH_FOR_DB = os.path.normpath(r"\\\\kari.local\\public\\all\\Analytics\\Marketplaceanalytics\\Федоров\\Дашбоард по рекламным кампаниям")',
                f'FOLDER_PATH_FOR_DB = os.path.normpath(r"{self.folder_path_for_db_var.get()}")'
            )
            
            # Заменяем FOLDER_PATH_DUDL
            modified_script = modified_script.replace(
                'FOLDER_PATH_DUDL = os.path.normpath(r"\\\\kari.local\\public\\all\\Агрегаторы\\Дашборд реклама WB_OZ")',
                f'FOLDER_PATH_DUDL = os.path.normpath(r"{self.folder_path_dudl_var.get()}")'
            )
            
            # Сохраняем временный скрипт
            temp_script_path = os.path.join(os.path.dirname(__file__), "_temp_DashboardAssemblyWB.py")
            with open(temp_script_path, 'w', encoding='utf-8') as f:
                f.write(modified_script)
            
            self.log_message("📝 Временный скрипт с измененными путями создан")
            self.log_message("🚀 Запуск скрипта... (окно консоли откроется отдельно)")
            
            # Запускаем скрипт в отдельном окне консоли
            if sys.platform == "win32":
                # Для Windows используем cmd с ключом /K чтобы окно не закрывалось
                process = subprocess.Popen(
                    ['cmd', '/c', 'start', 'cmd', '/K', 'python', temp_script_path],
                    shell=True
                )
            else:
                process = subprocess.Popen(['python', temp_script_path])
            
            self.log_message("✅ Скрипт запущен в отдельном окне консоли")
            self.log_message("📌 Следите за прогрессом в окне консоли")
            self.log_message("─" * 60)
            self.log_message("ℹ️ После завершения работы скрипта можно закрыть это окно")
            
            # Записываем в лог-файл
            logs = open('logs.log', 'a')
            logs.write(f'{datetime.now()} - DashboardAssemblyWB.py started via GUI\n')
            logs.close()
            
            messagebox.showinfo(
                "Скрипт запущен", 
                "Скрипт запущен в отдельном окне консоли.\n\n"
                "Следите за прогрессом в открывшемся окне.\n"
                "Это окно можно закрыть."
            )
            
        except Exception as e:
            error_message = f"❌ ОШИБКА: {str(e)}"
            self.log_message(error_message)
            messagebox.showerror("Ошибка выполнения", str(e))
            
            # Записываем ошибку в лог-файл
            logs = open('logs.log', 'a')
            logs.write(f'{datetime.now()} - ERROR in DashboardAssemblyWB.py: {str(e)}\n')
            logs.close()


def main():
    root = tk.Tk()
    app = DashboardAssemblyWBGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
