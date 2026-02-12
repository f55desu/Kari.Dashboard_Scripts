import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import importlib.util
from datetime import datetime


class DashboardOZONGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dashboard OZON Wrapper - Графический интерфейс")
        self.root.geometry("800x700")
        self.root.resizable(False, False)
        
        # Переменные для хранения путей
        self.downloads_folder_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin")
        self.folder_воронка_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\ВЫГРУЗКА воронка Озон")
        self.folder_затраты_Озон_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\Озон. Затраты из Аналитики")
        self.folder_затраты_Озон_NewFormat_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\Озон. Затраты из Аналитики New Format")
        
        # Переменная для выбора wrapper
        self.wrapper_choice = tk.StringVar(value="wrapper")
        
        # Создание интерфейса
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок
        title_frame = tk.Frame(self.root, bg="#4248f5", pady=15)
        title_frame.pack(fill=tk.X)
        
        # Контейнер для заголовка и кнопки справки
        title_container = tk.Frame(title_frame, bg="#4248f5")
        title_container.pack(fill=tk.X)
        
        title_label = tk.Label(
            title_container,
            text="Dashboard OZON Wrapper - Панель управления",
            font=("Arial", 16, "bold"),
            bg="#4248f5",
            fg="white"
        )
        title_label.pack(side=tk.LEFT, expand=True)
        
        # Кнопка справки
        help_button = tk.Button(
            title_container,
            text="❓ Справка",
            command=self.show_help,
            font=("Arial", 10, "bold"),
            bg="#5a67f5",
            fg="white",
            activebackground="#4248f5",
            activeforeground="white",
            cursor="hand2",
            padx=15,
            pady=5
        )
        help_button.pack(side=tk.RIGHT, padx=15)
        
        # Основная рамка для содержимого
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Раздел выбора wrapper ---
        wrapper_frame = tk.LabelFrame(main_frame, text="Выбор скрипта", font=("Arial", 12, "bold"), padx=10, pady=10)
        wrapper_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Radiobutton(
            wrapper_frame,
            text="Dashboard_OZON_Wrapper.py (основной)",
            variable=self.wrapper_choice,
            value="wrapper",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=2)
        
        tk.Radiobutton(
            wrapper_frame,
            text="Dashboard_OZON_Wrapper_Monday.py (понедельник)",
            variable=self.wrapper_choice,
            value="wrapper_monday",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=2)
        
        # --- Раздел настройки путей ---
        paths_frame = tk.LabelFrame(main_frame, text="Настройка путей к папкам", font=("Arial", 12, "bold"), padx=10, pady=10)
        paths_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Downloads folder
        self.create_path_row(paths_frame, "Downloads Folder:", self.downloads_folder_var, 0)
        
        # Folder воронка
        self.create_path_row(paths_frame, "Folder Воронка:", self.folder_воронка_var, 1)
        
        # Folder затраты Озон
        self.create_path_row(paths_frame, "Folder Затраты Озон:", self.folder_затраты_Озон_var, 2)
        
        # Folder затраты Озон NewFormat
        self.create_path_row(paths_frame, "Folder Затраты Озон NewFormat:", self.folder_затраты_Озон_NewFormat_var, 3)
        
        # --- Кнопка запуска ---
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        run_button = tk.Button(
            button_frame,
            text="▶️ Запустить выбранный скрипт",
            command=self.run_wrapper,
            font=("Arial", 12, "bold"),
            bg="#4248f5",
            fg="white",
            activebackground="#4248f5",
            activeforeground="white",
            pady=10,
            cursor="hand2"
        )
        run_button.pack(fill=tk.X)
        
        # --- Лог-консоль ---
        log_frame = tk.LabelFrame(main_frame, text="Лог выполнения", font=("Arial", 12, "bold"), padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Текстовое поле для логов
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg="#fef5e7", font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Скроллбар
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
    def create_path_row(self, parent, label_text, variable, row):
        """Создает строку с меткой, полем ввода и кнопкой выбора папки"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        label = tk.Label(frame, text=label_text, font=("Arial", 10), width=25, anchor=tk.W)
        label.pack(side=tk.LEFT, padx=(0, 5))
        
        entry = tk.Entry(frame, textvariable=variable, font=("Arial", 9))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_button = tk.Button(
            frame,
            text="📁 Обзор",
            command=lambda: self.browse_folder(variable),
            font=("Arial", 9),
            bg="#4248f5",
            fg="white",
            activebackground="#4248f5",
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
        help_window.title("Справка - Dashboard OZON Wrapper")
        help_window.geometry("700x600")
        help_window.resizable(True, True)
        
        # Заголовок окна справки
        header_frame = tk.Frame(help_window, bg="#4248f5", pady=10)
        header_frame.pack(fill=tk.X)
        
        header_label = tk.Label(
            header_frame,
            text="📖 Справка по использованию",
            font=("Arial", 14, "bold"),
            bg="#4248f5",
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

Это приложение предназначено для автоматизации обработки данных Ozon.
Оно запускает один из двух wrapper-скриптов для обработки файлов воронки и затрат.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ ПОРЯДОК РАБОТЫ

1. Выгрузите актуальные данные воронки и затрат из ЛК Озон в папку "Downloads Folder"
2. Выберите нужный скрипт (по умолчанию основной)
3. Проверьте все пути к папкам
4. Нажмите "▶️ Запустить выбранный скрипт"
5. Дождитесь завершения (следите за логом)
7. При успешном выполнении появится сообщение "✅ ВСЕ ЭТАПЫ ВЫПОЛНЕНЫ"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📂 ОПИСАНИЕ ПУТЕЙ К ПАПКАМ

1️⃣ Downloads Folder
   Назначение: Папка, куда сохраняются выгрузки из личного кабинета Ozon.
   Что хранится: Свежие файлы Аналитики продвижения и отчетов воронки.
   Пример: \\\\kari.local\\public\\all\\Analytics\\Marketplaceanalytics\\Taldykin

2️⃣ Folder Воронка (Папка воронка Озон)
   Назначение: Папка для хранения обработанных файлов воронки Ozon.
   Что хранится: Файлы analytics_report_*.xlsx и SKU статистики.
   Пример: \\\\kari.local\\...\\ВЫГРУЗКА воронка Озон

3️⃣ Folder Затраты Озон
   Назначение: Папка для хранения файлов затрат на рекламу Ozon (старый формат).
   Что хранится: Файлы Аналитика продвижения (старая версия).
   Пример: \\\\kari.local\\...\\Озон. Затраты из Аналитики

4️⃣ Folder Затраты Озон NewFormat
   Назначение: Папка для хранения файлов затрат на рекламу Ozon (новый формат).
   Что хранится: Аналитика продвижения (новый формат с расширенными данными).
   Пример: \\\\kari.local\\...\\Озон. Затраты из Аналитики New Format

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 ВЫБОР СКРИПТА

▪ Dashboard_OZON_Wrapper.py (основной)
  Используется для стандартной еженедельной обработки.
  Выполняет: поиск и конвертацию SKU статистики,
  копирование отчетов воронки.

▪ Dashboard_OZON_Wrapper_Monday.py (понедельник)
  Версия для понедельника, обрабатывает последние 3 файла.
  Отличие: работает с 3 файлами SKU и 3 отчетами за пятницу/субботу/воскресенье.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚠️ ВАЖНЫЕ ПРЕДУПРЕЖДЕНИЯ

❗ Перед запуском убедитесь, что:
   • Вы выгрузили актуальные данные воронки и затрат из ЛК Озон
   • АКТУАЛЬНЫЕ Выгрузки из Ozon сохранены в Downloads Folder
   • Все указанные папки существуют и доступны
   • У вас есть права на чтение и запись в этих папках
   • Файлы Excel не открыты в других программах

❗ Время выполнения:
   • Основной скрипт: 1-2 минуты
   • Скрипт понедельника: 3-5 минут (обработка 3 файлов)

❗ В случае ошибки:
   • Проверьте лог-консоль для деталей
   • Убедитесь, что все пути указаны правильно
   • Проверьте формат имен файлов (должны соответствовать шаблонам)
   • При повторении ошибки обратитесь к разработчику

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
            bg="#4248f5",
            fg="white",
            activebackground="#5a67f5",
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
        
    def run_wrapper(self):
        """Запускает выбранный wrapper скрипт"""
        # Очистка лога
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Получаем выбранный wrapper
        wrapper_choice = self.wrapper_choice.get()
        
        wrapper_files = {
            "wrapper": "Dashboard_OZON_Wrapper.py",
            "wrapper_monday": "Dashboard_OZON_Wrapper_Monday.py"
        }
        
        wrapper_file = wrapper_files[wrapper_choice]
        wrapper_path = os.path.join(os.path.dirname(__file__), wrapper_file)
        
        if not os.path.exists(wrapper_path):
            messagebox.showerror("Ошибка", f"Файл {wrapper_file} не найден!")
            self.log_message(f"❌ ОШИБКА: Файл {wrapper_file} не найден!")
            return
            
        self.log_message(f"▶️ Запуск скрипта: {wrapper_file}")
        self.log_message(f"📂 Downloads Folder: {self.downloads_folder_var.get()}")
        self.log_message(f"📂 Folder Воронка: {self.folder_воронка_var.get()}")
        self.log_message(f"📂 Folder Затраты Озон: {self.folder_затраты_Озон_var.get()}")
        self.log_message(f"📂 Folder Затраты Озон NewFormat: {self.folder_затраты_Озон_NewFormat_var.get()}")
        self.log_message("─" * 60)
        
        try:
            # Динамическая загрузка модуля
            spec = importlib.util.spec_from_file_location("wrapper_module", wrapper_path)
            wrapper_module = importlib.util.module_from_spec(spec)
            
            # Устанавливаем переменные окружения для модуля
            wrapper_module.folder_воронка = self.folder_воронка_var.get()
            wrapper_module.folder_затраты_Озон = self.folder_затраты_Озон_var.get()
            wrapper_module.folder_затраты_Озон_NewFormat = self.folder_затраты_Озон_NewFormat_var.get()
            wrapper_module.folder_показатели_по_дням = os.path.join(
                os.path.dirname(self.folder_воронка_var.get()), 
                "Показатели по дням"
            )
            
            # Загружаем модуль
            spec.loader.exec_module(wrapper_module)
            
            # Запуск функций в зависимости от выбранного wrapper
            if wrapper_choice == "wrapper":
                copyFrom = self.downloads_folder_var.get()
                
                self.log_message("🔄 Этап 1: Поиск последнего файла SKU статистики...")
                pattern = r"Аналитика продвижения_(\d{2}\.\d{2}\.\d{4})\.xlsx"
                file_sku = wrapper_module.get_latest_file_by_pattern(copyFrom, pattern)
                self.log_message(f"✅ Найден файл: {file_sku}")
                
                self.log_message("🔄 Этап 2: Конвертация SKU статистики...")
                file_sku = wrapper_module.copy_and_convert_sku_statistics(
                    filename=file_sku, 
                    downloads_path=copyFrom
                )
                self.log_message("✅ Этап 2 выполнен")
                
                self.log_message("🔄 Этап 3: Копирование файла отчета воронки...")
                pattern = r"analytics_report_\d{4}-\d{2}-\d{2}_\d{2}_\d{2}\.xlsx"
                file_report = wrapper_module.get_latest_file_by_pattern(copyFrom, pattern)
                wrapper_module.copy_to_directory(copyFrom, wrapper_module.folder_воронка, file_report)
                self.log_message("✅ Этап 3 выполнен")
                
            elif wrapper_choice == "wrapper_monday":
                copyFrom = self.downloads_folder_var.get()
                
                self.log_message("🔄 Этап 1: Обработка последних 3 файлов SKU...")
                file_sku_1, file_sku_2, file_sku_3 = wrapper_module.process_last_3_sku_files(copyFrom)
                self.log_message("✅ Этап 1 выполнен (3 файла обработаны)")
                
                self.log_message("🔄 Этап 2: Копирование последних 3 файлов отчетов...")
                pattern = r"analytics_report_\d{4}-\d{2}-\d{2}_\d{2}_\d{2}\.xlsx"
                file_report_1, file_report_2, file_report_3 = wrapper_module.get_latest_file_by_pattern(copyFrom, pattern)
                
                wrapper_module.copy_to_directory(copyFrom, wrapper_module.folder_воронка, file_report_3)
                self.log_message(f"✅ Скопирован файл пятницы: {file_report_3}")
                
                wrapper_module.copy_to_directory(copyFrom, wrapper_module.folder_воронка, file_report_2)
                self.log_message(f"✅ Скопирован файл субботы: {file_report_2}")
                
                wrapper_module.copy_to_directory(copyFrom, wrapper_module.folder_воронка, file_report_1)
                self.log_message(f"✅ Скопирован файл воскресенья: {file_report_1}")
                
                self.log_message("✅ Этап 2 выполнен")
            
            self.log_message("─" * 60)
            self.log_message("🎉 ВСЕ ЭТАПЫ УСПЕШНО ВЫПОЛНЕНЫ!")
            
            # Записываем в лог-файл
            logs = open('logs.log', 'a')
            logs.write(f'{datetime.now()} - {wrapper_file} completed via GUI\n')
            logs.close()
            
            messagebox.showinfo("Успех", "Скрипт успешно выполнен!")
            
        except Exception as e:
            error_message = f"❌ ОШИБКА: {str(e)}"
            self.log_message(error_message)
            messagebox.showerror("Ошибка выполнения", str(e))
            
            # Записываем ошибку в лог-файл
            logs = open('logs.log', 'a')
            logs.write(f'{datetime.now()} - ERROR in {wrapper_file}: {str(e)}\n')
            logs.close()


def main():
    root = tk.Tk()
    app = DashboardOZONGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
