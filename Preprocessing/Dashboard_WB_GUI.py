import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import importlib.util
from datetime import datetime


class DashboardGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dashboard WB Wrapper - Графический интерфейс")
        self.root.geometry("800x700")
        self.root.resizable(False, False)
        
        # Переменные для хранения путей
        self.downloads_folder_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Taldykin")
        self.folder_wb_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\ВЫГРУЗКА воронка ВБ")
        self.folder_zatraty_wb_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Затраты\Затраты ВБ")
        self.folder_week_var = tk.StringVar(value=r"\\kari.local\public\all\Analytics\Marketplaceanalytics\Федоров\Дашбоард по рекламным кампаниям\!!!_ИСХОДНИКИ ДЛЯ ДАШБОРДА_НЕ УДАЛЯТЬ_!!!\Показатели по неделям ВБ")
        
        # Переменная для выбора wrapper
        self.wrapper_choice = tk.StringVar(value="wrapper")
        
        # Создание интерфейса
        self.create_widgets()
        
    def create_widgets(self):
        # Заголовок
        title_frame = tk.Frame(self.root, bg="#8e44ad", pady=15)
        title_frame.pack(fill=tk.X)
        
        # Контейнер для заголовка и кнопки справки
        title_container = tk.Frame(title_frame, bg="#8e44ad")
        title_container.pack(fill=tk.X)
        
        title_label = tk.Label(
            title_container,
            text="Dashboard WB Wrapper - Панель управления",
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
        help_button.pack(side=tk.RIGHT, padx=15)
        
        # Основная рамка для содержимого
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Раздел выбора wrapper ---
        wrapper_frame = tk.LabelFrame(main_frame, text="Выбор скрипта", font=("Arial", 12, "bold"), padx=10, pady=10)
        wrapper_frame.pack(fill=tk.X, pady=(0, 15))
        
        tk.Radiobutton(
            wrapper_frame,
            text="Dashboard_WB_Wrapper.py (основной)",
            variable=self.wrapper_choice,
            value="wrapper",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=2)
        
        tk.Radiobutton(
            wrapper_frame,
            text="Dashboard_WB_Wrapper_Tuesday.py (вторник)",
            variable=self.wrapper_choice,
            value="wrapper_tuesday",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=2)
        
        tk.Radiobutton(
            wrapper_frame,
            text="Dashboard_WB_Wrapper_Monday.py (понедельник)",
            variable=self.wrapper_choice,
            value="wrapper_monday",
            font=("Arial", 10)
        ).pack(anchor=tk.W, pady=2)
        
        # --- Раздел настройки путей ---
        paths_frame = tk.LabelFrame(main_frame, text="Настройка путей к папкам", font=("Arial", 12, "bold"), padx=10, pady=10)
        paths_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Downloads folder
        self.create_path_row(paths_frame, "Downloads Folder:", self.downloads_folder_var, 0)
        
        # Folder WB
        self.create_path_row(paths_frame, "Folder WB:", self.folder_wb_var, 1)
        
        # Folder Zatraty WB
        self.create_path_row(paths_frame, "Folder Zatraty WB:", self.folder_zatraty_wb_var, 2)
        
        # Folder Week
        self.create_path_row(paths_frame, "Folder Week:", self.folder_week_var, 3)
        
        # --- Кнопка запуска ---
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        run_button = tk.Button(
            button_frame,
            text="▶️ Запустить выбранный скрипт",
            command=self.run_wrapper,
            font=("Arial", 12, "bold"),
            bg="#8e44ad",
            fg="white",
            activebackground="#8e13ae",
            activeforeground="white",
            pady=10,
            cursor="hand2"
        )
        run_button.pack(fill=tk.X)
        
        # --- Лог-консоль ---
        log_frame = tk.LabelFrame(main_frame, text="Лог выполнения", font=("Arial", 12, "bold"), padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Текстовое поле для логов
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED, bg="#ecf0f1", font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Скроллбар
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
    def create_path_row(self, parent, label_text, variable, row):
        """Создает строку с меткой, полем ввода и кнопкой выбора папки"""
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, pady=5)
        
        label = tk.Label(frame, text=label_text, font=("Arial", 10), width=20, anchor=tk.W)
        label.pack(side=tk.LEFT, padx=(0, 5))
        
        entry = tk.Entry(frame, textvariable=variable, font=("Arial", 9))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_button = tk.Button(
            frame,
            text="📁 Обзор",
            command=lambda: self.browse_folder(variable),
            font=("Arial", 9),
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
        help_window.title("Справка - Dashboard WB Wrapper")
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
        
        # Текстовое поле с прокруткой для справки
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

Это приложение предназначено для автоматизации обработки данных Wildberries.
Оно запускает один из трех wrapper-скриптов для обработки файлов воронки и затрат.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ ПОРЯДОК РАБОТЫ

1. Выгрузите актуальные данные воронки и затрат из ЛК Wildberries в папку "Downloads Folder"
2. Выберите нужный скрипт (по умолчанию основной)
3. Проверьте все пути к папкам (при необходимости измените)
4. Нажмите "▶️ Запустить выбранный скрипт"
5. Дождитесь завершения (следите за логом)
6. При успешном выполнении появится сообщение "✅ ВСЕ ЭТАПЫ ВЫПОЛНЕНЫ"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📂 ОПИСАНИЕ ПУТЕЙ К ПАПКАМ

1️⃣ Downloads Folder
   Назначение: Папка, куда сохраняются выгрузки из личного кабинета.
   Что хранится: Свежие файлы воронки, выгруженные с Wildberries.
   Пример: \\\\kari.local\\public\\all\\Analytics\\Marketplaceanalytics\\Taldykin

2️⃣ Folder WB (Папка воронка ВБ)
   Назначение: Папка для хранения обработанных файлов воронки.
   Что хранится: Файлы после унификации столбцов и первичной обработки.
   Пример: \\\\kari.local\\...\\ВЫГРУЗКА воронка ВБ

3️⃣ Folder Zatraty WB (Папка затраты ВБ)
   Назначение: Папка для хранения файлов затрат на рекламу Wildberries.
   Что хранится: История-затрат-Все.xlsx
   Пример: \\\\kari.local\\...\\Затраты\\Затраты ВБ

4️⃣ Folder Week (Папка показатели по неделям)
   Назначение: Папка для хранения еженедельных отчетов.
   Что хранится: Показатели по неделям ВБ.xlsx (данные за каждую неделю)
   Пример: \\\\kari.local\\...\\Показатели по неделям ВБ

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 ВЫБОР СКРИПТА

▪ Dashboard_WB_Wrapper.py (основной)
  Используется для стандартной еженедельной обработки.
  Выполняет: извлечение файла воронки, унификацию столбцов,
  копирование затрат, обновление недельного файла.

▪ Dashboard_WB_Wrapper_Tuesday.py (вторник)
  Специальная версия для вторника с обновлением Power Query.
  Дополнительно: дублирует файл недели, обновляет даты в запросах.

▪ Dashboard_WB_Wrapper_Monday.py (понедельник)
  Версия для понедельника, обрабатывает последние 3 файла воронки.
  Отличие: работает с тремя файлами одновременно для недельного анализа.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

⚠️ ВАЖНЫЕ ПРЕДУПРЕЖДЕНИЯ

❗ Перед запуском убедитесь, что:
   • Все указанные папки существуют и доступны
   • У вас есть права на чтение и запись в этих папках
   • Файлы Excel не открыты в других программах
   • Есть свободное место на диске (минимум 500 МБ)

❗ Время выполнения:
   • Основной скрипт: 5-7 минут
   • Скрипт вторника: 7-10 минут (из-за Power Query)
   • Скрипт понедельника: 15-20 минут (обработка 3 файлов)

❗ В случае ошибки:
   • Проверьте лог-консоль для деталей
   • Убедитесь, что все пути указаны правильно
   • Проверьте, что файлы не повреждены
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
        
    def run_wrapper(self):
        """Запускает выбранный wrapper скрипт"""
        # Очистка лога
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Получаем выбранный wrapper
        wrapper_choice = self.wrapper_choice.get()
        
        wrapper_files = {
            "wrapper": "Dashboard_WB_Wrapper.py",
            "wrapper_tuesday": "Dashboard_WB_Wrapper_Tuesday.py",
            "wrapper_monday": "Dashboard_WB_Wrapper_Monday.py"
        }
        
        wrapper_file = wrapper_files[wrapper_choice]
        wrapper_path = os.path.join(os.path.dirname(__file__), wrapper_file)
        
        if not os.path.exists(wrapper_path):
            messagebox.showerror("Ошибка", f"Файл {wrapper_file} не найден!")
            self.log_message(f"❌ ОШИБКА: Файл {wrapper_file} не найден!")
            return
            
        self.log_message(f"▶️ Запуск скрипта: {wrapper_file}")
        self.log_message(f"📂 Downloads Folder: {self.downloads_folder_var.get()}")
        self.log_message(f"📂 Folder WB: {self.folder_wb_var.get()}")
        self.log_message(f"📂 Folder Zatraty WB: {self.folder_zatraty_wb_var.get()}")
        self.log_message(f"📂 Folder Week: {self.folder_week_var.get()}")
        self.log_message("─" * 60)
        
        try:
            # Динамическая загрузка модуля
            spec = importlib.util.spec_from_file_location("wrapper_module", wrapper_path)
            wrapper_module = importlib.util.module_from_spec(spec)
            
            # Устанавливаем переменные окружения для модуля
            wrapper_module.downloads_folder = self.downloads_folder_var.get()
            wrapper_module.folder_wb = self.folder_wb_var.get()
            wrapper_module.folder_zatraty_wb = self.folder_zatraty_wb_var.get()
            wrapper_module.folder_week = self.folder_week_var.get()
            wrapper_module.folder_campaign_info = os.path.dirname(self.folder_wb_var.get())
            
            # Загружаем модуль
            spec.loader.exec_module(wrapper_module)
            
            # Запуск функций в зависимости от выбранного wrapper
            if wrapper_choice == "wrapper":
                self.log_message("🔄 Этап 1: Извлечение последнего файла воронки...")
                voronka_wb = wrapper_module.extract_latest_vb_file()
                self.log_message("✅ Этап 1 выполнен")
                
                self.log_message("🔄 Унификация столбцов...")
                wrapper_module.unify_columns(voronka_wb)
                self.log_message("✅ Унификация столбцов выполнена")
                
                self.log_message("🔄 Этап 2: Копирование файла История-затрат...")
                wrapper_module.copy_latest_cost_history()
                self.log_message("✅ Этап 2 выполнен")
                
                self.log_message("🔄 Этап 3: Обновление файла недели...")
                wrapper_module.update_latest_week_file()
                self.log_message("✅ Этап 3 выполнен")
                
                self.log_message("🔄 Этап 4: Обновление Затраты ВБ_2...")
                wrapper_module.update_zatraty_wb2()
                self.log_message("✅ Этап 4 выполнен")
                
            elif wrapper_choice == "wrapper_tuesday":
                self.log_message("🔄 Этап 1: Извлечение последнего файла воронки...")
                voronka_wb = wrapper_module.extract_latest_vb_file()
                self.log_message("✅ Этап 1 выполнен")
                
                self.log_message("🔄 Унификация столбцов...")
                wrapper_module.unify_columns(voronka_wb)
                self.log_message("✅ Унификация столбцов выполнена")
                
                self.log_message("🔄 Этап 2: Копирование файла История-затрат...")
                wrapper_module.copy_latest_cost_history()
                self.log_message("✅ Этап 2 выполнен")
                
                self.log_message("🔄 Этап 3: Дублирование файла недели...")
                new_file_path = wrapper_module.duplicate_latest_week_file()
                self.log_message("✅ Этап 3 выполнен")
                
                self.log_message("🔄 Этап 3.1: Извлечение дат из запроса...")
                old_dates = wrapper_module.extract_dates_from_query(new_file_path, wrapper_module.query_name)
                self.log_message("✅ Этап 3.1 выполнен")
                
                self.log_message("🔄 Этап 3.2-3.3: Обновление Power Query...")
                wrapper_module.update_power_query_steps(new_file_path, wrapper_module.query_name, old_dates)
                self.log_message("✅ Этап 3.2-3.3 выполнен")
                
                self.log_message("🔄 Этап 4: Обновление файла недели...")
                wrapper_module.update_latest_week_file()
                self.log_message("✅ Этап 4 выполнен")
                
                self.log_message("🔄 Этап 5: Обновление Затраты ВБ_2...")
                wrapper_module.update_zatraty_wb2()
                self.log_message("✅ Этап 5 выполнен")
                
            elif wrapper_choice == "wrapper_monday":
                self.log_message("🔄 Этап 1: Извлечение последних 3 файлов воронки...")
                voronka_wb_1, voronka_wb_2, voronka_wb_3 = wrapper_module.extract_latest_3_wb_files()
                self.log_message("✅ Этап 1 выполнен")
                
                self.log_message("🔄 Унификация столбцов для файлов...")
                wrapper_module.unify_columns(voronka_wb_1)
                wrapper_module.unify_columns(voronka_wb_2)
                wrapper_module.unify_columns(voronka_wb_3)
                self.log_message("✅ Унификация столбцов выполнена")
                
                self.log_message("🔄 Этап 2: Копирование файла История-затрат...")
                wrapper_module.copy_latest_cost_history()
                self.log_message("✅ Этап 2 выполнен")
                
                self.log_message("🔄 Этап 3: Обновление файла недели...")
                wrapper_module.update_latest_week_file()
                self.log_message("✅ Этап 3 выполнен")
                
                self.log_message("🔄 Этап 4: Обновление Затраты ВБ_2...")
                wrapper_module.update_zatraty_wb2()
                self.log_message("✅ Этап 4 выполнен")
            
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
    app = DashboardGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
