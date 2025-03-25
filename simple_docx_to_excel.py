import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
import platform
from docx_to_excel_logic import DocxToExcelProcessor

class SimpleDocxToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Таблицы из DOCX в Excel")
        self.root.geometry("600x600")
        
        # Переменные для хранения путей к файлам
        self.docx_path = None
        self.excel_path = None
        
        # Создаем экземпляр процессора для обработки файлов
        self.processor = DocxToExcelProcessor()
        
        # Создание интерфейса
        self.create_gui()
    
    def create_gui(self):
        """Создание простого интерфейса"""
        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Конвертация таблиц из DOCX в Excel", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Фрейм для выбора файла
        file_frame = ttk.LabelFrame(main_frame, text="Выбор DOCX файла", padding=10)
        file_frame.pack(fill=tk.X, pady=10)
        
        # Поле для отображения пути к файлу
        self.file_path_var = tk.StringVar()
        file_path_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_path_entry.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        # Кнопка выбора файла
        browse_button = ttk.Button(file_frame, text="Обзор...", command=self.select_file)
        browse_button.grid(row=0, column=1, padx=5, pady=5)
        
        file_frame.columnconfigure(0, weight=1)
        
        # Информационное поле
        info_frame = ttk.LabelFrame(main_frame, text="Статус", padding=10)
        info_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Текстовое поле для статуса
        self.status_text = tk.Text(info_frame, wrap=tk.WORD, height=8)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Полоса прокрутки
        scrollbar = ttk.Scrollbar(info_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.update_status("Выберите DOCX файл для начала обработки.")
        
        # Фрейм для кнопок
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # Кнопка обработки
        process_button = ttk.Button(
            button_frame, 
            text="Обработать (конвертировать + удалить столбцы A и C)", 
            command=self.process_file
        )
        process_button.pack(side=tk.LEFT, padx=5)
        
        # Кнопка выхода
        exit_button = ttk.Button(button_frame, text="Выход", command=self.root.destroy)
        exit_button.pack(side=tk.RIGHT, padx=5)
    
    def update_status(self, message):
        """Обновление статуса в текстовом поле"""
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(tk.END, message)
        self.status_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def select_file(self):
        """Выбор DOCX файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите DOCX файл",
            filetypes=[("DOCX файлы", "*.docx"), ("Все файлы", "*.*")]
        )
        if file_path:
            self.docx_path = file_path
            self.file_path_var.set(file_path)
            
            # Создаем путь для сохранения Excel
            base_name = os.path.basename(file_path)
            name, _ = os.path.splitext(base_name)
            self.excel_path = os.path.join(os.path.dirname(file_path), f"{name}.xlsx")
            
            self.update_status(f"Выбран файл: {file_path}\n\nНажмите кнопку 'Обработать' для конвертации таблиц и удаления столбцов A и C.")
    
    def process_file(self):
        """Обработка файла: конвертация DOCX в Excel и удаление столбцов A и C"""
        if not self.docx_path:
            messagebox.showerror("Ошибка", "Сначала выберите DOCX файл")
            return
        
        try:
            self.update_status("Обработка файла...\nШаг 1: Извлечение таблиц из DOCX...\n\nПравила обработки:\n"
                               "- Все таблицы из Word перенесутся в Excel\n"
                               "- Первая строка удаляется, если во второй ячейке НЕТ даты\n"
                               "- Первая строка сохраняется, если во второй ячейке ЕСТЬ дата\n"
                               "- Столбцы A и C будут удалены\n"
                               "- Даты во втором столбце будут приведены к формату ДД.ММ.ГГГГ\n"
                               "- Даты рождения в пятом столбце будут приведены к формату ДД.ММ.ГГГГ\n"
                               "- Даты окончания в восьмом столбце будут приведены к формату ДД.ММ.ГГГГ\n"
                               "- Информация о судах из столбцов D и E (бывшие F и G) будет перемещена в столбец I (бывший K)\n"
                               "- Все даты в столбце I будут отформатированы в виде ДД.ММ.ГГГГ\n"
                               "- Текст с информацией о судах будет отформатирован для улучшения читаемости")
            
            # Шаг 1: Извлечение таблиц из DOCX в Excel
            table_count = self.processor.convert_docx_to_excel(self.docx_path, self.excel_path)
            
            if table_count > 0:
                self.update_status(f"Шаг 1: Таблицы успешно извлечены ({table_count} шт.)\n\n"
                                   f"Шаг 2: Удаление первых строк и столбцов A и C, нормализация дат, "
                                   f"форматирование информации о судах...")
                
                # Шаг 2: Обработка файла Excel
                stats = self.processor.process_excel_file(self.excel_path)
                
                # Финальное сообщение
                self.update_status(
                    f"Обработка успешно завершена!\n\n"
                    f"- Извлечено таблиц: {table_count}\n"
                    f"- Обработано листов: {stats['sheets_processed']}\n"
                    f"- Удалено первых строк: {stats['rows_deleted']}\n"
                    f"- Нормализовано дат: {stats['total_dates_normalized']}\n"
                    f"- Перемещено текстовых блоков: {stats['text_moved']}\n"
                    f"- Перемещено записей о судах: {stats['court_info_moved']}\n"
                    f"- Отформатировано записей о судах: {stats['formatted_cells']}\n"
                    f"- Удалены столбцы: A и C\n\n"
                    f"Результат сохранен в: {self.excel_path}"
                )
                
                # Открываем Excel-файл
                self.open_file(self.excel_path)
                
                messagebox.showinfo("Успех", f"Обработка завершена. Таблицы сохранены в {self.excel_path}")
            else:
                self.update_status("В документе не найдено таблиц.")
                messagebox.showwarning("Предупреждение", "В документе не найдено таблиц")
        
        except Exception as e:
            self.update_status(f"Произошла ошибка при обработке файла:\n{str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
    
    def open_file(self, file_path):
        """Открытие файла в соответствующем приложении"""
        try:
            if platform.system() == 'Windows':
                os.startfile(file_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.call(['open', file_path])
            else:  # Linux и другие
                subprocess.call(['xdg-open', file_path])
            return True
        except Exception as e:
            print(f"Ошибка при открытии файла: {str(e)}")
            return False


if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleDocxToExcelApp(root)
    root.mainloop()
