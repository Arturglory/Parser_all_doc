import tkinter as tk
from tkinter import filedialog, scrolledtext
from Parser_all import DocumentParser


class ParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Parser_all")
        self.root.geometry("800x600")

        # Поле для ввода файлов
        self.file_label = tk.Label(root, text="Выбранные файлы:")
        self.file_label.pack(pady=5)

        self.files_text = tk.Text(root, height=5, width=80)
        self.files_text.pack(pady=5)

        self.select_files_button = tk.Button(root, text="Выбрать файлы", command=self.select_files)
        self.select_files_button.pack(pady=5)

        # Поле для ввода поискового запроса
        self.search_frame = tk.Frame(root)
        self.search_frame.pack(pady=5)

        self.search_label = tk.Label(self.search_frame, text="Введите фразу или число для поиска:")
        self.search_label.pack(side=tk.LEFT, padx=5)

        self.search_entry = tk.Entry(self.search_frame, width=60)
        self.search_entry.pack(side=tk.LEFT, padx=5)

        self.paste_button = tk.Button(self.search_frame, text="Вставить", command=self.paste_search_term)
        self.paste_button.pack(side=tk.LEFT, padx=5)

        # Кнопка запуска поиска
        self.search_button = tk.Button(root, text="Поиск", command=self.run_search)
        self.search_button.pack(pady=10)

        # Текстовое поле для вывода результатов
        self.result_label = tk.Label(root, text="Результаты:")
        self.result_label.pack(pady=5)

        self.result_text = scrolledtext.ScrolledText(root, height=20, width=80)
        self.result_text.pack(pady=5)

    def select_files(self):
        # Диалог выбора нескольких файлов
        file_paths = filedialog.askopenfilenames(
            title="Выберите файлы",
            filetypes=[
                ("All Supported Files",
                 "*.pdf *.xls *.xlsx *.docx *.txt *.csv *.odt *.json *.html *.xml *.rtf *.md *.zip *.rar"),
                ("PDF Files", "*.pdf"),
                ("Excel Files", "*.xls *.xlsx"),
                ("Word Files", "*.docx"),
                ("Text Files", "*.txt"),
                ("CSV Files", "*.csv"),
                ("ODT Files", "*.odt"),
                ("JSON Files", "*.json"),
                ("HTML Files", "*.html"),
                ("XML Files", "*.xml"),
                ("RTF Files", "*.rtf"),
                ("Markdown Files", "*.md"),
                ("ZIP Archives", "*.zip"),
                ("RAR Archives", "*.rar")
            ]
        )
        if file_paths:
            self.files_text.delete(1.0, tk.END)
            self.files_text.insert(tk.END, "\n".join(file_paths))
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Выбрано {len(file_paths)} файл(ов):\n" + "\n".join(file_paths) + "\n\n")

    def paste_search_term(self):
        # Вставка текста из буфера обмена в поле ввода
        try:
            clipboard_text = self.root.clipboard_get()
            self.search_entry.delete(0, tk.END)
            self.search_entry.insert(0, clipboard_text)
        except tk.TclError:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Ошибка: В буфере обмена нет текста.\n")

    def run_search(self):
        file_paths = self.files_text.get(1.0, tk.END).strip().split('\n')
        search_term = self.search_entry.get().strip()

        if not file_paths or not search_term:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Ошибка: Укажите файлы и поисковый запрос.\n")
            return

        parser = DocumentParser(file_paths)
        results = parser.parse_and_search([search_term])

        self.result_text.delete(1.0, tk.END)
        for file_name, result in results.items():
            self.result_text.insert(tk.END, f"\nФайл: {file_name}\n")
            if 'error' in result:
                self.result_text.insert(tk.END, f"Ошибка: {result['error']}\n")
            else:
                self.result_text.insert(tk.END, f"Тип файла: {result['type']}\n")
                if result['type'] == 'pdf':
                    self.result_text.insert(tk.END, f"Количество страниц: {result['pages']}\n")
                elif result['type'] in ['excel', 'csv']:
                    if 'sheets' in result:
                        self.result_text.insert(tk.END, f"Листы: {result['sheets']}\n")
                    self.result_text.insert(tk.END, f"Общее количество строк: {result['total_rows']}\n")
                elif result['type'] in ['zip', 'rar']:
                    self.result_text.insert(tk.END, "Результаты для вложенных файлов:\n")

                self.result_text.insert(tk.END, "Результаты поиска:\n")
                if not result['search_results']:
                    self.result_text.insert(tk.END, "Совпадений не найдено\n")
                else:
                    if result['type'] in ['zip', 'rar']:
                        for nested_file, nested_result in result['search_results'].items():
                            self.result_text.insert(tk.END, f"\n  Вложенный файл: {nested_file}\n")
                            if 'error' in nested_result:
                                self.result_text.insert(tk.END, f"    Ошибка: {nested_result['error']}\n")
                            else:
                                for term, findings in nested_result['search_results'].items():
                                    self.result_text.insert(tk.END, f"    Поисковый запрос: '{term}'\n")
                                    for finding in findings:
                                        if 'sheet' in finding:
                                            self.result_text.insert(tk.END,
                                                                    f"      Лист '{finding['sheet']}', Строка {finding['row']}, Колонка '{finding['column']}': значение '{finding['cell_value']}'\n")
                                        elif 'page' in finding:
                                            self.result_text.insert(tk.END,
                                                                    f"      Страница {finding['page']}: найдено {finding['count']} раз, позиции: {finding['positions']}\n")
                                        elif 'path' in finding:
                                            self.result_text.insert(tk.END,
                                                                    f"      Путь: {finding['path']}, Значение: '{finding['value']}'\n")
                                        else:
                                            self.result_text.insert(tk.END,
                                                                    f"      Найдено {finding['count']} раз, позиции: {finding['positions']}\n")
                    else:
                        for term, findings in result['search_results'].items():
                            self.result_text.insert(tk.END, f"\nПоисковый запрос: '{term}'\n")
                            for finding in findings:
                                if 'sheet' in finding:
                                    self.result_text.insert(tk.END,
                                                            f"  Лист '{finding['sheet']}', Строка {finding['row']}, Колонка '{finding['column']}': значение '{finding['cell_value']}'\n")
                                elif 'page' in finding:
                                    self.result_text.insert(tk.END,
                                                            f"  Страница {finding['page']}: найдено {finding['count']} раз, позиции: {finding['positions']}\n")
                                elif 'path' in finding:
                                    self.result_text.insert(tk.END,
                                                            f"  Путь: {finding['path']}, Значение: '{finding['value']}'\n")
                                else:
                                    self.result_text.insert(tk.END,
                                                            f"  Найдено {finding['count']} раз, позиции: {finding['positions']}\n")


if __name__ == "__main__":
    root = tk.Tk()
    app = ParserGUI(root)
    root.mainloop()