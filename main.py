import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
from PIL import Image, ImageDraw, ImageFont

import csv
import json
import os
import sys

# Для копирования изображения в буфер обмена (Windows)
if sys.platform.startswith("win"):
    import win32clipboard
    import io

SUB_PATH = "created_files"


class EvaluationApp:
    def __init__(self, master):
        self.master = master
        master.title("Оценочный лист")

        # Настройки окна
        master.geometry("800x600")
        master.resizable(False, False)

        # Загрузка данных о студентах
        self.load_student_data()
        # Загрузка вариантов студентов
        self.load_student_variants()

        # Создаем вкладки
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill="both", expand=True)

        # Вкладка с информацией о студенте
        self.info_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.info_frame, text="Информация о студенте")

        # Вкладка с критериями оценки
        self.criteria_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.criteria_frame, text="Критерии оценки")

        # Вкладка с дополнительными штрафами
        self.penalty_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.penalty_frame, text="Дополнительные штрафы")

        # Вкладка с генерацией отчета
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_frame, text="Генерация отчета")

        self.section_max_scores = {}

        self.create_info_tab()
        self.create_criteria_tab()    # Перенесено перед load_info_parameters()
        self.load_info_parameters()   # Загрузка сохраненных параметров
        self.create_penalty_tab()
        self.create_report_tab()

        # Строка состояния
        self.status_var = tk.StringVar()
        self.status_bar = tk.Label(
            master, textvariable=self.status_var, bd=1, relief="sunken", anchor="w"
        )
        self.status_bar.pack(side="bottom", fill="x")

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.save_info_parameters()
        self.master.destroy()

    def save_info_parameters(self):
        data = {
            "hw_name": self.hw_name_var.get(),
            "variant_count": self.variant_count_entry.get(),
            "group": self.group_var.get(),
            "student": self.student_var.get(),
            "variant": self.variant_entry.get(),
            "on_time": self.on_time.get(),
        }
        with open("info_parameters.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

    def load_info_parameters(self):
        if os.path.exists("info_parameters.json"):
            with open("info_parameters.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                self.hw_name_var.set(data.get("hw_name", ""))
                self.on_homework_selected(None)  # Обновляем критерии
                self.variant_count_entry.delete(0, tk.END)
                self.variant_count_entry.insert(0, data.get("variant_count", "29"))
                self.group_var.set(data.get("group", ""))
                # Обновление списка студентов на основе загруженной группы
                self.update_student_list(None)
                self.student_var.set(data.get("student", ""))
                # Обновление информации о студенте на основе загруженного имени
                self.update_student_info(None)
                self.variant_entry.delete(0, tk.END)
                self.variant_entry.insert(0, data.get("variant", ""))
                self.on_time.set(data.get("on_time", True))

    def load_student_data(self):
        self.student_data = []
        self.groups = set()
        if os.path.exists("student_list.csv"):
            with open("student_list.csv", encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile, delimiter=";")
                for row in reader:
                    user_info = row["Данные о пользователе"]
                    group = user_info.split(";")[-1].strip()
                    self.groups.add(group)
                    self.student_data.append(
                        {"Фамилия": row["Фамилия"], "Имя": row["Имя"], "Группа": group}
                    )
        else:
            tk.messagebox.showerror("Ошибка", "Файл student_list.csv не найден.")

        self.groups = sorted(list(self.groups))

    def load_student_variants(self):
        self.student_variants = {}
        if os.path.exists("student_variants.csv"):
            with open("student_variants.csv", encoding="utf-8") as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    key = row[0]  # 'Группа;Фамилия Имя'
                    variant = row[1]
                    self.student_variants[key] = variant

    def save_student_variants(self):
        with open("student_variants.csv", "w", encoding="utf-8", newline="") as csvfile:
            writer = csv.writer(csvfile)
            for key, variant in self.student_variants.items():
                writer.writerow([key, variant])

    def load_homework_names(self):
        try:
            with open("criteria.json", "r", encoding="utf-8") as f:
                self.criteria_data = json.load(f)
            self.homework_names = list(self.criteria_data.get("sections", {}).keys())
        except Exception as e:
            tk.messagebox.showerror("Ошибка", f"Не удалось загрузить критерии: {e}")
            self.homework_names = []

    def create_info_tab(self):
        # Загрузка названий домашних заданий
        self.load_homework_names()

        # Название домашней работы (выпадающий список)
        tk.Label(self.info_frame, text="Название домашней работы:").grid(
            row=0, column=0, sticky="w", pady=5, padx=5
        )
        self.hw_name_var = tk.StringVar()
        self.hw_name_combobox = ttk.Combobox(
            self.info_frame,
            textvariable=self.hw_name_var,
            values=self.homework_names,
            state="readonly",
        )
        self.hw_name_combobox.grid(row=0, column=1, sticky="w", pady=5)
        self.hw_name_combobox.bind("<<ComboboxSelected>>", self.on_homework_selected)

        # Количество вариантов
        tk.Label(self.info_frame, text="Количество вариантов:").grid(
            row=1, column=0, sticky="w", pady=5, padx=5
        )
        self.variant_count_entry = tk.Entry(self.info_frame, width=10)
        self.variant_count_entry.grid(row=1, column=1, sticky="w", pady=5)
        self.variant_count_entry.insert(0, "29")  # По умолчанию 29 вариантов

        # Группа
        tk.Label(self.info_frame, text="Группа:").grid(
            row=2, column=0, sticky="w", pady=5, padx=5
        )
        self.group_var = tk.StringVar()
        self.group_combobox = ttk.Combobox(
            self.info_frame,
            textvariable=self.group_var,
            values=self.groups,
            state="readonly",
        )
        self.group_combobox.grid(row=2, column=1, sticky="w", pady=5)
        self.group_combobox.bind("<<ComboboxSelected>>", self.update_student_list)

        # Студент
        tk.Label(self.info_frame, text="Студент:").grid(
            row=3, column=0, sticky="w", pady=5, padx=5
        )
        self.student_var = tk.StringVar()
        self.student_combobox = ttk.Combobox(
            self.info_frame, textvariable=self.student_var, values=[], state="readonly"
        )
        self.student_combobox.grid(row=3, column=1, sticky="w", pady=5)
        self.student_combobox.bind("<<ComboboxSelected>>", self.update_student_info)

        # Кнопки навигации
        nav_frame = ttk.Frame(self.info_frame)
        nav_frame.grid(row=4, column=1, sticky="w", pady=5)
        self.prev_button = ttk.Button(nav_frame, text="<<", command=self.prev_student)
        self.prev_button.pack(side="left", padx=5)
        self.next_button = ttk.Button(nav_frame, text=">>", command=self.next_student)
        self.next_button.pack(side="left", padx=5)

        # Вариант
        tk.Label(self.info_frame, text="Вариант:").grid(
            row=5, column=0, sticky="w", pady=5, padx=5
        )
        self.variant_entry = tk.Entry(self.info_frame, width=10)
        self.variant_entry.grid(row=5, column=1, sticky="w", pady=5)
        self.variant_entry.bind("<FocusOut>", self.save_variant)
        self.variant_entry.bind("<Return>", self.save_variant)

        # Сдано вовремя
        self.on_time = tk.BooleanVar(value=True)
        tk.Checkbutton(
            self.info_frame, text="Сдано вовремя", variable=self.on_time
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=5)

    def copy_to_clipboard(self):
        # Генерация отчета, если он еще не создан
        self.generate_report(save_to_file=False)

        # Копирование изображения в буфер обмена
        if sys.platform.startswith("win"):
            output = io.BytesIO()
            self.generated_image.convert("RGB").save(output, "BMP")
            data = output.getvalue()[14:]
            output.close()
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            self.status_var.set("Оценочный лист скопирован в буфер обмена.")
        else:
            self.status_var.set(
                "Копирование изображения в буфер обмена поддерживается только на Windows."
            )

    def on_homework_selected(self, event):
        selected_homework = self.hw_name_var.get()
        self.current_homework = selected_homework
        self.load_criteria_for_homework(selected_homework)
        # Нет необходимости обновлять штрафы, так как они общие

    def update_student_list(self, event):
        selected_group = self.group_var.get()
        students_in_group = sorted(
            [s for s in self.student_data if s["Группа"] == selected_group],
            key=lambda x: x["Фамилия"],
        )
        self.students_in_group = students_in_group  # Сохраняем для навигации
        self.student_names = [f"{s['Фамилия']} {s['Имя']}" for s in students_in_group]
        self.student_combobox["values"] = self.student_names
        if self.student_names:
            self.current_student_index = 0
            self.student_var.set(self.student_names[0])
            self.update_student_info(None)

    def update_student_info(self, event):
        selected_student_name = self.student_var.get()
        if selected_student_name in self.student_names:
            self.current_student_index = self.student_names.index(selected_student_name)
            self.calculate_variant()

    def calculate_variant(self):
        try:
            variant_count = int(self.variant_count_entry.get())
        except ValueError:
            variant_count = 29  # По умолчанию 29 вариантов

        student_number = self.current_student_index + 1  # Нумерация с 1
        key = f"{self.group_var.get()};{self.student_var.get()}"
        if key in self.student_variants:
            variant_number = self.student_variants[key]
        else:
            if student_number > variant_count:
                variant_number = student_number % variant_count
                if variant_number == 0:
                    variant_number = variant_count
            else:
                variant_number = student_number
        self.variant_entry.configure(state="normal")
        self.variant_entry.delete(0, tk.END)
        self.variant_entry.insert(0, str(variant_number))
        self.variant_entry.configure(state="normal")

    def save_variant(self, event):
        key = f"{self.group_var.get()};{self.student_var.get()}"
        variant_number = self.variant_entry.get()
        self.student_variants[key] = variant_number
        self.save_student_variants()

    def prev_student(self):
        if self.current_student_index > 0:
            self.current_student_index -= 1
            self.student_var.set(self.student_names[self.current_student_index])
            self.calculate_variant()
        self.reset_fields()

    def next_student(self):
        if self.current_student_index < len(self.student_names) - 1:
            self.current_student_index += 1
            self.student_var.set(self.student_names[self.current_student_index])
            self.calculate_variant()
        self.reset_fields()

    def create_criteria_tab(self):
        self.criteria_scores = {}

        canvas = tk.Canvas(self.criteria_frame)
        scrollbar = ttk.Scrollbar(
            self.criteria_frame, orient="vertical", command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Создаем фрейм внутри канваса
        self.criteria_inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.criteria_inner_frame, anchor="nw")

        # Привязка прокрутки колесиком мыши
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.criteria_inner_frame.bind(
            "<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel)
        )
        self.criteria_inner_frame.bind(
            "<Leave>", lambda e: canvas.unbind_all("<MouseWheel>")
        )

        self.criteria_inner_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Изначально загружаем критерии для выбранного домашнего задания
        # Если домашнее задание не выбрано, оставляем пустым
        self.current_criteria = []
        self.current_homework = self.hw_name_var.get()
        if self.current_homework:
            self.load_criteria_for_homework(self.current_homework)

    def create_penalty_tab(self):
        canvas = tk.Canvas(self.penalty_frame)
        scrollbar = ttk.Scrollbar(
            self.penalty_frame, orient="vertical", command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Создаем фрейм внутри канваса
        self.penalty_inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.penalty_inner_frame, anchor="nw")

        # Привязка прокрутки колесиком мыши
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.penalty_inner_frame.bind(
            "<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel)
        )
        self.penalty_inner_frame.bind(
            "<Leave>", lambda e: canvas.unbind_all("<MouseWheel>")
        )

        self.penalty_inner_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        # Создаем штрафы из JSON-файла
        self.create_penalties_from_json()

    def load_criteria_from_json(self):
        # Загрузка происходит в load_homework_names
        pass

    def load_criteria_for_homework(self, homework_name):
        # Очистка предыдущих критериев
        self.criteria_scores = {}
        for widget in self.criteria_inner_frame.winfo_children():
            widget.destroy()

        self.section_max_scores = {}
        sections = self.criteria_data.get("sections", {})
        if homework_name in sections:
            self.current_criteria = sections[homework_name]
            self.create_criteria()
        else:
            tk.messagebox.showerror("Ошибка", f"Критерии для '{homework_name}' не найдены.")
            self.current_criteria = []

    def create_criteria(self):
        for section in self.current_criteria:
            title = section.get("title", "")
            section_type = section.get("type", "")
            max_score = section.get("max_score", 0)
            self.section_max_scores[title] = max_score
            options = section.get("options", [])

            section_frame = ttk.Labelframe(self.criteria_inner_frame, text=title)
            section_frame.pack(fill="x", padx=10, pady=5)

            vars_list = []

            if section_type == "radio_with_subchecks":
                var = tk.IntVar(value=options[0].get("score", 0))

                for option in options:
                    score = option.get("score", 0)
                    text = option.get("text", "")
                    suboptions = option.get("suboptions", [])

                    rb = ttk.Radiobutton(
                        section_frame,
                        text=text,
                        variable=var,
                        value=score,
                        command=lambda opt=option, var_main=var, opts=options: self.radiobutton_callback(
                            opt, var_main, opts
                        ),
                    )
                    rb.pack(anchor="w")
                    vars_list.append((var, score))

                    if suboptions:
                        sub_frame = ttk.Frame(section_frame)
                        sub_frame.pack(anchor="w", padx=20)

                        option["suboption_vars"] = []
                        for subtext in suboptions:
                            var_cb = tk.BooleanVar()
                            cb = ttk.Checkbutton(
                                sub_frame,
                                text=subtext,
                                variable=var_cb,
                                command=lambda v_cb=var_cb, s=score, var_main=var: self.checkbox_callback(
                                    v_cb, s, var_main
                                ),
                            )
                            cb.pack(anchor="w")
                            option["suboption_vars"].append(var_cb)
                            vars_list.append(
                                (var_cb, 0)
                            )  # Значение 0, т.к. учитываем ниже

                # Сохраняем ссылки для дальнейшего использования
                self.criteria_scores[title] = {
                    "type": section_type,
                    "vars": vars_list,
                    "options": options,
                    "main_var": var,
                }

            elif section_type == "checkbox":
                vars_list = []
                for option in options:
                    score = option.get("score", 0)
                    text = option.get("text", "")
                    var_cb = tk.BooleanVar(value=score > 0)
                    cb = ttk.Checkbutton(
                        section_frame,
                        text=text,
                        variable=var_cb,
                        command=lambda: None,
                    )
                    cb.pack(anchor="w")
                    vars_list.append((var_cb, score))

                self.criteria_scores[title] = {
                    "type": section_type,
                    "vars": vars_list,
                }

    def create_penalties_from_json(self):
        # Очистка предыдущих штрафов
        for widget in self.penalty_inner_frame.winfo_children():
            widget.destroy()

        tk.Label(self.penalty_inner_frame, text="Дополнительные штрафы:").pack(
            anchor="w", pady=5
        )

        self.penalty_vars = []
        self.penalty_texts = []

        penalties = self.criteria_data.get("penalties", [])
        for penalty in penalties:
            text = penalty.get("text", "")
            score = penalty.get("score", 0)
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.penalty_inner_frame, text=text, variable=var)
            cb.pack(anchor="w")
            self.penalty_vars.append((var, score))
            self.penalty_texts.append(text)

        # Просрочка
        delay_info = self.criteria_data.get("delays", {})
        delay_text = delay_info.get("text", "")
        self.delay_penalty_per_day = delay_info.get("score_per_day", -2)

        delay_frame = ttk.Frame(self.penalty_inner_frame)
        delay_frame.pack(fill="x", pady=10)
        tk.Label(delay_frame, text=delay_text).pack(side="left")
        self.delay_entry = tk.Entry(delay_frame, width=5)
        self.delay_entry.pack(side="left", padx=5)
        self.delay_entry.insert(0, "0")

    def checkbox_callback(self, checkbox_var, score, main_var):
        if checkbox_var.get():
            main_var.set(score)
        else:
            pass

    def radiobutton_callback(self, current_option, var_main, options_list):
        # Сбрасываем все чекбоксы под всеми опциями
        for opt in options_list:
            if "suboption_vars" in opt:
                for sub_var in opt["suboption_vars"]:
                    sub_var.set(False)
        # Устанавливаем значение радиокнопки
        var_main.set(current_option.get("score", 0))

    def create_report_tab(self):
        ttk.Label(
            self.report_frame, text="Нажмите кнопку для формирования оценочного листа."
        ).pack(pady=10)
        self.generate_button = ttk.Button(
            self.report_frame,
            text="Сформировать и сохранить оценочный лист",
            command=self.generate_report,
        )
        self.generate_button.pack(pady=5)
        self.copy_button = ttk.Button(
            self.report_frame,
            text="Скопировать картинку в буфер обмена",
            command=self.copy_to_clipboard,
        )
        self.copy_button.pack(pady=5)
        # Добавляем кнопку очистки полей
        self.reset_button = ttk.Button(
            self.report_frame, text="Очистить все поля", command=self.reset_fields
        )
        self.reset_button.pack(pady=5)

    def reset_fields(self):
        # Сбрасываем поля на вкладке "Информация о студенте"
        self.variant_count_entry.delete(0, tk.END)
        self.variant_count_entry.insert(0, "29")
        self.on_time.set(True)

        # Сбрасываем критерии оценки
        for section, data in self.criteria_scores.items():
            if data["type"] == "radio_with_subchecks":
                data["main_var"].set(0)  # Устанавливаем на значение по умолчанию
                for option in data["options"]:
                    if "suboption_vars" in option:
                        for var_cb in option["suboption_vars"]:
                            var_cb.set(False)
            elif data["type"] == "checkbox":
                for var_cb, _ in data["vars"]:
                    var_cb.set(False)

        # Сбрасываем дополнительные штрафы
        for var, _ in self.penalty_vars:
            var.set(False)
        self.delay_entry.delete(0, tk.END)
        self.delay_entry.insert(0, "0")

        self.status_var.set("Все поля сброшены к значениям по умолчанию.")

    def generate_report(self, save_to_file=True):
        # Ваш существующий код для генерации отчета
        # ...

        # Проверка, что критерии загружены
        if not self.current_criteria:
            self.status_var.set("Пожалуйста, выберите домашнее задание.")
            return

        # Остальная часть метода остается без изменений
        # ...

    # Остальные методы остаются без изменений
    # ...


if __name__ == "__main__":
    root = tk.Tk()
    app = EvaluationApp(root)
    root.mainloop()
