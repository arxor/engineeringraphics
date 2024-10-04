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

        self.create_info_tab()
        self.load_info_parameters()  # Load saved parameters
        self.create_criteria_tab()
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
            "hw_name": self.hw_name_entry.get(),
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
                self.hw_name_entry.delete(0, tk.END)
                self.hw_name_entry.insert(0, data.get("hw_name", ""))
                self.variant_count_entry.delete(0, tk.END)
                self.variant_count_entry.insert(0, data.get("variant_count", "30"))
                self.group_var.set(data.get("group", ""))
                # Update the student list based on the loaded group
                self.update_student_list(None)
                self.student_var.set(data.get("student", ""))
                # Update student info based on the loaded student
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
                    # Извлекаем группу из поля 'Данные о пользователе'
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

    def create_info_tab(self):
        # Название домашней работы
        tk.Label(self.info_frame, text="Название домашней работы:").grid(
            row=0, column=0, sticky="w", pady=5, padx=5
        )
        self.hw_name_entry = tk.Entry(self.info_frame, width=30)
        self.hw_name_entry.grid(row=0, column=1, sticky="w", pady=5)

        # Количество вариантов
        tk.Label(self.info_frame, text="Количество вариантов:").grid(
            row=1, column=0, sticky="w", pady=5, padx=5
        )
        self.variant_count_entry = tk.Entry(self.info_frame, width=10)
        self.variant_count_entry.grid(row=1, column=1, sticky="w", pady=5)
        self.variant_count_entry.insert(0, "30")  # По умолчанию 30 вариантов

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
            variant_count = 30  # По умолчанию 30 вариантов

        student_number = self.current_student_index + 1  # Нумерация с 1
        key = f"{self.group_var.get()};{self.student_var.get()}"
        if key in self.student_variants:
            variant_number = self.student_variants[key]
        else:
            if student_number > 29:
                variant_number = student_number - 29
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

        # Создаем критерии
        self.create_criteria()

    def create_criteria(self):
        # Критерий 1
        self.create_section1()

        # Критерий 2
        self.create_section2()

        # Критерий 3
        self.create_section3()

        # Критерий 4
        self.create_section4()

    def create_section1(self):
        section_title = "1. Сходство итогового эскиза с изображением (0–2 балла)"
        section_frame = ttk.Labelframe(self.criteria_inner_frame, text=section_title)
        section_frame.pack(fill="x", padx=10, pady=5)

        var = tk.IntVar(value=2)
        vars_list = []

        def on_radiobutton_one_selected():
            for cb in self.major_discrepancies_vars:
                cb.set(False)

        def on_radiobutton_two_selected():
            for cb in self.minor_discrepancies_vars:
                cb.set(False)

        def on_radiobutton_selected():
            on_radiobutton_one_selected()
            on_radiobutton_two_selected()

        def on_checkbox_one_selected():
            var.set(1)
            on_radiobutton_one_selected()

        def on_checkbox_two_selected():
            var.set(0)
            on_radiobutton_two_selected()

        # 2 балла
        rb_full_match = ttk.Radiobutton(
            section_frame,
            text="2 балла — Полное соответствие изображению.",
            variable=var,
            value=2,
            command=on_radiobutton_selected,
        )
        rb_full_match.pack(anchor="w")
        vars_list.append((var, 2))

        # -1 балл
        rb_minor_discrepancies = ttk.Radiobutton(
            section_frame,
            text="-1 балл — Небольшие несовпадения:",
            variable=var,
            value=1,
            command=on_radiobutton_one_selected,
        )
        rb_minor_discrepancies.pack(anchor="w")
        vars_list.append((var, 1))

        # Подпункты для -1 балла
        self.minor_discrepancies_vars = []
        self.minor_discrepancies_texts = []

        sub_frame_minor = ttk.Frame(section_frame)
        sub_frame_minor.pack(anchor="w", padx=20)

        for text in [
            "Несовпадение одного из размеров.",
            "Небольшие отклонения в форме.",
            "Лишние элементы на эскизе.",
        ]:
            var_cb = tk.BooleanVar()
            cb = ttk.Checkbutton(
                sub_frame_minor,
                text=text,
                variable=var_cb,
                command=on_checkbox_one_selected,
            )
            cb.pack(anchor="w")
            self.minor_discrepancies_vars.append(var_cb)
            self.minor_discrepancies_texts.append(text)

        # -2 балла
        rb_major_discrepancies = ttk.Radiobutton(
            section_frame,
            text="-2 балла — Существенные расхождения:",
            variable=var,
            value=0,
            command=on_radiobutton_two_selected,
        )
        rb_major_discrepancies.pack(anchor="w")
        vars_list.append((var, 0))

        # Подпункты для -2 балла
        self.major_discrepancies_vars = []
        self.major_discrepancies_texts = []

        sub_frame_major = ttk.Frame(section_frame)
        sub_frame_major.pack(anchor="w", padx=20)

        for text in [
            "Отсутствие части эскиза.",
            "Эскиз только отдаленно напоминает вариант.",
            "Половина эскиза отсутствует.",
        ]:
            var_cb = tk.BooleanVar()
            cb = ttk.Checkbutton(
                sub_frame_major,
                text=text,
                variable=var_cb,
                command=on_checkbox_two_selected,
            )
            cb.pack(anchor="w")
            self.major_discrepancies_vars.append(var_cb)
            self.major_discrepancies_texts.append(text)

        self.criteria_scores[section_title] = vars_list

    def create_section2(self):
        section_title = "2. Правильность использования линий построения (0–4 балла)"
        section_frame = ttk.Labelframe(self.criteria_inner_frame, text=section_title)
        section_frame.pack(fill="x", padx=10, pady=5)

        var = tk.IntVar(value=4)
        vars_list = []

        def on_radiobutton_selected():
            for cb in self.line_errors_vars:
                cb.set(False)

        # 4 балла
        rb_all_correct = ttk.Radiobutton(
            section_frame,
            text="4 балла — Все необходимые линии построения присутствуют и правильно используются.",
            variable=var,
            value=4,
            command=on_radiobutton_selected,
        )
        rb_all_correct.pack(anchor="w")
        vars_list.append((var, 4))

        # -1 балл за каждую ошибку
        rb_errors = ttk.Radiobutton(
            section_frame,
            text="-1 балл за каждую из следующих ошибок:",
            variable=var,
            value=-1,
        )
        rb_errors.pack(anchor="w")
        vars_list.append((var, -1))

        # Список ошибок
        self.line_errors_vars = []
        self.line_errors_texts = []

        sub_frame_errors = ttk.Frame(section_frame)
        sub_frame_errors.pack(anchor="w", padx=20)

        for text in [
            "Отсутствуют осевые линии.",
            "Отсутствуют необходимые линии построения.",
            "Линии построения не влияют на эскиз.",
            "Избыточное количество линий построения.",
            "Неиспользуемые линии построения.",
            "Линии построения пересекаются некорректно.",
        ]:
            var_cb = tk.BooleanVar()
            cb = ttk.Checkbutton(
                sub_frame_errors,
                text=text,
                variable=var_cb,
                command=lambda: var.set(-1),
            )
            cb.pack(anchor="w")
            self.line_errors_vars.append(var_cb)
            self.line_errors_texts.append(text)

        # -4 балла
        rb_no_lines = ttk.Radiobutton(
            section_frame,
            text="-4 балла — Практически полное отсутствие линий построения.",
            variable=var,
            value=0,
            command=on_radiobutton_selected,
        )
        rb_no_lines.pack(anchor="w")
        vars_list.append((var, 0))

        self.criteria_scores[section_title] = vars_list

    def create_section3(self):
        section_title = (
            "3. Правильность использования линий изображения (обводки) (0–2 балла)"
        )
        section_frame = ttk.Labelframe(self.criteria_inner_frame, text=section_title)
        section_frame.pack(fill="x", padx=10, pady=5)

        vars_list = []

        # 2 балла
        var_correct_lines = tk.BooleanVar(value=True)
        cb_correct_lines = ttk.Checkbutton(
            section_frame,
            text="2 балла — Линии обводки использованы правильно.",
            variable=var_correct_lines,
        )
        cb_correct_lines.pack(anchor="w")
        vars_list.append((var_correct_lines, 2))

        # -1 балл за каждую ошибку
        self.obvodka_errors_vars = []
        self.obvodka_errors_texts = []

        for text in [
            "Лишние линии обводки (-1 балл).",
            "Отсутствуют необходимые линии обводки (-1 балл).",
            "Неправильное применение типов линий (-1 балл).",
            "Линии обводки пересекаются некорректно (-1 балл).",
            "Неправильное положение линий обводки (-1 балл).",
        ]:
            var_cb = tk.BooleanVar()
            cb = ttk.Checkbutton(section_frame, text=text, variable=var_cb)
            cb.pack(anchor="w")
            self.obvodka_errors_vars.append(var_cb)
            self.obvodka_errors_texts.append(text)
            vars_list.append((var_cb, -1))

        self.criteria_scores[section_title] = vars_list

    def create_section4(self):
        section_title = "4. Правильность использования размеров (0–2 балла)"
        section_frame = ttk.Labelframe(self.criteria_inner_frame, text=section_title)
        section_frame.pack(fill="x", padx=10, pady=5)

        vars_list = []

        # 2 балла
        var_correct_dimensions = tk.BooleanVar(value=True)
        cb_correct_dimensions = ttk.Checkbutton(
            section_frame,
            text="2 балла — Все размеры установлены корректно и соответствуют заданию.",
            variable=var_correct_dimensions,
        )
        cb_correct_dimensions.pack(anchor="w")
        vars_list.append((var_correct_dimensions, 2))

        # -1 балл за каждую ошибку
        self.dimension_errors_vars = []
        self.dimension_errors_texts = []

        for text in [
            "Отсутствуют некоторые размеры (-1 балл).",
            "Размеры установлены некорректно (-1 балл).",
            "Размеры пересекаются (-1 балл).",
            "Неправильное отмеривание углов (-1 балл).",
            "Размеры выходят за пределы листа (-1 балл).",
            "Использованы константные значения вместо размеров (-1 балл).",
        ]:
            var_cb = tk.BooleanVar()
            cb = ttk.Checkbutton(section_frame, text=text, variable=var_cb)
            cb.pack(anchor="w")
            self.dimension_errors_vars.append(var_cb)
            self.dimension_errors_texts.append(text)
            vars_list.append((var_cb, -1))

        self.criteria_scores[section_title] = vars_list

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

        # Создаем штрафы
        self.create_penalties()

    def create_penalties(self):
        tk.Label(self.penalty_inner_frame, text="Дополнительные штрафы:").pack(
            anchor="w", pady=5
        )

        self.penalties = [
            ("Неправильное название архива (-3 балла)", -3),
            ("Чертеж выходит за границы листа (-1 балл)", -1),
            ("Чертеж переопределен (избыточные зависимости) (-1 балл)", -1),
            (
                "Неиспользован необходимый массив или использован некорректно (-1 балл)",
                -1,
            ),
            ("Не указаны осевые линии при необходимости (-1 балл)", -1),
            ("Чертеж выполнен не в той области рабочего пространства (-1 балл)", -1),
            ("Обнаружен плагиат (оценка обнуляется)", 0),
        ]
        self.penalty_vars = []
        self.penalty_texts = []
        for text, value in self.penalties:
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.penalty_inner_frame, text=text, variable=var)
            cb.pack(anchor="w")
            self.penalty_vars.append((var, value))
            self.penalty_texts.append(text)

        # Просрочка
        delay_frame = ttk.Frame(self.penalty_inner_frame)
        delay_frame.pack(fill="x", pady=10)
        tk.Label(
            delay_frame, text="Дней просрочки сдачи задания (-2 балла за каждый день):"
        ).pack(side="left")
        self.delay_entry = tk.Entry(delay_frame, width=5)
        self.delay_entry.pack(side="left", padx=5)
        self.delay_entry.insert(0, "0")

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
        self.variant_count_entry.insert(0, "30")
        self.on_time.set(True)

        # Сбрасываем критерии оценки
        for section, vars_list in self.criteria_scores.items():
            if section.startswith("1."):
                var = vars_list[0][0]
                var.set(2)
                for var_cb in self.minor_discrepancies_vars:
                    var_cb.set(False)
                for var_cb in self.major_discrepancies_vars:
                    var_cb.set(False)
            elif section.startswith("2."):
                var = vars_list[0][0]
                var.set(4)
                for var_cb in self.line_errors_vars:
                    var_cb.set(False)
            elif section.startswith("3."):
                vars_list[0][0].set(True)
                for var_cb in self.obvodka_errors_vars:
                    var_cb.set(False)
            elif section.startswith("4."):
                vars_list[0][0].set(True)
                for var_cb in self.dimension_errors_vars:
                    var_cb.set(False)

        # Сбрасываем дополнительные штрафы
        for var, _ in self.penalty_vars:
            var.set(False)
        self.delay_entry.delete(0, tk.END)
        self.delay_entry.insert(0, "0")

        self.status_var.set("Все поля сброшены к значениям по умолчанию.")

    def generate_report(self, save_to_file=True):
        # Проверка заполнения обязательных полей
        if not self.student_var.get() or not self.group_var.get():
            self.status_var.set("Пожалуйста, выберите группу и студента.")
            return

        # Рассчитываем баллы по критериям
        total_score = 0
        section_scores = {}
        section_comments = {}
        for section, vars_list in self.criteria_scores.items():
            score = 0
            comments = []

            if section.startswith("1."):
                var = vars_list[0][0]
                score = var.get()

                # Учитываем подвыборы для -1 и -2 баллов
                if score == 1:
                    # Если выбрано -1 балл, вычитаем 1 балл
                    score = 2 - 1
                    # Добавляем комментарии из выбранных подпунктов
                    for var_cb, text in zip(
                        self.minor_discrepancies_vars, self.minor_discrepancies_texts
                    ):
                        if var_cb.get():
                            comments.append(text)
                elif score == 0:
                    # Если выбрано -2 балла, вычитаем 2 балла
                    score = 2 - 2
                    # Добавляем комментарии из выбранных подпунктов
                    for var_cb, text in zip(
                        self.major_discrepancies_vars, self.major_discrepancies_texts
                    ):
                        if var_cb.get():
                            comments.append(text)

            elif section.startswith("2."):
                var = vars_list[0][0]
                radio_value = var.get()
                if radio_value == 4:
                    score = 4
                elif radio_value == -1:
                    # Вычитаем по -1 баллу за каждую отмеченную ошибку
                    errors = sum(1 for var_cb in self.line_errors_vars if var_cb.get())
                    score = 4 - errors
                    # Добавляем комментарии из выбранных ошибок
                    for var_cb, text in zip(
                        self.line_errors_vars, self.line_errors_texts
                    ):
                        if var_cb.get():
                            comments.append(text)
                elif radio_value == 0:
                    score = 0

            elif section.startswith("3."):
                for var, value in vars_list:
                    if var.get():
                        score += value
                        if value == -1:
                            # Добавляем комментарии из ошибок
                            index = self.obvodka_errors_vars.index(var)
                            comments.append(self.obvodka_errors_texts[index])
                # Ограничиваем баллы в пределах 0-2
                score = max(0, min(2, score))

            elif section.startswith("4."):
                for var, value in vars_list:
                    if var.get():
                        score += value
                        if value == -1:
                            # Добавляем комментарии из ошибок
                            index = self.dimension_errors_vars.index(var)
                            comments.append(self.dimension_errors_texts[index])
                # Ограничиваем баллы в пределах 0-2
                score = max(0, min(2, score))

            section_scores[section] = score
            section_comments[section] = comments
            total_score += score

        # Учёт штрафов
        penalty_score = 0
        penalty_comments = []
        for (var, value), text in zip(self.penalty_vars, self.penalty_texts):
            if var.get():
                # Если выбран пункт про плагиат, оценка обнуляется
                if "плагиат" in text.lower() and var.get():
                    total_score = 0
                    penalty_comments.append(text)
                else:
                    penalty_score += value
                    penalty_comments.append(text)

        # Учёт просрочки
        try:
            delay_days = int(self.delay_entry.get())
            delay_penalty = -2 * delay_days
            penalty_score += delay_penalty
            if delay_days > 0:
                penalty_comments.append(
                    f"Просрочка сдачи задания на {delay_days} дней (-{2 * delay_days} балла)"
                )
        except ValueError:
            delay_days = 0
            delay_penalty = 0

        # Общий итог
        final_score = max(0, total_score + penalty_score)
        if not self.on_time.get():
            final_score = max(0, final_score)

        # Генерация изображения оценочного листа
        self.create_image(
            section_scores, section_comments, penalty_comments, final_score, delay_days
        )

        # Сохранение изображения
        if save_to_file:
            self.save_image()

    def create_image(
        self,
        section_scores,
        section_comments,
        penalty_comments,
        final_score,
        delay_days,
    ):
        # Настройки изображения
        img_width = 1200  # Увеличено разрешение
        img_height = 1500
        background_color = (255, 255, 255)
        text_color = (0, 0, 0)

        # Шрифты Gilroy
        title_font = ImageFont.truetype("gilroy-bold.ttf", 36)
        header_font = ImageFont.truetype("gilroy-medium.ttf", 24)
        text_font = ImageFont.truetype("gilroy-regular.ttf", 18)

        img = Image.new("RGB", (img_width, img_height), color=background_color)
        draw = ImageDraw.Draw(img)

        y_position = 20

        # Заголовок
        draw.text(
            (img_width // 2 - 150, y_position),
            "Оценочный лист",
            font=title_font,
            fill=text_color,
        )
        y_position += 50

        # Информация о студенте
        student_info = f"Студент: {self.student_var.get()}    Группа: {self.group_var.get()}    Вариант: {self.variant_entry.get()}"
        draw.text((50, y_position), student_info, font=text_font, fill=text_color)
        y_position += 30
        date_info = f"Сдано вовремя: {'Да' if self.on_time.get() else 'Нет'}    Дней просрочки: {delay_days}"
        draw.text((50, y_position), date_info, font=text_font, fill=text_color)
        y_position += 40

        # Критерии
        for section, score in section_scores.items():
            draw.text((50, y_position), section, font=header_font, fill=text_color)
            y_position += 30
            draw.text(
                (70, y_position), f"Баллы: {score}", font=text_font, fill=text_color
            )
            y_position += 30
            comments = section_comments.get(section, [])
            for comment in comments:
                draw.text(
                    (90, y_position), f"- {comment}", font=text_font, fill=text_color
                )
                y_position += 25
            y_position += 10

        # Штрафы
        draw.text(
            (50, y_position),
            "Дополнительные штрафы:",
            font=header_font,
            fill=text_color,
        )
        y_position += 30
        if penalty_comments:
            for comment in penalty_comments:
                draw.text(
                    (70, y_position), f"- {comment}", font=text_font, fill=text_color
                )
                y_position += 25
        else:
            draw.text((70, y_position), "Нет", font=text_font, fill=text_color)
            y_position += 25

        y_position += 10
        draw.line((50, y_position, img_width - 50, y_position), fill=text_color)
        y_position += 10

        # Итоговая оценка
        draw.text(
            (50, y_position),
            f"Итоговая оценка: {final_score} из 10",
            font=header_font,
            fill=text_color,
        )

        self.generated_image = img

    def save_image(self):
        # Создание папки с названием домашней работы
        hw_name = self.hw_name_entry.get()
        if not hw_name:
            self.status_var.set("Пожалуйста, введите название домашней работы.")
            return
        if not os.path.exists(SUB_PATH):
            os.makedirs(SUB_PATH)
        hw_name = os.path.join(SUB_PATH, hw_name)
        if not os.path.exists(hw_name):
            os.makedirs(hw_name)
        # Сохранение изображения с именем студента
        student_name = self.student_var.get().replace(" ", "_")
        filename = os.path.join(hw_name, f"{student_name}.png")
        self.generated_image.save(filename)
        self.status_var.set(f"Оценочный лист сохранен как '{filename}'.")

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


if __name__ == "__main__":
    root = tk.Tk()
    app = EvaluationApp(root)
    root.mainloop()
