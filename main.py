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

        # Создаем вкладки
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(fill="both", expand=True)

        # Вкладка с информацией о студенте
        self.info_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.info_frame, text="Информация о студенте")

        # Вкладка с критериями оценки
        self.criteria_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.criteria_frame, text="Критерии оценки")

        # Вкладка с дополнительными штрафами и поощрениями
        self.penalty_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.penalty_frame, text="Штрафы и поощрения")

        # Вкладка с генерацией отчета
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_frame, text="Генерация отчета")

        self.section_max_scores = {}
        self.current_criteria_source = None

        self.create_info_tab()
        self.create_criteria_tab()
        self.load_info_parameters()
        self.create_penalty_tab()
        self.create_report_tab()
        self.register_shortcuts()

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

    @staticmethod
    def _normalize_homework_name(hw_name):
        mapping = {
            "Домашнее задание №1 до 8": "Домашнее задание №1",
            "Домашнее задание №1 9-10": "Домашнее задание №1",
            "ДЗ_1": "Домашнее задание №1",
        }
        return mapping.get(hw_name, hw_name)

    def save_info_parameters(self):
        data = {
            "hw_name": self.hw_name_var.get(),
            "variant_count": self.variant_count_entry.get(),
            "group": self.group_var.get(),
            "student": self.student_var.get(),
            "variant": self.variant_entry.get(),
            "on_time": self.on_time.get(),
            "double_mode_enabled": self.double_mode_enabled.get() if hasattr(self, "double_mode_enabled") else False,
            "work_variant_is_eight": self.limit_to_eight.get() if hasattr(self, "limit_to_eight") else True,
        }
        with open("info_parameters.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)

    def load_info_parameters(self):
        if os.path.exists("info_parameters.json"):
            try:
                with open("info_parameters.json", "r", encoding="utf-8") as f:
                    data = json.load(f)
            except json.JSONDecodeError:
                with open("info_parameters.json", "r", encoding="utf-8-sig") as f:
                    data = json.load(f)
                hw_name = self._normalize_homework_name(data.get("hw_name", ""))
                self.hw_name_var.set(hw_name)
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
                if hasattr(self, "double_mode_enabled"):
                    self.double_mode_enabled.set(data.get("double_mode_enabled", False))
                if hasattr(self, "limit_to_eight"):
                    self.limit_to_eight.set(data.get("work_variant_is_eight", True))
                self._sync_double_mode_controls()

    def load_student_data(self):
        """Load student list from CSV, creating a scaffold file if it is absent."""
        self.student_data = []
        self.groups = set()
        self.student_lookup = {}
        filename = "student_list.csv"
        if not os.path.exists(filename):
            with open(filename, "w", encoding="utf-8", newline="") as csvfile:
                csvfile.write("ФИО;Группа;Номер Варианта\n")
            tk.messagebox.showwarning(
                "Нет данных о студентах",
                "Файл student_list.csv не найден. Создан шаблонный файл. "
                "Добавьте студентов и перезапустите приложение.",
            )
            self.groups = []
            return

        def _read_students(delimiter):
            with open(filename, encoding="utf-8-sig") as csvfile:
                reader = csv.DictReader(csvfile, delimiter=delimiter)
                for row in reader:
                    normalized_row = {
                        (key.strip().lstrip("\ufeff") if key else key): value
                        for key, value in row.items()
                    }
                    fio = (
                        normalized_row.get("ФИО")
                        or " ".join(
                            part
                            for part in [
                                normalized_row.get("Фамилия", "").strip(),
                                normalized_row.get("Имя", "").strip(),
                            ]
                            if part
                        )
                    ).strip()
                    group_name = (
                        normalized_row.get("Группа")
                        or normalized_row.get("Группы")
                        or normalized_row.get("Данные о пользователе", "")
                    ).strip()
                    variant_number = (
                        normalized_row.get("Номер Варианта")
                        or normalized_row.get("Вариант")
                        or ""
                    ).strip()
                    if not (fio and group_name):
                        continue
                    student_record = {
                        "ФИО": fio,
                        "Группа": group_name,
                        "Номер Варианта": variant_number,
                    }
                    self.groups.add(group_name)
                    self.student_data.append(student_record)
                    self.student_lookup[(group_name, fio)] = student_record

        # Попытка прочитать как CSV с разделителем ';', затем ','.
        try:
            _read_students(";")
        except csv.Error:
            self.student_data.clear()
            self.student_lookup.clear()
            _read_students(",")

        self.groups = sorted(self.groups)
        if not self.groups:
            tk.messagebox.showwarning(
                "Пустой список студентов",
                "Не удалось найти валидные записи в student_list.csv. "
                "Проверьте структуру файла (ФИО;Группа;Номер Варианта).",
            )

    def save_student_list(self):
        filename = "student_list.csv"
        fieldnames = ["ФИО", "Группа", "Номер Варианта"]
        with open(filename, "w", encoding="utf-8", newline="") as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=";")
            writer.writeheader()
            for record in self.student_data:
                writer.writerow(
                    {
                        "ФИО": record.get("ФИО", ""),
                        "Группа": record.get("Группа", ""),
                        "Номер Варианта": record.get("Номер Варианта", ""),
                    }
                )

    def load_homework_names(self):
        try:
            with open("criteria.json", "r", encoding="utf-8") as f:
                self.criteria_data = json.load(f)
            self.homework_names = list(self.criteria_data.get("sections", {}).keys())
        except Exception as e:
            self.criteria_data = {"sections": {}, "penalties": [], "rewards": [], "delays": {}}
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
        self.variant_count_entry.insert(0, "29")

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
        self.on_time.trace_add("write", self._on_on_time_toggle)
        tk.Checkbutton(
            self.info_frame, text="Сдано вовремя", variable=self.on_time
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=5)

        # Двойной режим оценивания
        self.double_mode_enabled = tk.BooleanVar(value=False)
        self.double_mode_enabled.trace_add("write", self._on_double_mode_toggle)
        tk.Checkbutton(
            self.info_frame,
            text="Включить двойной режим оценивания",
            variable=self.double_mode_enabled,
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=5)

        # Вариант работы (8 или 10 баллов)
        self.limit_to_eight = tk.BooleanVar(value=True)
        self.limit_to_eight.trace_add("write", self._on_limit_to_eight_toggle)
        self.student_variant_checkbox = tk.Checkbutton(
            self.info_frame,
            text="Студент выбрал вариант работы на макс. оценку 8",
            variable=self.limit_to_eight,
        )
        self.student_variant_checkbox.grid(
            row=8, column=0, columnspan=2, sticky="w", pady=5
        )
        self.student_variant_checkbox.configure(state="disabled")
        self._sync_double_mode_controls()

    def on_homework_selected(self, event):
        selected_homework = self.hw_name_var.get()
        self.current_homework = selected_homework
        self.load_criteria_for_homework(selected_homework)
        # Нет необходимости обновлять штрафы и поощрения, так как они общие

    def update_student_list(self, event):
        selected_group = self.group_var.get()
        students_in_group = sorted(
            [s for s in self.student_data if s["Группа"] == selected_group],
            key=lambda x: x["ФИО"],
        )
        self.students_in_group = students_in_group  # Сохраняем для навигации
        self.student_names = [s["ФИО"] for s in students_in_group]
        self.student_combobox["values"] = self.student_names
        if self.student_names:
            self.current_student_index = 0
            self.student_var.set(self.student_names[0])
            self.update_student_info(None)

    def set_criteria_to_max(self):
        # Устанавливаем критерии на максимальные баллы по умолчанию
        for section, data in self.criteria_scores.items():
            if data["type"] == "radio_with_subchecks":
                # Найдём опцию с максимальным score
                max_option = max(data["options"], key=lambda x: float(x.get("score", 0.0)))
                data["main_var"].set(float(max_option["score"]))
                # Сбрасываем субопции
                for option in data["options"]:
                    if "suboption_vars" in option:
                        for var_cb in option["suboption_vars"]:
                            var_cb.set(False)

            elif data["type"] == "checkbox":
                # Ставим True для всех чекбоксов с положительным score, чтобы получить максимум
                for var_cb, var_score in data["vars"]:
                    var_cb.set(var_score > 0)

    def update_student_info(self, event):
        selected_student_name = self.student_var.get()
        if selected_student_name in self.student_names:
            self.current_student_index = self.student_names.index(selected_student_name)
            self.calculate_variant()
            # После пересчёта варианта устанавливаем критерии на максимальные значения:
            self.set_criteria_to_max()
            if hasattr(self, "double_mode_enabled") and self.double_mode_enabled.get():
                self.limit_to_eight.set(True)

    def calculate_variant(self):
        raw_count = self.variant_count_entry.get().strip()
        try:
            variant_count = int(raw_count)
            if variant_count <= 0:
                raise ValueError
        except ValueError:
            if hasattr(self, "status_var"):
                self.status_var.set("Количество вариантов должно быть целым положительным числом.")
            self.variant_entry.configure(state="normal")
            self.variant_entry.delete(0, tk.END)
            return

        student_number = self.current_student_index + 1  # Нумерация с 1
        group = self.group_var.get()
        student_name = self.student_var.get()
        record = self.student_lookup.get((group, student_name))
        variant_number = ""
        if record:
            variant_number = record.get("Номер Варианта", "").strip()

        if not variant_number:
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
        group = self.group_var.get()
        student_name = self.student_var.get()
        variant_number = self.variant_entry.get().strip()
        if not variant_number.isdigit():
            if hasattr(self, "status_var"):
                self.status_var.set("Введите числовое значение варианта.")
            return
        record = self.student_lookup.get((group, student_name))
        if record is not None:
            record["Номер Варианта"] = variant_number
            self.save_student_list()

    def prev_student(self):
        if self.current_student_index > 0:
            self.current_student_index -= 1
            self.student_var.set(self.student_names[self.current_student_index])
            self.calculate_variant()
        self.reset_fields()
        # Устанавливаем критерии на максимальные значения
        self.set_criteria_to_max()

    def next_student(self):
        if self.current_student_index < len(self.student_names) - 1:
            self.current_student_index += 1
            self.student_var.set(self.student_names[self.current_student_index])
            self.calculate_variant()
        self.reset_fields()
        # Устанавливаем критерии на максимальные значения
        self.set_criteria_to_max()

    def create_criteria_tab(self):
        self.criteria_scores = {}

        canvas = tk.Canvas(self.criteria_frame)
        scrollbar = ttk.Scrollbar(
            self.criteria_frame, orient="vertical", command=canvas.yview
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self.criteria_inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.criteria_inner_frame, anchor="nw")

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

        # Создаем штрафы и поощрения из JSON-файла
        self.create_penalties_and_rewards_from_json()

    def load_criteria_for_homework(self, homework_name):
        self.current_homework = homework_name
        sections = self.criteria_data.get("sections", {})
        self.current_criteria_source = None
        if homework_name and homework_name in sections:
            self.current_criteria_source = sections[homework_name]
        elif homework_name:
            tk.messagebox.showerror(
                "Ошибка", f"Критерии для '{homework_name}' не найдены."
            )
        self._render_current_criteria()

    def _get_current_criteria_list(self):
        source = getattr(self, "current_criteria_source", None)
        if isinstance(source, dict):
            base_sections = source.get("base", [])
            extended_sections = source.get("extended", [])
            use_extended = hasattr(self, "limit_to_eight") and not self.limit_to_eight.get()
            return list(base_sections) + (list(extended_sections) if use_extended else [])
        if isinstance(source, list):
            return source
        return []

    def _render_current_criteria(self):
        if not hasattr(self, "criteria_inner_frame"):
            return
        for widget in self.criteria_inner_frame.winfo_children():
            widget.destroy()
        self.criteria_scores = {}
        self.section_max_scores = {}
        self.current_criteria = self._get_current_criteria_list()
        if not self.current_criteria:
            return
        self.create_criteria()
        self.set_criteria_to_max()

    def create_criteria(self):
        for section in self.current_criteria:
            title = section.get("title", "")
            section_type = section.get("type", "")
            max_score = float(section.get("max_score", 0))
            self.section_max_scores[title] = max_score
            options = section.get("options", [])

            section_frame = ttk.Labelframe(self.criteria_inner_frame, text=title)
            section_frame.pack(fill="x", padx=10, pady=5)

            vars_list = []

            if section_type == "radio_with_subchecks":
                # Используем DoubleVar для корректной работы с дробными значениями
                initial_score = float(options[0].get("score", 0)) if options else 0.0
                var = tk.DoubleVar(value=initial_score)

                for option in options:
                    score = float(option.get("score", 0.0))
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
                            vars_list.append((var_cb, 0.0))  # 0, так как учитываем их через вычитание

                self.criteria_scores[title] = {
                    "type": section_type,
                    "vars": vars_list,
                    "options": options,
                    "main_var": var,
                }

            elif section_type == "checkbox":
                for option in options:
                    score = float(option.get("score", 0.0))
                    text = option.get("text", "")
                    var_cb = tk.BooleanVar(value=(score > 0))  # Можно оставить False, если нужно
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

    def create_penalties_and_rewards_from_json(self):
        # Очистка предыдущих штрафов и поощрений
        for widget in self.penalty_inner_frame.winfo_children():
            widget.destroy()

        # Штрафы
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
        self.delay_penalty_per_day = delay_info.get("score_per_day", 0)

        delay_frame = ttk.Frame(self.penalty_inner_frame)
        delay_frame.pack(fill="x", pady=10)
        tk.Label(delay_frame, text=delay_text).pack(side="left")
        self.delay_entry = tk.Entry(delay_frame, width=5)
        self.delay_entry.pack(side="left", padx=5)
        self.delay_entry.insert(0, "0")
        self.delay_entry.bind("<KeyRelease>", self._on_delay_changed)
        self.delay_entry.bind("<FocusOut>", self._on_delay_changed)

        # Поощрения
        tk.Label(self.penalty_inner_frame, text="Поощрения:").pack(anchor="w", pady=10)

        self.reward_items = []

        rewards = self.criteria_data.get("rewards", [])
        for reward in rewards:
            text = reward.get("text", "").strip()
            score = float(reward.get("score", 0) or 0)
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.penalty_inner_frame, text=text, variable=var)
            cb.pack(anchor="w")
            self.reward_items.append({"var": var, "text": text, "score": score})

    def checkbox_callback(self, checkbox_var, score, main_var):
        if checkbox_var.get():
            main_var.set(score)
        # else:
        #     pass

    def radiobutton_callback(self, current_option, var_main, options_list):
        # Сбрасываем все субопции для всех вариантов
        for opt in options_list:
            if "suboption_vars" in opt:
                for sub_var in opt["suboption_vars"]:
                    sub_var.set(False)

    def _on_on_time_toggle(self, *_):
        if not hasattr(self, "delay_entry"):
            return
        current_delay = self.delay_entry.get().strip()
        if self.on_time.get():
            if current_delay and current_delay != "0":
                self.delay_entry.delete(0, tk.END)
                self.delay_entry.insert(0, "0")
        else:
            if not current_delay or current_delay == "0":
                self.delay_entry.delete(0, tk.END)
                self.delay_entry.insert(0, "1")

    def _on_double_mode_toggle(self, *_):
        if self.double_mode_enabled.get() and not self.limit_to_eight.get():
            self.limit_to_eight.set(True)
        self._sync_double_mode_controls()

    def _sync_double_mode_controls(self):
        if not hasattr(self, "student_variant_checkbox"):
            return
        if self.double_mode_enabled.get():
            self.student_variant_checkbox.configure(state="normal")
        else:
            self.student_variant_checkbox.configure(state="disabled")
            if not self.limit_to_eight.get():
                self.limit_to_eight.set(True)

    def _on_limit_to_eight_toggle(self, *_):
        if not getattr(self, "current_homework", None):
            return
        if not getattr(self, "criteria_data", None):
            return
        self._render_current_criteria()

    def _resolve_scoring_scale(self, section_total_max):
        section_total_max = float(section_total_max or 0.0)
        limit_to_eight = hasattr(self, "limit_to_eight") and self.limit_to_eight.get()
        double_mode_on = hasattr(self, "double_mode_enabled") and self.double_mode_enabled.get()

        if limit_to_eight:
            target_cap = 8.0
            effective_cap = min(section_total_max, target_cap) if section_total_max > 0 else target_cap
            return max(0.0, effective_cap), target_cap, 1.0

        if double_mode_on:
            target_cap = 10.0
            effective_cap = target_cap
            return max(0.0, effective_cap), target_cap, 1.0

        display_cap = section_total_max if section_total_max > 0 else 10.0
        effective_cap = section_total_max
        return max(0.0, effective_cap), display_cap, 1.0

    @staticmethod
    def _format_score(value):
        if isinstance(value, (int, float)):
            formatted = f"{float(value):.2f}".rstrip("0").rstrip(".")
            return formatted if formatted else "0"
        return str(value)

    def _on_delay_changed(self, event=None):
        value = self.delay_entry.get().strip() if hasattr(self, "delay_entry") else ""
        if not value:
            if self.on_time.get() is False:
                self.on_time.set(True)
            return
        try:
            days = int(value)
        except ValueError:
            return
        if days < 0:
            return
        if days > 0:
            if self.on_time.get():
                self.on_time.set(False)
        else:
            if not self.on_time.get():
                self.on_time.set(True)

    def register_shortcuts(self):
        self.master.bind("<Control-Left>", self._shortcut_prev_student)
        self.master.bind("<Control-Right>", self._shortcut_next_student)
        self.master.bind("<Control-Return>", self._shortcut_generate_report)
        self.master.bind("<Control-s>", self._shortcut_generate_report)
        self.master.bind("<Control-Shift-C>", self._shortcut_copy_report)

    def _shortcut_prev_student(self, event):
        self.prev_student()
        return "break"

    def _shortcut_next_student(self, event):
        self.next_student()
        return "break"

    def _shortcut_generate_report(self, event):
        self.generate_report()
        return "break"

    def _shortcut_copy_report(self, event):
        self.copy_to_clipboard()
        return "break"

    def create_report_tab(self):
        ttk.Label(
            self.report_frame, text="Нажмите кнопку для формирования оценочного листа."
        ).pack(pady=10)

        # Добавляем поле для комментария
        ttk.Label(self.report_frame, text="Комментарий:").pack(pady=5)
        self.comment_text = tk.Text(self.report_frame, height=5, width=60)
        self.comment_text.pack(pady=5)

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
        self.on_time.set(True)
        if hasattr(self, "limit_to_eight"):
            self.limit_to_eight.set(True)
        self._sync_double_mode_controls()

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

        # Сбрасываем поощрения
        for reward_item in getattr(self, "reward_items", []):
            reward_item["var"].set(False)
        for reward_item in getattr(self, "reward_items", []):
            reward_item["var"].set(False)

        # Сбрасываем комментарий
        self.comment_text.delete("1.0", tk.END)

        self.status_var.set("Все поля сброшены к значениям по умолчанию.")

    def generate_report(self, save_to_file=True):
        if not self.student_var.get() or not self.group_var.get():
            self.status_var.set("Пожалуйста, выберите группу и студента.")
            return

        # Рассчитываем баллы по критериям
        section_scores = {}
        section_comments = {}
        section_raw_total = 0.0

        for section, data in self.criteria_scores.items():
            max_score = float(self.section_max_scores.get(section, 0.0))
            comments = []
            score = 0.0

            if data["type"] == "radio_with_subchecks":
                main_var = data["main_var"]
                main_score = float(main_var.get())  # main_score теперь float
                selected_option = None

                for option in data["options"]:
                    opt_score = float(option.get("score", 0.0))
                    if abs(opt_score - main_score) < 1e-9:
                        selected_option = option
                        break

                if selected_option:
                    if "suboptions" in selected_option and selected_option["suboptions"]:
                        num_selected_checkboxes = 0
                        for var_cb, subtext in zip(
                            selected_option["suboption_vars"],
                            selected_option["suboptions"],
                        ):
                            if var_cb.get():
                                num_selected_checkboxes += 1
                                comments.append(subtext)

                        deduction = abs(main_score) * num_selected_checkboxes
                        deduction = min(deduction, max_score)
                        score = max_score - deduction
                    else:
                        # Если нет субопций, score равен main_score, если он положительный,
                        # или (max_score + main_score), если main_score отрицательный
                        if main_score >= 0:
                            score = main_score
                        else:
                            # Если main_score отрицательный, уменьшаем от max_score
                            # Пример: max_score=10, main_score=-0.5 -> score=10 - 0.5 = 9.5
                            score = max_score + main_score
                else:
                    score = 0.0

            elif data["type"] == "checkbox":
                # Суммируем баллы за выбранные чекбоксы
                score_sum = 0.0
                for var_cb, var_score in data["vars"]:
                    if var_cb.get():
                        score_sum += var_score
                        if var_score < 0:
                            # Можно добавить текст ошибки, если нужно
                            pass
                # Ограничиваем в пределах 0 и max_score
                score = max(0.0, min(max_score, score_sum))

            section_scores[section] = score
            section_comments[section] = comments
            section_raw_total += score

        # Учёт штрафов
        penalty_score = 0.0
        penalty_comments = []
        disqualified = False
        penalty_pairs = getattr(self, "penalty_vars", [])
        penalty_texts = getattr(self, "penalty_texts", [])
        for (var, value), text in zip(penalty_pairs, penalty_texts):
            val = float(value)
            if not var.get():
                continue
            if val <= -1000:
                disqualified = True
                penalty_comments.append(text)
            else:
                penalty_score += val
                penalty_comments.append(text)

        # Учёт просрочки
        delay_raw = self.delay_entry.get().strip()
        if not delay_raw:
            delay_days = 0
        else:
            try:
                delay_days = int(delay_raw)
            except ValueError:
                self.status_var.set("Количество дней просрочки должно быть целым числом.")
                return
            if delay_days < 0:
                self.status_var.set("Количество дней просрочки не может быть отрицательным.")
                return

        effective_delay_days = delay_days
        if not self.on_time.get() and effective_delay_days == 0:
            effective_delay_days = 1

        forced_zero_due_to_delay = False
        delay_penalty_comment_index = None
        if effective_delay_days > 0:
            forced_zero_due_to_delay = True
            delay_penalty_comment_index = len(penalty_comments)
            penalty_comments.append("")

        comment_text = self.comment_text.get("1.0", tk.END).strip()

        reward_comments = []
        reward_bonus = 0.0
        for reward_item in getattr(self, "reward_items", []):
            if reward_item["var"].get():
                reward_comments.append(reward_item["text"])
                reward_bonus += float(reward_item["score"])

        if disqualified:
            total_score = 0.0
            penalty_score = 0.0
            reward_bonus = 0.0

        max_total_score = sum(self.section_max_scores.values()) or 0.0
        effective_cap, _display_cap, scaling_factor = self._resolve_scoring_scale(max_total_score)
        result_display_cap = 10.0  # Всегда показываем итог как «… из 10»

        if scaling_factor != 1.0:
            section_scores = {
                key: value * scaling_factor for key, value in section_scores.items()
            }
            penalty_score *= scaling_factor
            reward_bonus *= scaling_factor
        total_score = sum(section_scores.values())
        limit_to_eight_active = hasattr(self, "limit_to_eight") and self.limit_to_eight.get()
        double_mode_active = hasattr(self, "double_mode_enabled") and self.double_mode_enabled.get()

        if limit_to_eight_active:
            base_cap = effective_cap if effective_cap > 0 else 8.0
            reference_max = max_total_score if max_total_score > 0 else base_cap
            lost_points = max(0.0, reference_max - total_score)
            adjusted_total = base_cap - lost_points + penalty_score + reward_bonus
            final_score = max(0.0, min(base_cap, adjusted_total))
        elif double_mode_active:
            base_cap = effective_cap if effective_cap > 0 else 10.0
            reference_max = max_total_score if max_total_score > 0 else base_cap
            scaling_ratio = (base_cap / reference_max) if reference_max > 0 else 1.0
            lost_points = max(0.0, reference_max - total_score)
            adjusted_total = base_cap - (lost_points * scaling_ratio) + penalty_score + reward_bonus
            final_score = max(0.0, min(base_cap, adjusted_total))
        elif effective_cap > 0:
            base_total = total_score + penalty_score + reward_bonus
            final_score = max(0.0, min(effective_cap, base_total))
        else:
            final_score = max(0.0, total_score + penalty_score + reward_bonus)

        if forced_zero_due_to_delay:
            final_score = 0.0

        if delay_penalty_comment_index is not None and 0 <= delay_penalty_comment_index < len(penalty_comments):
            penalty_comments[delay_penalty_comment_index] = (
                f"Работа сдана с просрочкой на {effective_delay_days} дн. Итоговая оценка 0 согласно правилам."
            )

        delay_days = effective_delay_days

        if hasattr(self, "status_var"):
            self.status_var.set(
                f"Рассчитан итог: {self._format_score(final_score)} из {self._format_score(result_display_cap)}."
            )

        # Генерация изображения с отчётом
        self.create_image(
            section_scores,
            section_comments,
            penalty_comments,
            final_score,
            result_display_cap,
            delay_days,
            comment_text,
            reward_comments,
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
        max_score_cap,
        delay_days,
        comment,
        reward_comments,
    ):
        # Настройки изображения
        img_width = 1200  # Увеличено разрешение
        img_height = 1500  # Увеличено высоту для размещения поощрений
        background_color = (255, 255, 255)
        text_color = (0, 0, 0)

        # Определение путей к шрифтам (предполагается, что шрифты находятся в той же директории, что и скрипт)
        base_path = os.path.dirname(os.path.abspath(__file__))
        gilroy_black_path = os.path.join(base_path, "gilroy-black.ttf")
        gilroy_bold_path = os.path.join(base_path, "gilroy-bold.ttf")
        gilroy_regular_path = os.path.join(base_path, "gilroy-regular.ttf")
        gilroy_medium_path = os.path.join(base_path, "gilroy-medium.ttf")
        segoe_emoji_path = os.path.join(base_path, "segoe-ui-emoji.ttf")

        # Загрузка шрифтов
        try:
            title_font = ImageFont.truetype(gilroy_black_path, 36)
            header_font = ImageFont.truetype(gilroy_bold_path, 24)
            text_font = ImageFont.truetype(gilroy_regular_path, 18)
            emoji_font = ImageFont.truetype(segoe_emoji_path, 18)
        except IOError as e:
            tk.messagebox.showerror(
                "Ошибка",
                f"Не удалось загрузить шрифты: {e}",
            )
            return

        # Создание нового изображения
        img = Image.new("RGB", (img_width, img_height), color=background_color)
        draw = ImageDraw.Draw(img)
        y_position = 20

        # Импорт внутри метода, чтобы избежать глобальных импортов
        import emoji

        def split_text_and_emojis(text):
            """
            Разделяет текст на сегменты: обычный текст и эмодзи.
            Возвращает список кортежей вида ('text', текст) или ('emoji', эмодзи).
            """
            emojis = emoji.emoji_list(text)
            segments = []
            last_end = 0
            for em in emojis:
                start, end = em["match_start"], em["match_end"]
                if start > last_end:
                    # Добавляем текст перед эмодзи
                    segments.append(("text", text[last_end:start]))
                # Добавляем эмодзи
                segments.append(("emoji", text[start:end]))
                last_end = end
            if last_end < len(text):
                # Добавляем оставшийся текст
                segments.append(("text", text[last_end:]))
            return segments

        def emoji_to_codepoints(emoji_char):
            """
            Преобразует эмодзи в строку кодовых точек, разделённых дефисами.
            Например, 😀 -> '1f600'
            """
            codepoints = [f"{ord(ch):x}" for ch in emoji_char]
            return "-".join(codepoints)

        def draw_text_with_emojis(
            draw, img, x, y, text, font_regular, font_emoji, fill
        ):
            """
            Рисует текст с эмодзи на изображении с использованием объекта draw.
            """
            current_x = x
            current_y = y
            max_height = 0

            segments = split_text_and_emojis(text)
            for typ, segment in segments:
                if typ == "text":
                    # Рисуем текст
                    draw.text(
                        (current_x, current_y), segment, font=font_regular, fill=fill
                    )
                    bbox = font_regular.getbbox(segment)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]
                    current_x += text_width
                    max_height = max(max_height, text_height)
                elif typ == "emoji":
                    codepoint_seq = emoji_to_codepoints(segment)
                    emoji_filename = os.path.join(
                        base_path, "emoji_images", f"{codepoint_seq}.png"
                    )
                    if os.path.exists(emoji_filename):
                        try:
                            emoji_image = Image.open(emoji_filename).convert("RGBA")
                            # Изменяем размер эмодзи, чтобы соответствовать высоте текста
                            # Используем метод getbbox для символа 'A' как репрезентативного
                            text_bbox = font_regular.getbbox("A")
                            text_height = int(1.5 * text_bbox[3] - text_bbox[1])
                            # Используем Image.Resampling.LANCZOS для Pillow >=10
                            if hasattr(Image, "Resampling"):
                                resample_filter = Image.Resampling.LANCZOS
                            else:
                                resample_filter = Image.LANCZOS
                            emoji_image = emoji_image.resize(
                                (text_height, text_height), resample=resample_filter
                            )
                            img.paste(emoji_image, (current_x, current_y), emoji_image)
                            current_x += (
                                text_height  # Смещаемся вправо на ширину эмодзи
                            )
                            max_height = max(max_height, text_height)
                        except Exception as e:
                            # В случае ошибки загрузки изображения эмодзи, рисуем его как текст
                            print(e)
                            draw.text(
                                (current_x, current_y),
                                segment,
                                font=font_regular,
                                fill=fill,
                            )
                            bbox = font_regular.getbbox(segment)
                            text_width = bbox[2] - bbox[0]
                            text_height = bbox[3] - bbox[1]
                            current_x += text_width
                            max_height = max(max_height, text_height)
                    else:
                        # pass
                        # # Если изображение эмодзи не найдено, рисуем его как текст
                        draw.text(
                            (current_x, current_y),
                            segment,
                            font=font_emoji,
                            fill=fill,
                        )
                        bbox = font_emoji.getbbox(segment)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                        current_x += text_width
                        max_height = max(max_height, text_height)
            return current_y + max_height + 5

        # Заголовок
        header_text = "Оценочный лист"
        # Центрируем заголовок
        header_width = (
            title_font.getbbox(header_text)[2] - title_font.getbbox(header_text)[0]
        )
        header_x = (img_width - header_width) // 2
        y_position = draw_text_with_emojis(
            draw,
            img,
            header_x,
            y_position,
            header_text,
            title_font,
            emoji_font,
            text_color,
        )
        y_position += 20  # Добавляем отступ после заголовка

        # Информация о студенте
        student_info = f"Студент: {self.student_var.get()}    Группа: {self.group_var.get()}    Вариант: {self.variant_entry.get()}"
        y_position = draw_text_with_emojis(
            draw, img, 50, y_position, student_info, text_font, emoji_font, text_color
        )
        y_position += 10

        if (
            hasattr(self, "double_mode_enabled")
            and self.double_mode_enabled.get()
            and hasattr(self, "limit_to_eight")
        ):
            cap_text = self._format_score(8 if self.limit_to_eight.get() else 10)
            variant_text = f"Вариант работы: максимум {cap_text} баллов"
            y_position = draw_text_with_emojis(
                draw, img, 50, y_position, variant_text, text_font, emoji_font, text_color
            )
            y_position += 10

        # Информация о сдаче
        date_info = f"Сдано вовремя: {'Да' if self.on_time.get() else 'Нет'}    Дней просрочки: {delay_days}"
        y_position = draw_text_with_emojis(
            draw, img, 50, y_position, date_info, text_font, emoji_font, text_color
        )
        y_position += 20

        # Критерии
        for section, score in section_scores.items():
            y_position = draw_text_with_emojis(
                draw, img, 50, y_position, section, header_font, emoji_font, text_color
            )
            y_position += 10
            score_text = f"Баллы: {self._format_score(score)}"
            y_position = draw_text_with_emojis(
                draw, img, 70, y_position, score_text, text_font, emoji_font, text_color
            )
            y_position += 5
            comments = section_comments.get(section, [])
            for comment_text in comments:
                comment_line = f"- {comment_text}"
                y_position = draw_text_with_emojis(
                    draw,
                    img,
                    90,
                    y_position,
                    comment_line,
                    text_font,
                    emoji_font,
                    text_color,
                )
                y_position += 5
            y_position += 10

        # Штрафы
        y_position = draw_text_with_emojis(
            draw,
            img,
            50,
            y_position,
            "Дополнительные штрафы:",
            header_font,
            emoji_font,
            text_color,
        )
        y_position += 10
        if penalty_comments:
            for comment_text in penalty_comments:
                comment_line = f"- {comment_text}"
                y_position = draw_text_with_emojis(
                    draw,
                    img,
                    70,
                    y_position,
                    comment_line,
                    text_font,
                    emoji_font,
                    text_color,
                )
                y_position += 5
        else:
            y_position = draw_text_with_emojis(
                draw, img, 70, y_position, "Нет", text_font, emoji_font, text_color
            )
            y_position += 5
        y_position += 10

        # Поощрения
        y_position = draw_text_with_emojis(
            draw,
            img,
            50,
            y_position,
            "И ещё кое-что:",
            header_font,
            emoji_font,
            text_color,
        )
        y_position += 10
        if reward_comments:
            for reward in reward_comments:
                y_position = draw_text_with_emojis(
                    draw, img, 70, y_position, reward, text_font, emoji_font, text_color
                )
                y_position += 5
        else:
            y_position = draw_text_with_emojis(
                draw, img, 70, y_position, "Нет", text_font, emoji_font, text_color
            )
            y_position += 5
        y_position += 10

        # Разделительная линия
        draw.line((50, y_position, img_width - 50, y_position), fill=text_color)
        y_position += 10

        # Итоговая оценка
        final_score_text = f"Итоговая оценка: {self._format_score(final_score)} из {self._format_score(max_score_cap)}"
        y_position = draw_text_with_emojis(
            draw,
            img,
            50,
            y_position,
            final_score_text,
            header_font,
            emoji_font,
            text_color,
        )
        y_position += 20

        # Комментарий
        if comment:
            y_position = draw_text_with_emojis(
                draw,
                img,
                50,
                y_position,
                "Комментарий:",
                header_font,
                emoji_font,
                text_color,
            )
            y_position += 10
            # Разделяем комментарий на строки
            lines = comment.split("\n")
            for line in lines:
                y_position = draw_text_with_emojis(
                    draw, img, 70, y_position, line, text_font, emoji_font, text_color
                )
                y_position += 5

        self.generated_image = img

    def save_image(self):
        # Создание папки с названием домашней работы
        hw_name = self.hw_name_var.get()
        if not hw_name:
            self.status_var.set("Пожалуйста, выберите название домашней работы.")
            return
        # Проверка, что критерии загружены
        if not self.current_criteria:
            self.status_var.set("Пожалуйста, выберите домашнее задание.")
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
