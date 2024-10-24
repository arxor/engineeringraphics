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

        # Вкладка с дополнительными штрафами и поощрениями
        self.penalty_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.penalty_frame, text="Штрафы и поощрения")

        # Вкладка с генерацией отчета
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.report_frame, text="Генерация отчета")

        self.section_max_scores = {}

        self.create_info_tab()
        self.create_criteria_tab()
        self.load_info_parameters()  # Загрузка сохраненных параметров
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
            with open("student_list.csv", "w", encoding="utf-8") as csvfile:
                reader = csv.DictReader(csvfile, delimiter=";")
                tk.messagebox.showerror(
                    "Ошибка",
                    "Не найден файл student_list.csv. Был создан новый пустой файл.",
                )

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
        # Нет необходимости обновлять штрафы и поощрения, так как они общие

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

        # Создаем штрафы и поощрения из JSON-файла
        self.create_penalties_and_rewards_from_json()

    def load_criteria_for_homework(self, homework_name):
        # Очистка предыдущих критериев
        self.criteria_scores = {}
        for widget in self.criteria_inner_frame.winfo_children():
            widget.destroy()

        self.section_max_scores = {}
        sections = self.criteria_data.get("sections", {})
        if homework_name is not None:
            if homework_name in sections:
                self.current_criteria = sections[homework_name]
                self.create_criteria()
            else:
                tk.messagebox.showerror(
                    "Ошибка", f"Критерии для '{homework_name}' не найдены."
                )
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
        self.delay_penalty_per_day = delay_info.get("score_per_day", -2)

        delay_frame = ttk.Frame(self.penalty_inner_frame)
        delay_frame.pack(fill="x", pady=10)
        tk.Label(delay_frame, text=delay_text).pack(side="left")
        self.delay_entry = tk.Entry(delay_frame, width=5)
        self.delay_entry.pack(side="left", padx=5)
        self.delay_entry.insert(0, "0")

        # Поощрения
        tk.Label(self.penalty_inner_frame, text="Поощрения:").pack(anchor="w", pady=10)

        self.reward_vars = []
        self.reward_texts = []

        rewards = self.criteria_data.get("rewards", [])
        for reward in rewards:
            text = reward.get("text", "")
            var = tk.BooleanVar(value=False)
            cb = ttk.Checkbutton(self.penalty_inner_frame, text=text, variable=var)
            cb.pack(anchor="w")
            self.reward_vars.append(var)
            self.reward_texts.append(text)

    def checkbox_callback(self, checkbox_var, score, main_var):
        if checkbox_var.get():
            main_var.set(score)
        # else:
        #     pass

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

        # Сбрасываем поощрения
        for var in self.reward_vars:
            var.set(False)

        # Сбрасываем комментарий
        self.comment_text.delete("1.0", tk.END)

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
        for section, data in self.criteria_scores.items():
            max_score = self.section_max_scores.get(section, 0)
            score = max_score  # Начинаем с максимального балла
            comments = []
            if data["type"] == "radio_with_subchecks":
                main_var = data["main_var"]
                main_score = main_var.get()
                selected_option = None
                # Находим выбранную опцию
                for option in data["options"]:
                    if option.get("score") == main_score:
                        selected_option = option
                        break
                if selected_option:
                    # Вычисляем общее количество выбранных чекбоксов
                    num_selected_checkboxes = 0
                    if "suboption_vars" in selected_option:
                        for var_cb, subtext in zip(
                            selected_option["suboption_vars"],
                            selected_option["suboptions"],
                        ):
                            if var_cb.get():
                                num_selected_checkboxes += 1
                                comments.append(subtext)
                    # Вычисляем общую сумму вычитаемых баллов
                    deduction = abs(main_score) * (
                        num_selected_checkboxes if -main_score != max_score else 1
                    )
                    # Ограничиваем вычитание максимальным баллом по критерию
                    deduction = min(deduction, max_score)
                    # Итоговый балл по критерию
                    score = max_score - deduction
            elif data["type"] == "checkbox":
                vars_list = data["vars"]
                score = 0  # Начинаем с 0 для типа checkbox
                for var_cb, var_score in vars_list:
                    if var_cb.get():
                        score += var_score
                        if var_score < 0:
                            comments.append(var_cb._name)  # Добавляем текст ошибки
                # Ограничиваем баллы в пределах допустимых значений
                score = max(0, min(max_score, score))
            section_scores[section] = score
            section_comments[section] = comments
            total_score += score
        # Учёт штрафов
        penalty_score = 0
        penalty_comments = []
        for (var, value), text in zip(self.penalty_vars, self.penalty_texts):
            if var.get():
                # Если выбран пункт про плагиат, оценка обнуляется
                if value <= -1000 and var.get():
                    total_score = 0
                    penalty_comments.append(text)
                else:
                    penalty_score += value
                    penalty_comments.append(text)
        # Учёт просрочки
        try:
            delay_days = int(self.delay_entry.get())
            delay_penalty = self.delay_penalty_per_day * delay_days
            penalty_score += delay_penalty
            if delay_days > 0:
                penalty_comments.append(
                    f"Просрочка сдачи задания на {delay_days} дней ({delay_penalty} балла)"
                )
        except ValueError:
            delay_days = 0
            delay_penalty = 0
        # Общий итог
        final_score = max(0, total_score + penalty_score)
        if not self.on_time.get():
            final_score = max(0, final_score)
        # Получаем комментарий
        comment_text = self.comment_text.get("1.0", tk.END).strip()
        # Получаем поощрения
        reward_comments = []
        for var, text in zip(self.reward_vars, self.reward_texts):
            if var.get():
                reward_comments.append(text)
        # Генерация изображения оценочного листа
        self.create_image(
            section_scores,
            section_comments,
            penalty_comments,
            final_score,
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
            score_text = f"Баллы: {score}"
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
        final_score_text = f"Итоговая оценка: {final_score} из 10"
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
