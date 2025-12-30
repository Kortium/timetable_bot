import base64
import colorsys
import hashlib
import json
import os
from datetime import datetime, timedelta
from math import ceil, floor

import svgwrite
from parse_xls import (  # Подключение специализированных функций из файла parse_xls.py
    extract_initials, shorten_group)
from PIL import ImageFont

# Задаем начальную и конечную даты семестра
SEM_START = datetime(2026, 2, 9)
SEM_END = datetime(2026, 6, 5)

# Получаем текущую директорию, где находится выполняемый скрипт
current_directory = os.path.dirname(os.path.abspath(__file__))
# Определяем путь к шрифту

# Функция для вычисления ширины текста при отображении
def get_text_width(text, font_size=10):
    FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" # Пример пути к шрифту
    try:
        font = ImageFont.truetype(FONT_PATH, font_size)  # Загружаем шрифт
    except:
        FONT_PATH = os.path.join(current_directory, "times.ttf")
        font = ImageFont.truetype(FONT_PATH, font_size)  # Загружаем шрифт
    width = font.getlength(text)*1.1  # Вычисляем длину текста с учетом шрифта
    return width

# Функция подготовки данных для расписания
def prepare_data(exercises, start_date, end_date):
    organized_data = {}
    weekday_time_spans = {
        1: set(),
        2: set(),
        3: set(),
        4: set(),
        5: set(),
        6: set()
    }
    delta = timedelta(days=1)

    # Выравниваем начальную дату на начало недели (понедельник)
    while start_date.weekday() != 0:
        start_date -= delta

    # Выравниваем конечную дату на начало недели (понедельник)
    while end_date.weekday() != 0:
        end_date += delta

    for exercise in exercises:
        # Пропускаем занятия вне заданного диапазона дат
        if exercise['date'] < start_date or exercise['date'] > end_date:
            continue
        date = exercise['date']
        time_start = exercise['time_start']
        time_end = exercise['time_end']
        time_str = f"{time_start}-{time_end}"  # Строковое представление временного промежутка

        # Собираем данные в структурированном виде
        if date not in organized_data:
            organized_data[date] = {}
        _, _, exercise_week_day = exercise["date"].isocalendar()
        weekday_time_spans[exercise_week_day].add((time_start, time_end))

        # Если занятие объединено с другим, добавляем временные промежутки
        if exercise["joined"]:
            weekday_time_spans[exercise_week_day].add((exercise['time_start_s'], exercise['time_end_s']))

        organized_data[date][time_str] = exercise

    # Сортируем временные промежутки для каждого дня недели
    for key, time_spans in weekday_time_spans.items():
        sorted_time_spans = sorted(list(time_spans))
        weekday_time_spans[key] = sorted_time_spans
    return organized_data, weekday_time_spans

# Функция для сериализации объектов datetime в JSON
def datetime_handler(dt_obj):
    if isinstance(dt_obj, datetime):
        return dt_obj.strftime('%d.%m.%y')  # Возвращаем отформатированную дату
    raise TypeError("Type not serializable")  # В случае ошибки

# Функция для сохранения расписания в формате JSON
def save_to_json(timetable_data, filename="data.json"):
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(timetable_data, f, default=datetime_handler, ensure_ascii=False, indent=4)

# Функция получения цвета для элемента расписания
def get_color(subject, exercise_type, group):
    if len(group) > 0:
        unique_str = subject + exercise_type + group[0]  # Создаем уникальную строку
    else:
        unique_str = subject + exercise_type
    hash_value = int(hashlib.md5(unique_str.encode()).hexdigest(), 16)  # Получаем хеш-значение
    # Преобразуем хеш в цвет в HSV, а затем в RGB
    hue = (hash_value % 360) / 360.0
    saturation = 0.6 + (hash_value % 40) / 100.0
    value = 0.5 + (hash_value % 50) / 100.0
    r, g, b = colorsys.hsv_to_rgb(hue, saturation, value)
    # Возвращаем цвет в формате HEX
    return '#{:02x}{:02x}{:02x}'.format(int(r*255), int(g*255), int(b*255))

def generate_date_list(start_date, end_date):
    delta = timedelta(days=1)  # Определяем шаг в один день

    current_date = start_date
    date_list = []

    # Перебираем дни от начальной до конечной даты
    while current_date <= end_date:
        # Добавляем в список только рабочие дни (понедельник - пятница)
        if 0 <= current_date.weekday() <= 5:
            date_list.append(current_date.strftime("%d.%m"))
        current_date += delta
    return date_list  # Возвращаем список дат

def get_font_size(name, name_column_width):
    font_size = 18  # Начинаем с размера шрифта 20
    # Уменьшаем размер шрифта до тех пор, пока текст не поместится в заданную ширину
    while font_size > 1:
        width = get_text_width(name, font_size)
        if width < name_column_width*0.9:  # Оставляем небольшой отступ в 10%
            return font_size
        font_size -= 1
    return None  # Если подходящий размер шрифта не найден

# Формируем текст для отображения в ячейке расписания
def form_text(subject, exercise_type, groups, room, height, width):
    average_char_width = 0.75  # Средняя ширина символа (используется как заглушка)
    max_font_size = 16  # Максимальный размер шрифта

    # Сокращаем текст в зависимости от количества групп
    if len(groups) > 26:
        elements = [
            f"{subject}",
            f"({exercise_type})",
            f"{groups[:12]}...",
            room
        ]
    elif len(groups) > 12:
        elements = [
            f"{subject}",
            f"({exercise_type})",
            f"{groups[:12]}",
            f"{groups[12:]}",
            room
        ]
    else:
        elements = [
            f"{subject}",
            f"({exercise_type})",
            f"{groups}",
            room
        ]

    # Функция попытки объединить элементы для уменьшения количества строк
    def try_merge_elements(elements):
        longest_length = max(get_text_width(element) for element in elements)
        merged = []
        i = 0
        # Пытаемся объединить соседние элементы, чтобы они поместились в строку
        while i < len(elements):
            current_element = elements[i]
            # Проверяем, можно ли объединить с соседним элементом
            if i + 1 < len(elements) and get_text_width(current_element) + get_text_width(elements[i + 1]) <= longest_length:
                current_element += ', ' + elements[i + 1]
                i += 2
            else:
                i += 1
            merged.append(current_element)
        return merged
    if len(groups) > 18:
        elements = try_merge_elements(elements)

    # Уменьшаем размер шрифта, пока текст не поместится в ячейку
    while max_font_size > 1:
        for i in range(len(elements), 0, -1):  # i - количество строк для текста
            combinations = [', '.join(elements[j:j+i]) for j in range(0, len(elements), i)]
            estimated_height = 1.5 * max_font_size * len(combinations)
            # Проверяем, что текст помещается по ширине и высоте
            if all(get_text_width(comb, max_font_size) <= width*0.9 for comb in combinations) and estimated_height <= height:
                return combinations, max_font_size

        max_font_size -= 1

    return None, None  # Если не удалось подобрать размер шрифта

# Конвертируем шрифт в base64, чтобы его можно было использовать в вебе или передать как данные
def font_to_base64(font_path):
    with open(font_path, 'rb') as font_file:
        # Кодируем содержимое файла шрифта в base64
        return base64.b64encode(font_file.read()).decode('utf-8')


class TableFormer:
    # Статические переменные для определения номеров дней недели и отступов по дням недели.
    weekday_numbers = {
        "Пн": 1,
        "Вт": 2,
        "Ср": 3,
        "Чт": 4,
        "Пт": 5,
        "Сб": 6
    }
    weekday_margin = {
        1: 0,
        2: 0,
        3: 0,
        4: 0,
        5: 0,
        6: 0
    }

    def __init__(self,
                 name,
                 start_date,
                 end_date,
                 exercises,
                 weekday_time_spans,
                 file_name, no_color=False):
        # Инициализация класса с заданными параметрами.
        delta = timedelta(days=1)

        # Корректировка начальной и конечной даты до ближайшего понедельника.
        while start_date.weekday() != 0:
            start_date -= delta
        while end_date.weekday() != 0:
            end_date += delta

        # Сохранение переданных параметров в атрибуты экземпляра.
        self.name = name  # Имя таблицы
        self.start_date = start_date  # Начальная дата
        self.end_date = end_date  # Конечная дата
        # Запись упражнений в словарь только в том случае, если они в пределах заданных дат.
        self.exercises = {key: value for key, value in exercises.items() if key >= self.start_date and key < self.end_date}
        self.weekday_time_spans = weekday_time_spans  # Расписание по времени и дням недели
        self.no_color = no_color

        # Установка постоянных отступов и размеров элементов таблицы.
        self.margin_top = 50
        self.margin_left = 20
        self.name_column = 128
        self.header_height = 20

        # Инициализация параметров ячейки.
        self.cell_height = 46

        # Расчет полной ширины и высоты.
        days_difference = (self.start_date - SEM_START).days
        week_number = days_difference // 7
        self.week_start_number = week_number
        self.week_count = (self.end_date - self.start_date).days // 7
        self.full_width = 297 * 3.78
        self.cell_width = (self.full_width - self.margin_left * 2 - self.name_column) / self.week_count
        self.full_row_height = ceil(self.full_width * 0.707) - self.margin_top - self.header_height

        # Создание объекта SVG.
        self.dwg = svgwrite.Drawing(file_name, profile='full', size=(f"{self.full_width}", f"{ceil(self.full_width * 0.707)}"))

    # Метод для рисования ячейки расписания.
    def draw_timetable_cell(self, x, y, exercise):
        if len(exercise["subject"]) < 8:
            subject = exercise["subject"]
        else:
            subject = extract_initials(exercise["subject"])
        group = exercise.get("group", exercise.get("professor"))
        room = exercise["room"]
        if self.no_color:
            fill_color = '#ffffff'
        else:
            fill_color = get_color(subject, exercise["type"], group)
        if exercise["joined"] and exercise["type"] == "ЛР":
            rect = self.dwg.rect(insert=(x, y), size=(self.cell_width, self.cell_height*2), fill=fill_color,
                                 fill_opacity=0.5, rx=10, ry=10, stroke='black')
        else:
            if exercise["joined"]:
                exercise["joined"] = False
                self.draw_timetable_cell(x, y + self.cell_height, exercise)
            rect = self.dwg.rect(insert=(x, y), size=(self.cell_width, self.cell_height), fill=fill_color,
                                 fill_opacity=0.5, rx=10, ry=10, stroke='black')
        self.dwg.add(rect)
        if exercise["joined"] and exercise["type"] == "ЛР":
            cell_text = form_text(subject, exercise['type'], shorten_group(group), room, self.cell_height*2, self.cell_width)
        else:
            cell_text = form_text(subject, exercise['type'], shorten_group(group), room, self.cell_height, self.cell_width)
        font_size = cell_text[1]
        line_spacing = 2
        for index in range(len(cell_text[0])):
            text_y = y + (index + 1) * (font_size + line_spacing)
            self.dwg.add(self.dwg.text(cell_text[0][index], insert=(x + 5, text_y), font_family="CustomFont",
                                       font_size=font_size))

    # Метод для рисования заголовка таблицы.
    def draw_header(self):
        font_size = get_font_size(self.name, self.name_column)
        self.dwg.add(self.dwg.text(self.name,
                                   insert=(self.margin_left, self.margin_top-self.header_height/2+font_size/2),
                                   font_family="CustomFont", font_size=font_size))

        current_x = self.margin_left + self.name_column
        current_y = self.margin_top

        colored_row = True
        font_size = 14
        for week_number in range(1, self.week_count + 1):
            self.dwg.add(self.dwg.rect(insert=(current_x, current_y-self.header_height),
                                       size=(self.cell_width, self.header_height),
                                       fill='white', stroke='black', rx=10, ry=10))
            txt = str(self.week_start_number + week_number)+" В" if (self.week_start_number + week_number) % 2 == 1 \
                else str(self.week_start_number + week_number)+" Н"
            self.dwg.add(self.dwg.text(txt,
                                       insert=(current_x+self.cell_width/2, current_y-self.header_height/2+font_size/2-1),
                                       text_anchor="middle", font_family="CustomFont", font_size=font_size))
            if colored_row:
                self.dwg.add(self.dwg.rect(insert=(current_x, current_y),
                                           size=(self.cell_width, self.full_row_height),
                                           fill='rgb(220, 220, 220)', fill_opacity=0.2, rx=10, ry=10))
            colored_row = not colored_row
            current_x += self.cell_width

    # Метод для рисования дней недели и временных интервалов.
    def draw_week_days_and_time_spans(self):
        date_list = generate_date_list(self.start_date, self.end_date)
        days_of_week = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб"]
        day_of_cells = {}

        needed_cell_rows = 6
        for day in days_of_week:
            if self.weekday_time_spans[self.weekday_numbers[day]]:
                needed_cell_rows += len(self.weekday_time_spans[self.weekday_numbers[day]])-1
                day_of_cells[day] = len(self.weekday_time_spans[self.weekday_numbers[day]])
            else:
                day_of_cells[day] = 1

        self.cell_height = ceil((self.full_row_height/(needed_cell_rows+3)))-1
        current_x = self.margin_left+self.name_column

        for week_number in range(1, self.week_count + 1):
            current_y = self.margin_top
            for day in days_of_week:
                self.dwg.add(self.dwg.rect(insert=(current_x, current_y), size=(self.cell_width, self.cell_height/2), fill='rgb(220, 220, 220)', fill_opacity=0.2, rx=10, ry=10, stroke='black'))
                self.dwg.add(self.dwg.text(date_list[(week_number-1)*6+days_of_week.index(day)], insert=(current_x+self.cell_width/2, current_y + self.cell_height/4 + 3), text_anchor="middle", font_family="CustomFont", font_size=10))
                current_y += self.cell_height*day_of_cells[day]+self.cell_height/2
            current_x += self.cell_width

        current_y = self.margin_top+self.cell_height/2
        for day in days_of_week:
            self.weekday_margin[self.weekday_numbers[day]] = current_y+self.cell_height/2
            self.dwg.add(self.dwg.rect(insert=(self.margin_left, current_y), size=(self.name_column/2, self.cell_height*day_of_cells[day]), fill='rgb(220, 220, 220)', fill_opacity=0.2, rx=10, ry=10, stroke='black'))
            self.dwg.add(self.dwg.text(day, insert=(self.margin_left+self.name_column/4, current_y+self.cell_height*day_of_cells[day]/2 + 6), text_anchor="middle", font_family="CustomFont", font_size=24))
            current_y += self.cell_height*day_of_cells[day]+self.cell_height/2
        current_x = self.margin_left+self.name_column/2
        current_y = self.margin_top+self.cell_height/2
        for day in days_of_week:
            if self.weekday_time_spans[self.weekday_numbers[day]]:
                for time_span in self.weekday_time_spans[self.weekday_numbers[day]]:
                    self.dwg.add(self.dwg.rect(insert=(current_x, current_y), size=(self.name_column/2, self.cell_height), fill='rgb(220, 220, 220)', fill_opacity=0.2, rx=10, ry=10, stroke='black'))
                    self.dwg.add(self.dwg.text(time_span[0], insert=(current_x+self.name_column/4, current_y+self.cell_height/2 - 3), text_anchor="middle", font_family="CustomFont", font_size=12))
                    self.dwg.add(self.dwg.text(time_span[1], insert=(current_x+self.name_column/4, current_y+self.cell_height/2 + 9), text_anchor="middle", font_family="CustomFont", font_size=12))
                    current_y += self.cell_height
                current_y += self.cell_height/2
            else:
                current_y += self.cell_height*1.5

    # Основной метод для рисования всего расписания.
    def draw_timetable(self):
        self.draw_header()
        self.draw_week_days_and_time_spans()

        for date, time_intervals in self.exercises.items():
            for time_interval, exercise in time_intervals.items():
                _, exercise_week_number, exercise_week_day = exercise["date"].isocalendar()
                _, start_week_number, _ = self.start_date.isocalendar()
                if start_week_number > exercise_week_number:
                    exercise_week_number = 53
                column_index = exercise_week_number-start_week_number
                x = self.margin_left + self.name_column + column_index * self.cell_width
                y = self.weekday_margin[exercise_week_day] + self.weekday_time_spans[exercise_week_day].index(
                    (exercise["time_start"], exercise["time_end"]))*self.cell_height-self.cell_height/2

                self.draw_timetable_cell(x, y, exercise)

    # Метод для сохранения сгенерированного SVG файла.
    def save(self):
        self.dwg.save()
