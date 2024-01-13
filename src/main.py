import os
import sys

# Получаем абсолютный путь к директории, где находится main.py
dir_path = os.path.dirname(os.path.realpath(__file__))

# Добавляем путь к папке scripts, чтобы мы могли импортировать из неё модули
scripts_path = os.path.join(dir_path, '..', 'scripts')
sys.path.append(scripts_path)

from datetime import datetime
from build_svg import prepare_data, TableFormer  # Импортируем необходимые функции и классы для работы со временем и создания SVG
from parse_xls import read_xlsx  # Импортируем функцию для чтения xlsx файла
import json  # Импортируем модуль json для работы с JSON данными
import cairosvg  # Импортируем модуль cairosvg для конвертации SVG в PDF

def datetime_serializer(obj):
    # Функция сериализации даты/времени для преобразования в JSON-совместимый формат
    if isinstance(obj, datetime):
        return obj.isoformat()  # Если объект является datetime, возвращаем его в формате ISO
    else:
        return obj  # В противном случае возвращаем объект без изменений

if __name__ == '__main__':
    file_name = "data/example.xlsx"
    name, exercises, errors = read_xlsx(file_name)  # Читаем данные из xlsx файла и получаем имя, занятия и ошибки
    start_date = datetime(2024, 2, 12)  # Начальная дата для генерации расписания
    end_date = datetime(2024, 6, 1)  # Конечная дата для генерации расписания
    # Подготавливаем данные для построения SVG
    exercises, weekday_time_spans = prepare_data(exercises, start_date, end_date)
    data = {
        'name': name,  # Имя преподавателя или название события
        'exercises': exercises  # Список упражнений
    }
    # Сериализуем данные для сохранения в JSON, используя нашу функцию datetime_serializer
    serialized_data = {datetime_serializer(key): datetime_serializer(value) for key, value in exercises.items()}
    with open('data.json', 'w') as outfile:  # Открываем файл для записи данных в формате JSON
        # Сохраняем сериализованные данные в файл с форматированием для удобочитаемости
        json.dump(serialized_data, outfile, default=datetime_serializer, indent=4, separators=(',', ': '), sort_keys=True)
    # Создаем экземпляр класса TableFormer для создания таблицы расписания в формате SVG
    svg_table_former = TableFormer(name, start_date, end_date, exercises, weekday_time_spans, "timetable.svg")
    svg_table_former.draw_timetable()  # Рисуем таблицу расписания
    svg_table_former.save()  # Сохраняем таблицу расписания
    # Конвертируем SVG таблицу в PDF документ
    cairosvg.svg2pdf(url="timetable.svg", write_to="timetable.pdf")
    print(errors)  # Выводим ошибки, если они есть
