import os
import sys
import traceback
import datetime

# Подключаем локально написанные модули из папки scripts
dir_path = os.path.dirname(os.path.realpath(__file__))
scripts_path = os.path.join(dir_path, '..', 'scripts')
sys.path.append(scripts_path)

# import cairosvg
from build_svg import TableFormer, prepare_data
from parse_xls import DocType, check_type, read_professor, read_student

def main():
    """
    Локальный скрипт для отладки методов parse_xls и build_svg
    без запуска Telegram бота. Использует тот же функционал,
    который применяется в main_telegram.py.
    """
    # Укажите путь к нужному файлу:
    # Например, example_professor.xlsx или example_student.xlsx
    file_name = os.path.join(dir_path, '..', 'data', 'example_professor.xlsx')
    
    # Задаём диапазон дат, который хотите обработать
    start_date = datetime.datetime(2025, 9, 1)
    end_date   = datetime.datetime(2025, 12, 31)

    try:
        # Определяем тип файла (преподаватель / студент)
        doc_type = check_type(file_name)
        
        if doc_type == DocType.PROFESSOR:
            # Если это файл преподавателя
            name, exercises, errors = read_professor(file_name)
            label = name  # Подпись для будущей таблицы (ФИО)
        elif doc_type == DocType.STUDENT:
            # Если это файл студента
            group, exercises, errors = read_student(file_name)
            label = group  # Подпись для будущей таблицы (№ группы)
        else:
            print("Невозможно определить тип документа (ни студент, ни преподаватель).")
            return

        # Готовим данные для формирования расписания
        exercises, weekday_time_spans = prepare_data(exercises, start_date, end_date)

        # Рисуем SVG-таблицу расписания
        svg_filename = "timetable.svg"
        pdf_filename = "timetable.pdf"

        svg_table_former = TableFormer(
            label,               # Например, ФИО преподавателя или номер группы
            start_date,
            end_date,
            exercises,
            weekday_time_spans,
            svg_filename
        )

        svg_table_former.draw_timetable()
        svg_table_former.save()

        # Преобразуем SVG → PDF
        # cairosvg.svg2pdf(url=svg_filename, write_to=pdf_filename)
        print(f"✔ Расписание сохранено в файлах: {svg_filename}, {pdf_filename}")

        # Если при разборе файла найдены ошибки (например, пересечения занятий)
        if errors and len(errors) > 0:
            print("При обработке файла возникли замечания:")
            for err in errors:
                print(" -", err)
        else:
            print("Ошибок при разборе файла не обнаружено.")

    except Exception as e:
        print("Произошла ошибка во время обработки данных:")
        print(e)
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
