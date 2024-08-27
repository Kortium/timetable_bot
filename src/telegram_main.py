import os
import sys

# Получаем абсолютный путь к директории, где находится main.py
dir_path = os.path.dirname(os.path.realpath(__file__))

# Добавляем путь к папке scripts, чтобы мы могли импортировать из неё модули
scripts_path = os.path.join(dir_path, '..', 'scripts')
sys.path.append(scripts_path)

from dotenv import load_dotenv
from build_svg import prepare_data, TableFormer  # Импорт функций для подготовки данных и формирования SVG таблицы
from parse_xls import read_professor, read_student, check_type, DocType  # Импорт функции для чтения данных из xlsx файла
import cairosvg  # Импорт модуля для конвертации SVG в PDF
from telegram import Update  # Импорт класса Update для обработки обновлений в Telegram
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackContext  # Импорт необходимых классов для работы бота
import datetime  # Импорт модуля datetime для работы с датами
import traceback

# Загрузка переменных окружения из файла .env
load_dotenv()

# Получение ID администратора и токена бота из переменных окружения
ADMIN_ID = os.getenv('ADMIN_ID')
TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN')

def notify_admin(context: CallbackContext, text: str) -> None:
    # Функция для отправки уведомлений администратору
    context.bot.send_message(chat_id=ADMIN_ID, text=text)

def start(update: Update, context: CallbackContext) -> None:
    # Обработчик команды /start для бота, отправляет приветственное сообщение
    update.message.reply_text("Отправьте мне файл и диапазон дат для обработки!")

def handle_document(update: Update, context: CallbackContext) -> None:
    # Обработчик получения документа от пользователя
    user_id = update.message.from_user.id
    user_name = update.message.from_user.first_name
    file = update.message.document.get_file()
    # Скачиваем присланный файл и сохраняем его
    file.download(f'recieved_timetable_{user_id}.xlsx')
    # Запрашиваем у пользователя диапазон дат
    update.message.reply_text("Файл получен! Теперь отправьте мне диапазон дат в формате 'ДД.ММ-ДД.ММ'.")
    # Уведомляем админа о получении файла
    notify_admin(context, f"Пользователь {user_name} (ID: {user_id}) отправил файл.")

def handle_text(update: Update, context: CallbackContext) -> None:
    # Обработчик текстовых сообщений от пользователя, в основном для дат
    text = update.message.text
    user_id = update.message.from_user.id
    user_name = update.message.from_user.first_name
    try:
        # Преобразуем текст в диапазон дат
        current_year = datetime.datetime.now().year
        start_date, end_date = [
            datetime.datetime.strptime(f"{current_year}.{date.strip()}", "%Y.%d.%m") for date in text.split('-')
        ]
    except:
        # Если не удалось преобразовать текст в диапазон дат, отправляем сообщение об ошибке
        update.message.reply_text("Ошибка в формате даты! Пожалуйста, используйте формат 'ДД.ММ-ДД.ММ'.")
    try:
        # Читаем данные из файла и подготавливаем их
        # Проверить тип документа, студент или преподаватель и в зависимости от этого по разному разбирать
        type = check_type(f'recieved_timetable_{user_id}.xlsx')
        if type == DocType.PROFESSOR:
            name, exercises, errors = read_professor(f'recieved_timetable_{user_id}.xlsx')
            exercises, weekday_time_spans = prepare_data(exercises, start_date, end_date)
            try:
                # Создаём расписание в виде SVG и конвертируем его в PDF
                svg_table_former = TableFormer(name, start_date, end_date, exercises, weekday_time_spans, f"timetable_{user_id}.svg")
                svg_table_former.draw_timetable()
                svg_table_former.save()
                cairosvg.svg2pdf(url=f"timetable_{user_id}.svg", write_to=f"timetable_{user_id}.pdf")
                processed_file_path = f"timetable_{user_id}.pdf"
                # Отправляем сформированный PDF пользователю
                update.message.reply_document(open(processed_file_path, 'rb'))
                if len(errors) > 0:
                    # Если в процессе обработки возникли ошибки, сообщаем об этом пользователю
                    update.message.reply_text("Ваше расписание готово!\nОбратите внимание что в изначальном расписании есть наложения:")
                    update.message.reply_text(errors)
                else:
                    # Если ошибок нет, сообщаем, что расписание готово
                    update.message.reply_text("Ваше расписание готово!")
                # Уведомляем админа о завершении формирования расписания
                notify_admin(context, f"Пользователь {user_name} сформировал расписание для {name}.")
            except:
                # В случае неизвестной ошибки сообщаем пользователю обратиться к разработчику
                update.message.reply_text("Неизвестная ошибка при формировании расписания. Обратитесь к разработчику.")
        if type == DocType.STUDENT:
            group, exercises, errors = read_student(f'recieved_timetable_{user_id}.xlsx')
            exercises, weekday_time_spans = prepare_data(exercises, start_date, end_date)
            try:
                # Создаём расписание в виде SVG и конвертируем его в PDF
                svg_table_former = TableFormer(group, start_date, end_date, exercises, weekday_time_spans, f"timetable_{user_id}.svg")
                svg_table_former.draw_timetable()
                svg_table_former.save()
                cairosvg.svg2pdf(url=f"timetable_{user_id}.svg", write_to=f"timetable_{user_id}.pdf")
                processed_file_path = f"timetable_{user_id}.pdf"
                # Отправляем сформированный PDF пользователю
                update.message.reply_document(open(processed_file_path, 'rb'))
                if len(errors) > 0:
                    # Если в процессе обработки возникли ошибки, сообщаем об этом пользователю
                    update.message.reply_text("Ваше расписание готово!\nОбратите внимание что в изначальном расписании есть наложения:")
                    if len(errors) < 4096:
                        update.message.reply_text(errors)
                    else:
                        update.message.reply_text(errors[0:4093]+'...')
                else:
                    # Если ошибок нет, сообщаем, что расписание готово
                    update.message.reply_text("Ваше расписание готово!")
                # Уведомляем админа о завершении формирования расписания
                notify_admin(context, f"Пользователь {user_name} сформировал расписание для группы {group}.")
            except Exception as e:
                error_traceback = traceback.format_exc()
                print(error_traceback)
                # В случае неизвестной ошибки сообщаем пользователю обратиться к разработчику
                update.message.reply_text("Неизвестная ошибка при формировании расписания. Обратитесь к разработчику.")
    except:
        # В случае ошибки при чтении файла отправляем сообщение пользователю
        update.message.reply_text("Возникла ошибка при разборе файла, проверьте ваш файл и загрузите его снова или обратитесь к разработчику.")

def handle_unknown_document(update: Update, context: CallbackContext):
    # Обработчик для неизвестных документов
    update.message.reply_text("Расписание принимается только в формате .xlsx")

def main():
    # Главная функция для запуска бота
    # Замените TELEGRAM_TOKEN на ваш токен в .env файлк, который вы получили от BotFather
    updater = Updater(TELEGRAM_TOKEN)

    dp = updater.dispatcher
    # Добавление обработчиков команд и сообщений
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(MessageHandler(Filters.document.mime_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_document))
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, handle_text))
    dp.add_handler(MessageHandler(Filters.document, handle_unknown_document))

    # Запуск бота
    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
