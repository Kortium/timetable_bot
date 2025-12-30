import os
import sys

# Получаем абсолютный путь к директории, где находится main.py
dir_path = os.path.dirname(os.path.realpath(__file__))

# Добавляем путь к папке scripts, чтобы мы могли импортировать из неё модули
scripts_path = os.path.join(dir_path, '..', 'scripts')
sys.path.append(scripts_path)

import datetime  # Импорт модуля datetime для работы с датами
import traceback

import cairosvg  # Импорт модуля для конвертации SVG в PDF
from build_svg import (  # Импорт функций для подготовки данных и формирования SVG таблицы
    TableFormer, prepare_data)
from dotenv import load_dotenv
from parse_xls import (  # Импорт функции для чтения данных из xlsx файла
    DocType, check_type, read_professor, read_student)
from telegram import \
    Update  # Импорт класса Update для обработки обновлений в Telegram
from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (  # Импорт необходимых классов для работы бота
    CallbackContext, CallbackQueryHandler, CommandHandler, Filters,
    MessageHandler, Updater)

# Загрузка переменных окружения из файла .env
load_dotenv()

# Получение ID администратора и токена бота из переменных окружения
ADMIN_ID = os.getenv('ADMIN_ID')
MODERATOR_ID = int(os.getenv('MODERATOR_ID'))
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
    keyboard = [
        [InlineKeyboardButton("Весь семестр", callback_data='all')],
        [InlineKeyboardButton("С текущей даты до конца семестра", callback_data='now')],
        [InlineKeyboardButton("Две недели с текущей даты", callback_data='short')],
        [InlineKeyboardButton("Первая половина семестра", callback_data='first_half')],
        [InlineKeyboardButton("Вторая половина семестра", callback_data='second_half')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    update.message.reply_text("Файл получен! Теперь отправьте мне диапазон дат в формате 'ДД.ММ-ДД.ММ'. Или воспользуйтесь встроенной клавиатурой.", reply_markup=reply_markup)
    # Уведомляем админа о получении файла
    notify_admin(context, f"Пользователь {user_name} (ID: {user_id}) отправил файл.")

def handle_text(update: Update, context: CallbackContext) -> None:
    keyboard = [
        [InlineKeyboardButton("Весь семестр", callback_data='all')],
        [InlineKeyboardButton("С текущей даты до конца семестра", callback_data='now')],
        [InlineKeyboardButton("Две недели с текущей даты", callback_data='short')],
        [InlineKeyboardButton("Первая половина семестра", callback_data='first_half')],
        [InlineKeyboardButton("Вторая половина семестра", callback_data='second_half')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    # Обработчик текстовых сообщений от пользователя, в основном для дат
    text = update.message.text
    user_id = update.message.from_user.id
    user_name = update.message.from_user.first_name
    try:
        # Преобразуем текст в диапазон дат
        current_year = datetime.datetime.now().year+1
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
                no_color = False
                if user_id==MODERATOR_ID:
                    no_color = True
                svg_table_former = TableFormer(name, start_date, end_date, exercises, weekday_time_spans, f"timetable_{user_id}.svg", no_color)
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
                        update.message.reply_text(errors, reply_markup=reply_markup)
                    else:
                        update.message.reply_text(errors[0:4000]+'...', reply_markup=reply_markup)
                else:
                    # Если ошибок нет, сообщаем, что расписание готово
                    update.message.reply_text("Ваше расписание готово!", reply_markup=reply_markup)
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
                no_color = False
                if user_id==MODERATOR_ID:
                    no_color = True
                svg_table_former = TableFormer(group, start_date, end_date, exercises, weekday_time_spans, f"timetable_{user_id}.svg", no_color)
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
                        update.message.reply_text(errors, reply_markup=reply_markup)
                    else:
                        update.message.reply_text(errors[0:4000]+'...', reply_markup=reply_markup)
                else:
                    # Если ошибок нет, сообщаем, что расписание готово
                    update.message.reply_text("Ваше расписание готово!", reply_markup=reply_markup)
                # Уведомляем админа о завершении формирования расписания
                notify_admin(context, f"Пользователь {user_name} сформировал расписание для группы {group}.")
            except Exception as e:
                error_traceback = traceback.format_exc()
                print(error_traceback)
                # В случае неизвестной ошибки сообщаем пользователю обратиться к разработчику
                update.message.reply_text("Неизвестная ошибка при формировании расписания. Обратитесь к разработчику.")
    except Exception as e:
        error_traceback = traceback.format_exc()
        print(error_traceback)
        # В случае ошибки при чтении файла отправляем сообщение пользователю
        update.message.reply_text("Возникла ошибка при разборе файла, проверьте ваш файл и загрузите его снова или обратитесь к разработчику.")

def auto_range(update: Update, context: CallbackContext) -> None:
    keyboard = [
        [InlineKeyboardButton("Весь семестр", callback_data='all')],
        [InlineKeyboardButton("С текущей даты до конца семестра", callback_data='now')],
        [InlineKeyboardButton("Две недели с текущей даты", callback_data='short')],
        [InlineKeyboardButton("Первая половина семестра", callback_data='first_half')],
        [InlineKeyboardButton("Вторая половина семестра", callback_data='second_half')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    query = update.callback_query
    query.answer()
    user_id = query.message.chat_id
    # Обрабатываем ответы на кнопки
    if query.data == 'all':
        text="09.02-05.06"
    elif query.data == 'now':
        current_date = datetime.datetime.now()
        formatted_date = current_date.strftime('%d.%m')
        text=f"{formatted_date}-05.06"
    elif query.data == 'short':
        current_date = datetime.datetime.now()
        end_date = current_date + datetime.timedelta(days=14)
        formatted_current_date = current_date.strftime('%d.%m')
        formatted_end_date = end_date.strftime('%d.%m')
        text = f"{formatted_current_date}-{formatted_end_date}"
    elif query.data == 'first_half':
        text="09.02-30.03"
    elif query.data == 'second_half':
        text="30.03-05.06"
    user_name = query.from_user.first_name
    try:
        # Преобразуем текст в диапазон дат
        current_year = datetime.datetime.now().year
        start_date, end_date = [
            datetime.datetime.strptime(f"{current_year}.{date.strip()}", "%Y.%d.%m") for date in text.split('-')
        ]
    except:
        # Если не удалось преобразовать текст в диапазон дат, отправляем сообщение об ошибке
        context.bot.send_message(chat_id=user_id, text="Ошибка в формате даты! Пожалуйста, используйте формат 'ДД.ММ-ДД.ММ'.")
    try:
        # Читаем данные из файла и подготавливаем их
        # Проверить тип документа, студент или преподаватель и в зависимости от этого по разному разбирать
        type = check_type(f'recieved_timetable_{user_id}.xlsx')
        if type == DocType.PROFESSOR:
            name, exercises, errors = read_professor(f'recieved_timetable_{user_id}.xlsx')
            exercises, weekday_time_spans = prepare_data(exercises, start_date, end_date)
            try:
                # Создаём расписание в виде SVG и конвертируем его в PDF
                no_color = False
                if user_id==MODERATOR_ID:
                    no_color = True
                svg_table_former = TableFormer(name, start_date, end_date, exercises, weekday_time_spans, f"timetable_{user_id}.svg", no_color)
                svg_table_former.draw_timetable()
                svg_table_former.save()
                cairosvg.svg2pdf(url=f"timetable_{user_id}.svg", write_to=f"timetable_{user_id}.pdf")
                processed_file_path = f"timetable_{user_id}.pdf"
                # Отправляем сформированный PDF пользователю
                context.bot.send_document(chat_id=user_id, document=open(processed_file_path, 'rb'))
                if len(errors) > 0:
                    # Если в процессе обработки возникли ошибки, сообщаем об этом пользователю
                    context.bot.send_message(chat_id=user_id, text="Ваше расписание готово!\nОбратите внимание что в изначальном расписании есть наложения:")
                    if len(errors) < 4096:
                        context.bot.send_message(chat_id=user_id, text=errors, reply_markup=reply_markup)
                    else:
                        context.bot.send_message(chat_id=user_id, text=errors[0:4000]+'...', reply_markup=reply_markup)
                else:
                    # Если ошибок нет, сообщаем, что расписание готово
                    context.bot.send_message(chat_id=user_id, text="Ваше расписание готово!", reply_markup=reply_markup)
                # Уведомляем админа о завершении формирования расписания
                notify_admin(context, f"Пользователь {user_name} сформировал расписание для {name}.")
            except:
                error_traceback = traceback.format_exc()
                print(error_traceback)
                # В случае неизвестной ошибки сообщаем пользователю обратиться к разработчику
                context.bot.send_message(chat_id=user_id, text="Неизвестная ошибка при формировании расписания. Обратитесь к разработчику.")
        if type == DocType.STUDENT:
            group, exercises, errors = read_student(f'recieved_timetable_{user_id}.xlsx')
            exercises, weekday_time_spans = prepare_data(exercises, start_date, end_date)
            try:
                # Создаём расписание в виде SVG и конвертируем его в PDF
                no_color = False
                if user_id==MODERATOR_ID:
                    no_color = True
                svg_table_former = TableFormer(group, start_date, end_date, exercises, weekday_time_spans, f"timetable_{user_id}.svg", no_color)
                svg_table_former.draw_timetable()
                svg_table_former.save()
                cairosvg.svg2pdf(url=f"timetable_{user_id}.svg", write_to=f"timetable_{user_id}.pdf")
                processed_file_path = f"timetable_{user_id}.pdf"
                # Отправляем сформированный PDF пользователю
                context.bot.send_document(chat_id=user_id, document=open(processed_file_path, 'rb'))
                if len(errors) > 0:
                    # Если в процессе обработки возникли ошибки, сообщаем об этом пользователю
                    context.bot.send_message(chat_id=user_id, text="Ваше расписание готово!\nОбратите внимание что в изначальном расписании есть наложения:")
                    if len(errors) < 4096:
                        context.bot.send_message(chat_id=user_id, text=errors, reply_markup=reply_markup)
                    else:
                        context.bot.send_message(chat_id=user_id, text=errors[0:4000]+'...', reply_markup=reply_markup)
                else:
                    # Если ошибок нет, сообщаем, что расписание готово
                    context.bot.send_message(chat_id=user_id, text="Ваше расписание готово!", reply_markup=reply_markup)
                # Уведомляем админа о завершении формирования расписания
                notify_admin(context, f"Пользователь {user_name} сформировал расписание для группы {group}.")
            except Exception as e:
                error_traceback = traceback.format_exc()
                print(error_traceback)
                # В случае неизвестной ошибки сообщаем пользователю обратиться к разработчику
                context.bot.send_message(chat_id=user_id, text="Неизвестная ошибка при формировании расписания. Обратитесь к разработчику.")
    except Exception as e:
        error_traceback = traceback.format_exc()
        print(error_traceback)
        # В случае ошибки при чтении файла отправляем сообщение пользователю
        context.bot.send_message(chat_id=user_id, text="Возникла ошибка при разборе файла, проверьте ваш файл и загрузите его снова или обратитесь к разработчику.")

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
    dp.add_handler(CallbackQueryHandler(auto_range))

    # Запуск бота
    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
