import os
import logging

import telebot
from datetime import datetime

import settings
from main import Checker

bot = telebot.TeleBot(settings.__TELEGRAM_TOKEN__, parse_mode=None)

now = datetime.now()

logging.basicConfig(filename='logg.log', level=logging.DEBUG, format='%(asctime)s %(message)s')

@bot.message_handler(commands=['help'])
def help(message):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    bot.send_message(message.from_user.id, """
    Привет! Здесь можно проверить файл с выгрузкой на задвоенные кадастры. Мне нужен только xlsx файл с выгрузкой, и информация в каком столбце кадастровый номер, название улицы, номер дома, номер квартиры. В ответ вы получите эксельку в которой будет результат проверки. Если что-то не получается пишите сюда d.zaguzin@groupstp.ru 🤙🏻🤙🏻🤙🏻🤙🏻
    """)

@bot.message_handler(commands=['start'])
def start(message):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    bot.send_message(message.from_user.id, 'Чтобы начать введите команду /run')


@bot.message_handler(commands=['run'])
def run(message):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    bot.send_message(message.from_user.id, 'Начнем, отправьте мне в личные сообщения xlsx файл с выгрузкой')
    bot.register_next_step_handler(message, load_file)


def load_file(message):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    try:
        file_name = message.document.file_name
        file_id_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_id_info.file_path)
        src = file_name
        with open(src, 'wb') as new_file:
            new_file.write(downloaded_file)
        bot.send_message(message.chat.id,
                         'Файл добавлен, теперь мне нужно название страницы')
        bot.register_next_step_handler(message, get_sheet_name, file_name)
    except Exception as ex:
        bot.send_message(message.chat.id, "[!] Что-то не так с файлом - {}".format(str(ex)))


def get_sheet_name(message, file_name):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    sheet_name = message.text
    bot.send_message(message.chat.id, """
        Теперь отправьте мне названия столбцов с кадастром, улицей, домом и квартирой. Если кадастр в столбце J улица в столбце E, дом в столбце F, а квартира в столбце G отправьте 
J E F G 
    """)
    bot.register_next_step_handler(message, column_checker, file_name, sheet_name)


def column_checker(message, file_name, sheet_name):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    column = message.text
    column = column.split()
    if len(column) != 4:
        bot.send_message(message.chat.id, "Вы как-то не так отправили столбцы")
        os.remove(file_name)
        return
    num = []
    for i in column:
        i = i.upper()
        if len(i) != 1 and 'A' > i > 'Z':
            bot.send_message(message.chat.id, "Вы как-то не так отправили столбцы")
            os.remove(file_name)
            return
        num.append(ord(i) - ord('A'))
    bot.send_message(message.chat.id, 'Отлично, сейчас проверим файл')
    main(message, file_name, sheet_name, num)


def main(message, file_name, sheet_name, num):
    logging.info("{} : {}".format(message.from_user.id, message.text))
    x = Checker(file_name, sheet_name, *num)
    try:
        status = x.load_book()
    except Exception:
        bot.send_message(message.chat.id, 'Ошибка загрузки, попробуйте еще раз')
        os.remove(file_name)
    if not status:
        bot.send_message(message.chat.id,
                         'В шагах выше вы указали неприавльную информацию, скорее всего неверное имя страницы или проблема с файлом эксельки')
        os.remove(file_name)
        return
    else:
        bot.send_message(message.chat.id, 'Все круто, работаем. Если все введено корректно, то результат будет точный')
    try:
        status = x.check()
    except Exception:
        bot.send_message(message.chat.id, 'Произошла ошибка, свяжитесь со мной d.zaguzin@groupstp.ru')
        try:
            os.remove(x.get_file_name())
        except Exception:
            pass
        try:
            os.remove(file_name)
        except Exception:
            pass
        return

    if status:
        if x.cnt > 0:
            bot.send_message(message.chat.id, 'Было найдено {} дублей :('.format(x.cnt))
            print(x.get_file_name())

            f = open(x.get_file_name(), 'rb')
            bot.send_document(message.chat.id, f, {'serverDownload': True})
        else:
            bot.send_message(message.chat.id, 'Дублей нет! Файл отправлен на дальнейшую обработку')
            with open('готовые.txt', 'a') as file:
                file.write(x.get_file_name())
        os.remove(x.get_file_name())
        os.remove(file_name)


bot.polling()
